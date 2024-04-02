package main

import (
	"bytes"
	"encoding/xml"
	"fmt"
	"math"
	"math/bits"
	"os"
	"slices"
	"strconv"
	"strings"
	"time"
	"unsafe"
)

type CharacterWidth int

const (
	cwZeroWidth CharacterWidth = iota
	cwNarrow
	cwWide
	cwAmbiguous
)

type ClusterBreak int

const (
	cbOther         ClusterBreak = iota // GB999
	cbControl                           // GB3, GB4, GB5 -- includes CR, LF
	cbExtend                            // GB9, GB9a -- includes SpacingMark
	cbRI                                // GB12, GB13
	cbPrepend                           // GB9b
	cbHangulL                           // GB6, GB7, GB8
	cbHangulV                           // GB6, GB7, GB8
	cbHangulT                           // GB6, GB7, GB8
	cbHangulLV                          // GB6, GB7, GB8
	cbHangulLVT                         // GB6, GB7, GB8
	cbInCBLinker                        // GB9c
	cbInCBConsonant                     // GB9c
	cbExtPic                            // GB11
	cbZWJ                               // GB9, GB11

	cbCount
)

// Ω is what UAX #29 writes as "÷". ÷ is not a valid identifier in Go.
var Ω uint8 = 0b11

// JoinRules doesn't quite follow UAX #29, as it states:
// > Note: Testing two adjacent characters is insufficient for determining a boundary.
//
// I completely agree, however it makes the implementation complex and slow and it only benefits what can be considered
// edge cases in the context of terminals. By using a lookup table anyway this results in a >100MB/s throughput,
// before adding any fast-passes whatsoever. This is at least 2x as fast as any standards conforming implementation.
//
// This affects the following rules:
// * GB9c: \p{InCB=Consonant} [\p{InCB=Extend}\p{InCB=Linker}]* \p{InCB=Linker} [\p{InCB=Extend}\p{InCB=Linker}]* × \p{InCB=Consonant}
//   "Do not break within certain combinations with Indic_Conjunct_Break (InCB)=Linker."
//   Our implementation does this:
//                     × \p{InCB=Linker}
//     \p{InCB=Linker} × \p{InCB=Consonant}
//   In other words, it doesn't check for a leading \p{InCB=Consonant} or a series of Extenders/Linkers in between.
//   I suspect that these simplified rules are sufficient for the vast majority of terminal use cases.
// * GB11: \p{Extended_Pictographic} Extend* ZWJ × \p{Extended_Pictographic}
//   "Do not break within emoji modifier sequences or emoji zwj sequences."
//   Our implementation does this:
//     ZWJ × \p{Extended_Pictographic}
//   In other words, it doesn't check whether the ZWJ is led by another \p{InCB=Extended_Pictographic}.
//   Again, I suspect that a trailing, standalone ZWJ is a rare occurence and joining it with any Emoji is fine.
// * GB12: sot (RI RI)* RI × RI
//   GB13: [^RI] (RI RI)* RI × RI
//   "Do not break within emoji flag sequences. That is, do not break between regional indicator
//   (RI) symbols if there is an odd number of RI characters before the break point."
//   Our implementation does this (this is not a real notation):
//     RI ÷ RI × RI ÷ RI
//   In other words, it joins any pair of RIs and then immediately aborts further RI joins.
//   Unlike the above two cases, this is a bit more risky, because it's much more likely to be encountered in practice.
//   Imagine a shell that doesn't understand graphemes for instance. You type 2 flags (= 4 RIs) and backspace.
//   You'll now have 3 RIs. If iterating through it forwards, you'd join the first two, then get 1 lone RI at the end,
//   whereas if you iterate backwards you'd join the last two, then get 1 lone RI at the start.
//   This assymmetry may have some subtle effects, but I suspect that it's still rare enough to not matter much.
//
// This is a great reference for the resulting table:
// https://www.unicode.org/Public/UCD/latest/ucd/auxiliary/GraphemeBreakTest.html
var JoinRules = [][][]uint8{
	// Base table
	{
		/* | leading       -> trailing codepoint                                                                                                                                                            */
		/* v               |  cbOther | cbControl | cbExtend |   cbRI   | cbPrepend | cbHangulL | cbHangulV | cbHangulT | cbHangulLV | cbHangulLVT | cbInCBLinker | cbInCBConsonant | cbExtPic |   cbZWJ  | */
		/* cbOther         | */ {Ω /* | */, Ω /*  | */, 0 /* | */, Ω /* | */, Ω /*  | */, Ω /*  | */, Ω /*  | */, Ω /*  |  */, Ω /*  |  */, Ω /*   |   */, 0 /*   |    */, Ω /*     | */, Ω /* | */, 0 /* | */},
		/* cbControl       | */ {Ω /* | */, Ω /*  | */, Ω /* | */, Ω /* | */, Ω /*  | */, Ω /*  | */, Ω /*  | */, Ω /*  |  */, Ω /*  |  */, Ω /*   |   */, Ω /*   |    */, Ω /*     | */, Ω /* | */, Ω /* | */},
		/* cbExtend        | */ {Ω /* | */, Ω /*  | */, 0 /* | */, Ω /* | */, Ω /*  | */, Ω /*  | */, Ω /*  | */, Ω /*  |  */, Ω /*  |  */, Ω /*   |   */, 0 /*   |    */, Ω /*     | */, Ω /* | */, 0 /* | */},
		/* cbRI            | */ {Ω /* | */, Ω /*  | */, 0 /* | */, 1 /* | */, Ω /*  | */, Ω /*  | */, Ω /*  | */, Ω /*  |  */, Ω /*  |  */, Ω /*   |   */, 0 /*   |    */, Ω /*     | */, Ω /* | */, 0 /* | */},
		/* cbPrepend       | */ {0 /* | */, Ω /*  | */, 0 /* | */, 0 /* | */, 0 /*  | */, 0 /*  | */, 0 /*  | */, 0 /*  |  */, 0 /*  |  */, 0 /*   |   */, 0 /*   |    */, 0 /*     | */, 0 /* | */, 0 /* | */},
		/* cbHangulL       | */ {Ω /* | */, Ω /*  | */, 0 /* | */, Ω /* | */, Ω /*  | */, 0 /*  | */, 0 /*  | */, Ω /*  |  */, 0 /*  |  */, 0 /*   |   */, 0 /*   |    */, Ω /*     | */, Ω /* | */, 0 /* | */},
		/* cbHangulV       | */ {Ω /* | */, Ω /*  | */, 0 /* | */, Ω /* | */, Ω /*  | */, Ω /*  | */, 0 /*  | */, 0 /*  |  */, Ω /*  |  */, Ω /*   |   */, 0 /*   |    */, Ω /*     | */, Ω /* | */, 0 /* | */},
		/* cbHangulT       | */ {Ω /* | */, Ω /*  | */, 0 /* | */, Ω /* | */, Ω /*  | */, Ω /*  | */, Ω /*  | */, 0 /*  |  */, Ω /*  |  */, Ω /*   |   */, 0 /*   |    */, Ω /*     | */, Ω /* | */, 0 /* | */},
		/* cbHangulLV      | */ {Ω /* | */, Ω /*  | */, 0 /* | */, Ω /* | */, Ω /*  | */, Ω /*  | */, 0 /*  | */, 0 /*  |  */, Ω /*  |  */, Ω /*   |   */, 0 /*   |    */, Ω /*     | */, Ω /* | */, 0 /* | */},
		/* cbHangulLVT     | */ {Ω /* | */, Ω /*  | */, 0 /* | */, Ω /* | */, Ω /*  | */, Ω /*  | */, Ω /*  | */, 0 /*  |  */, Ω /*  |  */, Ω /*   |   */, 0 /*   |    */, Ω /*     | */, Ω /* | */, 0 /* | */},
		/* cbInCBLinker    | */ {Ω /* | */, Ω /*  | */, 0 /* | */, Ω /* | */, Ω /*  | */, Ω /*  | */, Ω /*  | */, Ω /*  |  */, Ω /*  |  */, Ω /*   |   */, 0 /*   |    */, 0 /*     | */, Ω /* | */, 0 /* | */},
		/* cbInCBConsonant | */ {Ω /* | */, Ω /*  | */, 0 /* | */, Ω /* | */, Ω /*  | */, Ω /*  | */, Ω /*  | */, Ω /*  |  */, Ω /*  |  */, Ω /*   |   */, 0 /*   |    */, Ω /*     | */, Ω /* | */, 0 /* | */},
		/* cbExtPic        | */ {Ω /* | */, Ω /*  | */, 0 /* | */, Ω /* | */, Ω /*  | */, Ω /*  | */, Ω /*  | */, Ω /*  |  */, Ω /*  |  */, Ω /*   |   */, 0 /*   |    */, Ω /*     | */, Ω /* | */, 0 /* | */},
		/* cbZWJ           | */ {Ω /* | */, Ω /*  | */, 0 /* | */, Ω /* | */, Ω /*  | */, Ω /*  | */, Ω /*  | */, Ω /*  |  */, Ω /*  |  */, Ω /*   |   */, 0 /*   |    */, Ω /*     | */, 0 /* | */, 0 /* | */},
	},
	// Once we have encountered a Regional Indicator pair we'll enter this table.
	// It's a copy of the base table, but further Regional Indicator joins are forbidden.
	{
		/* | leading       -> trailing codepoint                                                                                                                                                            */
		/* v               |  cbOther | cbControl | cbExtend |   cbRI   | cbPrepend | cbHangulL | cbHangulV | cbHangulT | cbHangulLV | cbHangulLVT | cbInCBLinker | cbInCBConsonant | cbExtPic |   cbZWJ  | */
		/* cbOther         | */ {Ω /* | */, Ω /*  | */, 0 /* | */, Ω /* | */, Ω /*  | */, Ω /*  | */, Ω /*  | */, Ω /*  |  */, Ω /*  |  */, Ω /*   |   */, 0 /*   |    */, Ω /*     | */, Ω /* | */, 0 /* | */},
		/* cbControl       | */ {Ω /* | */, Ω /*  | */, Ω /* | */, Ω /* | */, Ω /*  | */, Ω /*  | */, Ω /*  | */, Ω /*  |  */, Ω /*  |  */, Ω /*   |   */, Ω /*   |    */, Ω /*     | */, Ω /* | */, Ω /* | */},
		/* cbExtend        | */ {Ω /* | */, Ω /*  | */, 0 /* | */, Ω /* | */, Ω /*  | */, Ω /*  | */, Ω /*  | */, Ω /*  |  */, Ω /*  |  */, Ω /*   |   */, 0 /*   |    */, Ω /*     | */, Ω /* | */, 0 /* | */},
		/* cbRI            | */ {Ω /* | */, Ω /*  | */, 0 /* | */, Ω /* | */, Ω /*  | */, Ω /*  | */, Ω /*  | */, Ω /*  |  */, Ω /*  |  */, Ω /*   |   */, 0 /*   |    */, Ω /*     | */, Ω /* | */, 0 /* | */},
		/* cbPrepend       | */ {0 /* | */, Ω /*  | */, 0 /* | */, 0 /* | */, 0 /*  | */, 0 /*  | */, 0 /*  | */, 0 /*  |  */, 0 /*  |  */, 0 /*   |   */, 0 /*   |    */, 0 /*     | */, 0 /* | */, 0 /* | */},
		/* cbHangulL       | */ {Ω /* | */, Ω /*  | */, 0 /* | */, Ω /* | */, Ω /*  | */, 0 /*  | */, 0 /*  | */, Ω /*  |  */, 0 /*  |  */, 0 /*   |   */, 0 /*   |    */, Ω /*     | */, Ω /* | */, 0 /* | */},
		/* cbHangulV       | */ {Ω /* | */, Ω /*  | */, 0 /* | */, Ω /* | */, Ω /*  | */, Ω /*  | */, 0 /*  | */, 0 /*  |  */, Ω /*  |  */, Ω /*   |   */, 0 /*   |    */, Ω /*     | */, Ω /* | */, 0 /* | */},
		/* cbHangulT       | */ {Ω /* | */, Ω /*  | */, 0 /* | */, Ω /* | */, Ω /*  | */, Ω /*  | */, Ω /*  | */, 0 /*  |  */, Ω /*  |  */, Ω /*   |   */, 0 /*   |    */, Ω /*     | */, Ω /* | */, 0 /* | */},
		/* cbHangulLV      | */ {Ω /* | */, Ω /*  | */, 0 /* | */, Ω /* | */, Ω /*  | */, Ω /*  | */, 0 /*  | */, 0 /*  |  */, Ω /*  |  */, Ω /*   |   */, 0 /*   |    */, Ω /*     | */, Ω /* | */, 0 /* | */},
		/* cbHangulLVT     | */ {Ω /* | */, Ω /*  | */, 0 /* | */, Ω /* | */, Ω /*  | */, Ω /*  | */, Ω /*  | */, 0 /*  |  */, Ω /*  |  */, Ω /*   |   */, 0 /*   |    */, Ω /*     | */, Ω /* | */, 0 /* | */},
		/* cbInCBLinker    | */ {Ω /* | */, Ω /*  | */, 0 /* | */, Ω /* | */, Ω /*  | */, Ω /*  | */, Ω /*  | */, Ω /*  |  */, Ω /*  |  */, Ω /*   |   */, 0 /*   |    */, 0 /*     | */, Ω /* | */, 0 /* | */},
		/* cbInCBConsonant | */ {Ω /* | */, Ω /*  | */, 0 /* | */, Ω /* | */, Ω /*  | */, Ω /*  | */, Ω /*  | */, Ω /*  |  */, Ω /*  |  */, Ω /*   |   */, 0 /*   |    */, Ω /*     | */, Ω /* | */, 0 /* | */},
		/* cbExtPic        | */ {Ω /* | */, Ω /*  | */, 0 /* | */, Ω /* | */, Ω /*  | */, Ω /*  | */, Ω /*  | */, Ω /*  |  */, Ω /*  |  */, Ω /*   |   */, 0 /*   |    */, Ω /*     | */, Ω /* | */, 0 /* | */},
		/* cbZWJ           | */ {Ω /* | */, Ω /*  | */, 0 /* | */, Ω /* | */, Ω /*  | */, Ω /*  | */, Ω /*  | */, Ω /*  |  */, Ω /*  |  */, Ω /*   |   */, 0 /*   |    */, Ω /*     | */, 0 /* | */, 0 /* | */},
	},
}

type HexInt int

func (h *HexInt) UnmarshalXMLAttr(attr xml.Attr) error {
	v, err := strconv.ParseUint(attr.Value, 16, 32)
	if err != nil {
		return err
	}
	*h = HexInt(v)
	return nil
}

type UCD struct {
	Description string `xml:"description"`
	Repertoire  struct {
		Group []struct {
			GeneralCategory      string `xml:"gc,attr"`
			GraphemeClusterBreak string `xml:"GCB,attr"`
			IndicConjunctBreak   string `xml:"InCB,attr"`
			ExtendedPictographic string `xml:"ExtPict,attr"`
			EastAsian            string `xml:"ea,attr"`

			// This maps the following tags:
			//   <char>, <reserved>, <surrogate>, <noncharacter>
			Char []struct {
				Codepoint      HexInt `xml:"cp,attr"`
				FirstCodepoint HexInt `xml:"first-cp,attr"`
				LastCodepoint  HexInt `xml:"last-cp,attr"`

				GeneralCategory      string `xml:"gc,attr"`
				GraphemeClusterBreak string `xml:"GCB,attr"`
				IndicConjunctBreak   string `xml:"InCB,attr"`
				ExtendedPictographic string `xml:"ExtPict,attr"`
				EastAsian            string `xml:"ea,attr"`
			} `xml:",any"`
		} `xml:"group"`
	} `xml:"repertoire"`
}

func main() {
	if err := run(); err != nil {
		fmt.Println(err)
		os.Exit(1)
	}
}

func run() error {
	if len(os.Args) <= 1 {
		fmt.Println(`Usage:
    go run CodepointWidthDetector_gen.go <path to ucd.nounihan.grouped.xml>

You can download the latest ucd.nounihan.grouped.xml from:
    https://www.unicode.org/Public/UCD/latest/ucdxml/ucd.nounihan.grouped.zip`)
		os.Exit(1)
	}

	data, err := os.ReadFile(os.Args[1])
	if err != nil {
		return fmt.Errorf("failed to read XML: %w", err)
	}

	ucd := &UCD{}
	err = xml.Unmarshal(data, ucd)
	if err != nil {
		return fmt.Errorf("failed to parse XML: %w", err)
	}

	values, err := extractValuesFromUCD(ucd)
	if err != nil {
		return err
	}

	// More stages = Less size. The trajectory roughly follows a+b*c^stages, where c < 1.
	// 4 still gives ~30% savings over 3 stages and going beyond 5 gives diminishing returns (<10%).
	trie := buildBestTrie(values, 2, 8, 4)
	rules := prepareRulesTable(JoinRules)
	totalSize := trie.TotalSize + rulesSize(rules)

	for cp, expected := range values {
		var v TrieType
		for _, s := range trie.Stages {
			v = s.Values[int(v)+((cp>>s.Shift)&s.Mask)]
		}
		if v != expected {
			return fmt.Errorf("trie sanity check failed for %U", cp)
		}
	}

	buf := &strings.Builder{}
	add := func(format string, a ...any) {
		_, _ = fmt.Fprintf(buf, format, a...)
	}

	add("// Generated by CodepointWidthDetector_gen.go\n")
	add("// on %s, from %s, %d bytes\n", time.Now().UTC().Format(time.RFC3339), ucd.Description, totalSize)
	add("// clang-format off\n")

	for i, s := range trie.Stages {
		width := 16
		if i != 0 {
			width = s.Mask + 1
		}
		add("static constexpr uint%d_t s_stage%d[] = {", s.Bits, i+1)
		for j, value := range s.Values {
			if j%width == 0 {
				add("\n   ")
			}
			add(" %#0*x,", s.Bits/4, value)
		}
		add("\n};\n")
	}

	add("static constexpr uint32_t s_joinRules[%d][%d] = {\n", len(rules), len(rules[0]))
	for _, table := range rules {
		add("    {\n")
		for _, r := range table {
			add("        %#032b,\n", r)
		}
		add("    },\n")
	}
	add("};\n")

	add("constexpr uint%d_t ucdLookup(const char32_t cp) noexcept\n", trie.Stages[len(trie.Stages)-1].Bits)
	add("{\n")
	for i, s := range trie.Stages {
		add("    const auto s%d = s_stage%d[", i+1, i+1)
		if i == 0 {
			add("cp >> %d", s.Shift)
		} else {
			add("s%d + ((cp >> %d) & %d)", i, s.Shift, s.Mask)
		}
		add("];\n")
	}
	add("    return s%d;\n", len(trie.Stages))
	add("}\n")

	add("constexpr uint8_t ucdGraphemeJoins(const uint8_t state, const uint8_t lead, const uint8_t trail) noexcept\n")
	add("{\n")
	add("    const auto l = lead & 15;\n")
	add("    const auto t = trail & 15;\n")
	add("    return (s_joinRules[state][l] >> (t * %d)) & %d;\n", bits.Len8(Ω), Ω)
	add("}\n")
	add("constexpr bool ucdGraphemeDone(const uint8_t state) noexcept\n")
	add("{\n")
	add("    return state == %d;\n", Ω)
	add("}\n")
	add("constexpr int ucdToCharacterWidth(const uint8_t val) noexcept\n")
	add("{\n")
	add("    return val >> 6;\n")
	add("}\n")
	add("// clang-format on\n")

	_, _ = os.Stdout.WriteString(buf.String())
	return nil
}

type TrieType uint32

func extractValuesFromUCD(ucd *UCD) ([]TrieType, error) {
	values := make([]TrieType, 1114112)
	fillRange(values, trieValue(cbOther, cwNarrow))

	for _, group := range ucd.Repertoire.Group {
		for _, char := range group.Char {
			generalCategory := coalesce(char.GeneralCategory, group.GeneralCategory)
			graphemeClusterBreak := coalesce(char.GraphemeClusterBreak, group.GraphemeClusterBreak)
			indicConjunctBreak := coalesce(char.IndicConjunctBreak, group.IndicConjunctBreak)
			extendedPictographic := coalesce(char.ExtendedPictographic, group.ExtendedPictographic)
			eastAsian := coalesce(char.EastAsian, group.EastAsian)

			firstCp, lastCp := int(char.FirstCodepoint), int(char.LastCodepoint)
			if char.Codepoint != 0 {
				firstCp, lastCp = int(char.Codepoint), int(char.Codepoint)
			}

			var cb ClusterBreak
			switch graphemeClusterBreak {
			case "XX": // Anything else
				cb = cbOther
			case "CR", "LF", "CN": // Carriage Return, Line Feed, Control
				// We ignore GB3 which demands that CR × LF do not break apart, because
				// a) these control characters won't normally reach our text storage
				// b) otherwise we're in a raw write mode and historically conhost stores them in separate cells
				cb = cbControl
			case "EX", "SM": // Extend, SpacingMark
				cb = cbExtend
			case "PP": // Prepend
				cb = cbPrepend
			case "ZWJ": // Zero Width Joiner
				cb = cbZWJ
			case "RI": // Regional Indicator
				cb = cbRI
			case "L": // Hangul Syllable Type L
				cb = cbHangulL
			case "V": // Hangul Syllable Type V
				cb = cbHangulV
			case "T": // Hangul Syllable Type T
				cb = cbHangulT
			case "LV": // Hangul Syllable Type LV
				cb = cbHangulLV
			case "LVT": // Hangul Syllable Type LVT
				cb = cbHangulLVT
			default:
				return nil, fmt.Errorf("unrecognized GCB %s for %U to %U", graphemeClusterBreak, firstCp, lastCp)
			}
			if extendedPictographic == "Y" {
				// Currently every single Extended_Pictographic codepoint happens to be GCB=XX.
				// This is fantastic for us because it means we can stuff it into the ClusterBreak enum
				// and treat it as an alias of EXTEND, but with the special GB11 properties.
				if cb != cbOther {
					return nil, fmt.Errorf("unexpected GCB %s with ExtPict=Y for %U to %U", graphemeClusterBreak, firstCp, lastCp)
				}
				cb = cbExtPic
			}
			switch indicConjunctBreak {
			case "None", "Extend":
				break
			case "Linker":
				// Similarly here, we can treat it as an alias for EXTEND, but with the GB9c properties.
				if cb != cbExtend {
					return nil, fmt.Errorf("unexpected GCB %s with InCB=Linker for %U to %U", graphemeClusterBreak, firstCp, lastCp)
				}
				cb = cbInCBLinker
			case "Consonant":
				// Similarly here, we can treat it as an alias for OTHER, but with the GB9c properties.
				if cb != cbOther {
					return nil, fmt.Errorf("unexpected GCB %s with InCB=Consonant for %U to %U", graphemeClusterBreak, firstCp, lastCp)
				}
				cb = cbInCBConsonant
			default:
				return nil, fmt.Errorf("unrecognized InCB %s for %U to %U", indicConjunctBreak, firstCp, lastCp)
			}

			var width CharacterWidth
			switch eastAsian {
			case "N", "Na", "H": // neutral, narrow, half-width
				width = cwNarrow
			case "F", "W": // full-width, wide
				width = cwWide
			case "A": // ambiguous
				width = cwAmbiguous
			default:
				return nil, fmt.Errorf("unrecognized ea %s for %U to %U", eastAsian, firstCp, lastCp)
			}
			// There's no "ea" attribute for "zero width" so we need to do that ourselves. This matches:
			//   Mc: Mark, spacing combining
			//   Me: Mark, enclosing
			//   Mn: Mark, non-spacing
			//   Cf: Control, format
			if strings.HasPrefix(generalCategory, "M") || generalCategory == "Cf" {
				width = cwZeroWidth
			}

			fillRange(values[firstCp:lastCp+1], trieValue(cb, width))
		}
	}

	// Box-drawing and block elements are ambiguous according to their EastAsian attribute,
	// but by convention terminals always consider them to be narrow.
	fillRange(values[0x2500:0x259F+1], trieValue(cbOther, cwNarrow))
	// U+FE0F Variation Selector-16 is used to turn unqualified Emojis into qualified ones.
	// By convention, this also turns them from being ambiguous, = narrow by default, into wide ones.
	fillRange(values[0xFE0F:0xFE0F+1], trieValue(cbExtend, cwWide))

	return values, nil
}

func trieValue(cb ClusterBreak, width CharacterWidth) TrieType {
	return TrieType(byte(cb) | byte(width)<<6)
}

func coalesce(a, b string) string {
	if a != "" {
		return a
	}
	return b
}

type Stage struct {
	Values []TrieType
	Shift  int
	Mask   int
	Bits   int
}

type Trie struct {
	Stages    []*Stage
	TotalSize int
}

func buildBestTrie(uncompressed []TrieType, minShift, maxShift, stages int) *Trie {
	delta := maxShift - minShift + 1
	results := make(chan *Trie)
	bestTrie := &Trie{TotalSize: math.MaxInt}

	iters := 1
	for i := 1; i < stages; i++ {
		iters *= delta
	}

	for i := 0; i < iters; i++ {
		go func(i int) {
			// Given minShift=2, maxShift=3, depth=3 this generates
			//   [2 2 2]
			//   [3 2 2]
			//   [2 3 2]
			//   [3 3 2]
			//   [2 2 3]
			//   [3 2 3]
			//   [2 3 3]
			//   [3 3 3]
			shifts := make([]int, stages-1)
			for j := range shifts {
				shifts[j] = minShift + i%delta
				i /= delta
			}
			results <- buildTrie(uncompressed, shifts)
		}(i)
	}

	for i := 0; i < iters; i++ {
		t := <-results
		if bestTrie.TotalSize > t.TotalSize {
			bestTrie = t
		}
	}
	return bestTrie
}

func buildTrie(uncompressed []TrieType, shifts []int) *Trie {
	var cumulativeShift int
	var stages []*Stage

	for _, shift := range shifts {
		chunkSize := 1 << shift
		cache := map[string]TrieType{}
		compressed := make([]TrieType, 0, len(uncompressed)/8)
		offsets := make([]TrieType, 0, len(uncompressed)/chunkSize)

		for i := 0; i < len(uncompressed); i += chunkSize {
			chunk := uncompressed[i:min(len(uncompressed), i+chunkSize)]
			// Cast the integer slice to a string so that it can be hashed.
			key := unsafe.String((*byte)(unsafe.Pointer(&chunk[0])), len(chunk)*int(unsafe.Sizeof(chunk[0])))
			offset, exists := cache[key]

			if !exists {
				// For a 4-stage trie searching for existing occurrences of chunk in compressed yields a ~10%
				// compression improvement. Checking for overlaps with the tail end of compressed yields another ~15%.
				// FYI I tried to shuffle the order of compressed chunks but found that this has a negligible impact.
				if existing := findExisting(compressed, chunk); existing != -1 {
					offset = TrieType(existing)
					cache[key] = offset
				} else {
					overlap := measureOverlap(compressed, chunk)
					compressed = append(compressed, chunk[overlap:]...)
					offset = TrieType(len(compressed) - len(chunk))
					cache[key] = offset
				}
			}

			offsets = append(offsets, offset)
		}

		stages = append(stages, &Stage{
			Values: compressed,
			Shift:  cumulativeShift,
			Mask:   chunkSize - 1,
		})

		uncompressed = offsets
		cumulativeShift += shift
	}

	stages = append(stages, &Stage{
		Values: uncompressed,
		Shift:  cumulativeShift,
		Mask:   math.MaxInt32,
	})
	slices.Reverse(stages)

	for _, s := range stages {
		m := slices.Max(s.Values)
		if m <= 0xff {
			s.Bits = 8
		} else if m <= 0xffff {
			s.Bits = 16
		} else {
			s.Bits = 32
		}
	}

	totalSize := 0
	for _, s := range stages {
		totalSize += (s.Bits / 8) * len(s.Values)
	}

	return &Trie{
		Stages:    stages,
		TotalSize: totalSize,
	}
}

// Finds needle in haystack. Returns -1 if it couldn't be found.
func findExisting(haystack, needle []TrieType) int {
	if len(haystack) == 0 || len(needle) == 0 {
		return -1
	}

	s := int(unsafe.Sizeof(TrieType(0)))
	h := unsafe.Slice((*byte)(unsafe.Pointer(&haystack[0])), len(haystack)*s)
	n := unsafe.Slice((*byte)(unsafe.Pointer(&needle[0])), len(needle)*s)
	i := 0

	for {
		i = bytes.Index(h[i:], n)
		if i == -1 {
			return -1
		}
		if i%s == 0 {
			return i / s
		}
	}
}

// Given two slices, this returns the amount by which `prev`s end overlaps with `next`s start.
// That is, given [0,1,2,3,4] and [2,3,4,5] this returns 3 because [2,3,4] is the "overlap".
func measureOverlap(prev, next []TrieType) int {
	for overlap := min(len(prev), len(next)); overlap >= 0; overlap-- {
		if slices.Equal(prev[len(prev)-overlap:], next[:overlap]) {
			return overlap
		}
	}
	return 0
}

func prepareRulesTable(rules [][][]uint8) [][]uint32 {
	compressed := make([][]uint32, len(rules))
	for i := range compressed {
		compressed[i] = make([]uint32, 16)
	}

	for prevIndex, table := range rules {
		for lead, row := range table {
			if len(row) > 16 {
				panic("can't pack row into 32 bits")
			}

			var nextIndices uint32
			for trail, nextIndex := range row {
				if nextIndex > Ω {
					panic("can't pack table index into 2 bits")
				}
				nextIndices |= uint32(nextIndex) << (trail * 2)
			}

			compressed[prevIndex][lead] = nextIndices
		}
	}

	return compressed
}

func rulesSize(rules [][]uint32) int {
	// Each rules item has the same length. Each item is 32 bits = 4 bytes.
	return len(rules) * len(rules[0]) * 4
}

func fillRange[T any](s []T, v T) {
	for i := range s {
		s[i] = v
	}
}
