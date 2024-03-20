package main

import (
	"encoding/xml"
	"fmt"
	"math"
	"os"
	"slices"
	"strconv"
	"strings"
	"time"
	"unsafe"
)

type CharacterWidth int

const (
	CW_ZERO_WIDTH CharacterWidth = iota
	CW_NARROW
	CW_WIDE
	CW_AMBIGUOUS
)

type ClusterBreak int

const (
	CB_OTHER ClusterBreak = iota
	CB_CONTROL
	CB_EXTEND
	CB_PREPEND
	CB_ZERO_WIDTH_JOINER
	CB_REGIONAL_INDICATOR
	CB_HANGUL_L
	CB_HANGUL_V
	CB_HANGUL_T
	CB_HANGUL_LV
	CB_HANGUL_LVT
	CB_CONJUNCT_LINKER
	CB_EXTENDED_PICTOGRAPHIC

	CB_COUNT
)

type TrieType uint32

func trieValue(cb ClusterBreak, width CharacterWidth) TrieType {
	return TrieType(byte(cb) | byte(width)<<6)
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
    go run CodepointWidthDetector_triegen.go <path to ucd.nounihan.grouped.xml>

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

	stages := buildTrie(values, []int{4, 4, 4})
	rules := buildJoinRules()

	for cp, expected := range values {
		var v TrieType
		for _, s := range stages {
			v = s.Values[int(v)+((cp>>s.Shift)&s.Mask)]
		}
		if v != expected {
			return fmt.Errorf("trie sanity check failed for %U", cp)
		}
	}

	totalSize := len(rules) * len(rules)
	for _, s := range stages {
		totalSize += (s.Bits / 8) * len(s.Values)
	}

	buf := &strings.Builder{}

	_, _ = fmt.Fprintf(buf, "// Generated by CodepointWidthDetector_triegen.go\n")
	_, _ = fmt.Fprintf(buf, "// on %v, from %s, %d bytes\n", time.Now().UTC().Format(time.RFC3339), ucd.Description, totalSize)
	_, _ = fmt.Fprintf(buf, "// clang-format off\n")

	for stageIndex, s := range stages {
		octals := s.Bits / 4
		width := 16
		if stageIndex != 0 {
			width = s.Mask + 1
		}

		_, _ = fmt.Fprintf(buf, "static constexpr uint%d_t s_stage%d[] = {", s.Bits, stageIndex+1)
		for valueIndex, value := range s.Values {
			if valueIndex%width == 0 {
				buf.WriteString("\n   ")
			}
			_, _ = fmt.Fprintf(buf, " 0x%0*x,", octals, value)
		}
		buf.WriteString("\n};\n")
	}

	_, _ = fmt.Fprintf(buf, "static constexpr uint8_t s_joinRules[%d][%d] = {", len(rules), len(rules))
	for _, row := range rules {
		buf.WriteString("\n   ")
		for _, val := range row {
			var i int
			if val {
				i = 1
			}
			_, _ = fmt.Fprintf(buf, " %d,", i)
		}
	}
	_, _ = fmt.Fprintf(buf, "\n};\n")

	_ = (`
{{end}}

[[msvc::forceinline]] constexpr uint8_t ucdLookup(const char32_t cp) noexcept
{
    const auto s1 = s_stage1[cp >> {stage1_shift}];
    const auto s2 = s_stage2[s1 + ((cp >> {stage2_shift}) & {stage2_mask})];
    const auto s3 = s_stage3[s2 + ((cp >> {stage3_shift}) & {stage3_mask})];
    const auto s4 = s_stage4[s3 + (cp & {stage4_mask})];
    return s4;
}
[[msvc::forceinline]] constexpr uint8_t ucdGraphemeJoins(const uint8_t lead, const uint8_t trail) noexcept
{
    const auto l = lead & 15;
    const auto t = trail & 15;
    return s_joinRules[l][t];
}
[[msvc::forceinline]] constexpr int ucdToCharacterWidth(const uint8_t val) noexcept
{
    return (val >> {CHARACTER_WIDTH_SHIFT}) & {CHARACTER_WIDTH_MASK};
}
// clang-format on`)

	_ = rules
	fmt.Println(buf.String())
	return nil
}

func extractValuesFromUCD(ucd *UCD) ([]TrieType, error) {
	values := make([]TrieType, 1114112)
	fillRange(values, trieValue(CB_OTHER, CW_NARROW))

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

			// <reserved cp="2065" age="unassigned" na="" gc="Cn" bc="BN" lb="XX" sc="Zzzz" scx="Zzzz" DI="Y" ODI="Y" Gr_Base="N" Pat_Syn="N" GCB="CN" CWKCF="Y" NFKC_CF="" vo="U" NFKC_SCF=""/>

			var (
				cb    ClusterBreak
				width CharacterWidth
			)

			switch graphemeClusterBreak {
			case "XX": // Anything else
				cb = CB_OTHER
			case "CR", "LF", "CN": // Carriage Return, Line Feed, Control
				// We ignore GB3 which demands that CR × LF do not break apart, because
				// a) these control characters won't normally reach our text storage
				// b) otherwise we're in a raw write mode and historically conhost stores them in separate cells
				cb = CB_CONTROL
			case "EX", "SM": // Extend, SpacingMark
				cb = CB_EXTEND
			case "PP": // Prepend
				cb = CB_PREPEND
			case "ZWJ": // Zero Width Joiner
				cb = CB_ZERO_WIDTH_JOINER
			case "RI": // Regional Indicator
				cb = CB_REGIONAL_INDICATOR
			case "L": // Hangul Syllable Type L
				cb = CB_HANGUL_L
			case "V": // Hangul Syllable Type V
				cb = CB_HANGUL_V
			case "T": // Hangul Syllable Type T
				cb = CB_HANGUL_T
			case "LV": // Hangul Syllable Type LV
				cb = CB_HANGUL_LV
			case "LVT": // Hangul Syllable Type LVT
				cb = CB_HANGUL_LVT
			default:
				return nil, fmt.Errorf("unrecognized GCB %s for %U to %U", graphemeClusterBreak, firstCp, lastCp)
			}

			if extendedPictographic == "Y" {
				// Currently every single Extended_Pictographic codepoint happens to be GCB=XX.
				// This is fantastic for us because it means we can stuff it into the ClusterBreak enum
				// and treat it as an alias of EXTEND, but with the special GB11 properties.
				if cb != CB_OTHER {
					return nil, fmt.Errorf("unexpected GCB %s for ExtPict for %U to %U", graphemeClusterBreak, firstCp, lastCp)
				}
				cb = CB_EXTENDED_PICTOGRAPHIC
			}

			if indicConjunctBreak == "Linker" {
				// Similarly here, we can treat it as an alias for EXTEND, but with the GB9c properties.
				if cb != CB_EXTEND {
					return nil, fmt.Errorf("unexpected GCB %s for InCB=Linker for %U to %U", graphemeClusterBreak, firstCp, lastCp)
				}
				cb = CB_CONJUNCT_LINKER
			}

			switch eastAsian {
			case "N", "Na", "H": // neutral, narrow, half-width
				width = CW_NARROW
			case "F", "W": // full-width, wide
				width = CW_WIDE
			case "A": // ambiguous
				width = CW_AMBIGUOUS
			default:
				return nil, fmt.Errorf("unrecognized ea %s for %U to %U", eastAsian, firstCp, lastCp)
			}

			// There's no "ea" attribute for "zero width" so we need to do that ourselves. This matches:
			//   Mc: Mark, spacing combining
			//   Me: Mark, enclosing
			//   Mn: Mark, non-spacing
			//   Cf: Control, format
			if strings.HasPrefix(generalCategory, "M") || generalCategory == "Cf" {
				width = CW_ZERO_WIDTH
			}

			fillRange(values[firstCp:lastCp+1], trieValue(cb, width))
		}
	}

	// Box-drawing and block elements require 1-cell alignment.
	// Most characters in this range have an ambiguous width otherwise.
	fillRange(values[0x2500:0x259F+1], trieValue(CB_OTHER, CW_NARROW))
	// hexagrams are historically narrow
	fillRange(values[0x4DC0:0x4DFF+1], trieValue(CB_OTHER, CW_NARROW))
	// narrow combining ligatures (split into left/right halves, which take 2 columns together)
	fillRange(values[0xFE20:0xFE2F+1], trieValue(CB_OTHER, CW_NARROW))

	return values, nil
}

type Stage struct {
	Values []TrieType
	Shift  int
	Mask   int
	Bits   int
	Index  int
}

func buildTrie(uncompressed []TrieType, shifts []int) []*Stage {
	var cumulativeShift int
	var stages []*Stage

	for _, shift := range shifts {
		chunkSize := 1 << shift
		cache := map[string]TrieType{}
		offsets := []TrieType(nil)
		compressed := []TrieType(nil)

		for i := 0; i < len(uncompressed); i += chunkSize {
			chunk := uncompressed[i:min(i+chunkSize, len(uncompressed))]
			// Cast the integer slice to a string so that it's hashable.
			key := unsafe.String((*byte)(unsafe.Pointer(&chunk[0])), len(chunk)*int(unsafe.Sizeof(chunk[0])))

			offset, exists := cache[key]
			if !exists {
				offset = TrieType(len(compressed))
				cache[key] = offset
				compressed = append(compressed, chunk...)
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
		Mask:   math.MaxUint32,
	})

	slices.Reverse(stages)

	for i, s := range stages {
		s.Bits = bitSize(slices.Max(s.Values))
		s.Index = i + 1
	}

	return stages
}

func buildJoinRules() [CB_COUNT][CB_COUNT]bool {
	// UAX #29 states:
	// > Note: Testing two adjacent characters is insufficient for determining a boundary.
	//
	// I completely agree, but I really hate it. So this code trades off correctness for simplicity
	// by using a simple lookup table anyway. Under most circumstances users won't notice,
	// because as far as I can see this only behaves different for degenerate ("invalid") Unicode.
	// It reduces our code complexity significantly and is way *way* faster.
	//
	// This is a great reference for the resulting table:
	//   https://www.unicode.org/Public/UCD/latest/ucd/auxiliary/GraphemeBreakTest.html

	// NOTE: We build the table in reverse, because rules with lower numbers take priority.
	// (This is primarily relevant for GB9b vs. GB4.)

	// Otherwise, break everywhere.
	// GB999: Any ÷ Any
	var rules [CB_COUNT][CB_COUNT]bool

	// Do not break within emoji flag sequences. That is, do not break between regional indicator
	// (RI) symbols if there is an odd number of RI characters before the break point.
	// GB13: [^RI] (RI RI)* RI × RI
	// GB12: sot (RI RI)* RI × RI
	//
	// We cheat here by not checking that the number of RIs is even. Meh!
	rules[CB_REGIONAL_INDICATOR][CB_REGIONAL_INDICATOR] = true

	// Do not break within emoji modifier sequences or emoji zwj sequences.
	// GB11: \p{Extended_Pictographic} Extend* ZWJ × \p{Extended_Pictographic}
	//
	// We cheat here by not checking that the ZWJ is preceded by an ExtPic. Meh!
	rules[CB_ZERO_WIDTH_JOINER][CB_EXTENDED_PICTOGRAPHIC] = true

	// Do not break within certain combinations with Indic_Conjunct_Break (InCB)=Linker.
	// GB9c: \p{InCB=Consonant} [\p{InCB=Extend}\p{InCB=Linker}]* \p{InCB=Linker} [\p{InCB=Extend}\p{InCB=Linker}]* × \p{InCB=Consonant}
	//
	// I'm sure GB9c is great for these languages, but honestly the definition is complete whack.
	// Just look at that chonker! This isn't a "cheat" like the others above, this is a reinvention:
	// We treat it as having both ClusterBreak.PREPEND and ClusterBreak.EXTEND properties.
	fillRange(rules[CB_CONJUNCT_LINKER][:], true)
	for _, lead := range rules {
		lead[CB_CONJUNCT_LINKER] = true
	}

	// Do not break before SpacingMarks, or after Prepend characters.
	// GB9b: Prepend ×
	fillRange(rules[CB_PREPEND][:], true)

	// Do not break before SpacingMarks, or after Prepend characters.
	// GB9a: × SpacingMark
	// Do not break before extending characters or ZWJ.
	// GB9: × (Extend | ZWJ)
	for _, lead := range rules {
		// CodepointWidthDetector_triegen.py sets SpacingMarks to ClusterBreak.EXTEND as well,
		// since they're entirely identical to GB9's Extend.
		lead[CB_EXTEND] = true
		lead[CB_ZERO_WIDTH_JOINER] = true
	}

	// Do not break Hangul syllable sequences.
	// GB8: (LVT | T) x T
	rules[CB_HANGUL_LVT][CB_HANGUL_T] = true
	rules[CB_HANGUL_T][CB_HANGUL_T] = true
	// GB7: (LV | V) x (V | T)
	rules[CB_HANGUL_LV][CB_HANGUL_T] = true
	rules[CB_HANGUL_LV][CB_HANGUL_V] = true
	rules[CB_HANGUL_V][CB_HANGUL_V] = true
	rules[CB_HANGUL_V][CB_HANGUL_T] = true
	// GB6: L x (L | V | LV | LVT)
	rules[CB_HANGUL_L][CB_HANGUL_L] = true
	rules[CB_HANGUL_L][CB_HANGUL_V] = true
	rules[CB_HANGUL_L][CB_HANGUL_LV] = true
	rules[CB_HANGUL_L][CB_HANGUL_LVT] = true

	// Do not break between a CR and LF. Otherwise, break before and after controls.
	// GB5: ÷ (Control | CR | LF)
	for _, lead := range rules {
		lead[CB_CONTROL] = false
	}
	// GB4: (Control | CR | LF) ÷
	fillRange(rules[CB_CONTROL][:], false)

	// We ignore GB3 which demands that CR × LF do not break apart, because
	// a) these control characters won't normally reach our text storage
	// b) otherwise we're in a raw write mode and historically conhost stores them in separate cells

	// We also ignore GB1 and GB2 which demand breaks at the start and end,
	// because that's not part of the loops in GraphemeNext/Prev and not this table.
	return rules
}

func coalesce(a, b string) string {
	if a != "" {
		return a
	}
	return b
}

func fillRange[T any](s []T, v T) {
	for i := range s {
		s[i] = v
	}
}

func bitSize(x TrieType) int {
	if x <= 0xff {
		return 8
	}
	if x <= 0xffff {
		return 16
	}
	return 32
}
