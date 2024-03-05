// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

#include "pch.h"
#include "Appearances.h"
#include "Appearances.g.cpp"
#include "AxisKeyValuePair.g.cpp"
#include "FeatureKeyValuePair.g.cpp"

#include "EnumEntry.h"
#include <LibraryResources.h>
#include "..\WinRTUtils\inc\Utils.h"

using namespace winrt::Windows::UI::Text;
using namespace winrt::Windows::UI::Xaml;
using namespace winrt::Windows::UI::Xaml::Controls;
using namespace winrt::Windows::UI::Xaml::Data;
using namespace winrt::Windows::UI::Xaml::Navigation;
using namespace winrt::Windows::Foundation;
using namespace winrt::Windows::Foundation::Collections;
using namespace winrt::Microsoft::Terminal::Settings::Model;

static constexpr std::array<std::wstring_view, 11> DefaultFeatures{
    L"rlig",
    L"locl",
    L"ccmp",
    L"calt",
    L"liga",
    L"clig",
    L"rnrn",
    L"kern",
    L"mark",
    L"mkmk",
    L"dist"
};

namespace winrt::Microsoft::Terminal::Settings::Editor::implementation
{
    static winrt::hstring getLocalizedStringByIndex(IDWriteLocalizedStrings* strings, UINT32 index)
    {
        UINT32 length = 0;
        THROW_IF_FAILED(strings->GetStringLength(index, &length));

        winrt::impl::hstring_builder builder{ length };
        THROW_IF_FAILED(strings->GetString(index, builder.data(), length + 1));

        return builder.to_hstring();
    }

    static UINT32 getLocalizedStringIndex(IDWriteLocalizedStrings* strings, const wchar_t* locale, UINT32 fallback)
    {
        UINT32 index;
        BOOL exists;
        if (FAILED(strings->FindLocaleName(locale, &index, &exists)) || !exists)
        {
            index = fallback;
        }
        return index;
    }

    bool Font::HasPowerlineCharacters()
    {
        if (!_hasPowerlineCharacters.has_value())
        {
            wil::com_ptr<IDWriteFont> font;
            BOOL exists = false;

            if (SUCCEEDED(_family->GetFont(0, font.addressof())))
            {
                // We're actually checking for the "Extended" PowerLine glyph set.
                // They're more fun.
                std::ignore = font->HasCharacter(0xE0B6, &exists);
            }

            _hasPowerlineCharacters = exists == TRUE;
        }

        return *_hasPowerlineCharacters;
    }

    IMap<winrt::hstring, winrt::hstring> Font::FontAxesTagsAndNames()
    {
        if (!_fontAxesTagsAndNames)
        {
            // Passing a collection by reference gives the _generate() function the liberty to return early,
            // while we still ensure that the WinRT collection is initialized to a non-null value.
            std::unordered_map<winrt::hstring, winrt::hstring> tagsAndNames;
            _generateFontAxesTagsAndNames(tagsAndNames);
            _fontFeaturesTagsAndNames = winrt::single_threaded_map(std::move(tagsAndNames));
        }

        return _fontAxesTagsAndNames;
    }

    void Font::_generateFontAxesTagsAndNames(std::unordered_map<winrt::hstring, winrt::hstring>& tagsAndNames)
    {
        wil::com_ptr<IDWriteFont> font;
        THROW_IF_FAILED(_family->GetFont(0, font.addressof()));

        wil::com_ptr<IDWriteFontFace> fontFace;
        THROW_IF_FAILED(font->CreateFontFace(fontFace.addressof()));

        const auto fontFace5 = fontFace.try_query<IDWriteFontFace5>();
        if (!fontFace5)
        {
            return;
        }

        const auto axesCount = fontFace5->GetFontAxisValueCount();
        if (axesCount == 0)
        {
            return;
        }

        std::vector<DWRITE_FONT_AXIS_VALUE> axesVector(axesCount);
        THROW_IF_FAILED(fontFace5->GetFontAxisValues(axesVector.data(), axesCount));

        wchar_t localeName[LOCALE_NAME_MAX_LENGTH];
        const auto localeToTry = GetUserDefaultLocaleName(localeName, LOCALE_NAME_MAX_LENGTH) ? localeName : L"en-US";

        wil::com_ptr<IDWriteFontResource> fontResource;
        THROW_IF_FAILED(fontFace5->GetFontResource(fontResource.addressof()));

        for (UINT32 i = 0; i < axesCount; ++i)
        {
            wil::com_ptr<IDWriteLocalizedStrings> names;
            THROW_IF_FAILED(fontResource->GetAxisNames(i, names.addressof()));

            // As per MSDN:
            // > The font author may not have supplied names for some font axes.
            // > The localized strings will be empty in that case.
            if (names->GetCount() == 0)
            {
                continue;
            }

            const auto idx = getLocalizedStringIndex(names.get(), localeToTry, 0);
            auto name = getLocalizedStringByIndex(names.get(), idx);
            tagsAndNames.emplace(_tagToString(axesVector[i].axisTag), std::move(name));
        }
    }

    IMap<hstring, hstring> Font::FontFeaturesTagsAndNames()
    {
        if (!_fontFeaturesTagsAndNames)
        {
            // Passing a collection by reference gives the _generate() function the liberty to return early,
            // while we still ensure that the WinRT collection is initialized to a non-null value.
            std::unordered_map<winrt::hstring, winrt::hstring> tagsAndNames;
            _generateFontFeaturesTagsAndNames(tagsAndNames);
            _fontFeaturesTagsAndNames = winrt::single_threaded_map(std::move(tagsAndNames));
        }

        return _fontFeaturesTagsAndNames;
    }

    void Font::_generateFontFeaturesTagsAndNames(std::unordered_map<winrt::hstring, winrt::hstring>& tagsAndNames)
    {
        wil::com_ptr<IDWriteFactory> factory;
        THROW_IF_FAILED(DWriteCreateFactory(DWRITE_FACTORY_TYPE_SHARED, __uuidof(factory), reinterpret_cast<::IUnknown**>(factory.addressof())));

        wil::com_ptr<IDWriteTextAnalyzer> textAnalyzer;
        THROW_IF_FAILED(factory->CreateTextAnalyzer(textAnalyzer.addressof()));
        const auto textAnalyzer2 = textAnalyzer.query<IDWriteTextAnalyzer2>();

        wil::com_ptr<IDWriteFont> font;
        THROW_IF_FAILED(_family->GetFont(0, font.addressof()));

        wil::com_ptr<IDWriteFontFace> fontFace;
        THROW_IF_FAILED(font->CreateFontFace(fontFace.addressof()));

        static constexpr DWRITE_SCRIPT_ANALYSIS scriptAnalysis{};
        UINT32 tagCount;
        // we have to call GetTypographicFeatures twice, first to get the actual count then to get the features
        if (FAILED(textAnalyzer2->GetTypographicFeatures(fontFace.get(), scriptAnalysis, L"en-us", 0, &tagCount, nullptr)))
        {
            return;
        }

        std::vector<DWRITE_FONT_FEATURE_TAG> tags{ tagCount };
        if (FAILED(textAnalyzer2->GetTypographicFeatures(fontFace.get(), scriptAnalysis, L"en-us", tagCount, &tagCount, tags.data())))
        {
            return;
        }

        for (const auto& tag : tags)
        {
            const auto tagString = _tagToString(tag);
            const auto formattedResourceString = fmt::format(L"Profile_FontFeature_{}", tagString);
            hstring localizedName{ tagString };

            // we have resource strings for common font features, see if one for this feature exists
            if (HasLibraryResourceWithName(formattedResourceString))
            {
                localizedName = GetLibraryResourceString(formattedResourceString);
            }

            tagsAndNames.emplace(tagString, localizedName);
        }
    }

    winrt::hstring Font::_tagToString(UINT32 tag)
    {
        const wchar_t buffer[] = {
            static_cast<wchar_t>((tag >> 0) & 0xFF),
            static_cast<wchar_t>((tag >> 8) & 0xFF),
            static_cast<wchar_t>((tag >> 16) & 0xFF),
            static_cast<wchar_t>((tag >> 24) & 0xFF),
        };
        return winrt::hstring{ &buffer[0], 4 };
    }

    AxisKeyValuePair::AxisKeyValuePair(winrt::hstring key, float value, IMap<winrt::hstring, float> baseMap) :
        _key{ std::move(key) },
        _value{ value },
        _baseMap{ std::move(baseMap) }
    {
    }

    IInspectable AxisKeyValuePair::Key() const
    {
        return winrt::box_value(_key);
    }

    void AxisKeyValuePair::Key(const IInspectable& boxedKey)
    {
        const auto key = winrt::unbox_value<winrt::hstring>(boxedKey);
        if (key != _key)
        {
            _baseMap.Remove(_key);
            _key = key;
            _baseMap.Insert(_key, _value);
            _PropertyChangedHandlers(*this, PropertyChangedEventArgs{ L"Key" });
        }
    }

    float AxisKeyValuePair::Value() const
    {
        return _value;
    }

    void AxisKeyValuePair::Value(float value)
    {
        if (value != _value)
        {
            _value = value;
            _baseMap.Insert(_key, _value);
            _PropertyChangedHandlers(*this, PropertyChangedEventArgs{ L"Value" });
        }
    }

    FeatureKeyValuePair::FeatureKeyValuePair(hstring key, uint32_t value, IMap<winrt::hstring, uint32_t> baseMap) :
        _key{ std::move(key) },
        _value{ value },
        _baseMap{ std::move(baseMap) }
    {
    }

    IInspectable FeatureKeyValuePair::Key()
    {
        return winrt::box_value(_key);
    }

    void FeatureKeyValuePair::Key(const IInspectable& boxedKey)
    {
        const auto key = winrt::unbox_value<winrt::hstring>(boxedKey);
        if (key != _key)
        {
            _baseMap.Remove(_key);
            _key = key;
            _baseMap.Insert(_key, _value);
            _PropertyChangedHandlers(*this, PropertyChangedEventArgs{ L"Key" });
        }
    }

    uint32_t FeatureKeyValuePair::Value()
    {
        return _value;
    }

    void FeatureKeyValuePair::Value(uint32_t value)
    {
        if (value != _value)
        {
            _value = value;
            _baseMap.Insert(_key, _value);
            _PropertyChangedHandlers(*this, PropertyChangedEventArgs{ L"Value" });
        }
    }

    AppearanceViewModel::AppearanceViewModel(const Model::AppearanceConfig& appearance) :
        _appearance{ appearance }
    {
        // Add a property changed handler to our own property changed event.
        // This propagates changes from the settings model to anybody listening to our
        //  unique view model members.
        PropertyChanged([this](auto&&, const PropertyChangedEventArgs& args) {
            const auto viewModelProperty{ args.PropertyName() };
            if (viewModelProperty == L"BackgroundImagePath")
            {
                // notify listener that all background image related values might have changed
                //
                // We need to do this so if someone manually types "desktopWallpaper"
                // into the path TextBox, we properly update the checkbox and stored
                // _lastBgImagePath. Without this, then we'll permanently hide the text
                // box, prevent it from ever being changed again.
                _NotifyChanges(L"UseDesktopBGImage", L"BackgroundImageSettingsVisible");
            }
        });

        // Cache the original BG image path. If the user clicks "Use desktop
        // wallpaper", then un-checks it, this is the string we'll restore to
        // them.
        if (BackgroundImagePath() != L"desktopWallpaper")
        {
            _lastBgImagePath = BackgroundImagePath();
        }
    }

    winrt::hstring AppearanceViewModel::FontFace() const
    {
        return _appearance.SourceProfile().FontInfo().FontFace();
    }

    void AppearanceViewModel::FontFace(const winrt::hstring& value)
    {
        const auto fontInfo = _appearance.SourceProfile().FontInfo();
        if (fontInfo.FontFace() == value)
        {
            return;
        }

        fontInfo.FontFace(value);
        _NotifyChanges(L"HasFontFace", L"FontFace");

        _refreshFontFaceDependents();
    }

    bool AppearanceViewModel::HasFontFace() const
    {
        return _appearance.SourceProfile().FontInfo().HasFontFace();
    }

    void AppearanceViewModel::ClearFontFace()
    {
        const auto fontInfo = _appearance.SourceProfile().FontInfo();
        const auto hadValue = fontInfo.HasFontFace();

        fontInfo.ClearFontFace();

        if (hadValue)
        {
            _NotifyChanges(L"HasFontFace", L"FontFace");
        }

        _refreshFontFaceDependents();
    }

    Model::FontConfig AppearanceViewModel::FontFaceOverrideSource() const
    {
        return _appearance.SourceProfile().FontInfo().FontFaceOverrideSource();
    }

    void AppearanceViewModel::_refreshFontFaceDependents()
    {
        wil::com_ptr<IDWriteFactory> factory;
        THROW_IF_FAILED(DWriteCreateFactory(DWRITE_FACTORY_TYPE_SHARED, __uuidof(factory), reinterpret_cast<::IUnknown**>(factory.addressof())));

        wil::com_ptr<IDWriteFontCollection> fontCollection;
        THROW_IF_FAILED(factory->GetSystemFontCollection(fontCollection.addressof(), FALSE));

        const auto fontFace = FontFace();
        std::wstring_view remainingFaceNames{ fontFace };
        std::wstring missingFonts;
        std::wstring proportionalFonts;

        til::iterate_font_families(remainingFaceNames, [&](wil::zwstring_view name) {
            UINT32 index = 0;
            BOOL exists = FALSE;
            THROW_IF_FAILED(fontCollection->FindFamilyName(name.c_str(), &index, &exists));

            BOOL proportional = FALSE;
            if (exists)
            {
                wil::com_ptr<IDWriteFontFamily> fontFamily;
                THROW_IF_FAILED(fontCollection->GetFontFamily(index, fontFamily.addressof()));

                wil::com_ptr<IDWriteFont> font;
                THROW_IF_FAILED(fontFamily->GetFirstMatchingFont(DWRITE_FONT_WEIGHT_NORMAL, DWRITE_FONT_STRETCH_NORMAL, DWRITE_FONT_STYLE_NORMAL, font.addressof()));

                const auto font1 = font.query<IDWriteFont1>();
                proportional = font1->IsMonospacedFont();
            }

            std::wstring* accumulator;
            if (!exists)
            {
                accumulator = &missingFonts;
            }
            else if (proportional)
            {
                accumulator = &proportionalFonts;
            }
            else
            {
                return;
            }

            if (!accumulator->empty())
            {
                accumulator->append(L", ");
            }
            accumulator->append(name);
        });

        MissingFontWarningMessage(winrt::hstring{ missingFonts });
        ProportionalFontWarningMessage(winrt::hstring{ proportionalFonts });
    }

    double AppearanceViewModel::LineHeight() const
    {
        const auto fontInfo = _appearance.SourceProfile().FontInfo();
        const auto cellHeight = fontInfo.CellHeight();
        const auto str = cellHeight.c_str();

        auto& errnoRef = errno; // Nonzero cost, pay it once.
        errnoRef = 0;

        wchar_t* end;
        const auto value = std::wcstod(str, &end);

        return str == end || errnoRef == ERANGE ? NAN : value;
    }

    void AppearanceViewModel::LineHeight(const double value)
    {
        std::wstring str;

        if (value >= 0.1 && value <= 10.0)
        {
            str = fmt::format(FMT_STRING(L"{:.6g}"), value);
        }

        const auto fontInfo = _appearance.SourceProfile().FontInfo();

        if (fontInfo.CellHeight() != str)
        {
            if (str.empty())
            {
                fontInfo.ClearCellHeight();
            }
            else
            {
                fontInfo.CellHeight(str);
            }
            _NotifyChanges(L"HasLineHeight", L"LineHeight");
        }
    }

    bool AppearanceViewModel::HasLineHeight() const
    {
        const auto fontInfo = _appearance.SourceProfile().FontInfo();
        return fontInfo.HasCellHeight();
    }

    void AppearanceViewModel::ClearLineHeight()
    {
        LineHeight(NAN);
    }

    Model::FontConfig AppearanceViewModel::LineHeightOverrideSource() const
    {
        const auto fontInfo = _appearance.SourceProfile().FontInfo();
        return fontInfo.CellHeightOverrideSource();
    }

    void AppearanceViewModel::SetFontWeightFromDouble(double fontWeight)
    {
        FontWeight(winrt::Microsoft::Terminal::UI::Converters::DoubleToFontWeight(fontWeight));
    }

    void AppearanceViewModel::SetBackgroundImageOpacityFromPercentageValue(double percentageValue)
    {
        BackgroundImageOpacity(winrt::Microsoft::Terminal::UI::Converters::PercentageValueToPercentage(percentageValue));
    }

    void AppearanceViewModel::SetBackgroundImagePath(winrt::hstring path)
    {
        BackgroundImagePath(path);
    }

    bool AppearanceViewModel::UseDesktopBGImage()
    {
        return BackgroundImagePath() == L"desktopWallpaper";
    }

    void AppearanceViewModel::UseDesktopBGImage(const bool useDesktop)
    {
        if (useDesktop)
        {
            // Stash the current value of BackgroundImagePath. If the user
            // checks and un-checks the "Use desktop wallpaper" button, we want
            // the path that we display in the text box to remain unchanged.
            //
            // Only stash this value if it's not the special "desktopWallpaper"
            // value.
            if (BackgroundImagePath() != L"desktopWallpaper")
            {
                _lastBgImagePath = BackgroundImagePath();
            }
            BackgroundImagePath(L"desktopWallpaper");
        }
        else
        {
            // Restore the path we had previously cached. This might be the
            // empty string.
            BackgroundImagePath(_lastBgImagePath);
        }
    }

    bool AppearanceViewModel::BackgroundImageSettingsVisible()
    {
        return BackgroundImagePath() != L"";
    }

    void AppearanceViewModel::ClearColorScheme()
    {
        ClearDarkColorSchemeName();
        _NotifyChanges(L"CurrentColorScheme");
    }

    Editor::ColorSchemeViewModel AppearanceViewModel::CurrentColorScheme()
    {
        const auto schemeName{ DarkColorSchemeName() };
        const auto allSchemes{ SchemesList() };
        for (const auto& scheme : allSchemes)
        {
            if (scheme.Name() == schemeName)
            {
                return scheme;
            }
        }
        // This Appearance points to a color scheme that was renamed or deleted.
        // Fallback to the first one in the list.
        return allSchemes.GetAt(0);
    }

    void AppearanceViewModel::CurrentColorScheme(const ColorSchemeViewModel& val)
    {
        DarkColorSchemeName(val.Name());
        LightColorSchemeName(val.Name());
    }

    void AppearanceViewModel::AddNewAxisKeyValuePair()
    {
    }

    void AppearanceViewModel::DeleteAxisKeyValuePair(winrt::hstring)
    {
    }

    // Method Description:
    // - Determines whether the currently selected font has any variable font axes
    bool AppearanceViewModel::AreFontAxesAvailable()
    {
        return true;
    }

    // Method Description:
    // - Determines whether the currently selected font has any variable font axes that have not already been set
    bool AppearanceViewModel::CanFontAxesBeAdded()
    {
        return true;
    }

    void AppearanceViewModel::AddNewFeatureKeyValuePair()
    {
    }

    void AppearanceViewModel::DeleteFeatureKeyValuePair(hstring)
    {
    }

    // Method Description:
    // - Determines whether the currently selected font has any font features
    bool AppearanceViewModel::AreFontFeaturesAvailable()
    {
        return true;
    }

    // Method Description:
    // - Determines whether the currently selected font has any font features that have not already been set
    bool AppearanceViewModel::CanFontFeaturesBeAdded()
    {
        return true;
    }

    DependencyProperty Appearances::_AppearanceProperty{ nullptr };

    winrt::hstring Appearances::FontAxisName(const winrt::hstring& key)
    {
        return key;
    }

    winrt::hstring Appearances::FontFeatureName(const winrt::hstring& key)
    {
        return key;
    }

    Appearances::Appearances()
    {
        InitializeComponent();

        {
            using namespace winrt::Windows::Globalization::NumberFormatting;
            // > .NET rounds to 12 significant digits when displaying doubles, so we will [...]
            // ...obviously not do that, because this is an UI element for humans. This prevents
            // issues when displaying 32-bit floats, because WinUI is unaware about their existence.
            IncrementNumberRounder rounder;
            rounder.Increment(1e-6);

            for (const auto& box : { _fontSizeBox(), _lineHeightBox() })
            {
                // BODGY: Depends on WinUI internals.
                box.NumberFormatter().as<DecimalFormatter>().NumberRounder(rounder);
            }
        }

        INITIALIZE_BINDABLE_ENUM_SETTING(CursorShape, CursorStyle, winrt::Microsoft::Terminal::Core::CursorStyle, L"Profile_CursorShape", L"Content");
        INITIALIZE_BINDABLE_ENUM_SETTING(AdjustIndistinguishableColors, AdjustIndistinguishableColors, winrt::Microsoft::Terminal::Core::AdjustTextMode, L"Profile_AdjustIndistinguishableColors", L"Content");
        INITIALIZE_BINDABLE_ENUM_SETTING_REVERSE_ORDER(BackgroundImageStretchMode, BackgroundImageStretchMode, winrt::Windows::UI::Xaml::Media::Stretch, L"Profile_BackgroundImageStretchMode", L"Content");

        // manually add Custom FontWeight option. Don't add it to the Map
        INITIALIZE_BINDABLE_ENUM_SETTING(FontWeight, FontWeight, uint16_t, L"Profile_FontWeight", L"Content");
        _CustomFontWeight = winrt::make<EnumEntry>(RS_(L"Profile_FontWeightCustom/Content"), winrt::box_value<uint16_t>(0u));
        _FontWeightList.Append(_CustomFontWeight);

        if (!_AppearanceProperty)
        {
            _AppearanceProperty =
                DependencyProperty::Register(
                    L"Appearance",
                    xaml_typename<Editor::AppearanceViewModel>(),
                    xaml_typename<Editor::Appearances>(),
                    PropertyMetadata{ nullptr, PropertyChangedCallback{ &Appearances::_ViewModelChanged } });
        }

        // manually keep track of all the Background Image Alignment buttons
        _BIAlignmentButtons.at(0) = BIAlign_TopLeft();
        _BIAlignmentButtons.at(1) = BIAlign_Top();
        _BIAlignmentButtons.at(2) = BIAlign_TopRight();
        _BIAlignmentButtons.at(3) = BIAlign_Left();
        _BIAlignmentButtons.at(4) = BIAlign_Center();
        _BIAlignmentButtons.at(5) = BIAlign_Right();
        _BIAlignmentButtons.at(6) = BIAlign_BottomLeft();
        _BIAlignmentButtons.at(7) = BIAlign_Bottom();
        _BIAlignmentButtons.at(8) = BIAlign_BottomRight();

        // apply automation properties to more complex setting controls
        for (const auto& biButton : _BIAlignmentButtons)
        {
            const auto tooltip{ ToolTipService::GetToolTip(biButton) };
            Automation::AutomationProperties::SetName(biButton, unbox_value<hstring>(tooltip));
        }

        const auto showAllFontsCheckboxTooltip{ ToolTipService::GetToolTip(ShowAllFontsCheckbox()) };
        Automation::AutomationProperties::SetFullDescription(ShowAllFontsCheckbox(), unbox_value<hstring>(showAllFontsCheckboxTooltip));

        const auto backgroundImgCheckboxTooltip{ ToolTipService::GetToolTip(UseDesktopImageCheckBox()) };
        Automation::AutomationProperties::SetFullDescription(UseDesktopImageCheckBox(), unbox_value<hstring>(backgroundImgCheckboxTooltip));

        INITIALIZE_BINDABLE_ENUM_SETTING(IntenseTextStyle, IntenseTextStyle, winrt::Microsoft::Terminal::Settings::Model::IntenseStyle, L"Appearance_IntenseTextStyle", L"Content");
    }

    IObservableVector<Editor::Font> Appearances::FilteredFontList()
    {
        if (!_filteredFonts)
        {
            _updateFilteredFontList();
        }
        return _filteredFonts;
    }

    // Method Description:
    // - Determines whether we should show the list of all the fonts, or we should just show monospace fonts
    bool Appearances::ShowAllFonts() const noexcept
    {
        return _ShowAllFonts;
    }

    void Appearances::ShowAllFonts(const bool value)
    {
        if (_ShowAllFonts != value)
        {
            _ShowAllFonts = value;
            _filteredFonts = nullptr;
            _PropertyChangedHandlers(*this, PropertyChangedEventArgs{ L"ShowAllFonts" });
            _PropertyChangedHandlers(*this, PropertyChangedEventArgs{ L"FilteredFontList" });
        }
    }

    void Appearances::FontFaceBox_GotFocus(const Windows::Foundation::IInspectable& sender, const RoutedEventArgs&)
    {
        _updateFontNameFilter({});
        sender.as<AutoSuggestBox>().IsSuggestionListOpen(true);
    }

    void Appearances::FontFaceBox_LostFocus(const IInspectable& sender, const RoutedEventArgs&)
    {
        const auto appearance = Appearance();
        const auto fontSpec = sender.as<AutoSuggestBox>().Text();

        if (fontSpec.empty())
        {
            appearance.ClearFontFace();
        }
        else
        {
            appearance.FontFace(fontSpec);
        }
    }

    void Appearances::FontFaceBox_SuggestionChosen(const AutoSuggestBox& sender, const AutoSuggestBoxSuggestionChosenEventArgs& args)
    {
        const auto font = unbox_value<Editor::Font>(args.SelectedItem());
        const auto fontName = font.Name();
        auto fontSpec = sender.Text();

        const std::wstring_view fontSpecView{ fontSpec };
        if (const auto idx = fontSpecView.rfind(L','); idx != std::wstring_view::npos)
        {
            const auto prefix = fontSpecView.substr(0, idx);
            const auto suffix = std::wstring_view{ fontName };
            fontSpec = winrt::hstring{ fmt::format(FMT_COMPILE(L"{}, {}"), prefix, suffix) };
        }
        else
        {
            fontSpec = fontName;
        }

        sender.Text(fontSpec);
    }

    void Appearances::FontFaceBox_TextChanged(const AutoSuggestBox& sender, const AutoSuggestBoxTextChangedEventArgs& args)
    {
        if (args.Reason() != AutoSuggestionBoxTextChangeReason::UserInput)
        {
            return;
        }

        const auto fontSpec = sender.Text();
        std::wstring_view filter{ fontSpec };

        // Find the last font name in the font, spec, list.
        if (const auto idx = filter.rfind(L','); idx != std::wstring_view::npos)
        {
            filter = filter.substr(idx + 1);
        }

        filter = til::trim(filter, L' ');
        _updateFontNameFilter(filter);
    }

    void Appearances::_updateFontNameFilter(std::wstring_view filter)
    {
        if (_fontNameFilter != filter)
        {
            _filteredFonts = nullptr;
            _fontNameFilter = filter;
            _PropertyChangedHandlers(*this, PropertyChangedEventArgs{ L"FilteredFontList" });
        }
    }

    void Appearances::_updateFilteredFontList()
    {
        _filteredFonts = _ShowAllFonts ? ProfileViewModel::CompleteFontList() : ProfileViewModel::MonospaceFontList();

        if (_fontNameFilter.empty())
        {
            return;
        }

        std::vector<Editor::Font> filtered;
        filtered.reserve(_filteredFonts.Size());

        for (const auto& font : _filteredFonts)
        {
            const auto name = font.Name();
            bool match = til::contains_linguistic_insensitive(name, _fontNameFilter);

            if (!match)
            {
                const auto localizedName = font.LocalizedName();
                match = localizedName != name && til::contains_linguistic_insensitive(localizedName, _fontNameFilter);
            }

            if (match)
            {
                filtered.emplace_back(font);
            }
        }

        _filteredFonts = winrt::single_threaded_observable_vector(std::move(filtered));
    }

    void Appearances::_ViewModelChanged(const DependencyObject& d, const DependencyPropertyChangedEventArgs& /*args*/)
    {
        const auto& obj{ d.as<Editor::Appearances>() };
        get_self<Appearances>(obj)->_UpdateWithNewViewModel();
    }

    void Appearances::_UpdateWithNewViewModel()
    {
        if (const auto appearance = Appearance())
        {
            const auto& biAlignmentVal{ static_cast<int32_t>(appearance.BackgroundImageAlignment()) };
            for (const auto& biButton : _BIAlignmentButtons)
            {
                biButton.IsChecked(biButton.Tag().as<int32_t>() == biAlignmentVal);
            }

            //FontAxesContainer().HelpText(appearance.AreFontAxesAvailable() ? RS_(L"Profile_FontAxesAvailable/Text") : RS_(L"Profile_FontAxesUnavailable/Text"));
            //FontFeaturesContainer().HelpText(appearance.AreFontFeaturesAvailable() ? RS_(L"Profile_FontFeaturesAvailable/Text") : RS_(L"Profile_FontFeaturesUnavailable/Text"));

            _ViewModelChangedRevoker = appearance.PropertyChanged(winrt::auto_revoke, [=](auto&&, const PropertyChangedEventArgs& args) {
                const auto settingName{ args.PropertyName() };
                if (settingName == L"CursorShape")
                {
                    _PropertyChangedHandlers(*this, PropertyChangedEventArgs{ L"CurrentCursorShape" });
                    _PropertyChangedHandlers(*this, PropertyChangedEventArgs{ L"IsVintageCursor" });
                }
                else if (settingName == L"DarkColorSchemeName" || settingName == L"LightColorSchemeName")
                {
                    _PropertyChangedHandlers(*this, PropertyChangedEventArgs{ L"CurrentColorScheme" });
                }
                else if (settingName == L"BackgroundImageStretchMode")
                {
                    _PropertyChangedHandlers(*this, PropertyChangedEventArgs{ L"CurrentBackgroundImageStretchMode" });
                }
                else if (settingName == L"BackgroundImageAlignment")
                {
                    _UpdateBIAlignmentControl(static_cast<int32_t>(appearance.BackgroundImageAlignment()));
                }
                else if (settingName == L"FontWeight")
                {
                    _PropertyChangedHandlers(*this, PropertyChangedEventArgs{ L"CurrentFontWeight" });
                    _PropertyChangedHandlers(*this, PropertyChangedEventArgs{ L"IsCustomFontWeight" });
                }
                else if (settingName == L"IntenseTextStyle")
                {
                    _PropertyChangedHandlers(*this, PropertyChangedEventArgs{ L"CurrentIntenseTextStyle" });
                }
                else if (settingName == L"AdjustIndistinguishableColors")
                {
                    _PropertyChangedHandlers(*this, PropertyChangedEventArgs{ L"CurrentAdjustIndistinguishableColors" });
                }
                else if (settingName == L"ShowProportionalFontWarning")
                {
                    _PropertyChangedHandlers(*this, PropertyChangedEventArgs{ L"ShowProportionalFontWarning" });
                }
                // YOU THERE ADDING A NEW APPEARANCE SETTING
                // Make sure you add a block like
                //
                //   else if (settingName == L"MyNewSetting")
                //   {
                //       _PropertyChangedHandlers(*this, PropertyChangedEventArgs{ L"CurrentMyNewSetting" });
                //   }
                //
                // To make sure that changes to the AppearanceViewModel will
                // propagate back up to the actual UI (in Appearances). The
                // CurrentMyNewSetting properties are the ones that are bound in
                // XAML. If you don't do this right (or only raise a property
                // changed for "MyNewSetting"), then things like the reset
                // button won't work right.
            });

            // make sure to send all the property changed events once here
            // we do this in the case an old appearance was deleted and then a new one is created,
            // the old settings need to be updated in xaml
            _PropertyChangedHandlers(*this, PropertyChangedEventArgs{ L"CurrentCursorShape" });
            _PropertyChangedHandlers(*this, PropertyChangedEventArgs{ L"IsVintageCursor" });
            _PropertyChangedHandlers(*this, PropertyChangedEventArgs{ L"CurrentColorScheme" });
            _PropertyChangedHandlers(*this, PropertyChangedEventArgs{ L"CurrentBackgroundImageStretchMode" });
            _UpdateBIAlignmentControl(static_cast<int32_t>(appearance.BackgroundImageAlignment()));
            _PropertyChangedHandlers(*this, PropertyChangedEventArgs{ L"CurrentFontWeight" });
            _PropertyChangedHandlers(*this, PropertyChangedEventArgs{ L"IsCustomFontWeight" });
            _PropertyChangedHandlers(*this, PropertyChangedEventArgs{ L"ShowAllFonts" });
            _PropertyChangedHandlers(*this, PropertyChangedEventArgs{ L"CurrentIntenseTextStyle" });
            _PropertyChangedHandlers(*this, PropertyChangedEventArgs{ L"CurrentAdjustIndistinguishableColors" });
            _PropertyChangedHandlers(*this, PropertyChangedEventArgs{ L"ShowProportionalFontWarning" });
        }
    }

    fire_and_forget Appearances::BackgroundImage_Click(const IInspectable&, const RoutedEventArgs&)
    {
        auto lifetime = get_strong();

        const auto parentHwnd{ reinterpret_cast<HWND>(WindowRoot().GetHostingWindow()) };
        auto file = co_await OpenImagePicker(parentHwnd);
        if (!file.empty())
        {
            Appearance().BackgroundImagePath(file);
        }
    }

    void Appearances::BIAlignment_Click(const IInspectable& sender, const RoutedEventArgs& /*e*/)
    {
        if (const auto& button{ sender.try_as<Primitives::ToggleButton>() })
        {
            if (const auto& tag{ button.Tag().try_as<int32_t>() })
            {
                // Update the Appearance's value and the control
                Appearance().BackgroundImageAlignment(static_cast<ConvergedAlignment>(*tag));
                _UpdateBIAlignmentControl(*tag);
            }
        }
    }

    // Method Description:
    // - Resets all of the buttons to unchecked, and checks the one with the provided tag
    // Arguments:
    // - val - the background image alignment (ConvergedAlignment) that we want to represent in the control
    void Appearances::_UpdateBIAlignmentControl(const int32_t val)
    {
        for (const auto& biButton : _BIAlignmentButtons)
        {
            if (const auto& biButtonAlignment{ biButton.Tag().try_as<int32_t>() })
            {
                biButton.IsChecked(biButtonAlignment == val);
            }
        }
    }

    void Appearances::DeleteAxisKeyValuePair_Click(const IInspectable& sender, const RoutedEventArgs& /*e*/)
    {
        if (const auto& button{ sender.try_as<Controls::Button>() })
        {
            if (const auto& tag{ button.Tag().try_as<winrt::hstring>() })
            {
                Appearance().DeleteAxisKeyValuePair(tag.value());
            }
        }
    }

    void Appearances::AddNewAxisKeyValuePair_Click(const IInspectable& /*sender*/, const RoutedEventArgs& /*e*/)
    {
        Appearance().AddNewAxisKeyValuePair();
    }

    void Appearances::DeleteFeatureKeyValuePair_Click(const IInspectable& sender, const RoutedEventArgs& /*e*/)
    {
        if (const auto& button{ sender.try_as<Controls::Button>() })
        {
            if (const auto& tag{ button.Tag().try_as<hstring>() })
            {
                Appearance().DeleteFeatureKeyValuePair(tag.value());
            }
        }
    }

    void Appearances::AddNewFeatureKeyValuePair_Click(const IInspectable& /*sender*/, const RoutedEventArgs& /*e*/)
    {
        Appearance().AddNewFeatureKeyValuePair();
    }

    bool Appearances::IsVintageCursor() const
    {
        return Appearance().CursorShape() == Core::CursorStyle::Vintage;
    }

    IInspectable Appearances::CurrentFontWeight() const
    {
        // if no value was found, we have a custom value
        const auto maybeEnumEntry{ _FontWeightMap.TryLookup(Appearance().FontWeight().Weight) };
        return maybeEnumEntry ? maybeEnumEntry : _CustomFontWeight;
    }

    void Appearances::CurrentFontWeight(const IInspectable& enumEntry)
    {
        if (auto ee = enumEntry.try_as<Editor::EnumEntry>())
        {
            if (ee != _CustomFontWeight)
            {
                const auto weight{ winrt::unbox_value<uint16_t>(ee.EnumValue()) };
                const Windows::UI::Text::FontWeight setting{ weight };
                Appearance().FontWeight(setting);

                // Appearance does not have observable properties
                // So the TwoWay binding doesn't update on the State --> Slider direction
                FontWeightSlider().Value(weight);
            }
            _PropertyChangedHandlers(*this, PropertyChangedEventArgs{ L"IsCustomFontWeight" });
        }
    }

    bool Appearances::IsCustomFontWeight()
    {
        // Use SelectedItem instead of CurrentFontWeight.
        // CurrentFontWeight converts the Appearance's value to the appropriate enum entry,
        // whereas SelectedItem identifies which one was selected by the user.
        return FontWeightComboBox().SelectedItem() == _CustomFontWeight;
    }

}
