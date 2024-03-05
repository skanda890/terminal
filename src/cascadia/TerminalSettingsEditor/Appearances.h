/*++
Copyright (c) Microsoft Corporation
Licensed under the MIT license.

Module Name:
- Appearances

Abstract:
- The classes defined in this module are responsible for encapsulating the appearance settings
  of profiles and presenting them in the settings UI

Author(s):
- Pankaj Bhojwani - May 2021

--*/

#pragma once

#include "Font.g.h"
#include "AxisKeyValuePair.g.h"
#include "FeatureKeyValuePair.g.h"
#include "Appearances.g.h"
#include "AppearanceViewModel.g.h"
#include "Utils.h"
#include "ViewModelHelpers.h"
#include "SettingContainer.h"
#include <LibraryResources.h>

namespace winrt::Microsoft::Terminal::Settings::Editor::implementation
{
    struct Font : FontT<Font>
    {
    public:
        Font(winrt::hstring name, winrt::hstring localizedName, wil::com_ptr<IDWriteFontFamily> family) :
            _Name{ std::move(name) },
            _LocalizedName{ std::move(localizedName) },
            _family{ std::move(family) }
        {
        }

        hstring ToString() { return _LocalizedName; }
        bool HasPowerlineCharacters();
        Windows::Foundation::Collections::IMap<winrt::hstring, winrt::hstring> FontAxesTagsAndNames();
        Windows::Foundation::Collections::IMap<winrt::hstring, winrt::hstring> FontFeaturesTagsAndNames();

        WINRT_PROPERTY(hstring, Name);
        WINRT_PROPERTY(hstring, LocalizedName);

    private:
        void _generateFontAxesTagsAndNames(std::unordered_map<winrt::hstring, winrt::hstring>& tagsAndNames);
        void _generateFontFeaturesTagsAndNames(std::unordered_map<winrt::hstring, winrt::hstring>& tagsAndNames);
        winrt::hstring _tagToString(UINT32 tag);

        Windows::Foundation::Collections::IMap<winrt::hstring, winrt::hstring> _fontAxesTagsAndNames;
        Windows::Foundation::Collections::IMap<winrt::hstring, winrt::hstring> _fontFeaturesTagsAndNames;
        wil::com_ptr<IDWriteFontFamily> _family;
        std::optional<bool> _hasPowerlineCharacters;
    };

    struct AxisKeyValuePair : AxisKeyValuePairT<AxisKeyValuePair>, ViewModelHelper<AxisKeyValuePair>
    {
        AxisKeyValuePair(winrt::hstring axisKey, float axisValue, Windows::Foundation::Collections::IMap<winrt::hstring, float> available);

        Windows::Foundation::Collections::IObservableVector<winrt::hstring> Available() const
        {
            return winrt::single_threaded_observable_vector<winrt::hstring>();
        }

        IInspectable Key() const;
        void Key(const IInspectable& axisKey);

        float Value() const;
        void Value(float axisValue);

        WINRT_CALLBACK(PropertyChanged, Windows::UI::Xaml::Data::PropertyChangedEventHandler);

    private:
        winrt::hstring _key;
        float _value;
        Windows::Foundation::Collections::IMap<winrt::hstring, float> _baseMap{ nullptr };
    };

    struct FeatureKeyValuePair : FeatureKeyValuePairT<FeatureKeyValuePair>, ViewModelHelper<FeatureKeyValuePair>
    {
        FeatureKeyValuePair(winrt::hstring featureKey, uint32_t featureValue, Windows::Foundation::Collections::IMap<winrt::hstring, uint32_t> available);
        
        Windows::Foundation::Collections::IObservableVector<winrt::hstring> Available() const
        {
            return winrt::single_threaded_observable_vector<winrt::hstring>();
        }

        IInspectable Key();
        void Key(const IInspectable& featureKey);

        uint32_t Value();
        void Value(uint32_t featureValue);

        WINRT_CALLBACK(PropertyChanged, Windows::UI::Xaml::Data::PropertyChangedEventHandler);

    private:
        winrt::hstring _key;
        uint32_t _value;
        Windows::Foundation::Collections::IMap<winrt::hstring, uint32_t> _baseMap{ nullptr };
    };

    struct AppearanceViewModel : AppearanceViewModelT<AppearanceViewModel>, ViewModelHelper<AppearanceViewModel>
    {
    public:
        AppearanceViewModel(const Model::AppearanceConfig& appearance);

        winrt::hstring FontFace() const;
        void FontFace(const winrt::hstring& value);
        bool HasFontFace() const;
        void ClearFontFace();
        Model::FontConfig FontFaceOverrideSource() const;

        double LineHeight() const;
        void LineHeight(const double value);
        bool HasLineHeight() const;
        void ClearLineHeight();
        Model::FontConfig LineHeightOverrideSource() const;

        void SetFontWeightFromDouble(double fontWeight);
        void SetBackgroundImageOpacityFromPercentageValue(double percentageValue);
        void SetBackgroundImagePath(winrt::hstring path);

        // background image
        bool UseDesktopBGImage();
        void UseDesktopBGImage(const bool useDesktop);
        bool BackgroundImageSettingsVisible();

        void ClearColorScheme();
        Editor::ColorSchemeViewModel CurrentColorScheme();
        void CurrentColorScheme(const Editor::ColorSchemeViewModel& val);

        void AddNewAxisKeyValuePair();
        void DeleteAxisKeyValuePair(winrt::hstring key);
        bool AreFontAxesAvailable();
        bool CanFontAxesBeAdded();

        void AddNewFeatureKeyValuePair();
        void DeleteFeatureKeyValuePair(winrt::hstring key);
        bool AreFontFeaturesAvailable();
        bool CanFontFeaturesBeAdded();

        WINRT_PROPERTY(bool, IsDefault, false);

        // These settings are not defined in AppearanceConfig, so we grab them
        // from the source profile itself. The reason we still want them in the
        // AppearanceViewModel is so we can continue to have the 'Text' grouping
        // we currently have in xaml, since that grouping has some settings that
        // are defined in AppearanceConfig and some that are not.
        OBSERVABLE_PROJECTED_SETTING(_appearance.SourceProfile().FontInfo(), FontSize);
        OBSERVABLE_PROJECTED_SETTING(_appearance.SourceProfile().FontInfo(), FontWeight);
        OBSERVABLE_PROJECTED_SETTING(_appearance.SourceProfile().FontInfo(), FontAxes);
        OBSERVABLE_PROJECTED_SETTING(_appearance.SourceProfile().FontInfo(), FontFeatures);
        OBSERVABLE_PROJECTED_SETTING(_appearance.SourceProfile().FontInfo(), EnableBuiltinGlyphs);

        OBSERVABLE_PROJECTED_SETTING(_appearance, RetroTerminalEffect);
        OBSERVABLE_PROJECTED_SETTING(_appearance, CursorShape);
        OBSERVABLE_PROJECTED_SETTING(_appearance, CursorHeight);
        OBSERVABLE_PROJECTED_SETTING(_appearance, DarkColorSchemeName);
        OBSERVABLE_PROJECTED_SETTING(_appearance, LightColorSchemeName);
        OBSERVABLE_PROJECTED_SETTING(_appearance, BackgroundImagePath);
        OBSERVABLE_PROJECTED_SETTING(_appearance, BackgroundImageOpacity);
        OBSERVABLE_PROJECTED_SETTING(_appearance, BackgroundImageStretchMode);
        OBSERVABLE_PROJECTED_SETTING(_appearance, BackgroundImageAlignment);
        OBSERVABLE_PROJECTED_SETTING(_appearance, IntenseTextStyle);
        OBSERVABLE_PROJECTED_SETTING(_appearance, AdjustIndistinguishableColors);
        WINRT_OBSERVABLE_PROPERTY(Windows::Foundation::Collections::IObservableVector<Editor::ColorSchemeViewModel>, SchemesList, _propertyChangedHandlers, nullptr);
        WINRT_OBSERVABLE_PROPERTY(winrt::hstring, MissingFontWarningMessage, _propertyChangedHandlers);
        WINRT_OBSERVABLE_PROPERTY(winrt::hstring, ProportionalFontWarningMessage, _propertyChangedHandlers);

        WINRT_OBSERVABLE_PROPERTY(Windows::Foundation::Collections::IObservableVector<Editor::AxisKeyValuePair>, FontAxesVector, _propertyChangedHandlers, nullptr);
        WINRT_OBSERVABLE_PROPERTY(Windows::Foundation::Collections::IObservableVector<winrt::hstring>, SelectedFontAxes, _propertyChangedHandlers, winrt::single_threaded_observable_vector<winrt::hstring>());
        WINRT_OBSERVABLE_PROPERTY(Windows::Foundation::Collections::IObservableVector<winrt::hstring>, RemainingFontAxes, _propertyChangedHandlers, winrt::single_threaded_observable_vector<winrt::hstring>());

        WINRT_OBSERVABLE_PROPERTY(Windows::Foundation::Collections::IObservableVector<winrt::hstring>, SelectedFontFeatures, _propertyChangedHandlers, winrt::single_threaded_observable_vector<winrt::hstring>());
        WINRT_OBSERVABLE_PROPERTY(Windows::Foundation::Collections::IObservableVector<winrt::hstring>, RemainingFontFeatures, _propertyChangedHandlers, winrt::single_threaded_observable_vector<winrt::hstring>());
        WINRT_OBSERVABLE_PROPERTY(Windows::Foundation::Collections::IObservableVector<Editor::FeatureKeyValuePair>, FontFeaturesVector, _propertyChangedHandlers, nullptr);

    private:
        void _refreshFontFaceDependents();

        Model::AppearanceConfig _appearance;
        winrt::hstring _lastBgImagePath;
    };

    struct Appearances : AppearancesT<Appearances>
    {
    public:
        static winrt::hstring FontAxisName(const winrt::hstring& key);
        static winrt::hstring FontFeatureName(const winrt::hstring& key);

        Appearances();

        // CursorShape visibility logic
        bool IsVintageCursor() const;

        Windows::Foundation::Collections::IObservableVector<Editor::Font> FilteredFontList();
        bool ShowAllFonts() const noexcept;
        void ShowAllFonts(bool value);

        void FontFaceBox_GotFocus(const Windows::Foundation::IInspectable& sender, const Windows::UI::Xaml::RoutedEventArgs& e);
        void FontFaceBox_LostFocus(const Windows::Foundation::IInspectable& sender, const Windows::UI::Xaml::RoutedEventArgs& e);
        void FontFaceBox_SuggestionChosen(const winrt::Windows::UI::Xaml::Controls::AutoSuggestBox&, const winrt::Windows::UI::Xaml::Controls::AutoSuggestBoxSuggestionChosenEventArgs&);
        void FontFaceBox_TextChanged(const winrt::Windows::UI::Xaml::Controls::AutoSuggestBox&, const winrt::Windows::UI::Xaml::Controls::AutoSuggestBoxTextChangedEventArgs&);
        fire_and_forget BackgroundImage_Click(const Windows::Foundation::IInspectable& sender, const Windows::UI::Xaml::RoutedEventArgs& e);
        void BIAlignment_Click(const Windows::Foundation::IInspectable& sender, const Windows::UI::Xaml::RoutedEventArgs& e);
        void DeleteAxisKeyValuePair_Click(const Windows::Foundation::IInspectable& sender, const Windows::UI::Xaml::RoutedEventArgs& e);
        void AddNewAxisKeyValuePair_Click(const Windows::Foundation::IInspectable& sender, const Windows::UI::Xaml::RoutedEventArgs& e);
        void DeleteFeatureKeyValuePair_Click(const Windows::Foundation::IInspectable& sender, const Windows::UI::Xaml::RoutedEventArgs& e);
        void AddNewFeatureKeyValuePair_Click(const Windows::Foundation::IInspectable& sender, const Windows::UI::Xaml::RoutedEventArgs& e);

        // manually bind FontWeight
        Windows::Foundation::IInspectable CurrentFontWeight() const;
        void CurrentFontWeight(const Windows::Foundation::IInspectable& enumEntry);
        bool IsCustomFontWeight();
        WINRT_PROPERTY(Windows::Foundation::Collections::IObservableVector<Microsoft::Terminal::Settings::Editor::EnumEntry>, FontWeightList);

        GETSET_BINDABLE_ENUM_SETTING(CursorShape, Microsoft::Terminal::Core::CursorStyle, Appearance().CursorShape);
        GETSET_BINDABLE_ENUM_SETTING(AdjustIndistinguishableColors, Microsoft::Terminal::Core::AdjustTextMode, Appearance().AdjustIndistinguishableColors);

        WINRT_CALLBACK(PropertyChanged, Windows::UI::Xaml::Data::PropertyChangedEventHandler);
        DEPENDENCY_PROPERTY(Editor::AppearanceViewModel, Appearance);
        WINRT_PROPERTY(Editor::ProfileViewModel, SourceProfile, nullptr);
        WINRT_PROPERTY(IHostedInWindow, WindowRoot, nullptr);
        GETSET_BINDABLE_ENUM_SETTING(BackgroundImageStretchMode, Windows::UI::Xaml::Media::Stretch, Appearance().BackgroundImageStretchMode);

        GETSET_BINDABLE_ENUM_SETTING(IntenseTextStyle, Microsoft::Terminal::Settings::Model::IntenseStyle, Appearance().IntenseTextStyle);

    private:
        Windows::UI::Xaml::Data::INotifyPropertyChanged::PropertyChanged_revoker _ViewModelChangedRevoker;
        std::array<Windows::UI::Xaml::Controls::Primitives::ToggleButton, 9> _BIAlignmentButtons;
        Windows::Foundation::Collections::IMap<uint16_t, Microsoft::Terminal::Settings::Editor::EnumEntry> _FontWeightMap;
        Editor::EnumEntry _CustomFontWeight{ nullptr };
        Windows::Foundation::Collections::IObservableVector<Editor::Font> _filteredFonts;
        std::wstring _fontNameFilter;
        bool _ShowAllFonts = false;

        static void _ViewModelChanged(const Windows::UI::Xaml::DependencyObject& d, const Windows::UI::Xaml::DependencyPropertyChangedEventArgs& e);

        void _updateFontNameFilter(std::wstring_view filter);
        void _updateFilteredFontList();
        void _UpdateBIAlignmentControl(const int32_t val);
        void _UpdateWithNewViewModel();
    };
};

namespace winrt::Microsoft::Terminal::Settings::Editor::factory_implementation
{
    BASIC_FACTORY(Appearances);
}
