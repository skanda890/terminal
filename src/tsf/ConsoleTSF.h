/*++

Copyright (c) Microsoft Corporation.
Licensed under the MIT license.

Module Name:

    TfContext.h

Abstract:

    This file defines the CConsoleTSF Interface Class.

Author:

Revision History:

Notes:

--*/

#pragma once

class CCompCursorPos;
class CicCategoryMgr;
class CicDisplayAttributeMgr;
class CConversionArea;

class CConsoleTSF :
    public ITfContextOwner,
    public ITfContextOwnerCompositionSink,
    public ITfInputProcessorProfileActivationSink,
    public ITfUIElementSink,
    public ITfCleanupContextSink,
    public ITfTextEditSink
{
public:
    CConsoleTSF(HWND hwndConsole, GetSuggestionWindowPos pfnPosition, GetTextBoxAreaPos pfnTextArea);
    virtual ~CConsoleTSF();

    // IUnknown methods
    STDMETHODIMP QueryInterface(REFIID riid, void** ppvObj) noexcept override;
    ULONG STDMETHODCALLTYPE AddRef() noexcept override;
    ULONG STDMETHODCALLTYPE Release() noexcept override;

    // ITfContextOwner
    STDMETHODIMP GetACPFromPoint(const POINT*, DWORD, LONG* pCP) noexcept override;
    STDMETHODIMP GetScreenExt(RECT* pRect) noexcept override;
    STDMETHODIMP GetTextExt(LONG, LONG, RECT* pRect, BOOL* pbClipped) noexcept override;
    STDMETHODIMP GetStatus(TF_STATUS* pTfStatus) noexcept override;
    STDMETHODIMP GetWnd(HWND* phwnd) noexcept override;
    STDMETHODIMP GetAttribute(REFGUID, VARIANT*) noexcept override;

    // ITfContextOwnerCompositionSink methods
    STDMETHODIMP OnStartComposition(ITfCompositionView* pComposition, BOOL* pfOk) noexcept override;
    STDMETHODIMP OnUpdateComposition(ITfCompositionView* pComposition, ITfRange* pRangeNew) noexcept override;
    STDMETHODIMP OnEndComposition(ITfCompositionView* pComposition) noexcept override;

    // ITfInputProcessorProfileActivationSink
    STDMETHODIMP OnActivated(DWORD dwProfileType, LANGID langid, REFCLSID clsid, REFGUID catid, REFGUID guidProfile, HKL hkl, DWORD dwFlags) noexcept override;

    // ITfUIElementSink methods
    STDMETHODIMP BeginUIElement(DWORD dwUIElementId, BOOL* pbShow) noexcept override;
    STDMETHODIMP UpdateUIElement(DWORD dwUIElementId) noexcept override;
    STDMETHODIMP EndUIElement(DWORD dwUIElementId) noexcept override;

    // ITfCleanupContextSink methods
    STDMETHODIMP OnCleanupContext(TfEditCookie ecWrite, ITfContext* pic) noexcept override;

    // ITfTextEditSink methods
    STDMETHODIMP OnEndEdit(ITfContext* pInputContext, TfEditCookie ecReadOnly, ITfEditRecord* pEditRecord) noexcept override;

    CConversionArea* GetConversionArea() const;
    ITfContext* GetInputContext() const;
    HWND GetConsoleHwnd() const;
    TfClientId GetTfClientId() const;
    bool IsInComposition() const;
    bool IsPendingCompositionCleanup() const;
    void OnCompositionCleanup(bool bSucceeded);
    void SetModifyingDocFlag(bool fSet);
    void SetFocus(bool fSet) const;

    // A workaround for a MS Korean IME scenario where the IME appends a whitespace
    // composition programmatically right after completing a keyboard input composition.
    // Since post-composition clean-up is an async operation, the programmatic whitespace
    // composition gets completed before the previous composition cleanup happened,
    // and this results in a double insertion of the first composition. To avoid that, we'll
    // store the length of the last completed composition here until it's cleaned up.
    // (for simplicity, this patch doesn't provide a generic solution for all possible
    // scenarios with subsequent synchronous compositions, only for the known 'append').
    long GetCompletedRangeLength() const { return _cchCompleted; }
    void SetCompletedRangeLength(long cch) { _cchCompleted = cch; }

private:
    struct CEditSessionObjectBase : ITfEditSession
    {
        virtual ~CEditSessionObjectBase() = default;

        explicit CEditSessionObjectBase(CConsoleTSF* tsf) noexcept;

        // IUnknown methods
        STDMETHODIMP QueryInterface(REFIID riid, void** ppvObj) noexcept override;
        ULONG STDMETHODCALLTYPE AddRef() noexcept override;
        ULONG STDMETHODCALLTYPE Release() noexcept override;

        CConsoleTSF* tsf = nullptr;
        ULONG referenceCount = 0;
    };

    template<HRESULT (CConsoleTSF::*Callback)(TfEditCookie)>
    struct CEditSessionObject : CEditSessionObjectBase
    {
        using CEditSessionObjectBase::CEditSessionObjectBase;

        // ITfEditSession method
        STDMETHODIMP DoEditSession(TfEditCookie ec) noexcept override
        {
            return (tsf->*Callback)(ec);
        }
    };
    
    [[nodiscard]] HRESULT CompComplete(TfEditCookie ec);
    [[nodiscard]] HRESULT EmptyCompositionRange(TfEditCookie ec);
    [[nodiscard]] HRESULT UpdateCompositionString(TfEditCookie ec);

    void _cleanup() const noexcept;
    [[nodiscard]] HRESULT _requestCompositionComplete();
    [[nodiscard]] HRESULT _requestCompositionCleanup();
    static bool _HasCompositionChanged(ITfContext* context, TfEditCookie ec, ITfEditRecord* editRecord);

    static [[nodiscard]] HRESULT _GetAllTextRange(TfEditCookie ec, ITfContext* ic, ITfRange** range, LONG* lpTextLength, TF_HALTCOND* lpHaltCond);
    [[nodiscard]] HRESULT _ClearTextInRange(TfEditCookie ec, ITfRange* range);
    [[nodiscard]] HRESULT _GetTextAndAttribute(TfEditCookie ec, ITfRange* range, std::wstring& CompStr, std::vector<TfGuidAtom> CompGuid, BOOL bInWriteSession, CicCategoryMgr* pCicCatMgr, CicDisplayAttributeMgr* pCicDispAttr);
    [[nodiscard]] HRESULT _GetTextAndAttribute(TfEditCookie ec, ITfRange* range, std::wstring& CompStr, std::vector<TfGuidAtom>& CompGuid, std::wstring& ResultStr, BOOL bInWriteSession, CicCategoryMgr* pCicCatMgr, CicDisplayAttributeMgr* pCicDispAttr);
    [[nodiscard]] HRESULT _GetTextAndAttributeGapRange(TfEditCookie ec, ITfRange* gap_range, LONG result_comp, std::wstring& CompStr, std::vector<TfGuidAtom>& CompGuid, std::wstring& ResultStr);
    [[nodiscard]] HRESULT _GetTextAndAttributePropertyRange(TfEditCookie ec, ITfRange* pPropRange, BOOL fDispAttribute, LONG result_comp, BOOL bInWriteSession, TF_DISPLAYATTRIBUTE da, TfGuidAtom guidatom, std::wstring& CompStr, std::vector<TfGuidAtom>& CompGuid, std::wstring& ResultStr);
    [[nodiscard]] HRESULT _GetNoDisplayAttributeRange(TfEditCookie ec, ITfRange* range, const GUID** guids, int guid_size, ITfRange* no_display_attribute_range);
    [[nodiscard]] HRESULT _GetCursorPosition(TfEditCookie ec, CCompCursorPos& CompCursorPos);
    [[nodiscard]] HRESULT _IsInterimSelection(TfEditCookie ec, ITfRange** pInterimRange, BOOL* pfInterim);
    [[nodiscard]] HRESULT _MakeCompositionString(TfEditCookie ec, ITfRange* FullTextRange, BOOL bInWriteSession, CicCategoryMgr* pCicCatMgr, CicDisplayAttributeMgr* pCicDispAttr);
    [[nodiscard]] HRESULT _MakeInterimString(TfEditCookie ec, ITfRange* FullTextRange, ITfRange* InterimRange, LONG lTextLength, BOOL bInWriteSession, CicCategoryMgr* pCicCatMgr, CicDisplayAttributeMgr* pCicDispAttr);
    [[nodiscard]] HRESULT _CreateCategoryAndDisplayAttributeManager(CicCategoryMgr** pCicCatMgr, CicDisplayAttributeMgr** pCicDispAttr);

    ULONG _referenceCount = 1;

    // Cicero stuff.
    TfClientId _tid = 0;
    wil::com_ptr<ITfThreadMgrEx> _threadMgrEx;
    wil::com_ptr<ITfDocumentMgr> _documentMgr;
    wil::com_ptr<ITfContext> _context;
    wil::com_ptr<ITfSource> _threadMgrExSource;
    wil::com_ptr<ITfSource> _contextSource;
    wil::com_ptr<ITfSourceSingle> _contextSourceSingle;

    // Event sink cookies.
    DWORD _dwContextOwnerCookie = 0;
    DWORD _dwUIElementSinkCookie = 0;
    DWORD _dwTextEditSinkCookie = 0;
    DWORD _dwActivationSinkCookie = 0;

    // Conversion area object for the languages.
    std::unique_ptr<CConversionArea> _pConversionArea;

    // Console info.
    HWND _hwndConsole = nullptr;
    GetSuggestionWindowPos _pfnPosition = nullptr;
    GetTextBoxAreaPos _pfnTextArea = nullptr;

    CEditSessionObject<&CConsoleTSF::CompComplete> _editSessionCompositionComplete{ this };
    CEditSessionObject<&CConsoleTSF::EmptyCompositionRange> _editSessionCompositionCleanup{ this };
    CEditSessionObject<&CConsoleTSF::UpdateCompositionString> _editSessionUpdateCompositionString{ this };

    // Miscellaneous flags
    bool _fModifyingDoc = false; // Set true, when calls ITfRange::SetText
    bool _fCompositionCleanupSkipped = false;

    int _cCompositions = 0;
    long _cchCompleted = 0; // length of completed composition waiting for cleanup
};
