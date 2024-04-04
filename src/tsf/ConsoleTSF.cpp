// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

#include "precomp.h"
#include "ConsoleTSF.h"

#include "TfConvArea.h"
#include "TfConvArea.h"
#include "TfDispAttr.h"
#include "TfCtxtComp.h"

/* 626761ad-78d2-44d2-be8b-752cf122acec */
static constexpr GUID GUID_APPLICATION = { 0x626761ad, 0x78d2, 0x44d2, { 0xbe, 0x8b, 0x75, 0x2c, 0xf1, 0x22, 0xac, 0xec } };
/* 183C627A-B46C-44ad-B797-82F6BEC82131 */
static constexpr GUID GUID_PROP_CONIME_TRACKCOMPOSITION = { 0x183c627a, 0xb46c, 0x44ad, { 0xb7, 0x97, 0x82, 0xf6, 0xbe, 0xc8, 0x21, 0x31 } };

class CicCategoryMgr
{
public:
    [[nodiscard]] HRESULT GetGUIDFromGUIDATOM(TfGuidAtom guidatom, GUID* pguid);
    [[nodiscard]] HRESULT InitCategoryInstance();

    ITfCategoryMgr* GetCategoryMgr();

private:
    wil::com_ptr_nothrow<ITfCategoryMgr> m_pcat;
};

[[nodiscard]] HRESULT CicCategoryMgr::GetGUIDFromGUIDATOM(TfGuidAtom guidatom, GUID* pguid)
{
    return m_pcat->GetGUID(guidatom, pguid);
}

[[nodiscard]] HRESULT CicCategoryMgr::InitCategoryInstance()
{
    //
    // Create ITfCategoryMgr instance.
    //
    return ::CoCreateInstance(CLSID_TF_CategoryMgr, nullptr, CLSCTX_ALL, IID_PPV_ARGS(&m_pcat));
}

ITfCategoryMgr* CicCategoryMgr::GetCategoryMgr()
{
    return m_pcat.get();
}

CConsoleTSF::CConsoleTSF(HWND hwndConsole, GetSuggestionWindowPos pfnPosition, GetTextBoxAreaPos pfnTextArea) :
    _hwndConsole(hwndConsole),
    _pfnPosition(pfnPosition),
    _pfnTextArea(pfnTextArea)
{
    try
    {
        // There's no point in calling TF_GetThreadMgr. ITfThreadMgr is a per-thread singleton.
        _threadMgrEx = wil::CoCreateInstance<ITfThreadMgrEx>(CLSID_TF_ThreadMgr, CLSCTX_INPROC_SERVER);

        THROW_IF_FAILED(_threadMgrEx->ActivateEx(&_tid, TF_TMAE_CONSOLE));
        THROW_IF_FAILED(_threadMgrEx->CreateDocumentMgr(_documentMgr.addressof()));

        TfEditCookie ecTmp;
        THROW_IF_FAILED(_documentMgr->CreateContext(_tid, 0, static_cast<ITfContextOwnerCompositionSink*>(this), _context.addressof(), &ecTmp));

        _threadMgrExSource = _threadMgrEx.query<ITfSource>();
        THROW_IF_FAILED(_threadMgrExSource->AdviseSink(IID_ITfInputProcessorProfileActivationSink, static_cast<ITfInputProcessorProfileActivationSink*>(this), &_dwActivationSinkCookie));
        THROW_IF_FAILED(_threadMgrExSource->AdviseSink(IID_ITfUIElementSink, static_cast<ITfUIElementSink*>(this), &_dwUIElementSinkCookie));

        _contextSource = _context.query<ITfSource>();
        THROW_IF_FAILED(_contextSource->AdviseSink(IID_ITfContextOwner, static_cast<ITfContextOwner*>(this), &_dwContextOwnerCookie));
        THROW_IF_FAILED(_contextSource->AdviseSink(IID_ITfTextEditSink, static_cast<ITfTextEditSink*>(this), &_dwTextEditSinkCookie));

        _contextSourceSingle = _context.query<ITfSourceSingle>();
        THROW_IF_FAILED(_contextSourceSingle->AdviseSingleSink(_tid, IID_ITfCleanupContextSink, static_cast<ITfCleanupContextSink*>(this)));

        THROW_IF_FAILED(_documentMgr->Push(_context.get()));

        // Collect the active keyboard layout info.
        if (const auto spITfProfilesMgr = wil::CoCreateInstanceNoThrow<ITfInputProcessorProfileMgr>(CLSID_TF_InputProcessorProfiles, CLSCTX_INPROC_SERVER))
        {
            TF_INPUTPROCESSORPROFILE ipp;
            if (SUCCEEDED(spITfProfilesMgr->GetActiveProfile(GUID_TFCAT_TIP_KEYBOARD, &ipp)))
            {
                std::ignore = CConsoleTSF::OnActivated(ipp.dwProfileType, ipp.langid, ipp.clsid, ipp.catid, ipp.guidProfile, ipp.hkl, ipp.dwFlags);
            }
        }
    }
    catch (...)
    {
        _cleanup();
        throw;
    }
}

CConsoleTSF::~CConsoleTSF()
{
    _cleanup();
}

void CConsoleTSF::_cleanup() const noexcept
{
    if (_contextSourceSingle)
    {
        std::ignore = _contextSourceSingle->UnadviseSingleSink(_tid, IID_ITfCleanupContextSink);
    }
    if (_contextSource)
    {
        std::ignore = _contextSource->UnadviseSink(_dwTextEditSinkCookie);
        std::ignore = _contextSource->UnadviseSink(_dwContextOwnerCookie);
    }
    if (_threadMgrExSource)
    {
        std::ignore = _threadMgrExSource->UnadviseSink(_dwUIElementSinkCookie);
        std::ignore = _threadMgrExSource->UnadviseSink(_dwActivationSinkCookie);
    }

    // Clear the Cicero reference to our document manager.
    if (_threadMgrEx && _documentMgr)
    {
        wil::com_ptr<ITfDocumentMgr> spPrevDocMgr;
        std::ignore = _threadMgrEx->AssociateFocus(_hwndConsole, nullptr, spPrevDocMgr.addressof());
    }

    // Dismiss the input context and document manager.
    if (_documentMgr)
    {
        std::ignore = _documentMgr->Pop(TF_POPF_ALL);
    }

    // Deactivate per-thread Cicero.
    if (_threadMgrEx)
    {
        std::ignore = _threadMgrEx->Deactivate();
    }
}

STDMETHODIMP CConsoleTSF::QueryInterface(REFIID riid, void** ppvObj) noexcept
{
    if (!ppvObj)
    {
        return E_POINTER;
    }

    if (IsEqualGUID(riid, IID_ITfCleanupContextSink))
    {
        *ppvObj = static_cast<ITfCleanupContextSink*>(this);
    }
    else if (IsEqualGUID(riid, IID_ITfContextOwnerCompositionSink))
    {
        *ppvObj = static_cast<ITfContextOwnerCompositionSink*>(this);
    }
    else if (IsEqualGUID(riid, IID_ITfUIElementSink))
    {
        *ppvObj = static_cast<ITfUIElementSink*>(this);
    }
    else if (IsEqualGUID(riid, IID_ITfContextOwner))
    {
        *ppvObj = static_cast<ITfContextOwner*>(this);
    }
    else if (IsEqualGUID(riid, IID_ITfInputProcessorProfileActivationSink))
    {
        *ppvObj = static_cast<ITfInputProcessorProfileActivationSink*>(this);
    }
    else if (IsEqualGUID(riid, IID_ITfTextEditSink))
    {
        *ppvObj = static_cast<ITfTextEditSink*>(this);
    }
    else if (IsEqualGUID(riid, IID_IUnknown))
    {
        *ppvObj = static_cast<IUnknown*>(static_cast<ITfContextOwner*>(this));
    }
    else
    {
        *ppvObj = nullptr;
        return E_NOINTERFACE;
    }

    AddRef();
    return S_OK;
}

ULONG STDMETHODCALLTYPE CConsoleTSF::AddRef() noexcept
{
    return InterlockedIncrement(&_referenceCount);
}

ULONG STDMETHODCALLTYPE CConsoleTSF::Release() noexcept
{
    const auto cr = InterlockedDecrement(&_referenceCount);
    if (cr == 0)
    {
        delete this;
    }
    return cr;
}

STDMETHODIMP CConsoleTSF::GetACPFromPoint(const POINT*, DWORD, LONG* pCP) noexcept
{
    if (pCP)
    {
        *pCP = 0;
    }

    return S_OK;
}

// This returns Rectangle of the text box of whole console.
// When a user taps inside the rectangle while hardware keyboard is not available,
// touch keyboard is invoked.
STDMETHODIMP CConsoleTSF::GetScreenExt(RECT* pRect) noexcept
try
{
    if (pRect)
    {
        *pRect = _pfnTextArea();
    }

    return S_OK;
}
CATCH_RETURN();

// This returns rectangle of current command line edit area.
// When a user types in East Asian language, candidate window is shown at this position.
// Emoji and more panel (Win+.) is shown at the position, too.
STDMETHODIMP CConsoleTSF::GetTextExt(LONG, LONG, RECT* pRect, BOOL* pbClipped) noexcept
try
{
    if (pRect)
    {
        *pRect = _pfnPosition();
    }

    if (pbClipped)
    {
        *pbClipped = FALSE;
    }

    return S_OK;
}
CATCH_RETURN();

STDMETHODIMP CConsoleTSF::GetStatus(TF_STATUS* pTfStatus) noexcept
{
    if (pTfStatus)
    {
        pTfStatus->dwDynamicFlags = 0;
        pTfStatus->dwStaticFlags = TF_SS_TRANSITORY;
    }
    return pTfStatus ? S_OK : E_INVALIDARG;
}

STDMETHODIMP CConsoleTSF::GetWnd(HWND* phwnd) noexcept
{
    *phwnd = _hwndConsole;
    return S_OK;
}

STDMETHODIMP CConsoleTSF::GetAttribute(const GUID&, VARIANT*) noexcept
{
    return E_NOTIMPL;
}

STDMETHODIMP CConsoleTSF::OnStartComposition(ITfCompositionView* pCompView, BOOL* pfOk) noexcept
try
{
    if (!_pConversionArea || (_cCompositions > 0 && !_fModifyingDoc))
    {
        *pfOk = FALSE;
        return S_OK;
    }

    // Ignore compositions triggered by our own edit sessions
    // (i.e. when the application is the composition owner)
    auto clsidCompositionOwner = GUID_APPLICATION;
    pCompView->GetOwnerClsid(&clsidCompositionOwner);
    if (IsEqualGUID(clsidCompositionOwner, GUID_APPLICATION))
    {
        *pfOk = TRUE;
        return S_OK;
    }

    _cCompositions++;
    if (_cCompositions != 1)
    {
        *pfOk = TRUE;
        return S_OK;
    }

    LOG_IF_FAILED(ImeStartComposition());
    *pfOk = TRUE;
    return S_OK;
}
CATCH_RETURN();

STDMETHODIMP CConsoleTSF::OnUpdateComposition(ITfCompositionView* /*pComp*/, ITfRange*) noexcept
{
    return S_OK;
}

STDMETHODIMP CConsoleTSF::OnEndComposition(ITfCompositionView* pCompView) noexcept
try
{
    if (_cCompositions <= 0 || !_pConversionArea)
    {
        return E_FAIL;
    }
    // Ignore compositions triggered by our own edit sessions
    // (i.e. when the application is the composition owner)
    auto clsidCompositionOwner = GUID_APPLICATION;
    pCompView->GetOwnerClsid(&clsidCompositionOwner);
    if (IsEqualGUID(clsidCompositionOwner, GUID_APPLICATION))
    {
        return S_OK;
    }

    _cCompositions--;
    if (_cCompositions != 0)
    {
        return S_OK;
    }

    std::ignore = _requestCompositionComplete();
    std::ignore = _requestCompositionCleanup();

    LOG_IF_FAILED(ImeEndComposition());
    return S_OK;
}
CATCH_RETURN();

STDMETHODIMP CConsoleTSF::OnActivated(DWORD /*dwProfileType*/, LANGID /*langid*/, REFCLSID /*clsid*/, REFGUID catid, REFGUID /*guidProfile*/, HKL /*hkl*/, DWORD dwFlags) noexcept
try
{
    if ((dwFlags & TF_IPSINK_FLAG_ACTIVE) == 0)
    {
        return S_OK;
    }
    if (!IsEqualGUID(catid, GUID_TFCAT_TIP_KEYBOARD))
    {
        // Don't care for non-keyboard profiles.
        return S_OK;
    }

    // Associate the document\context with the console window.
    if (!_pConversionArea)
    {
        _pConversionArea = std::make_unique<CConversionArea>();
        wil::com_ptr<ITfDocumentMgr> spPrevDocMgr;
        LOG_IF_FAILED(_threadMgrEx->AssociateFocus(_hwndConsole, _pConversionArea ? _documentMgr.get() : nullptr, spPrevDocMgr.addressof()));
    }

    return S_OK;
}
CATCH_RETURN();

STDMETHODIMP CConsoleTSF::BeginUIElement(DWORD /*dwUIElementId*/, BOOL* pbShow) noexcept
{
    *pbShow = TRUE;
    return S_OK;
}

STDMETHODIMP CConsoleTSF::UpdateUIElement(DWORD /*dwUIElementId*/) noexcept
{
    return S_OK;
}

STDMETHODIMP CConsoleTSF::EndUIElement(DWORD /*dwUIElementId*/) noexcept
{
    return S_OK;
}

STDMETHODIMP CConsoleTSF::OnCleanupContext(TfEditCookie ecWrite, ITfContext* pic) noexcept
{
    //
    // Remove GUID_PROP_COMPOSING
    //
    wil::com_ptr<ITfProperty> prop;
    if (SUCCEEDED(pic->GetProperty(GUID_PROP_COMPOSING, prop.addressof())))
    {
        wil::com_ptr<IEnumTfRanges> enumranges;
        if (SUCCEEDED(prop->EnumRanges(ecWrite, enumranges.addressof(), nullptr)))
        {
            wil::com_ptr<ITfRange> rangeTmp;
            while (enumranges->Next(1, rangeTmp.addressof(), nullptr) == S_OK)
            {
                VARIANT var;
                VariantInit(&var);
                prop->GetValue(ecWrite, rangeTmp.get(), &var);
                if ((var.vt == VT_I4) && (var.lVal != 0))
                {
                    prop->Clear(ecWrite, rangeTmp.get());
                }
            }
        }
    }
    return S_OK;
}

STDMETHODIMP CConsoleTSF::OnEndEdit(ITfContext* pInputContext, TfEditCookie ecReadOnly, ITfEditRecord* pEditRecord) noexcept
try
{
    if (!_cCompositions || !_pConversionArea || !_HasCompositionChanged(pInputContext, ecReadOnly, pEditRecord))
    {
        return S_OK;
    }

    // OnEndEdit() occurs asynchronously with other events and in the meantime the composition may have restarted already.
    if (_editSessionUpdateCompositionString.referenceCount)
    {
        return S_FALSE;
    }

    HRESULT hr = S_OK;
    RETURN_IF_FAILED(_context->RequestEditSession(_tid, &_editSessionUpdateCompositionString, TF_ES_READWRITE | TF_ES_ASYNCDONTCARE, &hr));
    RETURN_IF_FAILED(hr);

    return S_OK;
}
CATCH_RETURN()

CConversionArea* CConsoleTSF::GetConversionArea() const
{
    return _pConversionArea.get();
}

ITfContext* CConsoleTSF::GetInputContext() const
{
    return _context.get();
}

HWND CConsoleTSF::GetConsoleHwnd() const
{
    return _hwndConsole;
}

TfClientId CConsoleTSF::GetTfClientId() const
{
    return _tid;
}

bool CConsoleTSF::IsInComposition() const
{
    return (_cCompositions > 0);
}

bool CConsoleTSF::IsPendingCompositionCleanup() const
{
    return _editSessionCompositionCleanup.referenceCount || _fCompositionCleanupSkipped;
}

void CConsoleTSF::OnCompositionCleanup(bool bSucceeded)
{
    _fCompositionCleanupSkipped = !bSucceeded;
}

void CConsoleTSF::SetModifyingDocFlag(bool fSet)
{
    _fModifyingDoc = fSet;
}

void CConsoleTSF::SetFocus(bool fSet) const
{
    if (!fSet && _cCompositions)
    {
        // Close (terminate) any open compositions when losing the input focus.
        if (_context)
        {
            if (const auto spCompositionServices = _context.try_query<ITfContextOwnerCompositionServices>())
            {
                spCompositionServices->TerminateComposition(nullptr);
            }
        }
    }
}

CConsoleTSF::CEditSessionObjectBase::CEditSessionObjectBase(CConsoleTSF* tsf) noexcept :
    tsf{ tsf }
{
}

STDMETHODIMP CConsoleTSF::CEditSessionObjectBase::QueryInterface(const IID& riid, void** ppvObj) noexcept
{
    if (!ppvObj)
    {
        return E_POINTER;
    }

    if (IsEqualGUID(riid, IID_ITfEditSession))
    {
        *ppvObj = static_cast<ITfEditSession*>(this);
    }
    else if (IsEqualGUID(riid, IID_IUnknown))
    {
        *ppvObj = static_cast<IUnknown*>(this);
    }
    else
    {
        *ppvObj = nullptr;
        return E_NOINTERFACE;
    }

    AddRef();
    return S_OK;
}

ULONG STDMETHODCALLTYPE CConsoleTSF::CEditSessionObjectBase::AddRef() noexcept
{
    return InterlockedIncrement(&referenceCount);
}

ULONG STDMETHODCALLTYPE CConsoleTSF::CEditSessionObjectBase::Release() noexcept
{
    FAIL_FAST_IF(referenceCount == 0);
    return InterlockedDecrement(&referenceCount);
}

[[nodiscard]] HRESULT CConsoleTSF::CompComplete(TfEditCookie ec)
{
    const auto pic = GetInputContext();
    RETURN_HR_IF_NULL(E_FAIL, pic);

    // Get the whole text, finalize it, and set empty string in TOM
    wil::com_ptr_nothrow<ITfRange> spRange;
    LONG cch;

    RETURN_IF_FAILED(_GetAllTextRange(ec, pic, &spRange, &cch, nullptr));

    // Check if a part of the range has already been finalized but not removed yet.
    // Adjust the range appropriately to avoid inserting the same text twice.
    auto cchCompleted = GetCompletedRangeLength();
    if ((cchCompleted > 0) &&
        (cchCompleted < cch) &&
        SUCCEEDED(spRange->ShiftStart(ec, cchCompleted, &cchCompleted, NULL)))
    {
        assert(((cchCompleted > 0) && (cchCompleted < cch)));
        cch -= cchCompleted;
    }
    else
    {
        cchCompleted = 0;
    }

    // Get conversion area service.
    const auto conv_area = GetConversionArea();
    RETURN_HR_IF_NULL(E_FAIL, conv_area);

    // If there is no string in TextStore we don't have to do anything.
    if (!cch)
    {
        // Clear composition
        LOG_IF_FAILED(conv_area->ClearComposition());
        return S_OK;
    }

    auto hr = S_OK;
    try
    {
        const auto wstr = std::make_unique<WCHAR[]>(cch + 1);

        // Get the whole text, finalize it, and erase the whole text.
        if (SUCCEEDED(spRange->GetText(ec, TF_TF_IGNOREEND, wstr.get(), (ULONG)cch, (ULONG*)&cch)))
        {
            // Make Result String.
            hr = conv_area->DrawResult({ wstr.get(), static_cast<size_t>(cch) });
        }
    }
    CATCH_RETURN();

    // Update the stored length of the completed fragment.
    SetCompletedRangeLength(cchCompleted + cch);

    return hr;
}

[[nodiscard]] HRESULT CConsoleTSF::EmptyCompositionRange(TfEditCookie ec)
{
    if (!IsPendingCompositionCleanup())
    {
        return S_OK;
    }

    auto hr = E_FAIL;
    const auto pic = GetInputContext();
    if (pic != nullptr)
    {
        // Cleanup (empty the context range) after the last composition.

        hr = S_OK;
        const auto cchCompleted = GetCompletedRangeLength();
        if (cchCompleted != 0)
        {
            wil::com_ptr_nothrow<ITfRange> spRange;
            LONG cch;
            hr = _GetAllTextRange(ec, pic, &spRange, &cch, nullptr);
            if (SUCCEEDED(hr))
            {
                // Clean up only the completed part (which start is expected to coincide with the start of the full range).
                if (cchCompleted < cch)
                {
                    spRange->ShiftEnd(ec, (cchCompleted - cch), &cch, nullptr);
                }
                hr = _ClearTextInRange(ec, spRange.get());
                SetCompletedRangeLength(0); // cleaned up all completed text
            }
        }
    }
    OnCompositionCleanup(SUCCEEDED(hr));
    return hr;
}

[[nodiscard]] HRESULT CConsoleTSF::UpdateCompositionString(TfEditCookie ec)
{
    HRESULT hr;

    const auto pic = GetInputContext();
    if (pic == nullptr)
    {
        return E_FAIL;
    }

    // If the composition has been cancelled\finalized, no update necessary.
    if (!IsInComposition())
    {
        return S_OK;
    }

    BOOL bInWriteSession;
    if (FAILED(hr = pic->InWriteSession(GetTfClientId(), &bInWriteSession)))
    {
        return hr;
    }

    wil::com_ptr_nothrow<ITfRange> FullTextRange;
    LONG lTextLength;
    if (FAILED(hr = _GetAllTextRange(ec, pic, &FullTextRange, &lTextLength, nullptr)))
    {
        return hr;
    }

    wil::com_ptr_nothrow<ITfRange> InterimRange;
    auto fInterim = FALSE;
    if (FAILED(hr = _IsInterimSelection(ec, &InterimRange, &fInterim)))
    {
        return hr;
    }

    //
    // Create Cicero Category Manager and Display Attribute Manager
    //
    CicCategoryMgr pCicCat;
    CicDisplayAttributeMgr pDispAttr;
    if (SUCCEEDED(hr = pCicCat.InitCategoryInstance()))
    {
        if (const auto pcat = pCicCat.GetCategoryMgr())
        {
            hr = pDispAttr.InitDisplayAttributeInstance(pcat);
        }
    }

    if (SUCCEEDED(hr))
    {
        if (fInterim)
        {
            hr = _MakeInterimString(ec, FullTextRange.get(), InterimRange.get(), lTextLength, bInWriteSession, &pCicCat, &pDispAttr);
        }
        else
        {
            hr = _MakeCompositionString(ec, FullTextRange.get(), bInWriteSession, &pCicCat, &pDispAttr);
        }
    }

    return hr;
}

[[nodiscard]] HRESULT CConsoleTSF::_requestCompositionComplete()
{
    // The composition could have been finalized because of a caret move, therefore it must be
    // inserted synchronously while at the original caret position.(TF_ES_SYNC is ok for a nested RO session).
    auto hr = E_OUTOFMEMORY;
    RETURN_IF_FAILED(_context->RequestEditSession(_tid, &_editSessionCompositionComplete, TF_ES_READ | TF_ES_SYNC, &hr));
    RETURN_IF_FAILED(hr);
    return S_OK;
}

[[nodiscard]] HRESULT CConsoleTSF::_requestCompositionCleanup()
{
    // Cleanup (empty the context range) after the last composition, unless a new one has started.
    if (_editSessionCompositionCleanup.referenceCount)
    {
        return S_OK;
    }

    // NOTE: This used to specify TF_ES_ASYNC explicitly in the past, because:
    //   For the same reason, must use explicit TF_ES_ASYNC, or the request will be rejected otherwise.
    // but I found that TF_ES_ASYNCDONTCARE works just fine.
    auto hr = E_OUTOFMEMORY;
    RETURN_IF_FAILED(_context->RequestEditSession(_tid, &_editSessionCompositionCleanup, TF_ES_READWRITE | TF_ES_ASYNCDONTCARE, &hr));
    RETURN_IF_FAILED(hr);
    return S_OK;
}

static wil::com_ptr<ITfRange> getTrackCompositionProperty(ITfContext* context, TfEditCookie ec)
{
    wil::com_ptr<ITfProperty> Property;
    if (FAILED(context->GetProperty(GUID_PROP_CONIME_TRACKCOMPOSITION, &Property)))
    {
        return {};
    }

    wil::com_ptr<IEnumTfRanges> ranges;
    if (FAILED(Property->EnumRanges(ec, ranges.addressof(), NULL)))
    {
        return {};
    }

    VARIANT var{ .vt = VT_EMPTY };
    wil::com_ptr<ITfRange> range;
    while (ranges->Next(1, range.put(), nullptr) == S_OK)
    {
        if (SUCCEEDED(Property->GetValue(ec, range.get(), &var)) && V_VT(&var) == VT_I4 && V_I4(&var) != 0)
        {
            return range;
        }
        VariantClear(&var);
    }

    return {};
}

bool CConsoleTSF::_HasCompositionChanged(ITfContext* context, TfEditCookie ec, ITfEditRecord* editRecord)
{
    BOOL changed;
    if (SUCCEEDED(editRecord->GetSelectionStatus(&changed)) && changed)
    {
        return TRUE;
    }

    const auto rangeTrackComposition = getTrackCompositionProperty(context, ec);
    // If there is no track composition property,
    // the composition has been changed since we put it.
    if (!rangeTrackComposition)
    {
        return TRUE;
    }

    // Get the text range that does not include read only area for reconversion.
    wil::com_ptr<ITfRange> rangeAllText;
    LONG cch;
    if (FAILED(_GetAllTextRange(ec, context, rangeAllText.addressof(), &cch, nullptr)))
    {
        return FALSE;
    }

    // If the start position of the track composition range is not the beginning of IC,
    // the composition has been changed since we put it.
    LONG lResult;
    if (FAILED(rangeTrackComposition->CompareStart(ec, rangeAllText.get(), TF_ANCHOR_START, &lResult)))
    {
        return FALSE;
    }
    if (lResult != 0)
    {
        return TRUE;
    }

    if (FAILED(rangeTrackComposition->CompareEnd(ec, rangeAllText.get(), TF_ANCHOR_END, &lResult)))
    {
        return FALSE;
    }
    if (lResult != 0)
    {
        return TRUE;
    }

    // If the start position of the track composition range is not the beginning of IC,
    // the composition has been changed since we put it.
    //
    // If we find the changes in these property, we need to update hIMC.
    const GUID* guids[] = { &GUID_PROP_COMPOSING, &GUID_PROP_ATTRIBUTE };
    wil::com_ptr<IEnumTfRanges> EnumPropertyChanged;
    if (FAILED(editRecord->GetTextAndPropertyUpdates(TF_GTP_INCL_TEXT, guids, ARRAYSIZE(guids), EnumPropertyChanged.addressof())))
    {
        return FALSE;
    }

    wil::com_ptr<ITfRange> range;
    while (EnumPropertyChanged->Next(1, range.put(), nullptr) == S_OK)
    {
        BOOL empty;
        if (range->IsEmpty(ec, &empty) != S_OK || !empty)
        {
            return TRUE;
        }
    }
    return FALSE;
}

HRESULT CConsoleTSF::_GetAllTextRange(TfEditCookie ec, ITfContext* ic, ITfRange** range, LONG* lpTextLength, TF_HALTCOND* lpHaltCond)
{
    *lpTextLength = 0;

    wil::com_ptr_nothrow<ITfRange> rangeFull;
    RETURN_IF_FAILED(ic->GetStart(ec, rangeFull.addressof()));

    LONG cch = 0;
    RETURN_IF_FAILED(rangeFull->ShiftEnd(ec, LONG_MAX, &cch, lpHaltCond));
    RETURN_IF_FAILED(rangeFull->Clone(range));

    *lpTextLength = cch;
    return S_OK;
}

HRESULT CConsoleTSF::_ClearTextInRange(TfEditCookie ec, ITfRange* range)
{
    SetModifyingDocFlag(TRUE);
    const auto hr = range->SetText(ec, 0, nullptr, 0);
    SetModifyingDocFlag(FALSE);
    return hr;
}

HRESULT CConsoleTSF::_GetTextAndAttribute(TfEditCookie ec, ITfRange* range, std::wstring& CompStr, std::vector<TfGuidAtom> CompGuid, BOOL bInWriteSession, CicCategoryMgr* pCicCatMgr, CicDisplayAttributeMgr* pCicDispAttr)
{
    std::wstring ResultStr;
    return _GetTextAndAttribute(ec, range, CompStr, CompGuid, ResultStr, bInWriteSession, pCicCatMgr, pCicDispAttr);
}

[[nodiscard]] HRESULT CConsoleTSF::_GetCursorPosition(TfEditCookie ec, CCompCursorPos& CompCursorPos)
{
    const auto pic = GetInputContext();
    if (pic == nullptr)
    {
        return E_FAIL;
    }

    HRESULT hr;
    ULONG cFetched;

    TF_SELECTION sel;
    sel.range = nullptr;

    if (SUCCEEDED(hr = pic->GetSelection(ec, TF_DEFAULT_SELECTION, 1, &sel, &cFetched)))
    {
        wil::com_ptr_nothrow<ITfRange> start;
        LONG ich;
        TF_HALTCOND hc;

        hc.pHaltRange = sel.range;
        hc.aHaltPos = (sel.style.ase == TF_AE_START) ? TF_ANCHOR_START : TF_ANCHOR_END;
        hc.dwFlags = 0;

        if (SUCCEEDED(hr = _GetAllTextRange(ec, pic, &start, &ich, &hc)))
        {
            CompCursorPos.SetCursorPosition(ich);
        }

        sel.range->Release();
    }

    return hr;
}

//
// Get text and attribute in given range
//
//                                ITfRange::range
//   TF_ANCHOR_START
//    |======================================================================|
//                        +--------------------+          #+----------+
//                        |ITfRange::pPropRange|          #|pPropRange|
//                        +--------------------+          #+----------+
//                        |     GUID_ATOM      |          #
//                        +--------------------+          #
//    ^^^^^^^^^^^^^^^^^^^^                      ^^^^^^^^^^#
//    ITfRange::gap_range                       gap_range #
//                                                        #
//                                                        V
//                                                        ITfRange::no_display_attribute_range
//                                                   result_comp
//                                          +1   <-       0    ->     -1
//
[[nodiscard]] HRESULT CConsoleTSF::_GetTextAndAttribute(TfEditCookie ec, ITfRange* rangeIn, std::wstring& CompStr, std::vector<TfGuidAtom>& CompGuid, std::wstring& ResultStr, BOOL bInWriteSession, CicCategoryMgr* pCicCatMgr, CicDisplayAttributeMgr* pCicDispAttr)
{
    HRESULT hr;

    auto pic = GetInputContext();
    if (pic == nullptr)
    {
        return E_FAIL;
    }

    //
    // Get no display attribute range if there exist.
    // Otherwise, result range is the same to input range.
    //
    LONG result_comp;
    wil::com_ptr_nothrow<ITfRange> no_display_attribute_range;
    if (FAILED(hr = rangeIn->Clone(&no_display_attribute_range)))
    {
        return hr;
    }

    const GUID* guids[] = { &GUID_PROP_COMPOSING };
    const int guid_size = sizeof(guids) / sizeof(GUID*);

    if (FAILED(hr = _GetNoDisplayAttributeRange(ec, rangeIn, guids, guid_size, no_display_attribute_range.get())))
    {
        return hr;
    }

    wil::com_ptr_nothrow<ITfReadOnlyProperty> propComp;
    if (FAILED(hr = pic->TrackProperties(guids, guid_size, NULL, 0, &propComp)))
    {
        return hr;
    }

    wil::com_ptr_nothrow<IEnumTfRanges> enumComp;
    if (FAILED(hr = propComp->EnumRanges(ec, &enumComp, rangeIn)))
    {
        return hr;
    }

    wil::com_ptr_nothrow<ITfRange> range;
    while (enumComp->Next(1, &range, nullptr) == S_OK)
    {
        VARIANT var;
        auto fCompExist = FALSE;

        hr = propComp->GetValue(ec, range.get(), &var);
        if (S_OK == hr)
        {
            wil::com_ptr_nothrow<IEnumTfPropertyValue> EnumPropVal;
            if (wil::try_com_query_to(var.punkVal, &EnumPropVal))
            {
                TF_PROPERTYVAL tfPropertyVal;

                while (EnumPropVal->Next(1, &tfPropertyVal, nullptr) == S_OK)
                {
                    for (auto i = 0; i < guid_size; i++)
                    {
                        if (IsEqualGUID(tfPropertyVal.guidId, *guids[i]))
                        {
                            if ((V_VT(&tfPropertyVal.varValue) == VT_I4 && V_I4(&tfPropertyVal.varValue) != 0))
                            {
                                fCompExist = TRUE;
                                break;
                            }
                        }
                    }

                    VariantClear(&tfPropertyVal.varValue);

                    if (fCompExist)
                    {
                        break;
                    }
                }
            }
        }

        VariantClear(&var);

        ULONG ulNumProp;

        wil::com_ptr_nothrow<IEnumTfRanges> enumProp;
        wil::com_ptr_nothrow<ITfReadOnlyProperty> prop;
        if (FAILED(hr = pCicDispAttr->GetDisplayAttributeTrackPropertyRange(ec, pic, range.get(), &prop, &enumProp, &ulNumProp)))
        {
            return hr;
        }

        // use text range for get text
        wil::com_ptr_nothrow<ITfRange> textRange;
        if (FAILED(hr = range->Clone(&textRange)))
        {
            return hr;
        }

        // use text range for gap text (no property range).
        wil::com_ptr_nothrow<ITfRange> gap_range;
        if (FAILED(hr = range->Clone(&gap_range)))
        {
            return hr;
        }

        wil::com_ptr_nothrow<ITfRange> pPropRange;
        while (enumProp->Next(1, &pPropRange, nullptr) == S_OK)
        {
            // pick up the gap up to the next property
            gap_range->ShiftEndToRange(ec, pPropRange.get(), TF_ANCHOR_START);

            //
            // GAP range
            //
            no_display_attribute_range->CompareStart(ec, gap_range.get(), TF_ANCHOR_START, &result_comp);
            LOG_IF_FAILED(_GetTextAndAttributeGapRange(ec, gap_range.get(), result_comp, CompStr, CompGuid, ResultStr));

            //
            // Get display attribute data if some GUID_ATOM exist.
            //
            TF_DISPLAYATTRIBUTE da;
            auto guidatom = TF_INVALID_GUIDATOM;

            LOG_IF_FAILED(pCicDispAttr->GetDisplayAttributeData(pCicCatMgr->GetCategoryMgr(), ec, prop.get(), pPropRange.get(), &da, &guidatom, ulNumProp));

            //
            // Property range
            //
            no_display_attribute_range->CompareStart(ec, pPropRange.get(), TF_ANCHOR_START, &result_comp);

            // Adjust GAP range's start anchor to the end of property range.
            gap_range->ShiftStartToRange(ec, pPropRange.get(), TF_ANCHOR_END);

            //
            // Get property text
            //
            LOG_IF_FAILED(_GetTextAndAttributePropertyRange(ec, pPropRange.get(), fCompExist, result_comp, bInWriteSession, da, guidatom, CompStr, CompGuid, ResultStr));

        } // while

        // the last non-attr
        textRange->ShiftStartToRange(ec, gap_range.get(), TF_ANCHOR_START);
        textRange->ShiftEndToRange(ec, range.get(), TF_ANCHOR_END);

        BOOL fEmpty;
        while (textRange->IsEmpty(ec, &fEmpty) == S_OK && !fEmpty)
        {
            WCHAR wstr0[256 + 1];
            ULONG ulcch0 = ARRAYSIZE(wstr0) - 1;
            textRange->GetText(ec, TF_TF_MOVESTART, wstr0, ulcch0, &ulcch0);

            TfGuidAtom guidatom;
            guidatom = TF_INVALID_GUIDATOM;

            TF_DISPLAYATTRIBUTE da;
            da.bAttr = TF_ATTR_INPUT;

            try
            {
                CompGuid.insert(CompGuid.end(), ulcch0, guidatom);
                CompStr.append(wstr0, ulcch0);
            }
            CATCH_RETURN();
        }

        textRange->Collapse(ec, TF_ANCHOR_END);

    } // out-most while for GUID_PROP_COMPOSING

    //
    // set GUID_PROP_CONIME_TRACKCOMPOSITION
    //
    wil::com_ptr_nothrow<ITfProperty> PropertyTrackComposition;
    if (SUCCEEDED(hr = pic->GetProperty(GUID_PROP_CONIME_TRACKCOMPOSITION, &PropertyTrackComposition)))
    {
        VARIANT var;
        var.vt = VT_I4;
        var.lVal = 1;
        PropertyTrackComposition->SetValue(ec, rangeIn, &var);
    }

    return hr;
}

[[nodiscard]] HRESULT CConsoleTSF::_GetTextAndAttributeGapRange(TfEditCookie ec, ITfRange* gap_range, LONG result_comp, std::wstring& CompStr, std::vector<TfGuidAtom>& CompGuid, std::wstring& ResultStr)
{
    TfGuidAtom guidatom;
    guidatom = TF_INVALID_GUIDATOM;

    TF_DISPLAYATTRIBUTE da;
    da.bAttr = TF_ATTR_INPUT;

    BOOL fEmpty;
    WCHAR wstr0[256 + 1];
    ULONG ulcch0;

    while (gap_range->IsEmpty(ec, &fEmpty) == S_OK && !fEmpty)
    {
        wil::com_ptr_nothrow<ITfRange> backup_range;
        if (FAILED(gap_range->Clone(&backup_range)))
        {
            return E_FAIL;
        }

        //
        // Retrieve gap text if there exist.
        //
        ulcch0 = ARRAYSIZE(wstr0) - 1;
        if (FAILED(gap_range->GetText(ec, TF_TF_MOVESTART, wstr0, ulcch0, &ulcch0)))
        {
            return E_FAIL;
        }

        try
        {
            if (result_comp <= 0)
            {
                CompGuid.insert(CompGuid.end(), ulcch0, guidatom);
                CompStr.append(wstr0, ulcch0);
            }
            else
            {
                ResultStr.append(wstr0, ulcch0);
                LOG_IF_FAILED(_ClearTextInRange(ec, backup_range.get()));
            }
        }
        CATCH_RETURN();
    }

    return S_OK;
}

[[nodiscard]] HRESULT CConsoleTSF::_GetTextAndAttributePropertyRange(TfEditCookie ec, ITfRange* pPropRange, BOOL fCompExist, LONG result_comp, BOOL bInWriteSession, TF_DISPLAYATTRIBUTE da, TfGuidAtom guidatom, std::wstring& CompStr, std::vector<TfGuidAtom>& CompGuid, std::wstring& ResultStr)
{
    BOOL fEmpty;
    WCHAR wstr0[256 + 1];
    ULONG ulcch0;

    while (pPropRange->IsEmpty(ec, &fEmpty) == S_OK && !fEmpty)
    {
        wil::com_ptr_nothrow<ITfRange> backup_range;
        if (FAILED(pPropRange->Clone(&backup_range)))
        {
            return E_FAIL;
        }

        //
        // Retrieve property text if there exist.
        //
        ulcch0 = ARRAYSIZE(wstr0) - 1;
        if (FAILED(pPropRange->GetText(ec, TF_TF_MOVESTART, wstr0, ulcch0, &ulcch0)))
        {
            return E_FAIL;
        }

        try
        {
            // see if there is a valid disp attribute
            if (fCompExist == TRUE && result_comp <= 0)
            {
                if (guidatom == TF_INVALID_GUIDATOM)
                {
                    da.bAttr = TF_ATTR_INPUT;
                }
                CompGuid.insert(CompGuid.end(), ulcch0, guidatom);
                CompStr.append(wstr0, ulcch0);
            }
            else if (bInWriteSession)
            {
                // if there's no disp attribute attached, it probably means
                // the part of string is finalized.
                //
                ResultStr.append(wstr0, ulcch0);

                // it was a 'determined' string
                // so the doc has to shrink
                //
                LOG_IF_FAILED(_ClearTextInRange(ec, backup_range.get()));
            }
            else
            {
                //
                // Prevent infinite loop
                //
                break;
            }
        }
        CATCH_RETURN();
    }

    return S_OK;
}

[[nodiscard]] HRESULT CConsoleTSF::_GetNoDisplayAttributeRange(TfEditCookie ec, ITfRange* rangeIn, const GUID** guids, const int guid_size, ITfRange* no_display_attribute_range)
{
    auto pic = GetInputContext();
    if (pic == nullptr)
    {
        return E_FAIL;
    }

    wil::com_ptr_nothrow<ITfReadOnlyProperty> propComp;
    auto hr = pic->TrackProperties(guids, guid_size, // system property
                                   nullptr,
                                   0, // application property
                                   &propComp);
    if (FAILED(hr))
    {
        return hr;
    }

    wil::com_ptr_nothrow<IEnumTfRanges> enumComp;
    hr = propComp->EnumRanges(ec, &enumComp, rangeIn);
    if (FAILED(hr))
    {
        return hr;
    }

    wil::com_ptr_nothrow<ITfRange> pRange;

    while (enumComp->Next(1, &pRange, nullptr) == S_OK)
    {
        VARIANT var;
        auto fCompExist = FALSE;

        hr = propComp->GetValue(ec, pRange.get(), &var);
        if (S_OK == hr)
        {
            wil::com_ptr_nothrow<IEnumTfPropertyValue> EnumPropVal;
            if (wil::try_com_query_to(var.punkVal, &EnumPropVal))
            {
                TF_PROPERTYVAL tfPropertyVal;

                while (EnumPropVal->Next(1, &tfPropertyVal, nullptr) == S_OK)
                {
                    for (auto i = 0; i < guid_size; i++)
                    {
                        if (IsEqualGUID(tfPropertyVal.guidId, *guids[i]))
                        {
                            if ((V_VT(&tfPropertyVal.varValue) == VT_I4 && V_I4(&tfPropertyVal.varValue) != 0))
                            {
                                fCompExist = TRUE;
                                break;
                            }
                        }
                    }

                    VariantClear(&tfPropertyVal.varValue);

                    if (fCompExist)
                    {
                        break;
                    }
                }
            }
        }

        if (!fCompExist)
        {
            // Adjust GAP range's start anchor to the end of property range.
            no_display_attribute_range->ShiftStartToRange(ec, pRange.get(), TF_ANCHOR_START);
        }

        VariantClear(&var);
    }

    return S_OK;
}

[[nodiscard]] HRESULT CConsoleTSF::_IsInterimSelection(TfEditCookie ec, ITfRange** pInterimRange, BOOL* pfInterim)
{
    const auto pic = GetInputContext();
    if (pic == nullptr)
    {
        return E_FAIL;
    }

    ULONG cFetched;

    TF_SELECTION sel;
    sel.range = nullptr;

    *pfInterim = FALSE;
    if (pic->GetSelection(ec, TF_DEFAULT_SELECTION, 1, &sel, &cFetched) != S_OK)
    {
        // no selection. we can return S_OK.
        return S_OK;
    }

    if (sel.style.fInterimChar && sel.range)
    {
        HRESULT hr;
        if (FAILED(hr = sel.range->Clone(pInterimRange)))
        {
            sel.range->Release();
            return hr;
        }

        *pfInterim = TRUE;
    }

    sel.range->Release();

    return S_OK;
}

[[nodiscard]] HRESULT CConsoleTSF::_MakeCompositionString(TfEditCookie ec, ITfRange* FullTextRange, BOOL bInWriteSession, CicCategoryMgr* pCicCatMgr, CicDisplayAttributeMgr* pCicDispAttr)
{
    std::wstring CompStr;
    std::vector<TfGuidAtom> CompGuid;
    CCompCursorPos CompCursorPos;
    std::wstring ResultStr;
    auto fIgnorePreviousCompositionResult = FALSE;

    RETURN_IF_FAILED(_GetTextAndAttribute(ec, FullTextRange, CompStr, CompGuid, ResultStr, bInWriteSession, pCicCatMgr, pCicDispAttr));

    if (IsPendingCompositionCleanup())
    {
        // Don't draw the previous composition result if there was a cleanup session requested for it.
        fIgnorePreviousCompositionResult = TRUE;
        // Cancel pending cleanup, since the ResultStr was cleared from the composition in _GetTextAndAttribute.
        OnCompositionCleanup(TRUE);
    }

    RETURN_IF_FAILED(_GetCursorPosition(ec, CompCursorPos));

    // Get display attribute manager
    const auto dam = pCicDispAttr->GetDisplayAttributeMgr();
    RETURN_HR_IF_NULL(E_FAIL, dam);

    // Get category manager
    const auto cat = pCicCatMgr->GetCategoryMgr();
    RETURN_HR_IF_NULL(E_FAIL, cat);

    // Allocate and fill TF_DISPLAYATTRIBUTE
    try
    {
        // Get conversion area service.
        const auto conv_area = GetConversionArea();
        RETURN_HR_IF_NULL(E_FAIL, conv_area);

        if (!ResultStr.empty() && !fIgnorePreviousCompositionResult)
        {
            return conv_area->DrawResult(ResultStr);
        }
        if (!CompStr.empty())
        {
            const auto cchDisplayAttribute = CompGuid.size();
            std::vector<TF_DISPLAYATTRIBUTE> DisplayAttributes;
            DisplayAttributes.reserve(cchDisplayAttribute);

            for (size_t i = 0; i < cchDisplayAttribute; i++)
            {
                TF_DISPLAYATTRIBUTE da;
                ZeroMemory(&da, sizeof(da));
                da.bAttr = TF_ATTR_OTHER;

                GUID guid;
                if (SUCCEEDED(cat->GetGUID(CompGuid.at(i), &guid)))
                {
                    CLSID clsid;
                    wil::com_ptr_nothrow<ITfDisplayAttributeInfo> dai;
                    if (SUCCEEDED(dam->GetDisplayAttributeInfo(guid, &dai, &clsid)))
                    {
                        dai->GetAttributeInfo(&da);
                    }
                }

                DisplayAttributes.emplace_back(da);
            }

            return conv_area->DrawComposition(CompStr, DisplayAttributes, CompCursorPos.GetCursorPosition());
        }
    }
    CATCH_RETURN();

    return S_OK;
}

[[nodiscard]] HRESULT CConsoleTSF::_MakeInterimString(TfEditCookie ec, ITfRange* FullTextRange, ITfRange* InterimRange, LONG lTextLength, BOOL bInWriteSession, CicCategoryMgr* pCicCatMgr, CicDisplayAttributeMgr* pCicDispAttr)
{
    LONG lStartResult;
    LONG lEndResult;

    FullTextRange->CompareStart(ec, InterimRange, TF_ANCHOR_START, &lStartResult);
    RETURN_HR_IF(E_FAIL, lStartResult > 0);

    FullTextRange->CompareEnd(ec, InterimRange, TF_ANCHOR_END, &lEndResult);
    RETURN_HR_IF(E_FAIL, lEndResult < 0);

    if (lStartResult < 0)
    {
        // Make result string.
        RETURN_IF_FAILED(FullTextRange->ShiftEndToRange(ec, InterimRange, TF_ANCHOR_START));

        // Interim char assume 1 char length.
        // Full text length - 1 means result string length.
        lTextLength--;

        assert((lTextLength > 0));

        if (lTextLength > 0)
        {
            try
            {
                const auto wstr = std::make_unique<WCHAR[]>(lTextLength + 1);

                // Get the result text, finalize it, and erase the result text.
                if (SUCCEEDED(FullTextRange->GetText(ec, TF_TF_IGNOREEND, wstr.get(), (ULONG)lTextLength, (ULONG*)&lTextLength)))
                {
                    // Clear the TOM
                    LOG_IF_FAILED(_ClearTextInRange(ec, FullTextRange));
                }
            }
            CATCH_RETURN();
        }
    }

    // Make interim character
    std::wstring CompStr;
    std::vector<TfGuidAtom> CompGuid;
    std::wstring _tempResultStr;

    RETURN_IF_FAILED(_GetTextAndAttribute(ec, InterimRange, CompStr, CompGuid, _tempResultStr, bInWriteSession, pCicCatMgr, pCicDispAttr));

    // Get display attribute manager
    const auto dam = pCicDispAttr->GetDisplayAttributeMgr();
    RETURN_HR_IF_NULL(E_FAIL, dam);

    // Get category manager
    const auto cat = pCicCatMgr->GetCategoryMgr();
    RETURN_HR_IF_NULL(E_FAIL, cat);

    // Allocate and fill TF_DISPLAYATTRIBUTE
    try
    {
        // Get conversion area service.
        const auto conv_area = GetConversionArea();
        RETURN_HR_IF_NULL(E_FAIL, conv_area);

        if (!CompStr.empty())
        {
            const auto cchDisplayAttribute = CompGuid.size();
            std::vector<TF_DISPLAYATTRIBUTE> DisplayAttributes;
            DisplayAttributes.reserve(cchDisplayAttribute);

            for (size_t i = 0; i < cchDisplayAttribute; i++)
            {
                TF_DISPLAYATTRIBUTE da;
                ZeroMemory(&da, sizeof(da));
                da.bAttr = TF_ATTR_OTHER;
                GUID guid;
                if (SUCCEEDED(cat->GetGUID(CompGuid.at(i), &guid)))
                {
                    CLSID clsid;
                    wil::com_ptr_nothrow<ITfDisplayAttributeInfo> dai;
                    if (SUCCEEDED(dam->GetDisplayAttributeInfo(guid, &dai, &clsid)))
                    {
                        dai->GetAttributeInfo(&da);
                    }
                }

                DisplayAttributes.emplace_back(da);
            }

            return conv_area->DrawComposition(CompStr, // composition string (Interim string)
                                              DisplayAttributes); // display attributes
        }
    }
    CATCH_RETURN();

    return S_OK;
}

[[nodiscard]] HRESULT CConsoleTSF::_CreateCategoryAndDisplayAttributeManager(CicCategoryMgr** pCicCatMgr, CicDisplayAttributeMgr** pCicDispAttr)
{
    auto hr = E_OUTOFMEMORY;

    CicCategoryMgr* pTmpCat = nullptr;
    CicDisplayAttributeMgr* pTmpDispAttr = nullptr;

    // Create Cicero Category Manager
    pTmpCat = new (std::nothrow) CicCategoryMgr;
    if (pTmpCat)
    {
        if (SUCCEEDED(hr = pTmpCat->InitCategoryInstance()))
        {
            if (const auto pcat = pTmpCat->GetCategoryMgr())
            {
                // Create Cicero Display Attribute Manager
                pTmpDispAttr = new (std::nothrow) CicDisplayAttributeMgr;
                if (pTmpDispAttr)
                {
                    if (SUCCEEDED(hr = pTmpDispAttr->InitDisplayAttributeInstance(pcat)))
                    {
                        *pCicCatMgr = pTmpCat;
                        *pCicDispAttr = pTmpDispAttr;
                    }
                }
            }
        }
    }

    if (FAILED(hr))
    {
        if (pTmpCat)
        {
            delete pTmpCat;
        }
        if (pTmpDispAttr)
        {
            delete pTmpDispAttr;
        }
    }

    return hr;
}
