/*++

Copyright (c) Microsoft Corporation.
Licensed under the MIT license.

Module Name:

    TfDispAttr.h

Abstract:

    This file defines the CicDisplayAttributeMgr Class.

Author:

Revision History:

Notes:

--*/

#pragma once

class CicDisplayAttributeMgr
{
public:
    [[nodiscard]] HRESULT GetDisplayAttributeTrackPropertyRange(TfEditCookie ec, ITfContext* pic, ITfRange* pRange, ITfReadOnlyProperty** ppProp, IEnumTfRanges** ppEnum, ULONG* pulNumProp);
    [[nodiscard]] HRESULT GetDisplayAttributeData(ITfCategoryMgr* pcat, TfEditCookie ec, ITfReadOnlyProperty* pProp, ITfRange* pRange, TF_DISPLAYATTRIBUTE* pda, TfGuidAtom* pguid, ULONG ulNumProp);
    [[nodiscard]] HRESULT InitDisplayAttributeInstance(ITfCategoryMgr* pcat);
    ITfDisplayAttributeMgr* GetDisplayAttributeMgr();

private:
    wil::com_ptr_nothrow<ITfDisplayAttributeMgr> m_pDAM;
    std::vector<GUID> m_DispAttrProp;
};
