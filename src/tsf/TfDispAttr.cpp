/*++

Copyright (c) Microsoft Corporation.
Licensed under the MIT license.

Module Name:

    TfDispAttr.cpp

Abstract:

    This file implements the CicDisplayAttributeMgr Class.

Author:

Revision History:

Notes:

--*/

#include "precomp.h"
#include "TfDispAttr.h"

[[nodiscard]] HRESULT CicDisplayAttributeMgr::GetDisplayAttributeTrackPropertyRange(TfEditCookie ec, ITfContext* pic, ITfRange* pRange, ITfReadOnlyProperty** ppProp, IEnumTfRanges** ppEnum, ULONG* pulNumProp)
{
    const auto ulNumProp = static_cast<ULONG>(m_DispAttrProp.size());
    if (ulNumProp == 0)
    {
        return S_OK;
    }

    // TrackProperties wants an array of GUID *'s
    const auto ppguidProp = std::make_unique<const GUID*[]>(ulNumProp);
    for (size_t i = 0; i < ulNumProp; i++)
    {
        ppguidProp[i] = &m_DispAttrProp[i];
    }

    wil::com_ptr<ITfReadOnlyProperty> pProp;
    THROW_IF_FAILED(pic->TrackProperties(ppguidProp.get(), ulNumProp, nullptr, 0, pProp.addressof()));
    THROW_IF_FAILED(pProp->EnumRanges(ec, ppEnum, pRange));

    *ppProp = pProp.detach();
    *pulNumProp = ulNumProp;
    return S_OK;
}

[[nodiscard]] HRESULT CicDisplayAttributeMgr::GetDisplayAttributeData(ITfCategoryMgr* pcat, TfEditCookie ec, ITfReadOnlyProperty* pProp, ITfRange* pRange, TF_DISPLAYATTRIBUTE* pda, TfGuidAtom* pguid, ULONG /*ulNumProp*/)
{
    wil::unique_variant var;
    if (FAILED(pProp->GetValue(ec, pRange, &var)) || var.vt != VT_UNKNOWN)
    {
        return S_OK;
    }

    wil::com_ptr_nothrow<IEnumTfPropertyValue> pEnumPropertyVal;
    if (!wil::try_com_query_to(var.punkVal, pEnumPropertyVal.addressof()))
    {
        return S_OK;
    }

    TF_PROPERTYVAL tfPropVal;
    while (pEnumPropertyVal->Next(1, &tfPropVal, nullptr) == S_OK)
    {
        if (tfPropVal.varValue.vt != VT_I4) // expecting TfGuidAtom
        {
            continue;
        }

        const auto gaVal = static_cast<TfGuidAtom>(tfPropVal.varValue.lVal);

        GUID guid;
        wil::com_ptr_nothrow<ITfDisplayAttributeInfo> pDAI;
        if (SUCCEEDED(pcat->GetGUID(gaVal, &guid)) &&
            SUCCEEDED(m_pDAM->GetDisplayAttributeInfo(guid, pDAI.addressof(), nullptr)))
        {
            //
            // Issue: for simple apps.
            //
            // Small apps can not show multi underline. So
            // this helper function returns only one
            // DISPLAYATTRIBUTE structure.
            //
            if (pda)
            {
                pDAI->GetAttributeInfo(pda);
            }

            if (pguid)
            {
                *pguid = gaVal;
            }

            break;
        }
    }

    return S_OK;
}

[[nodiscard]] HRESULT CicDisplayAttributeMgr::InitDisplayAttributeInstance(ITfCategoryMgr* pcat)
{
    m_pDAM = wil::CoCreateInstance<ITfDisplayAttributeMgr>(CLSID_TF_DisplayAttributeMgr);

    wil::com_ptr_nothrow<IEnumGUID> pEnumProp;
    std::ignore = pcat->EnumItemsInCategory(GUID_TFCAT_DISPLAYATTRIBUTEPROPERTY, pEnumProp.addressof());

    //
    // make a database for Display Attribute Properties.
    //
    if (pEnumProp)
    {
        //
        // add System Display Attribute first.
        // so no other Display Attribute property overwrite it.
        //
        m_DispAttrProp.emplace_back(GUID_PROP_ATTRIBUTE);

        GUID guidProp;
        while (pEnumProp->Next(1, &guidProp, nullptr) == S_OK)
        {
            if (!IsEqualGUID(guidProp, GUID_PROP_ATTRIBUTE))
            {
                m_DispAttrProp.emplace_back(guidProp);
            }
        }
    }

    return S_OK;
}

ITfDisplayAttributeMgr* CicDisplayAttributeMgr::GetDisplayAttributeMgr()
{
    return m_pDAM.get();
}
