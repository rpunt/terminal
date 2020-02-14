/*++
Copyright (c) Microsoft Corporation
Licensed under the MIT license.

Module Name:
- IControlAccessibilityInfo.h

Abstract:
- This serves as the interface defining all information known by the control
  hosting the terminal renderer that is needed for the UI Automation Tree.

Author(s):
- Zoey Riordan (zorio) Feb-2020
--*/

#pragma once

#include <wtypes.h>

namespace Microsoft::Console::Types
{
    class IControlAccessibilityInfo
    {
    public:
        virtual COORD GetFontSize() = 0;
        virtual RECT GetBounds() = 0;
        virtual RECT GetPadding() = 0;
        virtual double GetScaleFactor() = 0;
        virtual void ChangeViewport(const SMALL_RECT NewWindow) = 0;
        virtual HRESULT GetHostUiaProvider(IRawElementProviderSimple** provider) = 0;
    };
}