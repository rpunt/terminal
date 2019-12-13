#pragma once

#include <wtypes.h>

namespace Microsoft::Console::Types
{
    class IControlInfo
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