#pragma once
#include <windows.h>
#include <unknwn.h>
#include <vector>

// ----------------------------------------------------------------------------
// GUIDs for Windows 11 Build 26200 (24H2)
// ----------------------------------------------------------------------------

static const GUID CLSID_ImmersiveShell = 
    { 0xC2F03A33, 0x21F5, 0x47FA, { 0xB4, 0xBB, 0x15, 0x63, 0x62, 0xA2, 0xF2, 0x39 } };

static const GUID CLSID_VirtualDesktopManagerInternal_24H2 = 
    { 0xC5E0CDCA, 0x7B6E, 0x41B2, { 0x9F, 0xC4, 0xD9, 0x39, 0x75, 0xCC, 0x46, 0x7B } };

static const GUID IID_IVirtualDesktopManagerInternal_24H2 = 
    { 0x53F5CA0B, 0x158F, 0x4124, { 0x90, 0x0C, 0x05, 0x71, 0x58, 0x06, 0x0B, 0x27 } };

static const GUID IID_IApplicationViewCollection = 
    { 0x1841C6D7, 0x4F9D, 0x42C0, { 0xAF, 0x41, 0x87, 0x47, 0x53, 0x8F, 0x10, 0xE5 } };

// ----------------------------------------------------------------------------
// Interface Definitions
// ----------------------------------------------------------------------------

struct IApplicationView : public IUnknown {
    // Inherits IUnknown. Full vtable omitted for brevity as we only pass pointers.
};

struct IVirtualDesktop : public IUnknown {
    virtual HRESULT STDMETHODCALLTYPE IsViewVisible(IApplicationView* view, BOOL* visible) = 0;
    virtual HRESULT STDMETHODCALLTYPE GetID(GUID* id) = 0;
    virtual HRESULT STDMETHODCALLTYPE GetName(HSTRING* name) = 0;
    virtual HRESULT STDMETHODCALLTYPE GetWallpaperPath(HSTRING* path) = 0;
    virtual HRESULT STDMETHODCALLTYPE IsRemote(BOOL* remote) = 0;
};

struct IObjectArray : public IUnknown {
    virtual HRESULT STDMETHODCALLTYPE GetCount(UINT* count) = 0;
    virtual HRESULT STDMETHODCALLTYPE GetAt(UINT index, REFIID riid, void** ppv) = 0;
};

struct IVirtualDesktopManagerInternal : public IUnknown {
    virtual HRESULT STDMETHODCALLTYPE GetCount(UINT* count) = 0;
    virtual HRESULT STDMETHODCALLTYPE MoveViewToDesktop(IApplicationView* view, IVirtualDesktop* desktop) = 0;
    virtual HRESULT STDMETHODCALLTYPE CanViewMoveToDesktop(IApplicationView* view, BOOL* canMove) = 0;
    virtual HRESULT STDMETHODCALLTYPE GetCurrentDesktop(IVirtualDesktop** desktop) = 0;
    virtual HRESULT STDMETHODCALLTYPE GetDesktops(IObjectArray** desktops) = 0;
    virtual HRESULT STDMETHODCALLTYPE GetAdjacentDesktop(IVirtualDesktop* desktop, int direction, IVirtualDesktop** adjacent) = 0;
    virtual HRESULT STDMETHODCALLTYPE SwitchDesktop(IVirtualDesktop* desktop) = 0;
    virtual HRESULT STDMETHODCALLTYPE SwitchDesktopAndMoveForegroundView(IVirtualDesktop* desktop) = 0;
    virtual HRESULT STDMETHODCALLTYPE CreateDesktop(IVirtualDesktop** desktop) = 0;
    virtual HRESULT STDMETHODCALLTYPE MoveDesktop(IVirtualDesktop* desktop, int nIndex) = 0;
    virtual HRESULT STDMETHODCALLTYPE RemoveDesktop(IVirtualDesktop* desktop, IVirtualDesktop* fallback) = 0;
    virtual HRESULT STDMETHODCALLTYPE FindDesktop(GUID* desktopId, IVirtualDesktop** desktop) = 0;
};

struct IApplicationViewCollection : public IUnknown {
    virtual HRESULT STDMETHODCALLTYPE GetViews(IObjectArray** views) = 0;
    virtual HRESULT STDMETHODCALLTYPE GetViewsByZOrder(IObjectArray** views) = 0;
    virtual HRESULT STDMETHODCALLTYPE GetViewsByAppUserModelId(PCWSTR id, IObjectArray** views) = 0;
    virtual HRESULT STDMETHODCALLTYPE GetViewForHwnd(HWND hwnd, IApplicationView** view) = 0;
    virtual HRESULT STDMETHODCALLTYPE GetViewForApplication(void* application, IApplicationView** view) = 0;
    virtual HRESULT STDMETHODCALLTYPE GetViewForAppUserModelId(PCWSTR id, IApplicationView** view) = 0;
    virtual HRESULT STDMETHODCALLTYPE GetViewInFocus(IApplicationView** view) = 0;
    virtual HRESULT STDMETHODCALLTYPE Unknown1(IApplicationView** view) = 0;
    virtual void STDMETHODCALLTYPE RefreshCollection() = 0;
    virtual HRESULT STDMETHODCALLTYPE RegisterForApplicationViewChanges(void* listener, DWORD* cookie) = 0;
    virtual HRESULT STDMETHODCALLTYPE UnregisterForApplicationViewChanges(DWORD cookie) = 0;
};
