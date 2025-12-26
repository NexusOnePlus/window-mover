#include <windows.h>
#include <iostream>
#include <string>
#include <winstring.h>
#include "VirtualDesktopInterfaces.h"

#pragma comment(lib, "runtimeobject.lib")

// ----------------------------------------------------------------------------
// Constants & Globals
// ----------------------------------------------------------------------------

#define WM_APP_MOVE_LEFT       (WM_APP + 1)
#define WM_APP_MOVE_RIGHT      (WM_APP + 2)
#define WM_APP_SWITCH_TO_INDEX (WM_APP + 3)
#define WM_APP_MOVE_TO_INDEX   (WM_APP + 4)

// Internal Desktop Enums
#define VD_DIRECTION_LEFT 3
#define VD_DIRECTION_RIGHT 4

IVirtualDesktopManagerInternal* pVDMInternal = nullptr;
IApplicationViewCollection* pAppViewCollection = nullptr;
IServiceProvider* pServiceProvider = nullptr;
HHOOK hKeyboardHook = NULL;
DWORD g_mainThreadId = 0;

// ----------------------------------------------------------------------------
// Initialization
// ----------------------------------------------------------------------------

void InitializeVirtualDesktopManager() {
    std::cout << "Initializing COM..." << std::endl;

    HRESULT hr = CoInitialize(NULL);
    if (FAILED(hr)) return;

    hr = CoCreateInstance(CLSID_ImmersiveShell, nullptr, CLSCTX_LOCAL_SERVER, IID_IServiceProvider, (void**)&pServiceProvider);
    if (FAILED(hr)) return;

    hr = pServiceProvider->QueryService(CLSID_VirtualDesktopManagerInternal_24H2, IID_IVirtualDesktopManagerInternal_24H2, (void**)&pVDMInternal);
    if (FAILED(hr)) {
        std::cerr << "Failed to get IVirtualDesktopManagerInternal." << std::endl;
        return;
    }

    hr = pServiceProvider->QueryService(IID_IApplicationViewCollection, IID_IApplicationViewCollection, (void**)&pAppViewCollection);
    if (FAILED(hr)) {
        std::cerr << "Failed to get IApplicationViewCollection." << std::endl;
    }

    std::cout << "VirtualDesktopManager initialized." << std::endl;
}

// ----------------------------------------------------------------------------
// Helpers
// ----------------------------------------------------------------------------

// Ensures enough desktops exist to reach targetIndex (0-based)
// Returns the count of desktops after operation
UINT EnsureDesktopCount(UINT targetIndex) {
    if (!pVDMInternal) return 0;

    IObjectArray* pDesktops = nullptr;
    HRESULT hr = pVDMInternal->GetDesktops(&pDesktops);
    if (FAILED(hr)) return 0;

    UINT count = 0;
    pDesktops->GetCount(&count);
    pDesktops->Release(); // Release initial get

    while (count <= targetIndex) {
        std::cout << "Creating Desktop " << count + 1 << "..." << std::endl;
        IVirtualDesktop* pNewDesktop = nullptr;
        hr = pVDMInternal->CreateDesktop(&pNewDesktop);
        if (FAILED(hr)) {
            std::cerr << "CreateDesktop failed." << std::endl;
            break;
        }
        if (pNewDesktop) pNewDesktop->Release();
        count++;
    }
    return count;
}

IVirtualDesktop* GetDesktopAtIndex(UINT index) {
    if (!pVDMInternal) return nullptr;
    IObjectArray* pDesktops = nullptr;
    pVDMInternal->GetDesktops(&pDesktops);
    if (!pDesktops) return nullptr;

    IVirtualDesktop* pDesktop = nullptr;
    // Using IUnknown IID as generic accessor
    pDesktops->GetAt(index, IID_IUnknown, (void**)&pDesktop);
    pDesktops->Release();
    return pDesktop;
}

// ----------------------------------------------------------------------------
// Logic
// ----------------------------------------------------------------------------

void SwitchToDesktopAtIndex(int targetIndex) {
    std::cout << "--- SwitchToDesktopAtIndex (" << targetIndex << ") ---" << std::endl;
    if (!pVDMInternal) return;

    EnsureDesktopCount(targetIndex);
    
    IVirtualDesktop* pTarget = GetDesktopAtIndex(targetIndex);
    if (pTarget) {
        pVDMInternal->SwitchDesktop(pTarget);
        pTarget->Release();
        std::cout << "Switched to Desktop " << targetIndex + 1 << std::endl;
    }
}

void MoveWindowToDesktopAtIndex(int targetIndex) {
    std::cout << "--- MoveWindowToDesktopAtIndex (" << targetIndex << ") ---" << std::endl;
    if (!pVDMInternal || !pAppViewCollection) return;

    HWND hwnd = GetForegroundWindow();
    if (!hwnd) return;

    IApplicationView* pView = nullptr;
    if (FAILED(pAppViewCollection->GetViewForHwnd(hwnd, &pView))) return;

    EnsureDesktopCount(targetIndex);

    IVirtualDesktop* pTarget = GetDesktopAtIndex(targetIndex);
    if (pTarget) {
        pVDMInternal->MoveViewToDesktop(pView, pTarget);
        pVDMInternal->SwitchDesktop(pTarget);
        pTarget->Release();
        std::cout << "Moved window and switched to Desktop " << targetIndex + 1 << std::endl;
    }
    pView->Release();
}

void MoveWindowRelative(int direction) {
    std::cout << "--- MoveWindowRelative (" << direction << ") ---" << std::endl;
    if (!pVDMInternal || !pAppViewCollection) return;

    HWND hwnd = GetForegroundWindow();
    if (!hwnd) return;

    IApplicationView* pView = nullptr;
    if (FAILED(pAppViewCollection->GetViewForHwnd(hwnd, &pView))) return;

    IVirtualDesktop* pCurrent = nullptr;
    if (FAILED(pVDMInternal->GetCurrentDesktop(&pCurrent))) {
        pView->Release();
        return;
    }

    IVirtualDesktop* pAdjacent = nullptr;
    int dirCode = (direction == -1) ? VD_DIRECTION_LEFT : VD_DIRECTION_RIGHT;
    HRESULT hr = pVDMInternal->GetAdjacentDesktop(pCurrent, dirCode, &pAdjacent);

    // If moving right and no adjacent desktop exists, create one
    if ((FAILED(hr) || !pAdjacent) && direction == 1) {
        std::cout << "No adjacent desktop on right. Creating new one..." << std::endl;
        IVirtualDesktop* pNew = nullptr;
        if (SUCCEEDED(pVDMInternal->CreateDesktop(&pNew))) {
            pAdjacent = pNew; // Use the new one
        }
    }

    if (pAdjacent) {
        pVDMInternal->MoveViewToDesktop(pView, pAdjacent);
        pVDMInternal->SwitchDesktop(pAdjacent);
        pAdjacent->Release();
        std::cout << "Success." << std::endl;
    } else {
        std::cerr << "Could not find or create adjacent desktop." << std::endl;
    }

    pCurrent->Release();
    pView->Release();
}

// ----------------------------------------------------------------------------
// Hooks
// ----------------------------------------------------------------------------

LRESULT CALLBACK LowLevelKeyboardProc(int nCode, WPARAM wParam, LPARAM lParam) {
    if (nCode == HC_ACTION) {
        if (wParam == WM_KEYDOWN || wParam == WM_SYSKEYDOWN) {
            KBDLLHOOKSTRUCT* p = (KBDLLHOOKSTRUCT*)lParam;
            
            bool win = (GetAsyncKeyState(VK_LWIN) & 0x8000) || (GetAsyncKeyState(VK_RWIN) & 0x8000);
            bool shift = (GetAsyncKeyState(VK_SHIFT) & 0x8000);
            bool alt = (GetAsyncKeyState(VK_MENU) & 0x8000);

            if (win) {
                // Win + Shift + Left/Right -> Move Window Relative
                if (shift && !alt) {
                    if (p->vkCode == VK_LEFT) {
                        PostThreadMessage(g_mainThreadId, WM_APP_MOVE_LEFT, 0, 0);
                        return 1;
                    }
                    if (p->vkCode == VK_RIGHT) {
                        PostThreadMessage(g_mainThreadId, WM_APP_MOVE_RIGHT, 0, 0);
                        return 1;
                    }
                }
                // Win + Number -> Switch Only
                else if (!shift && !alt && p->vkCode >= '1' && p->vkCode <= '9') {
                    int index = p->vkCode - '1';
                    PostThreadMessage(g_mainThreadId, WM_APP_SWITCH_TO_INDEX, (WPARAM)index, 0);
                    return 1; 
                }
                // Win + Alt + Number -> Move Window & Switch
                else if (!shift && alt && p->vkCode >= '1' && p->vkCode <= '9') {
                    int index = p->vkCode - '1';
                    PostThreadMessage(g_mainThreadId, WM_APP_MOVE_TO_INDEX, (WPARAM)index, 0);
                    return 1;
                }
            }
        }
    }
    return CallNextHookEx(hKeyboardHook, nCode, wParam, lParam);
}

// ----------------------------------------------------------------------------
// Main
// ----------------------------------------------------------------------------

int main() {
    g_mainThreadId = GetCurrentThreadId();
    InitializeVirtualDesktopManager();

    if (!pVDMInternal) return 1;

    hKeyboardHook = SetWindowsHookEx(WH_KEYBOARD_LL, LowLevelKeyboardProc, GetModuleHandle(NULL), 0);
    
    std::cout << "DesktopMover Running:" << std::endl;
    std::cout << " [Win + 1..9]         Switch to Desktop" << std::endl;
    std::cout << " [Win + Alt + 1..9]   Move Window & Switch" << std::endl;
    std::cout << " [Win + Shift + <|>]  Move Window Left/Right (Creates on Right)" << std::endl;

    MSG msg;
    while (GetMessage(&msg, NULL, 0, 0)) {
        switch (msg.message) {
            case WM_APP_MOVE_LEFT:       MoveWindowRelative(-1); break;
            case WM_APP_MOVE_RIGHT:      MoveWindowRelative(1); break;
            case WM_APP_SWITCH_TO_INDEX: SwitchToDesktopAtIndex((int)msg.wParam); break;
            case WM_APP_MOVE_TO_INDEX:   MoveWindowToDesktopAtIndex((int)msg.wParam); break;
            default:
                TranslateMessage(&msg);
                DispatchMessage(&msg);
        }
    }

    UnhookWindowsHookEx(hKeyboardHook);
    CoUninitialize();
    return 0;
}