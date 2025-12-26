#include <windows.h>
#include <iostream>
#include <string>
#include <winstring.h>
#include "VirtualDesktopInterfaces.h"

#pragma comment(lib, "runtimeobject.lib")

// ----------------------------------------------------------------------------
// Constants & Globals
// ----------------------------------------------------------------------------

#define WM_APP_MOVE_LEFT  (WM_APP + 1)
#define WM_APP_MOVE_RIGHT (WM_APP + 2)

// Internal Desktop Enums (Standard for Windows 10/11)
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
    if (FAILED(hr)) {
        std::cerr << "CoInitialize failed: " << std::hex << hr << std::endl;
        return;
    }

    hr = CoCreateInstance(CLSID_ImmersiveShell, nullptr, CLSCTX_LOCAL_SERVER, IID_IServiceProvider, (void**)&pServiceProvider);
    if (FAILED(hr)) {
        std::cerr << "CoCreateInstance(ImmersiveShell) failed: " << std::hex << hr << std::endl;
        return;
    }

    hr = pServiceProvider->QueryService(CLSID_VirtualDesktopManagerInternal_24H2, IID_IVirtualDesktopManagerInternal_24H2, (void**)&pVDMInternal);
    if (FAILED(hr)) {
        std::cerr << "QueryService(IVirtualDesktopManagerInternal) failed. Check GUIDs. HR=" << std::hex << hr << std::endl;
        return;
    }
    std::cout << "Got IVirtualDesktopManagerInternal." << std::endl;

    hr = pServiceProvider->QueryService(IID_IApplicationViewCollection, IID_IApplicationViewCollection, (void**)&pAppViewCollection);
    if (FAILED(hr)) {
        std::cerr << "QueryService(IApplicationViewCollection) failed. HR=" << std::hex << hr << std::endl;
    } else {
        std::cout << "Got IApplicationViewCollection." << std::endl;
    }

    std::cout << "VirtualDesktopManager initialized successfully!" << std::endl;
}

// ----------------------------------------------------------------------------
// Desktop Management Logic
// ----------------------------------------------------------------------------

void MoveWindowToDesktop(int direction) {
    std::cout << "--- MoveWindowToDesktop (" << direction << ") ---" << std::endl;

    if (!pVDMInternal || !pAppViewCollection) {
        std::cerr << "Error: Interfaces not initialized." << std::endl;
        return;
    }

    HWND hwnd = GetForegroundWindow();
    if (!hwnd) {
        std::cerr << "Error: No foreground window." << std::endl;
        return;
    }
    std::cout << "Foreground Window HWND: " << std::hex << hwnd << std::dec << std::endl;

    IApplicationView* pView = nullptr;
    HRESULT hr = pAppViewCollection->GetViewForHwnd(hwnd, &pView);
    if (FAILED(hr)) {
        std::cerr << "GetViewForHwnd failed: " << std::hex << hr << std::endl;
        return;
    }

    IVirtualDesktop* pCurrentDesktop = nullptr;
    hr = pVDMInternal->GetCurrentDesktop(&pCurrentDesktop);
    if (FAILED(hr)) {
        std::cerr << "GetCurrentDesktop failed: " << std::hex << hr << std::endl;
        pView->Release();
        return;
    }
    std::cout << "Got Current Desktop." << std::endl;

    IVirtualDesktop* pAdjacentDesktop = nullptr;
    int dirCode = (direction == -1) ? VD_DIRECTION_LEFT : VD_DIRECTION_RIGHT;

    hr = pVDMInternal->GetAdjacentDesktop(pCurrentDesktop, dirCode, &pAdjacentDesktop);
    if (FAILED(hr) || !pAdjacentDesktop) {
        // Expected error when trying to move beyond the first/last desktop
        std::cerr << "GetAdjacentDesktop failed/none. HR=" << std::hex << hr << std::endl;
        pCurrentDesktop->Release();
        pView->Release();
        return;
    }
    std::cout << "Got Adjacent Desktop." << std::endl;

    hr = pVDMInternal->MoveViewToDesktop(pView, pAdjacentDesktop);
    if (SUCCEEDED(hr)) {
        std::cout << "SUCCESS: Window moved!" << std::endl;

        hr = pVDMInternal->SwitchDesktop(pAdjacentDesktop);
        if (SUCCEEDED(hr)) {
            std::cout << "SUCCESS: Switched to new desktop!" << std::endl;
        } else {
            std::cerr << "SwitchDesktop failed: " << std::hex << hr << std::endl;
        }
    } else {
        std::cerr << "MoveViewToDesktop failed: " << std::hex << hr << std::endl;
    }

    pAdjacentDesktop->Release();
    pCurrentDesktop->Release();
    pView->Release();
}

// ----------------------------------------------------------------------------
// Keyboard Hook
// ----------------------------------------------------------------------------

LRESULT CALLBACK LowLevelKeyboardProc(int nCode, WPARAM wParam, LPARAM lParam) {
    if (nCode == HC_ACTION) {
        if (wParam == WM_KEYDOWN || wParam == WM_SYSKEYDOWN) {
            KBDLLHOOKSTRUCT* p = (KBDLLHOOKSTRUCT*)lParam;

            if (p->vkCode == VK_LEFT || p->vkCode == VK_RIGHT) {
                bool shiftPressed = (GetAsyncKeyState(VK_SHIFT) & 0x8000);
                bool winPressed = (GetAsyncKeyState(VK_LWIN) & 0x8000) || (GetAsyncKeyState(VK_RWIN) & 0x8000);

                if (shiftPressed && winPressed) {
                    std::cout << "[HOOK] Win+Shift+" << (p->vkCode == VK_LEFT ? "LEFT" : "RIGHT") << " detected." << std::endl;

                    if (p->vkCode == VK_LEFT) {
                        PostThreadMessage(g_mainThreadId, WM_APP_MOVE_LEFT, 0, 0);
                    } else {
                        PostThreadMessage(g_mainThreadId, WM_APP_MOVE_RIGHT, 0, 0);
                    }
                    return 1; // Consume key to prevent Windows default behavior
                }
            }
        }
    }
    return CallNextHookEx(hKeyboardHook, nCode, wParam, lParam);
}

// ----------------------------------------------------------------------------
// Main Entry Point
// ----------------------------------------------------------------------------

int main() {
    g_mainThreadId = GetCurrentThreadId();
    std::cout << "Main Thread ID: " << g_mainThreadId << std::endl;

    InitializeVirtualDesktopManager();

    if (!pVDMInternal) {
        std::cerr << "Critical Error: Initialization failed. Exiting." << std::endl;
        return 1;
    }

    hKeyboardHook = SetWindowsHookEx(WH_KEYBOARD_LL, LowLevelKeyboardProc, GetModuleHandle(NULL), 0);
    if (!hKeyboardHook) {
        std::cerr << "Critical Error: Failed to install keyboard hook!" << std::endl;
        return 1;
    }

    std::cout << "Listening for Win + Shift + Left/Right..." << std::endl;
    std::cout << "Logs enabled (Press Ctrl+C to exit)." << std::endl;

    MSG msg;
    while (GetMessage(&msg, NULL, 0, 0)) {
        if (msg.message == WM_APP_MOVE_LEFT) {
            MoveWindowToDesktop(-1);
        } else if (msg.message == WM_APP_MOVE_RIGHT) {
            MoveWindowToDesktop(1);
        } else {
            TranslateMessage(&msg);
            DispatchMessage(&msg);
        }
    }

    UnhookWindowsHookEx(hKeyboardHook);
    CoUninitialize();
    return 0;
}
