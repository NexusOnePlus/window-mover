#include <windows.h>
#include <shellapi.h>
#include <string>
#include <winstring.h>
#include "VirtualDesktopInterfaces.h"

#pragma comment(lib, "runtimeobject.lib")

#define WM_APP_TRAYMSG         (WM_APP + 100)
#define ID_TRAY_ICON           1
#define ID_MENU_EXIT           2001
#define ID_MENU_HELP           2002

#define WM_APP_MOVE_LEFT       (WM_APP + 1)
#define WM_APP_MOVE_RIGHT      (WM_APP + 2)
#define WM_APP_SWITCH_TO_INDEX (WM_APP + 3)
#define WM_APP_MOVE_TO_INDEX   (WM_APP + 4)

#define VD_DIRECTION_LEFT 3
#define VD_DIRECTION_RIGHT 4

IVirtualDesktopManagerInternal* pVDMInternal = nullptr;
IApplicationViewCollection* pAppViewCollection = nullptr;
IServiceProvider* pServiceProvider = nullptr;
HHOOK hKeyboardHook = NULL;
HWND hHiddenWindow = NULL;
NOTIFYICONDATA nid = {};

// ----------------------------------------------------------------------------
// Initialization
// ----------------------------------------------------------------------------

bool InitializeVirtualDesktopManager() {
    HRESULT hr = CoInitialize(NULL);
    if (FAILED(hr)) return false;

    hr = CoCreateInstance(CLSID_ImmersiveShell, nullptr, CLSCTX_LOCAL_SERVER, IID_IServiceProvider, (void**)&pServiceProvider);
    if (FAILED(hr)) return false;

    hr = pServiceProvider->QueryService(CLSID_VirtualDesktopManagerInternal_24H2, IID_IVirtualDesktopManagerInternal_24H2, (void**)&pVDMInternal);
    if (FAILED(hr)) return false;

    hr = pServiceProvider->QueryService(IID_IApplicationViewCollection, IID_IApplicationViewCollection, (void**)&pAppViewCollection);
    
    return true;
}

// ----------------------------------------------------------------------------
// Helpers
// ----------------------------------------------------------------------------

UINT EnsureDesktopCount(UINT targetIndex) {
    if (!pVDMInternal) return 0;

    IObjectArray* pDesktops = nullptr;
    HRESULT hr = pVDMInternal->GetDesktops(&pDesktops);
    if (FAILED(hr)) return 0;

    UINT count = 0;
    pDesktops->GetCount(&count);
    pDesktops->Release();

    while (count <= targetIndex) {
        IVirtualDesktop* pNewDesktop = nullptr;
        hr = pVDMInternal->CreateDesktop(&pNewDesktop);
        if (FAILED(hr)) break;
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
    pDesktops->GetAt(index, IID_IUnknown, (void**)&pDesktop);
    pDesktops->Release();
    return pDesktop;
}

// ----------------------------------------------------------------------------
// Desktop Logic
// ----------------------------------------------------------------------------

void SwitchToDesktopAtIndex(int targetIndex) {
    if (!pVDMInternal) return;
    EnsureDesktopCount(targetIndex);
    IVirtualDesktop* pTarget = GetDesktopAtIndex(targetIndex);
    if (pTarget) {
        pVDMInternal->SwitchDesktop(pTarget);
        pTarget->Release();
    }
}

void MoveWindowToDesktopAtIndex(int targetIndex) {
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
    }
    pView->Release();
}

void MoveWindowRelative(int direction) {
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

    if ((FAILED(hr) || !pAdjacent) && direction == 1) {
        IVirtualDesktop* pNew = nullptr;
        if (SUCCEEDED(pVDMInternal->CreateDesktop(&pNew))) {
            pAdjacent = pNew;
        }
    }

    if (pAdjacent) {
        pVDMInternal->MoveViewToDesktop(pView, pAdjacent);
        pVDMInternal->SwitchDesktop(pAdjacent);
        pAdjacent->Release();
    }

    pCurrent->Release();
    pView->Release();
}

// ----------------------------------------------------------------------------
// Keyboard Hook
// ----------------------------------------------------------------------------

LRESULT CALLBACK LowLevelKeyboardProc(int nCode, WPARAM wParam, LPARAM lParam) {
    if (nCode == HC_ACTION) {
        if (wParam == WM_KEYDOWN || wParam == WM_SYSKEYDOWN) {
            KBDLLHOOKSTRUCT* p = (KBDLLHOOKSTRUCT*)lParam;
            
            bool win = (GetAsyncKeyState(VK_LWIN) & 0x8000) || (GetAsyncKeyState(VK_RWIN) & 0x8000);
            bool shift = (GetAsyncKeyState(VK_SHIFT) & 0x8000);
            bool alt = (GetAsyncKeyState(VK_MENU) & 0x8000);

            if (win) {
                if (shift && !alt) {
                    if (p->vkCode == VK_LEFT) {
                        PostMessage(hHiddenWindow, WM_APP_MOVE_LEFT, 0, 0);
                        return 1;
                    }
                    if (p->vkCode == VK_RIGHT) {
                        PostMessage(hHiddenWindow, WM_APP_MOVE_RIGHT, 0, 0);
                        return 1;
                    }
                }
                else if (!shift && !alt && p->vkCode >= '1' && p->vkCode <= '9') {
                    int index = p->vkCode - '1';
                    PostMessage(hHiddenWindow, WM_APP_SWITCH_TO_INDEX, (WPARAM)index, 0);
                    return 1; 
                }
                else if (!shift && alt && p->vkCode >= '1' && p->vkCode <= '9') {
                    int index = p->vkCode - '1';
                    PostMessage(hHiddenWindow, WM_APP_MOVE_TO_INDEX, (WPARAM)index, 0);
                    return 1;
                }
            }
        }
    }
    return CallNextHookEx(hKeyboardHook, nCode, wParam, lParam);
}

// ----------------------------------------------------------------------------
// Window Procedure
// ----------------------------------------------------------------------------

LRESULT CALLBACK WndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
    switch (msg) {
        case WM_APP_TRAYMSG:
            if (lParam == WM_RBUTTONUP) {
                POINT pt;
                GetCursorPos(&pt);
                HMENU hMenu = CreatePopupMenu();
                AppendMenu(hMenu, MF_STRING, ID_MENU_HELP, "Shortcuts Help");
                AppendMenu(hMenu, MF_SEPARATOR, 0, NULL);
                AppendMenu(hMenu, MF_STRING, ID_MENU_EXIT, "Exit");
                
                SetForegroundWindow(hwnd);
                int cmd = TrackPopupMenu(hMenu, TPM_RETURNCMD | TPM_NONOTIFY, pt.x, pt.y, 0, hwnd, NULL);
                
                if (cmd == ID_MENU_EXIT) {
                    DestroyWindow(hwnd);
                } else if (cmd == ID_MENU_HELP) {
                    MessageBox(NULL, 
                        "Win + 1..9 : Switch Desktop\n"
                        "Win + Alt + 1..9 : Move Window & Switch\n"
                        "Win + Shift + Arrows : Move Window Relative", 
                        "Desktop Mover", MB_OK | MB_ICONINFORMATION);
                }
                DestroyMenu(hMenu);
            }
            break;

        case WM_APP_MOVE_LEFT:       MoveWindowRelative(-1); break;
        case WM_APP_MOVE_RIGHT:      MoveWindowRelative(1); break;
        case WM_APP_SWITCH_TO_INDEX: SwitchToDesktopAtIndex((int)wParam); break;
        case WM_APP_MOVE_TO_INDEX:   MoveWindowToDesktopAtIndex((int)wParam); break;

        case WM_DESTROY:
            PostQuitMessage(0);
            break;
        
        default:
            return DefWindowProc(hwnd, msg, wParam, lParam);
    }
    return 0;
}

// ----------------------------------------------------------------------------
// Entry Point
// ----------------------------------------------------------------------------

int WINAPI WinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, LPSTR lpCmdLine, int nCmdShow) {
    CreateMutex(NULL, TRUE, "DesktopMover_Mutex_App");
    if (GetLastError() == ERROR_ALREADY_EXISTS) {
        MessageBox(NULL, "Desktop Mover is already running.", "Error", MB_OK | MB_ICONERROR);
        return 1;
    }

    if (!InitializeVirtualDesktopManager()) {
        MessageBox(NULL, "Failed to initialize Virtual Desktop Manager.\nEnsure Windows 11 (24H2).", "Error", MB_OK | MB_ICONERROR);
        return 1;
    }

    WNDCLASS wc = {0};
    wc.lpfnWndProc = WndProc;
    wc.hInstance = hInstance;
    wc.lpszClassName = "DesktopMoverTrayClass";
    RegisterClass(&wc);

    hHiddenWindow = CreateWindow(wc.lpszClassName, "Desktop Mover", 0, 0, 0, 0, 0, NULL, 0, hInstance, 0);

    nid.cbSize = sizeof(NOTIFYICONDATA);
    nid.hWnd = hHiddenWindow;
    nid.uID = ID_TRAY_ICON;
    nid.uFlags = NIF_ICON | NIF_MESSAGE | NIF_TIP;
    nid.uCallbackMessage = WM_APP_TRAYMSG;
    nid.hIcon = LoadIcon(NULL, IDI_APPLICATION);
    strcpy(nid.szTip, "Desktop Mover");

    Shell_NotifyIcon(NIM_ADD, &nid);

    hKeyboardHook = SetWindowsHookEx(WH_KEYBOARD_LL, LowLevelKeyboardProc, GetModuleHandle(NULL), 0);
    if (!hKeyboardHook) {
        MessageBox(NULL, "Failed to install keyboard hook.", "Error", MB_OK | MB_ICONERROR);
        return 1;
    }

    MSG msg;
    while (GetMessage(&msg, NULL, 0, 0)) {
        TranslateMessage(&msg);
        DispatchMessage(&msg);
    }

    Shell_NotifyIcon(NIM_DELETE, &nid);
    UnhookWindowsHookEx(hKeyboardHook);
    CoUninitialize();

    return 0;
}