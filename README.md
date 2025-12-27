# Window Mover (Windows 11 24H2)

A lightweight C++ tool for managing virtual desktops via custom keyboard shortcuts. It automatically creates desktops when you attempt to access a non-existent one.

## Keyboard Shortcuts

| Shortcut | Action |
|----------|--------|
| `Win + 1..9` | Switch to the corresponding desktop. Creates it if it doesn't exist. |
| `Win + Alt + 1..9` | Move the active window to the chosen desktop and switch to it. |
| `Win + Shift + →` | Move window to the right. Creates a new desktop if on the last one. |
| `Win + Shift + ←` | Move window to the left. |

## Compilation Requirements

Requires **MinGW-w64** (g++) with Windows library support.

### Compilation Command

Run the following command in your terminal from the project root:

```bash
g++ -o DesktopMover.exe DesktopMover.cpp -lole32 -loleaut32 -lruntimeobject -luuid -mwindows -static
```

### Libraries Used
- `ole32` & `oleaut32`: For COM object handling.
- `runtimeobject`: Required for Windows Runtime interfaces.
- `uuid`: Definitions for interface GUIDs.
- `-static`: Ensures the executable is standalone and doesn't depend on GCC DLLs.

## Technical Notes
This project uses internal Windows interfaces (`IVirtualDesktopManagerInternal`). These interfaces are undocumented and subject to change in future Windows updates. It is currently configured for **Windows 11 Build 26200 (24H2)**.
