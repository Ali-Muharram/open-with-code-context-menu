import winreg
import os
import shutil

def get_vscode_path():
    # 1. Try to find the path from the registry
    try:
        key = winreg.OpenKey(
            winreg.HKEY_LOCAL_MACHINE,
            r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\Code.exe"
        )
        value, _ = winreg.QueryValueEx(key, "")
        winreg.CloseKey(key)
        if os.path.exists(value):
            return value
    except FileNotFoundError:
        pass

    # 2. Search in common installation paths
    common_paths = [
        os.path.expanduser(r"~\AppData\Local\Programs\Microsoft VS Code\Code.exe"),
        r"C:\Program Files\Microsoft VS Code\Code.exe",
        r"C:\Program Files (x86)\Microsoft VS Code\Code.exe"
    ]
    for path in common_paths:
        if os.path.exists(path):
            return path

    # 3. Search in PATH environment variable
    code_path = shutil.which("code")
    if code_path:
        possible_exe = os.path.join(os.path.dirname(code_path), "Code.exe")
        if os.path.exists(possible_exe):
            return possible_exe

    return None

def add_registry_key(root, path, name, value):
    try:
        key = winreg.CreateKey(root, path)
        winreg.SetValueEx(key, name, 0, winreg.REG_SZ, value)
        winreg.CloseKey(key)
        print(f"[✔] Added key: {path} -> {name} = {value}")
    except Exception as e:
        print(f"[✘] Error in {path}: {e}")

vscode_path = get_vscode_path()
if vscode_path and os.path.exists(vscode_path):
    print(f"[✔] Found VS Code at: {vscode_path}")

    # Add "Open with Code" when right-clicking on folder background
    add_registry_key(
        winreg.HKEY_CLASSES_ROOT,
        r"Directory\Background\shell\OpenWithCode",
        "",
        "Open with Code"
    )
    add_registry_key(
        winreg.HKEY_CLASSES_ROOT,
        r"Directory\Background\shell\OpenWithCode",
        "Icon",
        f"\"{vscode_path}\""
    )
    add_registry_key(
        winreg.HKEY_CLASSES_ROOT,
        r"Directory\Background\shell\OpenWithCode\command",
        "",
        f"\"{vscode_path}\" \"%V\""
    )

    # Add "Open with Code" when right-clicking on a folder
    add_registry_key(
        winreg.HKEY_CLASSES_ROOT,
        r"Directory\shell\OpenWithCode",
        "",
        "Open with Code"
    )
    add_registry_key(
        winreg.HKEY_CLASSES_ROOT,
        r"Directory\shell\OpenWithCode",
        "Icon",
        f"\"{vscode_path}\""
    )
    add_registry_key(
        winreg.HKEY_CLASSES_ROOT,
        r"Directory\shell\OpenWithCode\command",
        "",
        f"\"{vscode_path}\" \"%1\""
    )

    print("\n✅ Registry updated successfully!")
else:
    print("[✘] Could not find VS Code on this system.")
