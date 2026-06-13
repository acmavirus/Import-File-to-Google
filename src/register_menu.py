import os
import sys
import winreg
import argparse

# Extensions to add to the context menu
EXTENSIONS = ['.docx', '.xlsx', '.xls', '.csv']
MENU_NAME = "ImportToSheet"
MENU_TEXT = "Import to Google Sheets"

def register(exe_path=None, script_path=None):
    python_exe = sys.executable
    
    if exe_path:
        # Register the compiled executable directly
        command_str = f'"{os.path.abspath(exe_path)}" "%1"'
        icon_path = os.path.abspath(exe_path)
    else:
        # Dev mode: register python script
        if not script_path:
            script_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "import_to_sheet.py")
        command_str = f'"{python_exe}" "{os.path.abspath(script_path)}" "%1"'
        icon_path = python_exe
        
    print(f"Registering context menu action:")
    print(f"Command: {command_str}")
    
    success_count = 0
    for ext in EXTENSIONS:
        sub_key_path = f"Software\\Classes\\SystemFileAssociations\\{ext}\\shell\\{MENU_NAME}"
        
        try:
            # Create shell key
            key = winreg.CreateKey(winreg.HKEY_CURRENT_USER, sub_key_path)
            winreg.SetValueEx(key, "", 0, winreg.REG_SZ, MENU_TEXT)
            
            # Set Icon
            winreg.SetValueEx(key, "Icon", 0, winreg.REG_SZ, icon_path)
            winreg.CloseKey(key)
            
            # Create command key
            command_key_path = f"{sub_key_path}\\command"
            cmd_key = winreg.CreateKey(winreg.HKEY_CURRENT_USER, command_key_path)
            winreg.SetValueEx(cmd_key, "", 0, winreg.REG_SZ, command_str)
            winreg.CloseKey(cmd_key)
            
            print(f"Successfully registered for {ext}")
            success_count += 1
        except Exception as e:
            print(f"Failed to register for {ext}: {e}")
            
    return success_count == len(EXTENSIONS)

def unregister():
    print("Unregistering context menu action...")
    for ext in EXTENSIONS:
        sub_key_path = f"Software\\Classes\\SystemFileAssociations\\{ext}\\shell\\{MENU_NAME}"
        try:
            try:
                winreg.DeleteKey(winreg.HKEY_CURRENT_USER, f"{sub_key_path}\\command")
            except FileNotFoundError:
                pass
            winreg.DeleteKey(winreg.HKEY_CURRENT_USER, sub_key_path)
            print(f"Successfully unregistered for {ext}")
        except FileNotFoundError:
            pass
        except Exception as e:
            print(f"Failed to unregister for {ext}: {e}")
    print("Context menu unregistered.")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Register/Unregister Windows Explorer Context Menu for Import to Sheets")
    parser.add_argument("--register", action="store_true", help="Register context menu")
    parser.add_argument("--unregister", action="store_true", help="Unregister context menu")
    parser.add_argument("--exe", type=str, help="Path to compiled executable")
    parser.add_argument("--script", type=str, help="Path to python script")
    
    args = parser.parse_args()
    
    if args.unregister:
        unregister()
    else:
        register(exe_path=args.exe, script_path=args.script)
