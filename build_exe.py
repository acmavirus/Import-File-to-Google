import os
import subprocess
import sys
import shutil

def run_command(cmd):
    print(f"Running command: {' '.join(cmd)}")
    result = subprocess.run(cmd, capture_output=True, text=True)
    if result.returncode != 0:
        print(f"Error running command:")
        print(result.stdout)
        print(result.stderr)
        return False
    return True

def main():
    python_exe = sys.executable
    pyinstaller = os.path.join(os.path.dirname(python_exe), "Scripts", "pyinstaller.exe")
    if not os.path.exists(pyinstaller):
        # Fallback to run module
        pyinstaller_cmd = [python_exe, "-m", "PyInstaller"]
    else:
        pyinstaller_cmd = [pyinstaller]

    print("--- STEP 1: Cleaning previous build artifacts ---")
    for folder in ["build", "dist"]:
        if os.path.exists(folder):
            print(f"Removing {folder} folder...")
            shutil.rmtree(folder, ignore_errors=True)

    # 1. Build main application EXE
    print("\n--- STEP 2: Compiling core application (ImportToSheet.exe) ---")
    cmd_app = pyinstaller_cmd + [
        "--onefile",
        "--noconsole",
        "--name", "ImportToSheet",
        "--exclude-module", "sqlalchemy",
        "src/import_to_sheet.py"
    ]
    if not run_command(cmd_app):
        print("Failed to compile core application.")
        sys.exit(1)

    app_exe = os.path.join("dist", "ImportToSheet.exe")
    if not os.path.exists(app_exe):
        print(f"Error: Compiled core executable not found at {app_exe}")
        sys.exit(1)

    print(f"Successfully compiled: {app_exe}")

    # 2. Build Installer wizard EXE containing the core EXE as data resource
    print("\n--- STEP 3: Compiling graphical Installer Wizard (Setup_ImportToSheet.exe) ---")
    cmd_setup = pyinstaller_cmd + [
        "--onefile",
        "--noconsole",
        "--add-data", f"{app_exe};.",
        "--name", "Setup_ImportToSheet",
        "--exclude-module", "sqlalchemy",
        "src/setup_wizard.py"
    ]
    if not run_command(cmd_setup):
        print("Failed to compile Installer Wizard.")
        sys.exit(1)

    setup_exe = os.path.join("dist", "Setup_ImportToSheet.exe")
    if not os.path.exists(setup_exe):
        print(f"Error: Compiled Setup executable not found at {setup_exe}")
        sys.exit(1)

    print(f"Successfully compiled Setup wizard: {setup_exe}")
    print("\n=======================================================")
    print(" BUILD SUCCESSFUL! ")
    print(f" Distributed setup file is ready at: {os.path.abspath(setup_exe)}")
    print("=======================================================")

if __name__ == "__main__":
    main()
