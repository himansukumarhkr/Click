import subprocess
import sys
import os

def install_pyinstaller():
    try:
        import PyInstaller
        print("PyInstaller is already installed.")
    except ImportError:
        print("PyInstaller not found. Installing...")
        try:
            subprocess.check_call([sys.executable, "-m", "pip", "install", "pyinstaller"])
            print("PyInstaller installed successfully.")
        except subprocess.CalledProcessError as e:
            print(f"Error: Failed to install PyInstaller. {e}")
            sys.exit(1)

def build_executable():
    print("Starting the build process for Windows...")

    # --- PyInstaller Command Configuration ---
    script_name = os.path.join("src", "main.py")
    exe_name = "Click!"
    icon_path = os.path.join("src", "assets", "app_icon.ico")
    splash_path = os.path.join("src", "assets", "splash.png")

    assets_path = f"src{os.path.sep}assets;src{os.path.sep}assets"

    command = [
        sys.executable,
        "-m", "PyInstaller",
        "--noconfirm",
        "--onefile",
        "--windowed",
        f"--name={exe_name}",
        f"--icon={icon_path}",
        f"--splash={splash_path}",
        f"--add-data={assets_path}",
        "--paths=.",
        "--hidden-import=PIL._tkinter_finder",
        "--hidden-import=customtkinter",
        "--collect-all=customtkinter",
        script_name
    ]

    print("\nRunning PyInstaller with the following command:")
    print(" ".join(f'"{arg}"' if " " in arg else arg for arg in command))
    print("\n" + "="*50)

    try:
        subprocess.check_call(command)
        print("="*50)
        print("\nBuild successful!")
        print(f"Executable created at: {os.path.join('dist', f'{exe_name}.exe')}")
    except subprocess.CalledProcessError as e:
        print("="*50)
        print(f"\nError: Build failed with exit code {e.returncode}.")
        print("Please check the output above for more details.")
        sys.exit(1)
    except FileNotFoundError:
        print("\nError: PyInstaller command not found. Make sure it's installed and in your PATH.")
        sys.exit(1)

if __name__ == "__main__":
    install_pyinstaller()
    build_executable()
