import PyInstaller.__main__
import os

def build():
    script_name = "ScreenshotTool.py"
    icon_name = "icon.ico"
    
    args = [
        script_name,
        "--onefile",
        "--noconsole",
        "--name=ScreenshotTool",
        "--clean",
    ]
    
    if os.path.exists(icon_name):
        args.append(f"--icon={icon_name}")
        print(f"Icon found: {icon_name}")
    else:
        print(f"Warning: {icon_name} not found. Using default icon.")

    if os.path.exists("config.toon"):
        args.append("--add-data=config.toon;.")

    print("Building executable...")
    PyInstaller.__main__.run(args)
    print("Build complete! Check the 'dist' folder.")

if __name__ == "__main__":
    build()