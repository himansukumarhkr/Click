Click! 📸
Open-Source Evidence Capture Tool for Professionals

Author: Himansu Kumar
Version: v1.0
Release Date: 31 January 2026
License: Open Source (MIT / Apache-2.0 / GPL-3.0)

Click! is a high-performance, open-source Windows utility designed for Software Testers (QA), Developers, Business Analysts, and Technical Professionals who need to capture screen evidence instantly without breaking workflow.

Click! transforms manual documentation from hours into minutes using global hotkeys, instant persistence, and intelligent organization.

🚀 Why Click!?
Feature	Benefit
Zero-Interruption Workflow	Capture runs silently in background
Instant Persistence	Saved immediately — crash-safe
Smart Organization	Timestamped files + window titles
Portable	No installation required
Open Source	Fully auditable, community-driven
⌨️ Command Center (Global Hotkeys)
Action	Key Combination	Description
Capture	~ (Tilde)	Capture screen + clipboard into DOCX
Undo	Ctrl + Alt + ~	Delete last capture from disk

✔ Works system-wide, even when Click! is not focused.

🛠 Features at a Glance
✅ DUAL OUTPUT MODES
   ├── 📄 Microsoft Word (.docx)
   └── 🖼️ JPEG Image Archive

✅ SMART FILE MANAGEMENT
   ├── Timestamp-based auto-naming
   ├── Active window title logging
   └── Automatic file rotation

✅ WORKFLOW FRIENDLY
   ├── Clipboard auto-copy
   ├── Modern dark-mode GUI
   └── Scrollable capture history

✅ ENGINEERED FOR PROFESSIONAL USE
   ├── DPI-aware (4K / High-DPI)
   ├── Multi-keyboard layout support
   ├── Native Windows API integration
   └── Low-latency capture pipeline

📥 Getting Started (Under 1 Minute)
1️⃣ Clone or download the repository
2️⃣ Run build.py
3️⃣ Run Click!.exe OR start from source
4️⃣ Choose output folder (default provided)
5️⃣ Press ~ to Capture screenshots.

Filename Format

DD-MM-YYYY.docx

🖥️ GUI Overview
┌─────────────────────────────────────────┐
│  Click! — Evidence Capture Tool        │
├─────────────────────────────────────────┤
│  [Scrollable Capture History]          │
│  ├── Click_Capture_14-29-15.docx       │
│  │   [Open] [Copy Path]                │
│  ├── Click_Capture_14-28-42.docx       │
│  └── [More captures...]                │
│                                         │
│  📁 Open Folder | 📋 Copy Path | 🗑️ Clear│
└─────────────────────────────────────────┘

🔒 Privacy & Security
✅ Yes	❌ Never
Fully offline	No cloud uploads
No telemetry	No tracking
Local processing	No background services
Transparent source	No hidden behavior

🔍 All behavior is auditable via source code.

🎯 Ideal For
👨‍💻 Software Testers (QA)
🔧 Developers & Engineers
📊 Business Analysts
📱 Support Engineers
📈 Project Managers
🎓 Trainers & Educators


If you document bugs, workflows, processes, or tutorials, Click! fits perfectly.

💻 System Requirements
Component	Requirement
OS	Windows 10 / 11 (64-bit)
RAM	~100 MB
Disk	~50 MB
Viewer	Word or LibreOffice
Display	Any (DPI-aware)
Admin Rights	Not required

🧩 Build From Source
pip install -r requirements.txt
python main.py

Output:

/dist/Click!.exe

🛠 Technical Architecture
CORE COMPONENTS
├── GUI Framework: CustomTkinter
├── Screen Capture: Pillow (ImageGrab)
├── Document Export: python-docx
├── Hotkeys: Windows API (ctypes)
├── Compiler: PyInstaller
└── DPI Handling: Per-monitor DPI awareness

DEPENDENCIES
├── Python 3.x
├── customtkinter
├── Pillow
├── python-docx
├── ctypes (kernel32, user32)
└── threading (async hotkey listener)

📈 Performance Characteristics
├── Startup time: ~2–3 seconds
├── Capture latency: <500ms
├── Memory usage: ~80–120MB
├── Disk writes: Immediate
└── Packaging: PyInstaller bootloader

📦 Repository Structure
/click
 ├── src/
 ├── assets/
 ├── requirements.txt
 ├── README.md
 ├── LICENSE
 └── build/

📈 Version History
v1.0.0 (31 Jan 2026)
├── Global hotkeys
├── DOCX export
├── DPI awareness
├── Clipboard integration
├── Multi-keyboard support
└── PyInstaller standalone build

🤝 Contributing

Contributions are welcome 🎉

✔ Bug reports
✔ Feature requests
✔ Performance improvements
✔ UI/UX enhancements
✔ Documentation updates
Please open an issue or submit a pull request.

Made in India

Developed by: Himansu Kumar
Technologies: Python, CustomTkinter, Windows API, PyInstaller

🚀 Quick Reference
~              → Capture
Ctrl + Alt + ~ → Undo last capture

Output: Desktop\Evidence
Formats: DOCX + JPEG


Click! — Evidence, captured instantly.
Open source. Transparent. Built for professionals.
