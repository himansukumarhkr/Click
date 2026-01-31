# Click! 📸
# The Ultimate Evidence Capture Tool for Professionals

**Author:** Himansu Kumar  
**Version:** v1.0  
**Release Date:** 31 January 2026  
**License:** Proprietary Freeware (see `CLICK_EULA.txt`)

---

Click! is a **high-performance, standalone Windows utility** designed for **Software Testers (QA), Developers, and Business Analysts** who need to **document screen activity rapidly**. Click! turns hours of manual documentation into **minutes of effortless capture**.

---

## 🚀 Why Use Click!?

| Feature | Benefit |
|---------|---------|
| **Zero-Interruption Workflow** | No window switching - capture in background while you work |
| **Instant Persistence** | Saves immediately to disk - survives system crashes |
| **Intelligent Organization** | Auto-naming + active window title capture |
| **One-File Portability** | Single `.exe` - no installation required |
| **Memory Efficient** | Nuitka-compiled Python - won't lag your work apps |

---

## ⌨️ Command Center (Hotkeys)

| **Action** | **Key Combination** | **Description** |
|------------|-------------------|-----------------|
| **Capture** | `~` (Tilde) | Captures screen + clipboard → instant DOCX save |
| **Undo** | `Ctrl + Alt + ~` | Deletes last capture from disk completely |

**Works globally** - even when Click! isn't focused!

---

## 🛠 Features at a Glance

```
✅ DUAL OUTPUT MODES
   ├── 📄 Microsoft Word (.docx) - Professional reports
   └── 🖼️ JPEG Folder - Raw image archives

✅ SMART FILE MANAGEMENT
   ├── Auto-naming: Click_Capture_2026-01-31_14-29-15.docx
   ├── Context aware: Logs active window title
   └── Auto-rotate: Creates Part 2 when file limit reached

✅ WORKFLOW INTEGRATION
   ├── Auto-copy to clipboard (Teams/Slack/Jira ready)
   ├── Modern Dark Mode GUI
   └── Scrollable capture history

✅ PROFESSIONAL FEATURES
   ├── DPI-aware (4K/high-DPI displays)
   ├── Multi-keyboard layout support
   ├── Windows API integration
   └── Efficient memory management
```

---

## 📥 Getting Started (30 seconds)

```
1️⃣ DOWNLOAD: Grab Click!.exe from Releases
2️⃣ LAUNCH: Double-click (no install needed)
3️⃣ SETUP: Choose Documents/Click! Captures/
4️⃣ GO: Press ~ anywhere to start capturing!
```

**Default Output Location:** `C:\Users\[YourName]\Documents\Click! Captures\`

**Filename Format:** `Click_Capture_2026-01-31_14-29-15.docx`

---

## 🖥️ GUI Overview

```
┌─────────────────────────────────────────┐
│  Click! - Evidence Capture Tool        │
├─────────────────────────────────────────┤
│  [Scrollable Capture History]          │
│  ├── Click_Capture_14-29-15.docx       │
│  │   [Open] [Copy Path]                │
│  ├── Click_Capture_14-28-42.docx       │
│  │   [Open] [Copy Path]                │
│  └── [More captures...]                │
│                                         │
│  [Buttons]                              │
│  📁 Open Folder | 📋 Copy Path | 🗑️ Clear│
└─────────────────────────────────────────┘
```

---

## 🔒 Privacy & Security

| ✅ **Safe** | ❌ **Never** |
|------------|-------------|
| **100% Offline** - Local storage only | No cloud uploads |
| **No Telemetry** - Zero tracking | No usage analytics |
| **Clean Exit** - Auto cleanup | No persistent data |
| **Local Processing** - Your machine only | No external servers |

---

## 🎯 Perfect For

```
👨‍💻 Software Testers (QA)
   └── Bug reproduction with visual evidence

🔧 Developers
   └── Code review documentation

📊 Business Analysts
   └── Process capture and workflow mapping

📱 Support Engineers
   └── Issue replication and troubleshooting

📈 Project Managers
   └── Meeting screenshots and status updates

🎓 Trainers & Educators
   └── Tutorial creation and step-by-step guides
```

---

## 💻 System Requirements

| Component | Minimum Specification |
|-----------|----------------------|
| **Operating System** | Windows 10/11 (64-bit) |
| **RAM** | 100MB free memory |
| **Disk Space** | 50MB |
| **Document Viewer** | Microsoft Word or LibreOffice |
| **Display** | Any (DPI-aware) |
| **Admin Rights** | Not required |

---

## 🔧 Troubleshooting

| **Issue** | **Solution** |
|-----------|-------------|
| **Hotkeys don't work** | • Run as Administrator once<br>• Close Snipping Tool or similar apps<br>• Check antivirus hasn't blocked it |
| **No DOCX files created** | • Check `Documents\Click! Captures\` folder<br>• Verify Word/LibreOffice is installed<br>• Check folder write permissions |
| **Blurry on 4K/high-DPI** | Already handled - DPI-aware by design |
| **Antivirus alert** | Add Click!.exe to antivirus exceptions |
| **Non-US keyboard** | Auto-detects keyboard layout (OEM_3 keycode) |

---

## 📱 Distribution Package

Your complete Click! package should include:

```
Click!.exe           ← Main application
README.md           ← This documentation
CLICK_EULA.txt      ← Legal terms (REQUIRED - READ BEFORE USE)
icon.ico            ← Application icon (optional)
```

---

## ⚖️ License Summary

**✅ FREE TO USE FOREVER**
- Personal use - unlimited
- Professional/commercial use - unlimited
- Install on multiple machines - OK

**❌ STRICTLY PROHIBITED:**
- Reverse engineering / decompiling the software
- Modifying or creating derivative works
- Redistributing (even for free)
- Removing copyright notices
- Developing competing products

**⚠️ DISCLAIMER:**
- Provided "AS IS" without warranty
- Use at your own risk
- Author not liable for damages

**📄 Full Legal Terms:** See `CLICK_EULA.txt` (mandatory reading before use)

---

## 🛠️ Technical Architecture

```
CORE COMPONENTS
├── GUI Framework: CustomTkinter + Tkinter Canvas
├── Image Capture: PIL/Pillow ImageGrab
├── Document Export: python-docx
├── Hotkey System: Windows API (ctypes)
├── Compiler: Nuitka (standalone executable)
└── DPI Handling: Per-monitor DPI awareness

LIBRARIES & DEPENDENCIES
├── Python 3.x (embedded)
├── customtkinter
├── Pillow (PIL)
├── python-docx
├── Windows ctypes (kernel32, user32)
└── Threading for async hotkey listener

PERFORMANCE
├── Memory footprint: ~50-100MB
├── Startup time: <2 seconds
├── Capture latency: <500ms
└── File save: Instant (async)
```

---

## 📈 Version History

```
v1.0.0 (31 January 2026) ✨
├── ✅ Global hotkey listener (~ & Ctrl+Alt+~)
├── ✅ Auto-scrolling GUI with capture history
├── ✅ DOCX export with embedded screenshots
├── ✅ Multi-keyboard layout support
├── ✅ DPI awareness (4K/high-DPI)
├── ✅ Nuitka compilation (standalone)
├── ✅ Output folder management
├── ✅ Clipboard auto-copy integration
└── ✅ Active window title capture
```

---

## 🆘 Support & Contact

- **Issues:** Report bugs via email or issue tracker
- **Feature Requests:** Contact author
- **Documentation:** This README + in-app help
- **No Warranty:** Use at your own risk

**Email:** [Your contact email]  
**Website:** [Your website/GitHub/GitLab]

---

## 🌟 Use Cases & Workflows

### Software Testing (QA)
```
1. Open application under test
2. Press ~ to capture each test step
3. Document automatically created with:
   - Screenshots of each step
   - Active window titles
   - Timestamps
4. Submit DOCX as bug report attachment
```

### Developer Documentation
```
1. Write code
2. Press ~ at key implementation points
3. Captures code + comments + results
4. Generate technical documentation automatically
```

### Business Process Mapping
```
1. Perform business process
2. Press ~ at each decision point
3. Auto-generates visual process flow
4. Share DOCX with stakeholders
```

---

## 🎉 Made in Mumbai, India 🇮🇳

**Developed by:** Himansu Kumar  
**Release Date:** 31 January 2026  
**Technologies:** Python, CustomTkinter, Nuitka, Windows API  
**License:** Proprietary Freeware

```
       ,     ,
      (\____/)
       (_oo_)
         (O)
       __||__    \)
    []/______\[] /
    / \______/ \/
   /    /__\ 
  /\   /____\ 
```

---

## 🚀 Quick Reference Card

**HOTKEYS**
```
~              → Capture screen + clipboard
Ctrl+Alt+~     → Undo last capture
```

**OUTPUT**
```
Location: Documents\Click! Captures\
Format:   Click_Capture_YYYY-MM-DD_HH-MM-SS.docx
Contains: Screenshot + Clipboard + Timestamp
```

**GUI BUTTONS**
```
📁 Open Folder    → Navigate to captures folder
📋 Copy Path      → Copy latest file path
🗑️ Clear History  → Clear GUI log (files remain)
```

---

**Click! - Your evidence, captured instantly** 🚀

*Turn documentation into automation. Turn hours into minutes.*

---

*For complete legal terms and conditions, see `CLICK_EULA.txt` included with this software.*