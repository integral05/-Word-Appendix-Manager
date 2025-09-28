# 📄 Word Appendix Manager

<div align="center">
  
![License](https://img.shields.io/badge/license-MIT-blue.svg)
![Python](https://img.shields.io/badge/python-3.8%2B-blue.svg)
![Platform](https://img.shields.io/badge/platform-Windows-lightgrey.svg)
![Status](https://img.shields.io/badge/status-Active-green.svg)

**Professional PDF Appendix Management Tool for Microsoft Word Documents**

*Seamlessly integrate PDF files as formatted appendices in Word documents with automated numbering, table of contents, and professional styling.*

[🚀 Quick Start](#-quick-start) • [📋 Features](#-features) • [🔧 Installation](#-installation) • [📖 Usage](#-usage) • [🏗️ Architecture](#️-project-architecture)

</div>

---

## 🎯 Project Overview

Word Appendix Manager is a **Windows desktop application** built with Python and PySide6 that automates the process of adding PDF files as appendices to Microsoft Word documents. It eliminates the manual, time-consuming task of formatting and inserting multiple PDF documents while maintaining professional document standards.

### ✨ What it Does:
- 📎 **Converts PDF pages** to high-quality images
- 📄 **Embeds images** directly into Word documents
- 🔢 **Auto-numbers** appendices (A, B, C... or 1, 2, 3...)
- 📑 **Updates** table of contents automatically
- 🎨 **Maintains** professional formatting and styling
- 💾 **Creates** backup copies before processing

### 🎮 Interactive Demo

<details>
<summary>👆 <strong>Click to see application screenshots</strong></summary>

```
┌─────────────────────────────────────────────────────────┐
│  🎯 Word Appendix Manager v1.0                         │
├─────────────────────────────────────────────────────────┤
│                                                         │
│  📄 Document Selection                                  │
│  ┌─────────────────────────────────────────────────┐   │
│  │ ◉ Open Documents  ○ Browse File                │   │
│  │ [Select Document ▼] [🔄 Refresh]              │   │
│  └─────────────────────────────────────────────────┘   │
│                                                         │
│  📎 Add PDF Files                    🔍 PDF Preview     │
│  ┌─────────────────────┐            ┌─────────────────┐ │
│  │                     │            │                 │ │
│  │   Drop PDF files    │            │  [PDF Preview]  │ │
│  │       here          │            │                 │ │
│  │                     │            │  ← Page 1 of 4→ │ │
│  │  [📁 Browse Files]  │            └─────────────────┘ │
│  └─────────────────────┘                                │
│                                                         │
│  📋 Appendices                       ⚙️  Settings       │
│  • Appendix A (2 pages)             Numbering: A, B, C  │
│  • Appendix B (4 pages)             ☑ Auto-backup      │
│  • Appendix C (21 pages)                               │
│                                                         │
│  [🔍 Preview Changes] [▶️ Process Appendices]          │
└─────────────────────────────────────────────────────────┘
```

</details>

---

## 📋 Features

### 🔥 Core Features
- [x] **PDF Integration** - Convert PDF pages to Word-compatible images
- [x] **Smart Document Detection** - Auto-detect open Word documents
- [x] **Drag & Drop Interface** - Easy PDF file management
- [x] **Live Preview** - Preview PDFs before processing
- [x] **Custom Numbering** - Multiple numbering schemes (A-Z, 1-9, I-X)
- [x] **Backup Creation** - Automatic document backups
- [x] **Progress Tracking** - Real-time processing feedback

### 🎨 User Experience
- [x] **Modern GUI** - Clean, professional interface
- [x] **Responsive Design** - Adapts to different screen sizes
- [x] **Error Handling** - Graceful error recovery and user feedback
- [x] **Threading** - Non-blocking operations for smooth performance
- [x] **Logging** - Comprehensive logging for debugging

### 🔧 Technical Features
- [x] **COM Integration** - Direct Microsoft Word automation
- [x] **Multi-threaded Processing** - Background PDF processing
- [x] **Memory Optimization** - Efficient handling of large PDF files
- [x] **Cross-version Support** - Works with multiple Word versions

---

## 📦 Dependencies & Packages

<details>
<summary>👆 <strong>Click to explore package details</strong></summary>

### Core Dependencies

| Package | Version | Purpose | Usage in Project |
|---------|---------|---------|------------------|
| **PySide6** | `>=6.0.0` | Qt GUI Framework | Main UI, widgets, dialogs, event handling |
| **PyMuPDF (fitz)** | `>=1.23.0` | PDF Processing | PDF reading, page extraction, thumbnail generation |
| **pywin32** | `>=306` | Windows COM | Microsoft Word automation and integration |
| **python-docx** | `>=0.8.11` | Word Documents | Fallback Word document manipulation |
| **Pillow** | `>=9.0.0` | Image Processing | Image conversion, resizing, format handling |

### Development Dependencies

| Package | Purpose | Files Using It |
|---------|---------|----------------|
| **pathlib** | Path handling | All modules for file operations |
| **logging** | Application logging | `utils/logger.py`, all components |
| **threading** | Background processing | `gui/widgets/pdf_preview_widget.py` |
| **json** | Configuration management | `utils/config.py` |
| **datetime** | Timestamps and logging | `utils/logger.py` |

</details>

---

## 🏗️ Project Architecture

<details>
<summary>👆 <strong>Click to explore project structure</strong></summary>

```
word_appendix_manager/
├── 📁 src/                          # Source code
│   ├── 📁 core/                     # Business logic
│   │   ├── 📄 word_manager.py       # Word document automation
│   │   ├── 📄 pdf_handler.py        # PDF processing and conversion
│   │   └── 📄 appendix_manager.py   # Appendix integration logic
│   │
│   ├── 📁 gui/                      # User interface
│   │   ├── 📄 main_window.py        # Main application window
│   │   ├── 📄 controller.py         # Application controller/coordinator
│   │   │
│   │   ├── 📁 widgets/              # Reusable UI components
│   │   │   ├── 📄 document_selector.py      # Word document selection
│   │   │   ├── 📄 drag_drop_area.py         # PDF drag & drop zone
│   │   │   ├── 📄 appendix_list_widget.py   # Appendix management
│   │   │   └── 📄 pdf_preview_widget.py     # PDF thumbnail preview
│   │   │
│   │   └── 📁 dialogs/              # Modal dialogs
│   │       ├── 📄 settings_dialog.py        # Application settings
│   │       ├── 📄 preview_dialog.py         # Document preview
│   │       └── 📄 appendix_edit_dialog.py   # Appendix editing
│   │
│   ├── 📁 utils/                    # Utility modules
│   │   ├── 📄 logger.py             # Logging configuration
│   │   ├── 📄 config.py             # Application configuration
│   │   └── 📄 exceptions.py         # Custom exceptions
│   │
│   └── 📄 main.py                   # Application entry point
│
├── 📁 assets/                       # Static resources
│   ├── 📄 icon.png                  # Application icon
│   └── 📁 images/                   # UI images and graphics
│
├── 📁 logs/                         # Application logs
├── 📁 temp/                         # Temporary files
├── 📁 output/                       # Processed documents
├── 📁 config/                       # Configuration files
│
├── 📄 requirements.txt              # Python dependencies
├── 📄 setup.py                      # Installation script
├── 📄 README.md                     # This file
└── 📄 LICENSE                       # MIT License
```

</details>

### 📁 Key File Descriptions

<details>
<summary>👆 <strong>Click to understand each file's purpose</strong></summary>

#### 🎯 Core Business Logic
- **`core/word_manager.py`** - Handles all Microsoft Word automation using COM, document opening, content insertion
- **`core/pdf_handler.py`** - PDF processing, page extraction, image conversion, thumbnail generation
- **`core/appendix_manager.py`** - Coordinates appendix integration, numbering, formatting

#### 🖥️ User Interface
- **`gui/main_window.py`** - Main application window, menu bar, status bar, layout management
- **`gui/controller.py`** - Application state management, coordinates between UI and business logic
- **`gui/widgets/document_selector.py`** - Word document detection and selection interface
- **`gui/widgets/pdf_preview_widget.py`** - PDF thumbnail display with page navigation
- **`gui/widgets/drag_drop_area.py`** - Drag & drop functionality for PDF files
- **`gui/widgets/appendix_list_widget.py`** - Appendix management, reordering, settings

#### 🔧 Utilities & Configuration
- **`utils/logger.py`** - Application-wide logging, file and console output
- **`utils/config.py`** - Settings management, user preferences, defaults
- **`utils/exceptions.py`** - Custom exception classes for error handling

</details>

---

## 🔧 Installation

### 📋 Prerequisites
- **Windows 10/11** (required for COM integration)
- **Python 3.8+** 
- **Microsoft Word** (2016 or later recommended)
- **Git** (for cloning)

### 🚀 Quick Installation

<details>
<summary>👆 <strong>Method 1: One-Command Setup (Recommended)</strong></summary>

```bash
# Clone and setup in one go
git clone https://github.com/yourusername/word-appendix-manager.git
cd word-appendix-manager
python -m pip install --upgrade pip
pip install -r requirements.txt
python src/main.py
```

</details>

<details>
<summary>👆 <strong>Method 2: Step-by-Step Setup</strong></summary>

```bash
# 1. Clone the repository
git clone https://github.com/yourusername/word-appendix-manager.git
cd word-appendix-manager

# 2. Create virtual environment (recommended)
python -m venv wamenv
wamenv\Scripts\activate  # Windows
# source wamenv/bin/activate  # Linux/Mac (if supported in future)

# 3. Install dependencies
pip install --upgrade pip
pip install -r requirements.txt

# 4. Verify installation
python -c "import PySide6, fitz, win32com.client; print('✅ All dependencies installed!')"

# 5. Run the application
python src/main.py
```

</details>

<details>
<summary>👆 <strong>Method 3: Development Setup</strong></summary>

```bash
# Clone with development tools
git clone https://github.com/yourusername/word-appendix-manager.git
cd word-appendix-manager

# Install in development mode
pip install -e .
pip install -r requirements-dev.txt

# Run tests (if available)
python -m pytest tests/

# Run with debugging
python src/main.py --debug
```

</details>

### 🛠️ Troubleshooting Installation

<details>
<summary>👆 <strong>Click if you encounter issues</strong></summary>

#### Common Issues:

**❌ `ImportError: No module named 'win32com'`**
```bash
pip install pywin32
# After installation, run:
python Scripts/pywin32_postinstall.py -install
```

**❌ `ModuleNotFoundError: No module named 'fitz'`**
```bash
pip install PyMuPDF
```

**❌ `ImportError: No module named 'PySide6'`**
```bash
pip install PySide6
```

**❌ Word COM not working:**
- Ensure Microsoft Word is installed
- Run as administrator once to register COM
- Check Windows COM+ registration

</details>

---

## 📖 Usage

### 🎮 Interactive Usage Guide

<details>
<summary>👆 <strong>Step 1: Start the Application</strong></summary>

```bash
# Navigate to project directory
cd word-appendix-manager

# Run the application
python src/main.py
```

**Expected Output:**
```
============================================================
🎯 WORD APPENDIX MANAGER v1.0
📋 Professional PDF Appendix Tool for Word Documents
⚡ Built with PySide6 & Python
============================================================
🚀 Word Appendix Manager v1.0 - Started Successfully!
📄 Ready to process Word documents with PDF appendices
============================================================
```

</details>

<details>
<summary>👆 <strong>Step 2: Select a Word Document</strong></summary>

**Option A: Use Open Documents**
1. Open Microsoft Word
2. Open or create a document
3. In the app, keep "Open Documents" selected
4. Click "🔄 Refresh" 
5. Select your document from dropdown
6. Click "✓ Select Document"

**Option B: Browse for File**
1. Click "Browse File" radio button
2. Click "📁 Browse"
3. Select a Word document (.docx or .doc)
4. Click "✓ Select Document"

</details>

<details>
<summary>👆 <strong>Step 3: Add PDF Files</strong></summary>

**Method 1: Drag & Drop**
```
1. Drag PDF files from Windows Explorer
2. Drop them onto the drop zone
3. Files will be automatically validated
```

**Method 2: Browse Button**
```
1. Click "📁 Browse Files" 
2. Select one or multiple PDF files
3. Click "Open"
```

**Result:** PDFs appear in the Appendices list with page counts

</details>

<details>
<summary>👆 <strong>Step 4: Configure Settings</strong></summary>

**Numbering Options:**
- `Alphabetical (A, B, C...)` - Default
- `Numeric (1, 2, 3...)`
- `Roman (I, II, III...)`

**Additional Settings:**
- ☑ `Auto-create backup` - Creates backup before processing
- Page orientation handling
- Image quality settings

</details>

<details>
<summary>👆 <strong>Step 5: Process Appendices</strong></summary>

```
1. Click "🔍 Preview Changes" (optional)
   - See what will be added to your document
   
2. Click "▶️ Process Appendices"
   - Progress bar shows processing status
   - Each PDF page is converted and inserted
   
3. Check your Word document
   - New appendix sections added
   - Table of contents updated
   - Professional formatting applied
```

**Processing Time:** ~5-10 seconds per PDF page

</details>

### 📊 Usage Examples

<details>
<summary>👆 <strong>Example 1: Academic Paper</strong></summary>

```
Scenario: Research paper with 3 supporting documents
- Research_Paper.docx (main document)
- Survey_Results.pdf (2 pages)
- Statistical_Analysis.pdf (5 pages)  
- Interview_Transcripts.pdf (12 pages)

Result: Paper with Appendix A, B, C containing all supporting materials
```

</details>

<details>
<summary>👆 <strong>Example 2: Business Proposal</strong></summary>

```
Scenario: Business proposal with supporting documents
- Proposal.docx (main document)
- Financial_Projections.pdf (3 pages)
- Market_Research.pdf (8 pages)
- Team_Resumes.pdf (6 pages)

Result: Professional proposal with organized appendices
```

</details>

---

## 🧪 Testing

<details>
<summary>👆 <strong>Click to see testing commands</strong></summary>

### Manual Testing
```bash
# Test core imports
python -c "from src.core.word_manager import WordManager; print('✅ Core working!')"

# Test GUI components  
python -c "from src.gui.main_window import MainWindow; print('✅ GUI working!')"

# Test PDF handling
python -c "from src.core.pdf_handler import PDFHandler; print('✅ PDF handler working!')"
```

### Integration Testing
```bash
# Run with test PDFs (create test_files/ directory)
mkdir test_files
# Copy sample PDFs to test_files/
python src/main.py
```

</details>

---

## 🚀 Performance

### 📊 Benchmarks

<details>
<summary>👆 <strong>Processing Performance</strong></summary>

| PDF Size | Pages | Processing Time | Memory Usage |
|----------|-------|-----------------|--------------|
| Small    | 1-5   | ~2-5 seconds    | ~50-100 MB   |
| Medium   | 6-20  | ~10-30 seconds  | ~100-300 MB  |
| Large    | 21+   | ~30+ seconds    | ~300+ MB     |

**Optimization Features:**
- Background processing prevents UI freezing
- Memory cleanup after each document
- Progressive loading for large files
- Thread-safe operations

</details>

---

## 🤝 Contributing

<details>
<summary>👆 <strong>How to Contribute</strong></summary>

```bash
# 1. Fork the repository
# 2. Create feature branch
git checkout -b feature/amazing-feature

# 3. Make changes and commit
git commit -m "Add amazing feature"

# 4. Push to branch
git push origin feature/amazing-feature

# 5. Open Pull Request
```

**Areas for Contribution:**
- 🐛 Bug fixes and error handling
- ✨ New features and enhancements
- 📝 Documentation improvements
- 🧪 Test coverage expansion
- 🎨 UI/UX improvements

</details>

---

## 📝 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

---

## 🆘 Support

<details>
<summary>👆 <strong>Getting Help</strong></summary>

**📧 Issues & Questions:**
- [GitHub Issues](https://github.com/yourusername/word-appendix-manager/issues)
- [Discussions](https://github.com/yourusername/word-appendix-manager/discussions)

**📋 Before Reporting Issues:**
1. Check existing issues
2. Include error logs from `logs/` directory
3. Specify Windows version and Word version
4. Provide sample files if possible

**🔍 Debug Mode:**
```bash
python src/main.py --debug
```

</details>

---

<div align="center">

## 🌟 Star History

[![Star History Chart](https://api.star-history.com/svg?repos=yourusername/word-appendix-manager&type=Date)](https://star-history.com/integral05/word-appendix-manager&Date)

**Made with ❤️ by Integral**

*If this project helped you, please consider giving it a ⭐ star!*

</div> 
