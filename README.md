## 🎯 About This Project

PDF Converter is a desktop application designed to simplify the process of converting Office documents to PDF format while preserving all original formatting. Unlike online converters or basic tools, this application uses LibreOffice's industry-standard conversion engine to ensure:

- **Perfect formatting preservation** (bullets, numbering, tables, images, styles)
- **Batch processing** of multiple files simultaneously
- **No file size limits** (unlike online converters)
- **Complete privacy** (files never leave your computer)
- **Free and open source** with no hidden costs

## 🔧 How It Works

The application provides a user-friendly graphical interface built with Python's tkinter that orchestrates LibreOffice's powerful command-line conversion capabilities:

1. **User Interface**: Clean, dark-themed GUI for selecting files and settings
2. **Conversion Engine**: LibreOffice runs in headless mode for high-quality PDF generation
3. **Batch Processing**: Multi-threaded conversion with progress tracking
4. **Error Handling**: Comprehensive error reporting and recovery

## 🚀 Why Use This?

### **For End Users:**
- **Simplicity**: No technical knowledge required - just click and convert
- **Speed**: Convert multiple files in one go, no manual processing
- **Quality**: Professional-grade PDF output identical to Office's "Save as PDF."
- **Free**: No subscriptions, no watermarks, no limits

### **For Developers:**
- **Clean Code**: Well-structured, commented Python code
- **Extensible**: Easy to add new features or file formats
- **Cross-Platform**: Works on Linux, Windows, and macOS
- **Open Source**: MIT licensed - use, modify, distribute freely

## 📊 Key Features

| Feature | Description |
|---------|-------------|
| **Format Preservation** | Maintains bullets, numbering, tables, images, and styles |
| **Batch Conversion** | Convert dozens of files simultaneously |
| **Progress Tracking** | Real-time status updates and completion indicators |
| **Resizable Interface** | Adjustable window to show more files |
| **Output Management** | Choose custom output folders |
| **Error Handling** | Clear error messages and recovery options |
| **Dark Theme** | Easy on the eyes during extended use |

## 🎨 User Experience

The application features a thoughtfully designed interface:

- **Intuitive Layout**: Logical flow from file selection to conversion
- **Visual Feedback**: Icons, colors, and status messages guide the user
- **Responsive Design**: Window resizing with proper element scaling
- **Accessible**: Large buttons and clear labels for ease of use

## 🔄 Comparison with Alternatives

| Feature | This App | Online Converters | Office Built-in |
|---------|----------|-------------------|-----------------|
| **Format Quality** | ✅ Perfect | ⚠️ Variable | ✅ Perfect |
| **Privacy** | ✅ Local | ❌ Upload required | ✅ Local |
| **Batch Processing** | ✅ Yes | ⚠️ Limited | ❌ No |
| **File Size Limits** | ✅ None | ❌ Often limited | ✅ None |
| **Cost** | ✅ Free | ⚠️ Freemium | ✅ Free* |

*Office subscription required for Microsoft Office

## 🛠️ Technical Excellence

- **Modern Architecture**: Separation of UI, business logic, and system integration
- **Error Resilience**: Graceful handling of file errors, missing dependencies
- **Performance**: Multi-threading prevents UI freezing during conversion
- **Maintainability**: Clean code structure with clear separation of concerns

## 🌟 Perfect For

- **Students**: Convert assignments and reports to PDF for submission
- **Professionals**: Prepare documents for clients while preserving formatting
- **Administrators**: Batch convert archives of Office documents
- **Developers**: Example of Python GUI + system integration
- **Open Source Enthusiasts**: Clean, well-documented Python application

## 📈 Project Status

**Stable & Production Ready** - Used by hundreds of users with excellent feedback on conversion quality and reliability.

**Version 1.3** includes:
- ✅ Centered button layout
- ✅ Improved window resizing
- ✅ Better error handling
- ✅ Enhanced user interface
- ✅ Credit attribution

## 🖼️ Screenshot
<img width="602" height="591" alt="image" src="https://github.com/user-attachments/assets/5511e485-d32a-4d1d-a7b9-bb7b4f8a85cf" />

## System Requirements

### Minimum Requirements
| Component | Requirement |
|-----------|-------------|
| **Operating System** | Linux (Ubuntu 18.04+, Debian 10+, Fedora 32+, Arch), Windows 10+, macOS 10.15+ |
| **Python** | 3.6 or higher |
| **LibreOffice** | 6.0 or higher |
| **RAM** | 2 GB minimum |
| **Storage** | 100 MB free space |
| **Display** | 1024x768 resolution minimum |

## Quick Installation

### For Ubuntu/Debian (One Command)
```bash
# Run this single command to install everything
sudo apt update && sudo apt install python3 python3-tk libreoffice -y && wget -O pdf_converter.py https://raw.githubusercontent.com/leolynn7/pdf-converter/main/pdf_converter.py && python3 pdf_converter.py
