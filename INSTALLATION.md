# 📦 Complete Installation Guide

## Table of Contents
- [System Requirements](#system-requirements)
- [Quick Installation](#quick-installation)
- [Detailed Platform Instructions](#detailed-platform-instructions)
  - [Ubuntu/Debian Linux](#ubuntudebian-linux)
  - [Fedora Linux](#fedora-linux)
  - [Arch Linux](#arch-linux)
  - [Windows 10/11](#windows-1011)
  - [macOS](#macos)
- [Verification Steps](#verification-steps)
- [Troubleshooting](#troubleshooting)
- [Advanced Installation](#advanced-installation)
- [Updating](#updating)

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

### Recommended Specifications
| Component | Recommendation |
|-----------|---------------|
| **Operating System** | Latest Ubuntu LTS or Windows 11/macOS 13+ |
| **Python** | 3.8 or higher |
| **LibreOffice** | 7.4 or higher |
| **RAM** | 4 GB or more |
| **Storage** | 500 MB free space |
| **Display** | 1366x768 or higher |

## Quick Installation

### For Ubuntu/Debian (One Command)
```bash
# Run this single command to install everything
sudo apt update && sudo apt install python3 python3-tk libreoffice -y && wget -O pdf_converter.py https://raw.githubusercontent.com/leolynn7/pdf-converter/main/pdf_converter.py && python3 pdf_converter.py
