# AI Companion

[![Build Status](https://img.shields.io/badge/build-passing-brightgreen)]()
[![License](https://img.shields.io/badge/license-MIT-blue)]()
[![.NET](https://img.shields.io/badge/.NET-8.0-purple)]()
[![IBM Granite](https://img.shields.io/badge/AI-IBM%20Granite-blue)]()

Voice-controlled computer assistance powered by IBM Granite AI technology.

## Project Overview

AI Companion is a desktop application that enables complete computer control through voice commands 
and artificial intelligence assistance. The system is designed for users with disabilities, 
elderly users, and computer beginners who need accessible computer interaction.

## Architecture

The system consists of three main components:

**Desktop Application (C# + WPF)** handles voice input/output, companion avatar display, 
screenshot capture, and Windows UI Automation for computer control.

**AI Engine (Python + PyTorch)** runs IBM Granite 3.3-2B model for natural language understanding,
Tesseract OCR for screen text extraction, and ChromaDB for conversation memory.

**Communication Layer (gRPC)** provides high-speed binary protocol communication between 
the desktop application and AI engine on localhost:50051.

## Technology Stack

- C# .NET 8 with WPF for desktop application
- Python 3.11 with PyTorch for AI processing
- IBM Granite 3.3-2B-Instruct model (Apache 2.0 license)
- gRPC for inter-process communication
- Windows UI Automation API for computer control
- Windows Speech Recognition for voice input
- System.Speech.Synthesis for text-to-speech output

## System Requirements

**Minimum (Local AI Mode):**
- Windows 10/11 64-bit
- Intel i5 8th gen or AMD Ryzen 5 2600
- NVIDIA GTX 1060 6GB
- 8GB RAM
- 10GB storage

**Recommended:**
- Windows 11 64-bit
- Intel i7 10th gen or higher
- NVIDIA RTX 4060 8GB
- 16GB RAM
- 20GB SSD

## Project Structure

```
AICompanion/
├── src/
│   └── AICompanion.Desktop/
│       ├── Models/              # Data models and entities
│       ├── Services/            # Business logic services
│       │   ├── Voice/           # Speech recognition and synthesis
│       │   ├── Screen/          # Screenshot and OCR processing
│       │   ├── Automation/      # Windows UI Automation
│       │   └── Communication/   # gRPC client for AI engine
│       ├── Views/               # WPF views (XAML)
│       ├── ViewModels/          # MVVM view models
│       ├── Controls/            # Custom WPF controls
│       ├── Helpers/             # Utility classes
│       ├── Resources/           # Images, assets, localization
│       └── Configuration/       # App settings and constants
├── tests/
│   └── AICompanion.Tests/       # Unit and integration tests
└── docs/                        # Documentation
```

## Building the Project

1. Install Visual Studio 2022 with .NET 8 SDK
2. Clone this repository
3. Open AICompanion.sln in Visual Studio
4. Restore NuGet packages
5. Build and run

## IBM Partnership

This project is developed in partnership with IBM, utilizing their Granite AI models 
and optional Watson cloud services for enhanced voice recognition.

## License

This project uses IBM Granite models under Apache 2.0 license.

## References

- IBM Granite Model: https://huggingface.co/ibm-granite/granite-3.3-2b-instruct
- gRPC Documentation: https://grpc.io/docs/languages/csharp/
- Windows UI Automation: https://docs.microsoft.com/en-us/dotnet/framework/ui-automation/
- WPF Documentation: https://docs.microsoft.com/en-us/dotnet/desktop/wpf/

## Author

Developed as an academic project demonstrating AI-powered accessibility solutions.
