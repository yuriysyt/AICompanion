# Changelog

All notable changes to the AI Companion project are documented in this file.

## [0.3.0] - 2026-01-20

### Added
- System tray integration for minimal desktop presence
- Global keyboard shortcuts for accessibility (Ctrl+Shift+A, S, H)
- Animated avatar control with emotion state transitions
- Conversation history tracking for context-aware responses

### Changed
- Improved error handling with retry logic
- Enhanced string processing utilities for command parsing

### Documentation
- Added AI Engine setup guide
- Updated README with project badges

## [0.2.0] - 2025-12-19

### Added
- Unit tests for VoiceCommand model
- AppSettings for user configuration persistence
- MainViewModel with full command processing pipeline
- MainWindow with floating companion UI design

### Changed
- Refined WPF styling with purple/blue gradient theme

## [0.1.0] - 2025-12-01

### Added
- Initial project structure and solution files
- gRPC protocol definition for AI engine communication
- VoiceCommand, ActionResult, and ScreenContext models
- VoiceRecognitionService with Windows Speech API
- TextToSpeechService for voice output
- ScreenCaptureService for display capture
- UIAutomationService for Windows application control
- AIEngineClient for gRPC communication
- WPF application entry point with dependency injection
- README with project overview and architecture

### Technical
- Target .NET 8.0-windows with WPF
- CommunityToolkit.Mvvm for MVVM pattern
- Serilog for structured logging
- gRPC.Net.Client for AI communication
