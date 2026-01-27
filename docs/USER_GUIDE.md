# AI Companion User Guide

Welcome to AI Companion, your voice-controlled desktop assistant.

## Getting Started

### First Launch

When you first start AI Companion, the companion avatar will appear in a floating window. The system will initialize voice recognition and attempt to connect to the AI engine.

### Activating the Assistant

You have three ways to activate the assistant:

1. **Wake Word**: Say "Hey Assistant" to start a voice command
2. **Push-to-Talk**: Hold the microphone button while speaking
3. **Keyboard Shortcut**: Press Ctrl+Shift+A to activate listening

### Basic Voice Commands

Here are some commands you can try:

**Opening Applications**
- "Open Notepad"
- "Launch Calculator"
- "Start Microsoft Word"

**Window Management**
- "Close this window"
- "Minimize the window"

**Clicking and Typing**
- "Click the Save button"
- "Type hello world"

**Getting Help**
- "What can you do?"
- "Help me with this"

### Using Microsoft Word (если установлен)

1. Скажите: **"Open Word"** или **"Open Microsoft Word"**
2. Дождитесь открытия Word
3. Кликните в область документа (чтобы фокус был на тексте)
4. Диктуйте:
   - **"Type: Hello world"**
   - **"Select all"**
   - **"Bold"**
   - **"Save"**

## Understanding the Avatar

The companion avatar displays different emotions to show what it is doing:

- **Neutral**: Ready and waiting for your command
- **Listening**: Actively listening to your voice (green indicator glows)
- **Thinking**: Processing your command (spinner appears)
- **Speaking**: Providing verbal feedback
- **Happy**: Successfully completed your request
- **Confused**: Had trouble understanding or completing the request

## Keyboard Shortcuts

For accessibility, these global shortcuts work even when the window is not focused:

| Shortcut | Action |
|----------|--------|
| Ctrl+Shift+A | Activate listening mode |
| Ctrl+Shift+S | Stop current operation |
| Ctrl+Shift+H | Show/hide the companion window |

## System Tray

When minimized to the system tray:

- Double-click the icon to show the window
- Right-click for quick actions menu
- Balloon notifications alert you to important events

## Settings

Access settings by clicking the gear icon. You can configure:

- IBM Watson API key (cloud STT)
- Voice engine (Watson or Windows Speech)

Tip: If Watson is not configured, the app can use Windows Speech (offline).

## Troubleshooting

### Voice Not Recognized

- Check that your microphone is connected and working
- Speak clearly and at a moderate pace
- Try adjusting the recognition confidence threshold in settings

### AI Not Responding

- Ensure the Python AI engine is running
- Check the connection status indicator
- Restart the application if the issue persists

### Commands Not Executing

- Some applications may not support UI Automation
- Try using more specific element names
- Check that the target application is visible on screen

## Privacy

AI Companion processes all voice commands locally on your computer. No audio or commands are sent to external servers unless you explicitly enable cloud features.

Conversation history is stored locally and can be cleared at any time through settings.
