# ElevenLabs TTS Integration Policy

## Current Status: Not Implemented

ElevenLabs Text-to-Speech is **not implemented** in the current version of AI Companion.

## Reasons for Non-Implementation

1. **API Key Requirement**: ElevenLabs requires a paid API key for production use. Free tier has strict limits.

2. **Online Dependency**: ElevenLabs is a cloud service requiring active internet connection, which contradicts the offline-first accessibility goal.

3. **Demo Stability Risk**: Relying on external cloud service during live demonstration introduces failure risk.

4. **Windows TTS Sufficiency**: Windows built-in TTS (System.Speech.Synthesis) provides reliable, offline voice output with acceptable quality for accessibility purposes.

## Future Implementation Guide

To add ElevenLabs TTS support in a future version:

### 1. Service Configuration
```csharp
public class ElevenLabsConfig
{
    public string ApiKey { get; set; }
    public string VoiceId { get; set; } = "21m00Tcm4TlvDq8ikWAM"; // Rachel
    public string ModelId { get; set; } = "eleven_monolingual_v1";
}
```

### 2. Settings Storage
- Store API key in `appsettings.json` (encrypted)
- Add toggle in Settings dialog
- Persist preference in local config

### 3. Fallback Logic
```csharp
public async Task SpeakAsync(string text)
{
    if (_useElevenLabs && _elevenLabsService.IsAvailable)
    {
        try
        {
            await _elevenLabsService.SpeakAsync(text);
        }
        catch
        {
            // Fallback to Windows TTS on failure
            await _windowsTts.SpeakAsync(text);
        }
    }
    else
    {
        await _windowsTts.SpeakAsync(text);
    }
}
```

### 4. UI Toggle Location
The placeholder toggle is already in MainWindow.xaml in the Voice Settings section. Enable it by:
1. Creating `ElevenLabsService.cs` in `Services/Voice/`
2. Adding configuration dialog similar to Watson setup
3. Implementing the API client with NAudio for playback

## Recommendation

Keep Windows TTS as default for:
- **Accessibility users** who need consistent, reliable output
- **Offline environments** where internet is unavailable
- **Demonstrations** where network issues could cause failures

ElevenLabs should be offered as an **optional premium feature** for users who want higher quality voices and have stable internet connections.
