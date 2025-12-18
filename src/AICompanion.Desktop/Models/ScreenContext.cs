using System;
using System.Collections.Generic;

namespace AICompanion.Desktop.Models
{
    /*
        ScreenContext captures the current state of the user's screen at a moment in time.
        
        This class combines multiple sources of screen information to give the AI engine
        a complete understanding of what the user is seeing. The context includes:
        
        1. A screenshot image for visual reference
        2. OCR-extracted text from all visible elements
        3. UI Automation tree data describing interactive elements
        4. Information about the currently focused window and element
        
        The AI uses this context to understand commands like "click that button" or
        "read what's on the screen" by correlating user intent with visible elements.
        
        Reference: https://docs.microsoft.com/en-us/dotnet/framework/ui-automation/
    */
    public class ScreenContext
    {
        /*
            Screenshot image data encoded as PNG bytes.
            The image is captured at the current screen resolution and compressed
            to approximately 2-3MB for efficient transmission to the AI engine.
        */
        public byte[] ScreenshotData { get; set; } = Array.Empty<byte>();

        /*
            Timestamp when the screenshot was captured.
            The system aims for less than 0.5 seconds between capture and processing.
        */
        public DateTime CapturedAt { get; set; } = DateTime.UtcNow;

        /*
            Screen resolution at capture time (width x height).
            Used to calculate element positions and scaling.
        */
        public int ScreenWidth { get; set; }
        public int ScreenHeight { get; set; }

        /*
            Text extracted from the screen using Tesseract OCR.
            This provides a textual description of all visible content,
            which the AI model processes to understand the screen state.
        */
        public string ExtractedText { get; set; } = string.Empty;

        /*
            List of UI elements discovered through Windows UI Automation.
            Each element includes its name, type, position, and whether
            it can be interacted with (clicked, typed into, etc.).
        */
        public List<UIElementInfo> DiscoveredElements { get; set; } = new();

        /*
            Information about the currently focused window.
            Helps the AI understand which application the user is working with.
        */
        public WindowInfo? ActiveWindow { get; set; }

        /*
            The currently focused UI element, if any.
            Useful for commands like "press enter" or "what is this?"
        */
        public UIElementInfo? FocusedElement { get; set; }

        /*
            Number of open windows on the desktop.
            Provides context for window management commands.
        */
        public int OpenWindowCount { get; set; }

        /*
            Creates a text summary of the screen state for the AI prompt.
            This method formats the context information in a way that
            the IBM Granite model can understand and process.
        */
        public string ToPromptContext()
        {
            var summary = $"Screen Resolution: {ScreenWidth}x{ScreenHeight}\n";
            summary += $"Active Window: {ActiveWindow?.Title ?? "None"}\n";
            summary += $"Open Windows: {OpenWindowCount}\n";
            summary += $"Visible Text: {ExtractedText}\n";
            
            if (DiscoveredElements.Count > 0)
            {
                summary += "Interactive Elements:\n";
                foreach (var element in DiscoveredElements)
                {
                    summary += $"  - {element.ElementType}: \"{element.Name}\" at ({element.X}, {element.Y})\n";
                }
            }
            
            return summary;
        }
    }

    /*
        UIElementInfo describes a single interactive element on the screen.
        
        Windows UI Automation provides structured access to application controls,
        allowing the system to programmatically interact with buttons, text fields,
        menus, and other UI components without relying solely on image recognition.
    */
    public class UIElementInfo
    {
        /*
            Unique identifier for this element within the automation tree.
            Used to target specific elements for interaction.
        */
        public string AutomationId { get; set; } = string.Empty;

        /*
            Human-readable name of the element, typically the button text
            or label associated with the control.
        */
        public string Name { get; set; } = string.Empty;

        /*
            Type of UI control: Button, TextBox, Menu, ListItem, etc.
        */
        public string ElementType { get; set; } = string.Empty;

        /*
            Position of the element on screen in pixels.
        */
        public int X { get; set; }
        public int Y { get; set; }
        public int Width { get; set; }
        public int Height { get; set; }

        /*
            Whether the element is currently enabled and can receive input.
        */
        public bool IsEnabled { get; set; }

        /*
            Whether the element is visible on screen (not hidden or off-screen).
        */
        public bool IsVisible { get; set; }

        /*
            List of supported interaction patterns.
            For example: "Invoke" for buttons, "Value" for text fields.
        */
        public List<string> SupportedPatterns { get; set; } = new();
    }

    /*
        WindowInfo contains details about an application window.
        
        This information helps the AI understand which application the user
        is working with and enables window management commands like
        "minimize this" or "switch to Chrome".
    */
    public class WindowInfo
    {
        /*
            Window handle for programmatic access.
        */
        public IntPtr Handle { get; set; }

        /*
            Text displayed in the window's title bar.
        */
        public string Title { get; set; } = string.Empty;

        /*
            Name of the process that owns this window.
        */
        public string ProcessName { get; set; } = string.Empty;

        /*
            Whether this window is currently in focus.
        */
        public bool IsActive { get; set; }

        /*
            Whether the window is minimized.
        */
        public bool IsMinimized { get; set; }

        /*
            Whether the window is maximized.
        */
        public bool IsMaximized { get; set; }

        /*
            Window position and size on screen.
        */
        public int X { get; set; }
        public int Y { get; set; }
        public int Width { get; set; }
        public int Height { get; set; }
    }
}
