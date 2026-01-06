using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Forms;
using AICompanion.Desktop.Models;
using Microsoft.Extensions.Logging;

namespace AICompanion.Desktop.Services.Screen
{
    /*
        ScreenCaptureService handles capturing screenshots of the user's display.
        
        Screenshots provide visual context for the AI engine to understand what
        the user is currently seeing. Combined with OCR text extraction and
        UI Automation data, the screenshot enables the AI to make informed
        decisions about which elements to interact with.
        
        The service captures the entire primary display at its native resolution
        and compresses the image to PNG format for efficient transmission to
        the Python AI engine via gRPC.
        
        Performance target: capture and compress in under 0.5 seconds.
        
        Reference: https://docs.microsoft.com/en-us/dotnet/api/system.drawing
    */
    public class ScreenCaptureService
    {
        private readonly ILogger<ScreenCaptureService> _logger;

        public ScreenCaptureService(ILogger<ScreenCaptureService> logger)
        {
            _logger = logger;
        }

        /*
            Captures the current state of the primary display.
            
            This method creates a bitmap of the entire screen and encodes it
            as PNG bytes for transmission. The ScreenContext object includes
            both the image data and metadata about screen dimensions.
        */
        public async Task<ScreenContext> CaptureScreenAsync()
        {
            return await Task.Run(() =>
            {
                var context = new ScreenContext();
                
                try
                {
                    _logger.LogDebug("Capturing screen");
                    var startTime = DateTime.UtcNow;
                    
                    /*
                        Get the bounds of the primary screen.
                        For multi-monitor setups, this captures only the main display.
                        Future versions may support capturing specific monitors.
                    */
                    var screenBounds = System.Windows.Forms.Screen.PrimaryScreen?.Bounds ?? new Rectangle(0, 0, 1920, 1080);
                    
                    context.ScreenWidth = screenBounds.Width;
                    context.ScreenHeight = screenBounds.Height;
                    
                    /*
                        Create a bitmap with the same dimensions as the screen.
                        Using 32-bit ARGB format for full color fidelity.
                    */
                    using var bitmap = new System.Drawing.Bitmap(screenBounds.Width, screenBounds.Height, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
                    
                    /*
                        Use Graphics.CopyFromScreen to capture the display contents.
                        This is the most efficient method for full-screen capture
                        on Windows systems.
                    */
                    using var graphics = Graphics.FromImage(bitmap);
                    graphics.CopyFromScreen(screenBounds.Location, Point.Empty, screenBounds.Size);
                    
                    /*
                        Encode the bitmap as PNG bytes.
                        PNG provides lossless compression which is good for text
                        readability while keeping file sizes reasonable (2-3MB typical).
                    */
                    using var memoryStream = new MemoryStream();
                    bitmap.Save(memoryStream, ImageFormat.Png);
                    context.ScreenshotData = memoryStream.ToArray();
                    
                    context.CapturedAt = DateTime.UtcNow;
                    
                    var captureTime = (DateTime.UtcNow - startTime).TotalMilliseconds;
                    _logger.LogDebug("Screen captured in {Time:F0}ms, size: {Size:N0} bytes", 
                        captureTime, context.ScreenshotData.Length);
                }
                catch (Exception ex)
                {
                    _logger.LogError(ex, "Failed to capture screen");
                    context.ScreenshotData = Array.Empty<byte>();
                }
                
                return context;
            });
        }

        /*
            Captures a specific region of the screen.
            
            Useful when the AI wants to focus on a particular window or
            UI element without processing the entire display.
        */
        public async Task<byte[]> CaptureRegionAsync(Rectangle region)
        {
            return await Task.Run(() =>
            {
                try
                {
                    _logger.LogDebug("Capturing region: {Region}", region);
                    
                    using var bitmap = new Bitmap(region.Width, region.Height, PixelFormat.Format32bppArgb);
                    using var graphics = Graphics.FromImage(bitmap);
                    graphics.CopyFromScreen(region.Location, Point.Empty, region.Size);
                    
                    using var memoryStream = new MemoryStream();
                    bitmap.Save(memoryStream, ImageFormat.Png);
                    return memoryStream.ToArray();
                }
                catch (Exception ex)
                {
                    _logger.LogError(ex, "Failed to capture region");
                    return Array.Empty<byte>();
                }
            });
        }

        /*
            Saves a screenshot to disk for debugging purposes.
            
            Screenshots are saved to the application's logs folder with
            timestamps for later analysis of problematic interactions.
        */
        public async Task SaveScreenshotAsync(byte[] imageData, string filename)
        {
            await Task.Run(() =>
            {
                try
                {
                    var logsFolder = Path.Combine(
                        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                        "AICompanion", "Screenshots");
                    
                    Directory.CreateDirectory(logsFolder);
                    
                    var filepath = Path.Combine(logsFolder, filename);
                    File.WriteAllBytes(filepath, imageData);
                    
                    _logger.LogDebug("Screenshot saved to: {Path}", filepath);
                }
                catch (Exception ex)
                {
                    _logger.LogError(ex, "Failed to save screenshot");
                }
            });
        }
    }
}
