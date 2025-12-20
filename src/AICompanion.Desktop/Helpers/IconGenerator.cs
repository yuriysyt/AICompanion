using System;
using System.IO;
using System.Windows;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using WpfColor = System.Windows.Media.Color;
using WpfPoint = System.Windows.Point;
using WpfBrushes = System.Windows.Media.Brushes;
using WpfPen = System.Windows.Media.Pen;

namespace AICompanion.Desktop.Helpers
{
    /// <summary>
    /// Generates application icon programmatically.
    /// Creates a modern AI companion icon with gradient background and avatar face.
    /// </summary>
    public static class IconGenerator
    {
        public static BitmapSource CreateAppIcon(int size = 256)
        {
            var visual = new DrawingVisual();

            using (var context = visual.RenderOpen())
            {
                // Background gradient (purple-blue)
                var gradientBrush = new LinearGradientBrush(
                    WpfColor.FromRgb(0x66, 0x7E, 0xEA),
                    WpfColor.FromRgb(0x76, 0x4B, 0xA2),
                    45);

                // Draw rounded rectangle background
                var rect = new Rect(0, 0, size, size);
                var radius = size * 0.2;
                context.DrawRoundedRectangle(gradientBrush, null, rect, radius, radius);

                // Draw face circle (lighter)
                var faceBrush = new SolidColorBrush(WpfColor.FromArgb(40, 255, 255, 255));
                var faceCenter = new WpfPoint(size / 2, size / 2);
                context.DrawEllipse(faceBrush, null, faceCenter, size * 0.35, size * 0.35);

                // Draw eyes (white)
                var eyeBrush = WpfBrushes.White;
                var leftEyeCenter = new WpfPoint(size * 0.35, size * 0.42);
                var rightEyeCenter = new WpfPoint(size * 0.65, size * 0.42);
                var eyeRadiusX = size * 0.08;
                var eyeRadiusY = size * 0.10;

                context.DrawEllipse(eyeBrush, null, leftEyeCenter, eyeRadiusX, eyeRadiusY);
                context.DrawEllipse(eyeBrush, null, rightEyeCenter, eyeRadiusX, eyeRadiusY);

                // Draw pupils (dark)
                var pupilBrush = new SolidColorBrush(WpfColor.FromRgb(0x1D, 0x1D, 0x1F));
                var pupilRadius = size * 0.03;
                context.DrawEllipse(pupilBrush, null, new WpfPoint(leftEyeCenter.X + 2, leftEyeCenter.Y + 2), pupilRadius, pupilRadius * 1.2);
                context.DrawEllipse(pupilBrush, null, new WpfPoint(rightEyeCenter.X + 2, rightEyeCenter.Y + 2), pupilRadius, pupilRadius * 1.2);

                // Draw smile
                var smilePen = new WpfPen(WpfBrushes.White, size * 0.025);
                smilePen.StartLineCap = PenLineCap.Round;
                smilePen.EndLineCap = PenLineCap.Round;

                var smileGeometry = new StreamGeometry();
                using (var sgc = smileGeometry.Open())
                {
                    sgc.BeginFigure(new WpfPoint(size * 0.32, size * 0.62), false, false);
                    sgc.QuadraticBezierTo(
                        new WpfPoint(size * 0.5, size * 0.75),
                        new WpfPoint(size * 0.68, size * 0.62),
                        true, false);
                }
                smileGeometry.Freeze();
                context.DrawGeometry(null, smilePen, smileGeometry);

                // Draw microphone indicator (small circle at bottom)
                var micBrush = new SolidColorBrush(WpfColor.FromRgb(0x34, 0xC7, 0x59));
                context.DrawEllipse(micBrush, null, new WpfPoint(size * 0.75, size * 0.8), size * 0.08, size * 0.08);
            }

            var bitmap = new RenderTargetBitmap(size, size, 96, 96, PixelFormats.Pbgra32);
            bitmap.Render(visual);
            bitmap.Freeze();

            return bitmap;
        }

        public static void SaveAsIco(BitmapSource bitmap, string path)
        {
            // Create multiple sizes for ICO
            var sizes = new[] { 16, 32, 48, 256 };

            using var stream = new FileStream(path, FileMode.Create);
            using var writer = new BinaryWriter(stream);

            // ICO header
            writer.Write((short)0);      // Reserved
            writer.Write((short)1);      // Type (1 = ICO)
            writer.Write((short)sizes.Length); // Number of images

            var imageDataList = new System.Collections.Generic.List<byte[]>();
            var offset = 6 + (16 * sizes.Length); // Header + directory entries

            // Write directory entries
            foreach (var size in sizes)
            {
                var resized = ResizeBitmap(bitmap, size);
                var pngData = GetPngBytes(resized);
                imageDataList.Add(pngData);

                writer.Write((byte)(size == 256 ? 0 : size)); // Width (0 = 256)
                writer.Write((byte)(size == 256 ? 0 : size)); // Height
                writer.Write((byte)0);   // Color palette
                writer.Write((byte)0);   // Reserved
                writer.Write((short)1);  // Color planes
                writer.Write((short)32); // Bits per pixel
                writer.Write(pngData.Length); // Size of image data
                writer.Write(offset);    // Offset to image data
                offset += pngData.Length;
            }

            // Write image data
            foreach (var data in imageDataList)
            {
                writer.Write(data);
            }
        }

        private static BitmapSource ResizeBitmap(BitmapSource source, int size)
        {
            var scale = (double)size / source.PixelWidth;
            var transform = new ScaleTransform(scale, scale);

            var resized = new TransformedBitmap(source, transform);
            resized.Freeze();

            return resized;
        }

        private static byte[] GetPngBytes(BitmapSource bitmap)
        {
            var encoder = new PngBitmapEncoder();
            encoder.Frames.Add(BitmapFrame.Create(bitmap));

            using var stream = new MemoryStream();
            encoder.Save(stream);
            return stream.ToArray();
        }
    }
}
