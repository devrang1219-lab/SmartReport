using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Globalization;
using System.IO;
using System.Text.RegularExpressions;
using Tesseract;

namespace WindowsFormsApp1
{
    public class FlirOcrReader : IDisposable
    {
        private readonly TesseractEngine engine;

        public FlirOcrReader()
        {
            string tessPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "tessdata");
            engine = new TesseractEngine(tessPath, "eng", EngineMode.Default);

            engine.SetVariable("tessedit_char_whitelist", "0123456789.");
        }

        public FlirOcrReader(string tessDataPath)
        {
            engine = new TesseractEngine(tessDataPath, "eng", EngineMode.Default);
            engine.SetVariable("tessedit_char_whitelist", "0123456789.");
        }

        public FlirResult Read(string imageFile)
        {
            using (Bitmap bmp = new Bitmap(imageFile))
            {
                FlirResult result = new FlirResult();

                using (Bitmap left = CropLeft(bmp))
                using (Bitmap right = CropRightBottom(bmp))
                using (Bitmap left2 = Preprocess(left, 4))
                using (Bitmap right2 = Preprocess(right, 6))
                {
#if DEBUG
                    string dir = Path.Combine(
                        Path.GetDirectoryName(imageFile),
                        "OcrDebug");

                    Directory.CreateDirectory(dir);

                    left.Save(Path.Combine(dir,
                        Path.GetFileNameWithoutExtension(imageFile) + "_left.png"),
                        System.Drawing.Imaging.ImageFormat.Png);

                    right.Save(Path.Combine(dir,
                        Path.GetFileNameWithoutExtension(imageFile) + "_right.png"),
                        System.Drawing.Imaging.ImageFormat.Png);

                    left2.Save(Path.Combine(dir,
                        Path.GetFileNameWithoutExtension(imageFile) + "_left_pre.png"),
                        System.Drawing.Imaging.ImageFormat.Png);

                    right2.Save(Path.Combine(dir,
                        Path.GetFileNameWithoutExtension(imageFile) + "_right_pre.png"),
                        System.Drawing.Imaging.ImageFormat.Png);
#endif

                    result.LeftRawText = Ocr(left2);
                    Debug.WriteLine($"result.LeftRawText : {result.LeftRawText}");
                    result.RightRawText = Ocr(right2);
                    Debug.WriteLine($"result.RightRawText : {result.RightRawText}");

                    result.Items = ParseItems(result.LeftRawText);
                    result.ScaleMinTemperature = ParseTemperature(result.RightRawText);
                }

                return result;
            }
        }

        private string Ocr(Bitmap bmp)
        {
            using (Pix pix = PixConverter.ToPix(bmp))
            using (Page page = engine.Process(pix, PageSegMode.SingleBlock))
            {
                return page.GetText();
            }
        }

        private Dictionary<string, double> ParseItems(string text)
        {
            Dictionary<string, double> dic = new Dictionary<string, double>();

            Regex regex = new Regex(@"([A-Za-z]{2}\d+).*?([\d]+\.[\d]+)");

            foreach (Match m in regex.Matches(text))
            {
                double value;

                if (double.TryParse(
                    m.Groups[2].Value,
                    NumberStyles.Any,
                    CultureInfo.InvariantCulture,
                    out value))
                {
                    dic[m.Groups[1].Value] = value;
                }
            }

            return dic;
        }

        private double ParseTemperature(string text)
        {
            Match m = Regex.Match(text, @"[\d]+\.[\d]+");

            if (!m.Success)
                return double.NaN;

            return double.Parse(m.Value, CultureInfo.InvariantCulture);
        }

        private Bitmap CropLeft(Bitmap bmp)
        {
            Rectangle r = new Rectangle(
                (int)(bmp.Width * 0.2),
                (int)(bmp.Height * 0.01),
                (int)(bmp.Width * 0.11),
                (int)(bmp.Height * 0.23));

            return bmp.Clone(r, bmp.PixelFormat);
        }

        private Bitmap CropRightBottom(Bitmap bmp)
        {
            Rectangle r = new Rectangle(
                (int)(bmp.Width * 0.86),
                (int)(bmp.Height * 0.9),
                (int)(bmp.Width * 0.14),
                (int)(bmp.Height * 0.09));

            return bmp.Clone(r, bmp.PixelFormat);
        }

        private Bitmap Preprocess(Bitmap src, int scale)
        {
            Bitmap bmp = new Bitmap(src.Width * scale, src.Height * scale);

            using (Graphics g = Graphics.FromImage(bmp))
            {
                g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                g.DrawImage(src, 0, 0, bmp.Width, bmp.Height);
            }

            for (int y = 0; y < bmp.Height; y++)
            {
                for (int x = 0; x < bmp.Width; x++)
                {
                    Color c = bmp.GetPixel(x, y);

                    int gray = (int)(0.299 * c.R + 0.587 * c.G + 0.114 * c.B);

                    bmp.SetPixel(x, y, Color.FromArgb(gray, gray, gray));
                }
            }

            return bmp;
        }

        public void Dispose()
        {
            engine?.Dispose();
        }
    }

    public class FlirResult
    {
        public Dictionary<string, double> Items { get; set; } = new Dictionary<string, double>();

        public double ScaleMinTemperature { get; set; }

        public string LeftRawText { get; set; }

        public string RightRawText { get; set; }
    }
}