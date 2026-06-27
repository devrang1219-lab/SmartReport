using PdfiumViewer;
using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace SmartReport
{
    public class ImageInserter : IDisposable
    {
        private readonly Excel.Worksheet _ws;
        private readonly string _imagePath;
        private readonly bool _tempFile;


        public ImageInserter(Excel.Worksheet ws, string filePath)
        {
            _ws = ws;

            if (Path.GetExtension(filePath).ToLower() == ".pdf")
            {
                _imagePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".png");

                int dpi = 300;
                int width = 2480;
                int height = 3508;

                using (var pdf = PdfDocument.Load(filePath))
                using (Image img = pdf.Render(0,
                    width,
                    height,
                    dpi,
                    dpi,
                    true))
                {
                    Console.WriteLine($"{img.Width} x {img.Height}");
                    img.Save(_imagePath, System.Drawing.Imaging.ImageFormat.Png);
                }

                _tempFile = true;
            }
            else
            {
                _imagePath = filePath;
            }
        }

        public Excel.Shape Insert(string cellAddress, ImageInsertOptions option)
        {
            if (option == null)
                option = new ImageInsertOptions();

            Excel.Range cell = _ws.Range[cellAddress];

            string imagePath = _imagePath;

            // Crop
            if (option.Crop)
            {
                imagePath = CropImage(
                    imagePath,
                    option.CropLeft,
                    option.CropTop,
                    option.CropRight,
                    option.CropBottom);
            }

            Excel.Shape pic;

            if (option.Width.HasValue && option.Height.HasValue)
            {
                pic = _ws.Shapes.AddPicture(
                    imagePath,
                    Office.MsoTriState.msoFalse,
                    Office.MsoTriState.msoTrue,
                    (float)cell.Left,
                    (float)cell.Top,
                    (float)option.Width.Value,
                    (float)option.Height.Value);
            }
            else
            {
                pic = _ws.Shapes.AddPicture(
                    imagePath,
                    Office.MsoTriState.msoFalse,
                    Office.MsoTriState.msoTrue,
                    (float)cell.Left,
                    (float)cell.Top,
                    -1,
                    -1);

                if (option.Scale.HasValue)
                {
                    pic.LockAspectRatio = Office.MsoTriState.msoTrue;
                    pic.ScaleWidth((float)option.Scale.Value, Office.MsoTriState.msoTrue);
                    pic.ScaleHeight((float)option.Scale.Value, Office.MsoTriState.msoTrue);
                }
            }

            return pic;
        }

        public void Dispose()
        {
            if (_tempFile && File.Exists(_imagePath))
                File.Delete(_imagePath);
        }

        private string CropWhiteMargin(string imagePath, byte threshold = 250)
        {
            using (Bitmap bmp = new Bitmap(imagePath))
            {
                int left = bmp.Width;
                int right = 0;
                int top = bmp.Height;
                int bottom = 0;

                for (int y = 0; y < bmp.Height; y++)
                {
                    for (int x = 0; x < bmp.Width; x++)
                    {
                        Color c = bmp.GetPixel(x, y);

                        // 흰색이 아닌 픽셀
                        if (c.R < threshold || c.G < threshold || c.B < threshold)
                        {
                            if (x < left) left = x;
                            if (x > right) right = x;
                            if (y < top) top = y;
                            if (y > bottom) bottom = y;
                        }
                    }
                }

                if (left >= right || top >= bottom)
                    return imagePath;

                Rectangle crop = Rectangle.FromLTRB(left, top, right + 1, bottom + 1);

                using (Bitmap cropped = bmp.Clone(crop, bmp.PixelFormat))
                {
                    string temp = Path.Combine(
                        Path.GetTempPath(),
                        Guid.NewGuid() + ".png");

                    cropped.Save(temp, ImageFormat.Png);

                    return temp;
                }
            }
        }

        private string CropImage(
            string imagePath,
            int left,
            int top,
            int right,
            int bottom)
        {
            using (Bitmap source = new Bitmap(imagePath))
            {
                int width = source.Width - left - right;
                int height = source.Height - top - bottom;

                if (width <= 0 || height <= 0)
                    throw new ArgumentException("Crop 크기가 이미지보다 큽니다.");

                Rectangle cropRect = new Rectangle(left, top, width, height);

                using (Bitmap cropped = new Bitmap(width, height))
                using (Graphics g = Graphics.FromImage(cropped))
                {
                    g.DrawImage(
                        source,
                        new Rectangle(0, 0, width, height),
                        cropRect,
                        GraphicsUnit.Pixel);

                    string tempPath = Path.Combine(
                        Path.GetTempPath(),
                        Guid.NewGuid().ToString() + ".png");

                    cropped.Save(tempPath, ImageFormat.Png);

                    return tempPath;
                }
            }
        }
    }

    public class ImageInsertOptions
    {
        // 위치
        public string CellAddress { get; set; }

        // 크기
        public double? Scale { get; set; }
        public double? Width { get; set; }
        public double? Height { get; set; }

        // Crop
        public bool Crop { get; set; } = false;
        public int CropLeft { get; set; }
        public int CropTop { get; set; }
        public int CropRight { get; set; }
        public int CropBottom { get; set; }
    }
}