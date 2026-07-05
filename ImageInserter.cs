using PdfiumViewer;
using System;
using System.Collections.Generic;
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
        private readonly List<string> _imagePaths = new List<string>(); // 삽입할 이미지들
        private readonly List<string> _tempFiles = new List<string>();  // 삭제할 임시파일들
        public int ImageCount => _imagePaths.Count;


        public ImageInserter(Excel.Worksheet ws, string filePath)
        {
            _ws = ws;


            int dpi = 300;
            int width = 2480;
            int height = 3508;

            if (Path.GetExtension(filePath).Equals(".pdf", StringComparison.OrdinalIgnoreCase))
            {
                using (var pdf = PdfDocument.Load(filePath))
                {
                    for (int page = 0; page < pdf.PageCount; page++)
                    {
                        string tempPath = Path.Combine(
                            Path.GetTempPath(),
                            Guid.NewGuid() + ".png");

                        using (Image img = pdf.Render(
                            page,
                            width,
                            height,
                            dpi,
                            dpi,
                            true))
                        {
                            img.Save(tempPath, ImageFormat.Png);
                        }

                        _imagePaths.Add(tempPath);
                        _tempFiles.Add(tempPath);
                    }
                }
            }
            else
            {
                _imagePaths.Add(filePath);
            }
        }

        public Excel.Shape Insert(int page, string cellAddress, ImageInsertOptions option = null)
        {
            if (option == null)
                option = new ImageInsertOptions();

            if (page < 0 || page >= _imagePaths.Count)
                throw new ArgumentOutOfRangeException(nameof(page));

            string imagePath = GetImagePath(page, option);

            Excel.Range cell = _ws.Range[cellAddress];

            Excel.Shape picture = _ws.Shapes.AddPicture(
                imagePath,
                Office.MsoTriState.msoFalse,
                Office.MsoTriState.msoTrue,
                (float)cell.Left,
                (float)cell.Top,
                -1,
                -1);

            if (option.Width.HasValue && option.Height.HasValue)
            {
                picture.LockAspectRatio = Office.MsoTriState.msoFalse;
                picture.Width = (float)option.Width.Value;
                picture.Height = (float)option.Height.Value;
            }
            else if (option.Scale.HasValue)
            {
                picture.LockAspectRatio = Office.MsoTriState.msoTrue;
                picture.ScaleWidth((float)option.Scale.Value, Office.MsoTriState.msoTrue);
                picture.ScaleHeight((float)option.Scale.Value, Office.MsoTriState.msoTrue);
            }

            return picture;
        }

        private bool NeedCrop(ImageInsertOptions option)
        {
            return option.CropLeft > 0 ||
                   option.CropTop > 0 ||
                   option.CropRight > 0 ||
                   option.CropBottom > 0;
        }

        private string GetImagePath(int page, ImageInsertOptions option)
        {
            string imagePath = _imagePaths[page];

            if (NeedCrop(option))
            {
                imagePath = CropImage(
                    imagePath,
                    option.CropLeft,
                    option.CropTop,
                    option.CropRight,
                    option.CropBottom);
            }

            return imagePath;
        }

        public Excel.Shape Insert(string cellAddress, ImageInsertOptions option)
        {
            return Insert(0, cellAddress, option);
        }

        public Excel.Shape Insert(string cellAddress, double scale)
        {
            return Insert(cellAddress, new ImageInsertOptions
            {
                Scale = scale
            });
        }

        public Excel.Shape Insert(string cellAddress, double width, double height)
        {
            return Insert(cellAddress, new ImageInsertOptions
            {
                Width = width,
                Height = height
            });
        }

        public void Dispose()
        {
            foreach (string file in _tempFiles)
            {
                try
                {
                    if (File.Exists(file))
                        File.Delete(file);
                }
                catch { }
            }
        }


        public Excel.Shape InsertFit(
            string cellFrom,
            string cellTo,
            ImageInsertOptions option = null)
        {
            return InsertFit(0, cellFrom, cellTo, option);
        }

        public Excel.Shape InsertFit(
            int page,
            string cellFrom,
            string cellTo,
            ImageInsertOptions option = null)
        {
            if(option == null)
                option = new ImageInsertOptions();

            if (page < 0 || page >= _imagePaths.Count)
                throw new ArgumentOutOfRangeException(nameof(page));

            string imagePath = GetImagePath(page, option);

            Excel.Range area = _ws.Range[cellFrom, cellTo];

            //float left = (float)area.Left;
            //float top = (float)area.Top;
            //float width = (float)area.Width;
            //float height = (float)area.Height;

            float left = (float)_ws.Range[cellFrom].Left;
            float top = (float)_ws.Range[cellFrom].Top;

            float width =
                (float)(_ws.Range[cellTo].Left + _ws.Range[cellTo].Width - left);

            float height =
                (float)(_ws.Range[cellTo].Top + _ws.Range[cellTo].Height - top);

            Excel.Shape pic = _ws.Shapes.AddPicture(
                imagePath,
                Office.MsoTriState.msoFalse,
                Office.MsoTriState.msoTrue,
                left,
                top,
                -1,
                -1);

            if (option.KeepAspectRatio)
            {
                pic.LockAspectRatio = Office.MsoTriState.msoTrue;

                double scaleX = width / pic.Width;
                double scaleY = height / pic.Height;
                double scale = Math.Min(scaleX, scaleY);

                pic.ScaleWidth((float)scale, Office.MsoTriState.msoTrue);
                pic.ScaleHeight((float)scale, Office.MsoTriState.msoTrue);

                // 가운데 정렬
                pic.Left = left + (width - pic.Width) / 2f + option.GapLeft;
                pic.Top = top + (height - pic.Height) / 2f + option.GapTop;
                pic.Width -= option.GapLeft + option.GapRight;
                pic.Height -= option.GapTop + option.GapBottom;
            }
            else
            {
                pic.LockAspectRatio = Office.MsoTriState.msoFalse;
                pic.Left = left + option.GapLeft;
                pic.Top = top + option.GapTop;
                pic.Width = width - option.GapLeft - option.GapRight;
                pic.Height = height - option.GapTop - option.GapBottom;
            }

            return pic;
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


        // 비율 유지
        public bool KeepAspectRatio { get; set; } = false;

        // Crop
        public int CropLeft { get; set; }
        public int CropTop { get; set; }
        public int CropRight { get; set; }
        public int CropBottom { get; set; }

        // Gap (Point)
        public float GapLeft { get; set; } = 2.8f;
        public float GapTop { get; set; } = 2.8f;
        public float GapRight { get; set; } = 2.8f;
        public float GapBottom { get; set; } = 2.8f;
    }
}