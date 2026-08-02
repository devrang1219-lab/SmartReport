using PdfiumViewer;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;


//using Tesseract;
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

        private bool IsRotatedImage(Excel.Shape shape)
        {
            double rot = shape.Rotation % 360;
            if (rot < 0) rot += 360;

            bool rotated90 =
                Math.Abs(rot - 90) < 1 ||
                Math.Abs(rot - 270) < 1;

            if (rotated90) return true;
            return false;
        }


        public Excel.Shape InsertFit(
            string cellFrom,
            string cellTo,
            ImageInsertOptions option = null)
        {
            return InsertFit(0, cellFrom, cellTo, option);
        }

        public string ResizeImage(string path, int maxWidth, int maxHeight)
        {
            using (Image src = Image.FromFile(path))
            {
                int width = src.Width;
                int height = src.Height;

                double ratio = Math.Min(
                    (double)maxWidth / width,
                    (double)maxHeight / height);

                // 원본보다 크게 만들지 않음
                if (ratio >= 1.0)
                    return path;

                int newWidth = (int)Math.Round(width * ratio);
                int newHeight = (int)Math.Round(height * ratio);

                using (Bitmap bmp = new Bitmap(newWidth, newHeight))
                {
                    using (Graphics g = Graphics.FromImage(bmp))
                    {
                        g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                        g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality;
                        g.PixelOffsetMode = System.Drawing.Drawing2D.PixelOffsetMode.HighQuality;
                        g.CompositingQuality = System.Drawing.Drawing2D.CompositingQuality.HighQuality;

                        g.DrawImage(src, 0, 0, newWidth, newHeight);
                    }

                    string newPath = Path.Combine(
                        Path.GetDirectoryName(path),
                        "resize_" + Path.GetFileName(path));

                    // JPEG 품질 설정 (90 정도 추천)
                    ImageCodecInfo jpgEncoder = ImageCodecInfo.GetImageEncoders()
                        .First(c => c.FormatID == ImageFormat.Jpeg.Guid);

                    EncoderParameters ep = new EncoderParameters(1);
                    ep.Param[0] = new EncoderParameter(System.Drawing.Imaging.Encoder.Quality, 90L);

                    bmp.Save(newPath, jpgEncoder, ep);

                    return newPath;
                }
            }
        }

        public static string FixImageOrientation(string path)
        {
            using (Image img = Image.FromFile(path))
            {
                const int OrientationId = 0x0112;

                if (!img.PropertyIdList.Contains(OrientationId))
                    return path;

                var prop = img.GetPropertyItem(OrientationId);

                int orientation = BitConverter.ToUInt16(prop.Value, 0);

                RotateFlipType rotateFlip = RotateFlipType.RotateNoneFlipNone;

                switch (orientation)
                {
                    case 2:
                        rotateFlip = RotateFlipType.RotateNoneFlipX;
                        break;

                    case 3:
                        rotateFlip = RotateFlipType.Rotate180FlipNone;
                        break;

                    case 4:
                        rotateFlip = RotateFlipType.Rotate180FlipX;
                        break;

                    case 5:
                        rotateFlip = RotateFlipType.Rotate90FlipX;
                        break;

                    case 6:
                        rotateFlip = RotateFlipType.Rotate90FlipNone;
                        break;

                    case 7:
                        rotateFlip = RotateFlipType.Rotate270FlipX;
                        break;

                    case 8:
                        rotateFlip = RotateFlipType.Rotate270FlipNone;
                        break;
                }

                if (rotateFlip != RotateFlipType.RotateNoneFlipNone)
                {
                    img.RotateFlip(rotateFlip);

                    // EXIF 방향 제거
                    img.RemovePropertyItem(OrientationId);
                }

                string newPath = Path.Combine(
                    Path.GetDirectoryName(path),
                    "fixed_" + Path.GetFileName(path));

                img.Save(newPath, ImageFormat.Jpeg);

                return newPath;
            }
        }

        private void DeleteTempFile(string path, string originalPath)
        {
            if (string.IsNullOrEmpty(path))
                return;

            if (string.Equals(
                path,
                originalPath,
                StringComparison.OrdinalIgnoreCase))
                return;

            try
            {
                if (File.Exists(path))
                    File.Delete(path);
            }
            catch (Exception ex)
            {
                Debug.WriteLine(
                    $"임시 이미지 삭제 실패: {path}, {ex.Message}");
            }
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

            Excel.Shape pic = null;
            string tempFixed = null;
            string tempResized = null;
            string imagePath = null;

            try
            {

                imagePath = GetImagePath(page, option);

                tempFixed = FixImageOrientation(imagePath);

                //tempResized = ResizeImage(
                //    tempFixed,
                //    1024);

                Excel.Range area = _ws.Range[cellFrom, cellTo];
                Debug.WriteLine($"cell from: {cellFrom}, to: {cellTo}");

                Debug.WriteLine($"image name: {imagePath}");
                Debug.WriteLine($"area: {area.Left}, {area.Top}, {area.Width}, {area.Height}");

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

                //Point(1/72인치)를 사용하고 일반적인 화면은 96dpi이므로 pixel = point × 96 / 72 정도로 계산
                //화질을 위해 조금 더 크게 계산 1.3배 정도로 계산
                int pixelWidth = (int)Math.Ceiling(width * 96.0 / 72.0 * 1.3);
                int pixelHeight = (int)Math.Ceiling(height * 96.0 / 72.0 * 1.3);

                tempResized = ResizeImage(
                                tempFixed,
                                pixelWidth,
                                pixelHeight);

                pic = _ws.Shapes.AddPicture(
                tempResized,
                Office.MsoTriState.msoFalse,
                Office.MsoTriState.msoTrue,
                left,
                top,
                -1,
                -1);

                Debug.WriteLine($"After AddPicture");
                Debug.WriteLine($"Left={pic.Left}");
                Debug.WriteLine($"Top={pic.Top}");
                Debug.WriteLine($"Width={pic.Width}");
                Debug.WriteLine($"Height={pic.Height}");
                Debug.WriteLine($"Rotation={pic.Rotation}");

                Debug.WriteLine($"image left: {left}, top: {top}, width: {width}, height: {height}");
                Debug.WriteLine($"area left : {_ws.Range[cellFrom].Left}, top: {_ws.Range[cellFrom].Top}," +
                    $"width: {_ws.Range[cellTo].Width}, height: {_ws.Range[cellTo].Height}");

                Debug.WriteLine($"option.KeepAspectRatio: {option.KeepAspectRatio}");

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

                    if (IsRotatedImage(pic))
                    {

                        Debug.WriteLine($"Rotated Image: {pic.Name}, Rotation: {pic.Rotation}");
                        pic.Width = height - option.GapLeft - option.GapRight;
                        pic.Height = width - option.GapTop - option.GapBottom;
                    }
                }

                Debug.WriteLine($"pic left: {pic.Left}, top: {pic.Top}, width: {pic.Width}, height: {pic.Height}");

            }
            catch
            {

            }
            finally
            {

                DeleteTempFile(tempFixed, imagePath);
                DeleteTempFile(tempResized, imagePath);
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

                System.Drawing.Rectangle crop = System.Drawing.Rectangle.FromLTRB(left, top, right + 1, bottom + 1);

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

                System.Drawing.Rectangle cropRect = new System.Drawing.Rectangle(left, top, width, height);

                using (Bitmap cropped = new Bitmap(width, height))
                using (Graphics g = Graphics.FromImage(cropped))
                {
                    g.DrawImage(
                        source,
                        new System.Drawing.Rectangle(0, 0, width, height),
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