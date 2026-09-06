using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Tesseract;

namespace WindowsFormsApp1.Ocr
{
    public class OcrReader : IDisposable
    {
        private readonly TesseractEngine _engine;

        public OcrReader(
            string whitelist = "0123456789.-:%/ ")
        {
            string tessPath = Path.Combine(
                AppDomain.CurrentDomain.BaseDirectory,
                "Ocr",
                "tessdata");

            _engine = new TesseractEngine(
                tessPath,
                "eng",
                EngineMode.Default);

            SetWhitelist(whitelist);
        }

        public OcrReader(
            string tessDataPath,
            string whitelist)
        {
            _engine = new TesseractEngine(
                tessDataPath,
                "eng",
                EngineMode.Default);

            SetWhitelist(whitelist);
        }

        private void SetWhitelist(string whitelist)
        {
            if (!string.IsNullOrEmpty(whitelist))
            {
                _engine.SetVariable(
                    "tessedit_char_whitelist",
                    whitelist);
            }

            _engine.SetVariable(
                "preserve_interword_spaces",
                "1");
        }

        public OcrReadResult Read(
            Bitmap bitmap,
            PageSegMode pageSegMode = PageSegMode.SingleLine)
        {
            if (bitmap == null)
                throw new ArgumentNullException(nameof(bitmap));

            using (Pix pix = PixConverter.ToPix(bitmap))
            using (Page page = _engine.Process(
                pix,
                pageSegMode))
            {
                return new OcrReadResult
                {
                    Text = page.GetText()?.Trim() ?? "",
                    Confidence = page.GetMeanConfidence()
                };
            }
        }

        public void Dispose()
        {
            _engine?.Dispose();
        }
    }

    public class OcrReadResult
    {
        public string Text { get; set; }

        /// <summary>
        /// Tesseract confidence.
        /// 일반적으로 0.0 ~ 1.0
        /// </summary>
        public float Confidence { get; set; }

        public bool HasText
        {
            get
            {
                return !string.IsNullOrWhiteSpace(Text);
            }
        }

        public override string ToString()
        {
            return $"{Text} ({Confidence:P1})";
        }
    }
}
