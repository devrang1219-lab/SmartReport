using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WindowsFormsApp1.SortImage
{
    public partial class PreviewForm : Form
    {
        private readonly List<ImageItem> _images;

        private int _index;

        private PictureBox pictureBox;

        private Label lbInfo;

        private Button btnPrev;
        private Button btnNext;
        private Button btnFit;
        private Button btn100;

        private float _zoom = 1.0f;

        private Panel imagePanel;


        public PreviewForm(
            List<ImageItem> images,
            int startIndex)
        {
            _images = images;

            _index = startIndex;

            InitializeUI();

            ShowImage();
        }


        // =========================================================
        // UI
        // =========================================================

        private void InitializeUI()
        {
            Text = "이미지 크게 보기";

            Width = 1300;
            Height = 900;

            StartPosition =
                FormStartPosition.CenterParent;

            KeyPreview = true;

            KeyDown += PreviewForm_KeyDown;


            // -----------------------
            // 상단
            // -----------------------

            var top =
                new Panel
                {
                    Dock = DockStyle.Top,
                    Height = 50
                };


            Controls.Add(top);


            btnPrev =
                new Button
                {
                    Text = "◀ 이전",
                    Left = 10,
                    Top = 8,
                    Width = 90,
                    Height = 34
                };

            btnPrev.Click +=
                (s, e) => Previous();

            top.Controls.Add(btnPrev);


            btnNext =
                new Button
                {
                    Text = "다음 ▶",
                    Left = 110,
                    Top = 8,
                    Width = 90,
                    Height = 34
                };

            btnNext.Click +=
                (s, e) => Next();

            top.Controls.Add(btnNext);


            btnFit =
                new Button
                {
                    Text = "화면 맞춤",
                    Left = 220,
                    Top = 8,
                    Width = 90,
                    Height = 34
                };

            btnFit.Click +=
                (s, e) => FitImage();

            top.Controls.Add(btnFit);


            btn100 =
                new Button
                {
                    Text = "100%",
                    Left = 320,
                    Top = 8,
                    Width = 70,
                    Height = 34
                };

            btn100.Click +=
                (s, e) => Zoom100();

            top.Controls.Add(btn100);


            lbInfo =
                new Label
                {
                    AutoSize = true,
                    Left = 420,
                    Top = 17
                };

            top.Controls.Add(lbInfo);


            // =====================================================
            // 이미지 영역
            // =====================================================

            imagePanel =
                new Panel
                {
                    Dock = DockStyle.Fill,
                    AutoScroll = true,
                    BackColor =
                        Color.FromArgb(
                            30,
                            30,
                            30)
                };


            Controls.Add(imagePanel);


            pictureBox =
                new PictureBox
                {
                    SizeMode =
                        PictureBoxSizeMode.Zoom,

                    BackColor =
                        Color.FromArgb(
                            30,
                            30,
                            30)
                };


            pictureBox.MouseWheel +=
                PictureBox_MouseWheel;

            pictureBox.MouseEnter +=
                (s, e) =>
                pictureBox.Focus();


            imagePanel.Controls.Add(
                pictureBox);


            Shown +=
                (s, e) => FitImage();


            Resize +=
                (s, e) =>
                {
                    if (_zoom == 0)
                        FitImage();
                };
        }


        // =========================================================
        // 이미지 표시
        // =========================================================

        private void ShowImage()
        {
            if (_images.Count == 0)
                return;


            if (_index < 0)
                _index = 0;

            if (_index >= _images.Count)
                _index =
                    _images.Count - 1;


            if (pictureBox.Image != null)
            {
                Image old =
                    pictureBox.Image;

                pictureBox.Image = null;

                old.Dispose();
            }


            ImageItem item =
                _images[_index];


            pictureBox.Image =
                FormSortImage.LoadImageWithoutLock(
                    item.FullPath);


            Text =
                $"{item.Order:D3} - {item.FileName}";


            lbInfo.Text =
                $"{_index + 1} / {_images.Count}    {item.FileName}";


            FitImage();
        }


        // =========================================================
        // 이전
        // =========================================================

        private void Previous()
        {
            if (_index <= 0)
                return;

            _index--;

            ShowImage();
        }


        // =========================================================
        // 다음
        // =========================================================

        private void Next()
        {
            if (_index >=
                _images.Count - 1)
                return;

            _index++;

            ShowImage();
        }


        // =========================================================
        // 화면 맞춤
        // =========================================================

        private void FitImage()
        {
            if (pictureBox.Image == null)
                return;


            _zoom = 0;


            int panelWidth =
                Math.Max(
                    imagePanel.ClientSize.Width - 20,
                    100);

            int panelHeight =
                Math.Max(
                    imagePanel.ClientSize.Height - 20,
                    100);


            double imageRatio =
                (double)pictureBox.Image.Width /
                pictureBox.Image.Height;


            double panelRatio =
                (double)panelWidth /
                panelHeight;


            int width;
            int height;


            if (imageRatio >
                panelRatio)
            {
                width =
                    panelWidth;

                height =
                    (int)(
                        width /
                        imageRatio);
            }
            else
            {
                height =
                    panelHeight;

                width =
                    (int)(
                        height *
                        imageRatio);
            }


            pictureBox.Size =
                new Size(
                    width,
                    height);


            pictureBox.Location =
                new Point(
                    Math.Max(
                        0,
                        (imagePanel.ClientSize.Width - width) / 2),

                    Math.Max(
                        0,
                        (imagePanel.ClientSize.Height - height) / 2));
        }


        // =========================================================
        // 100%
        // =========================================================

        private void Zoom100()
        {
            if (pictureBox.Image == null)
                return;


            _zoom = 1.0f;


            pictureBox.Size =
                new Size(
                    pictureBox.Image.Width,
                    pictureBox.Image.Height);


            pictureBox.Location =
                new Point(0, 0);
        }


        // =========================================================
        // Mouse Wheel Zoom
        // =========================================================

        private void PictureBox_MouseWheel(
            object sender,
            MouseEventArgs e)
        {
            if (pictureBox.Image == null)
                return;


            // 화면맞춤 상태에서 확대 시작
            if (_zoom == 0)
            {
                _zoom =
                    (float)pictureBox.Width /
                    pictureBox.Image.Width;
            }


            if (e.Delta > 0)
            {
                _zoom *= 1.15f;
            }
            else
            {
                _zoom /= 1.15f;
            }


            if (_zoom < 0.05f)
                _zoom = 0.05f;

            if (_zoom > 8.0f)
                _zoom = 8.0f;


            int width =
                (int)(
                    pictureBox.Image.Width *
                    _zoom);

            int height =
                (int)(
                    pictureBox.Image.Height *
                    _zoom);


            pictureBox.Size =
                new Size(
                    width,
                    height);


            pictureBox.Location =
                new Point(0, 0);


            lbInfo.Text =
                $"{_index + 1} / {_images.Count}    " +
                $"{_zoom * 100:F0}%";
        }


        // =========================================================
        // 키보드
        // =========================================================

        private void PreviewForm_KeyDown(
            object sender,
            KeyEventArgs e)
        {
            switch (e.KeyCode)
            {
                case Keys.Left:

                    Previous();

                    e.Handled = true;

                    break;


                case Keys.Right:

                    Next();

                    e.Handled = true;

                    break;


                case Keys.Escape:

                    Close();

                    e.Handled = true;

                    break;


                case Keys.Home:

                    _index = 0;

                    ShowImage();

                    break;


                case Keys.End:

                    _index =
                        _images.Count - 1;

                    ShowImage();

                    break;
            }
        }


        // =========================================================
        // 종료
        // =========================================================

        protected override void OnFormClosed(
            FormClosedEventArgs e)
        {
            if (pictureBox.Image != null)
            {
                Image image =
                    pictureBox.Image;

                pictureBox.Image = null;

                image.Dispose();
            }


            base.OnFormClosed(e);
        }
    }
}
