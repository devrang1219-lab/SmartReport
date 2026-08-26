using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace WindowsFormsApp1.SortImage
{
    public partial class FormSortImage : Form
    {
        private readonly List<ImageItem> _images
            = new List<ImageItem>();

        private string _currentFolder;
            

        // 드래그 여부
        private bool _dragging;

        public FormSortImage()
        {
            InitializeForm();
            InitializeUI();
        }


        // =========================================================
        // Form 기본설정
        // =========================================================

        private void InitializeForm()
        {
            Text = "이미지 순서 정리";
            Width = 1500;
            Height = 900;

            StartPosition = FormStartPosition.CenterScreen;

            MinimumSize = new Size(1000, 650);
        }


        // =========================================================
        // UI 생성
        // =========================================================

        private void InitializeUI()
        {
            // -----------------------------------
            // 최상위
            // -----------------------------------

            this.root = new TableLayoutPanel();

            this.root.Dock = DockStyle.Fill;
            this.root.RowCount = 3;
            this.root.ColumnCount = 1;

            this.root.RowStyles.Add(
                new RowStyle(SizeType.Absolute, 55));

            this.root.RowStyles.Add(
                new RowStyle(SizeType.Percent, 100));

            this.root.RowStyles.Add(
                new RowStyle(SizeType.Absolute, 65));

            Controls.Add(this.root);


            // =====================================================
            // 상단
            // =====================================================

            topPanel = new Panel();

            topPanel.Dock = DockStyle.Fill;

            root.Controls.Add(topPanel, 0, 0);


            lblFolder = new Label();

            lblFolder.Text = "폴더";
            lblFolder.AutoSize = true;
            lblFolder.Left = 10;
            lblFolder.Top = 19;

            topPanel.Controls.Add(lblFolder);


            tbFolder = new TextBox();

            tbFolder.Left = 55;
            tbFolder.Top = 13;
            tbFolder.Width = 550;
            tbFolder.Anchor =
                AnchorStyles.Left |
                AnchorStyles.Top |
                AnchorStyles.Right;

            topPanel.Controls.Add(tbFolder);


            btnFolder = new Button();

            btnFolder.Text = "폴더 선택";
            btnFolder.Left = 615;
            btnFolder.Top = 11;
            btnFolder.Width = 100;
            btnFolder.Height = 32;

            btnFolder.Click += BtnFolder_Click;

            topPanel.Controls.Add(btnFolder);


            btnRefresh = new Button();

            btnRefresh.Text = "새로고침";
            btnRefresh.Left = 725;
            btnRefresh.Top = 11;
            btnRefresh.Width = 90;
            btnRefresh.Height = 32;

            btnRefresh.Click += BtnRefresh_Click;

            topPanel.Controls.Add(btnRefresh);


            lblSort = new Label();

            lblSort.Text = "정렬";
            lblSort.Left = 840;
            lblSort.Top = 19;
            lblSort.AutoSize = true;

            topPanel.Controls.Add(lblSort);


            cbSort = new ComboBox();

            cbSort.Left = 880;
            cbSort.Top = 13;
            cbSort.Width = 150;

            cbSort.DropDownStyle =
                ComboBoxStyle.DropDownList;

            cbSort.Items.Add("사용자 지정");
            cbSort.Items.Add("파일명 ↑");
            cbSort.Items.Add("파일명 ↓");
            cbSort.Items.Add("수정일 ↑");
            cbSort.Items.Add("수정일 ↓");

            cbSort.SelectedIndex = 0;

            cbSort.SelectedIndexChanged += CbSort_SelectedIndexChanged;

            topPanel.Controls.Add(cbSort);


            // =====================================================
            // 가운데 Split
            // =====================================================

            mainSplit = new SplitContainer();

            mainSplit.Dock = DockStyle.Fill;

            mainSplit.SplitterDistance = 220;
            mainSplit.FixedPanel = FixedPanel.Panel1;

            root.Controls.Add(mainSplit, 0, 1);


            // =====================================================
            // 왼쪽 폴더 Tree
            // =====================================================

            tvFolders = new TreeView();

            tvFolders.Dock = DockStyle.Fill;
            tvFolders.HideSelection = false;

            tvFolders.AfterSelect += TvFolders_AfterSelect;

            mainSplit.Panel1.Controls.Add(tvFolders);


            // =====================================================
            // 가운데 + 오른쪽 Split
            // =====================================================

            imageSplit = new SplitContainer();

            imageSplit.Dock = DockStyle.Fill;

            imageSplit.SplitterDistance = 750;

            mainSplit.Panel2.Controls.Add(imageSplit);


            // =====================================================
            // 썸네일 ListView
            // =====================================================

            lvImages = new ListView();

            lvImages.Dock = DockStyle.Fill;

            lvImages.View = View.LargeIcon;

            lvImages.MultiSelect = true;

            lvImages.HideSelection = false;

            lvImages.AllowDrop = true;

            lvImages.BackColor = Color.White;

            imageList = new ImageList();

            imageList.ImageSize = new Size(150, 110);

            imageList.ColorDepth =
                ColorDepth.Depth32Bit;

            lvImages.LargeImageList = imageList;


            lvImages.SelectedIndexChanged
                += LvImages_SelectedIndexChanged;

            lvImages.DoubleClick
                += LvImages_DoubleClick;

            lvImages.ItemDrag
                += LvImages_ItemDrag;

            lvImages.DragEnter
                += LvImages_DragEnter;

            lvImages.DragOver
                += LvImages_DragOver;

            lvImages.DragDrop
                += LvImages_DragDrop;

            imageSplit.Panel1.Controls.Add(lvImages);


            // =====================================================
            // 오른쪽 Preview
            // =====================================================

            previewPanel = new TableLayoutPanel();

            previewPanel.Dock = DockStyle.Fill;

            previewPanel.RowCount = 4;
            previewPanel.ColumnCount = 1;

            previewPanel.RowStyles.Add(
                new RowStyle(SizeType.Absolute, 35));

            previewPanel.RowStyles.Add(
                new RowStyle(SizeType.Percent, 100));

            previewPanel.RowStyles.Add(
                new RowStyle(SizeType.Absolute, 50));

            previewPanel.RowStyles.Add(
                new RowStyle(SizeType.Absolute, 45));

            imageSplit.Panel2.Controls.Add(previewPanel);


            lbFileName = new Label();

            lbFileName.Dock = DockStyle.Fill;
            lbFileName.TextAlign =
                ContentAlignment.MiddleCenter;

            lbFileName.Text = "이미지를 선택하세요.";

            previewPanel.Controls.Add(
                lbFileName, 0, 0);


            pbPreview = new PictureBox();

            pbPreview.Dock = DockStyle.Fill;

            pbPreview.SizeMode =
                PictureBoxSizeMode.Zoom;

            pbPreview.BackColor =
                Color.FromArgb(40, 40, 40);

            pbPreview.DoubleClick += PbPreview_DoubleClick;

            previewPanel.Controls.Add(
                pbPreview, 0, 1);


            // 이전 다음
            previewButtonPanel = new Panel();

            previewButtonPanel.Dock = DockStyle.Fill;

            btnPrev = new Button();

            btnPrev.Text = "◀ 이전";
            btnPrev.Width = 100;
            btnPrev.Height = 34;
            btnPrev.Left = 30;
            btnPrev.Top = 7;

            btnPrev.Click += BtnPrev_Click;

            previewButtonPanel.Controls.Add(btnPrev);


            btnNext = new Button();

            btnNext.Text = "다음 ▶";
            btnNext.Width = 100;
            btnNext.Height = 34;
            btnNext.Left = 140;
            btnNext.Top = 7;

            btnNext.Click += BtnNext_Click;

            previewButtonPanel.Controls.Add(btnNext);


            previewPanel.Controls.Add(
                previewButtonPanel, 0, 2);


            lbInfo = new Label();

            lbInfo.Dock = DockStyle.Fill;
            lbInfo.TextAlign =
                ContentAlignment.MiddleCenter;

            previewPanel.Controls.Add(
                lbInfo, 0, 3);


            // =====================================================
            // 하단
            // =====================================================

            bottomPanel = new Panel();

            bottomPanel.Dock = DockStyle.Fill;

            root.Controls.Add(bottomPanel, 0, 2);


            btnUp = new Button();

            btnUp.Text = "▲ 위로";
            btnUp.Left = 10;
            btnUp.Top = 14;
            btnUp.Width = 80;
            btnUp.Height = 35;

            btnUp.Click += BtnUp_Click;

            bottomPanel.Controls.Add(btnUp);


            btnDown = new Button();

            btnDown.Text = "▼ 아래로";
            btnDown.Left = 100;
            btnDown.Top = 14;
            btnDown.Width = 80;
            btnDown.Height = 35;

            btnDown.Click += BtnDown_Click;

            bottomPanel.Controls.Add(btnDown);


            btnDelete = new Button();

            btnDelete.Text = "목록 제외";
            btnDelete.Left = 190;
            btnDelete.Top = 14;
            btnDelete.Width = 90;
            btnDelete.Height = 35;

            btnDelete.Click += BtnDelete_Click;

            bottomPanel.Controls.Add(btnDelete);


            lblThumb = new Label();

            lblThumb.Text = "썸네일";
            lblThumb.Left = 310;
            lblThumb.Top = 24;
            lblThumb.AutoSize = true;

            bottomPanel.Controls.Add(lblThumb);


            tbThumbSize = new TrackBar();

            tbThumbSize.Left = 365;
            tbThumbSize.Top = 5;
            tbThumbSize.Width = 180;

            tbThumbSize.Minimum = 80;
            tbThumbSize.Maximum = 250;
            tbThumbSize.Value = 150;

            tbThumbSize.TickFrequency = 20;

            tbThumbSize.ValueChanged += TbThumbSize_ValueChanged;

            bottomPanel.Controls.Add(tbThumbSize);


            chkKeepOriginalName = new CheckBox();

            chkKeepOriginalName.Text =
                "순번 뒤에 원본 파일명 유지";

            chkKeepOriginalName.Checked = true;

            chkKeepOriginalName.Left = 560;
            chkKeepOriginalName.Top = 22;
            chkKeepOriginalName.AutoSize = true;

            bottomPanel.Controls.Add(chkKeepOriginalName);


            btnMakeFolder = new Button();

            btnMakeFolder.Text = "정렬본 만들기";

            btnMakeFolder.Width = 140;
            btnMakeFolder.Height = 40;

            btnMakeFolder.Top = 11;

            btnMakeFolder.Anchor =
                AnchorStyles.Top |
                AnchorStyles.Right;

            btnMakeFolder.Left =
                bottomPanel.Width - 160;

            btnMakeFolder.Click += BtnMakeFolder_Click;

            bottomPanel.Controls.Add(btnMakeFolder);


            //bottomPanel.Resize += (s, e) =>
            //{
            //    btnMakeFolder.Left =
            //        bottomPanel.ClientSize.Width -
            //        btnMakeFolder.Width - 15;
            //};
        }


        // =========================================================
        // 폴더 선택
        // =========================================================

        private void BtnFolder_Click(object sender, EventArgs e)
        {
            using (var dialog = new FolderBrowserDialog())
            {
                dialog.Description =
                    "이미지가 있는 폴더를 선택하세요.";

                if (dialog.ShowDialog() != DialogResult.OK)
                    return;

                LoadRootFolder(dialog.SelectedPath);
            }
        }


        // =========================================================
        // 루트 폴더 설정
        // =========================================================

        private void LoadRootFolder(string path)
        {
            if (!Directory.Exists(path))
                return;

            tbFolder.Text = path;

            tvFolders.Nodes.Clear();

            var root =
                new TreeNode(
                    new DirectoryInfo(path).Name);

            root.Tag = path;

            tvFolders.Nodes.Add(root);

            AddSubFolders(root);

            root.Expand();

            tvFolders.SelectedNode = root;
        }


        // =========================================================
        // 하위폴더
        // =========================================================

        private void AddSubFolders(TreeNode parent)
        {
            string path =
                Convert.ToString(parent.Tag);

            try
            {
                foreach (string dir in
                    Directory.GetDirectories(path))
                {
                    var di =
                        new DirectoryInfo(dir);

                    var node =
                        new TreeNode(di.Name);

                    node.Tag = dir;

                    parent.Nodes.Add(node);

                    // 한 단계 정도만 미리 추가
                    try
                    {
                        if (Directory.GetDirectories(dir).Length > 0)
                        {
                            node.Nodes.Add(
                                new TreeNode("..."));
                        }
                    }
                    catch
                    {
                    }
                }
            }
            catch
            {
            }
        }


        // =========================================================
        // Tree 폴더 선택
        // =========================================================

        private void TvFolders_AfterSelect(
            object sender,
            TreeViewEventArgs e)
        {
            string path =
                Convert.ToString(e.Node.Tag);

            if (string.IsNullOrWhiteSpace(path))
                return;

            // ... 임시 노드였다면 실제 하위폴더 구성
            if (e.Node.Nodes.Count == 1 &&
                e.Node.Nodes[0].Text == "...")
            {
                e.Node.Nodes.Clear();

                AddSubFolders(e.Node);
            }

            LoadImages(path);
        }


        // =========================================================
        // 이미지 불러오기
        // =========================================================

        private void LoadImages(string folder)
        {
            if (!Directory.Exists(folder))
                return;

            _currentFolder = folder;

            tbFolder.Text = folder;

            ClearPreview();

            DisposeThumbnailImages();

            _images.Clear();

            lvImages.Items.Clear();

            imageList.Images.Clear();

            try
            {
                var files =
                    Directory.GetFiles(folder)
                    .Where(IsImageFile)
                    .OrderBy(
                        x => Path.GetFileName(x),
                        StringComparer.CurrentCultureIgnoreCase)
                    .ToList();


                foreach (string file in files)
                {
                    var fi =
                        new FileInfo(file);

                    _images.Add(
                        new ImageItem
                        {
                            FileName = fi.Name,
                            FullPath = fi.FullName,
                            LastWriteTime =
                                fi.LastWriteTime
                        });
                }

                Reorder();

                cbSort.SelectedIndex = 0;
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    ex.Message,
                    "오류",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }


        // =========================================================
        // 이미지 확장자
        // =========================================================

        private bool IsImageFile(string file)
        {
            string ext =
                Path.GetExtension(file)
                .ToLowerInvariant();

            return ext == ".jpg"
                || ext == ".jpeg"
                || ext == ".png"
                || ext == ".bmp"
                || ext == ".gif"
                || ext == ".tif"
                || ext == ".tiff";
        }


        // =========================================================
        // 번호 재계산
        // =========================================================

        private void Reorder()
        {
            for (int i = 0; i < _images.Count; i++)
            {
                _images[i].Order = i + 1;
            }

            RefreshListView();
        }


        // =========================================================
        // ListView 다시 표시
        // =========================================================

        private void RefreshListView()
        {
            lvImages.BeginUpdate();

            try
            {
                DisposeThumbnailImages();

                lvImages.Items.Clear();
                imageList.Images.Clear();

                for (int i = 0; i < _images.Count; i++)
                {
                    ImageItem item = _images[i];

                    Image thumb =
                        CreateThumbnailWithNumber(
                            item.FullPath,
                            item.Order);

                    imageList.Images.Add(
                        item.FullPath,
                        thumb);

                    var lvItem =
                        new ListViewItem();

                    lvItem.Text =
                        $"{item.Order:D3}\r\n{item.FileName}";

                    lvItem.ImageKey =
                        item.FullPath;

                    lvItem.Tag = item;

                    lvImages.Items.Add(lvItem);
                }

                lbInfo.Text =
                    $"{_images.Count:N0}개 이미지";
            }
            finally
            {
                lvImages.EndUpdate();
            }
        }


        // =========================================================
        // 번호가 붙은 썸네일 생성
        // =========================================================

        private Image CreateThumbnailWithNumber(
            string path,
            int number)
        {
            int width =
                imageList.ImageSize.Width;

            int height =
                imageList.ImageSize.Height;

            Bitmap bitmap =
                new Bitmap(width, height);

            using (Graphics g =
                Graphics.FromImage(bitmap))
            {
                g.Clear(Color.White);

                using (Image image =
                    LoadImageWithoutLock(path))
                {
                    Rectangle target =
                        CalculateZoomRectangle(
                            image.Width,
                            image.Height,
                            width,
                            height);

                    g.DrawImage(
                        image,
                        target);
                }


                // 번호 배경
                Rectangle numberRect =
                    new Rectangle(
                        5,
                        5,
                        45,
                        28);

                using (Brush b =
                    new SolidBrush(
                        Color.FromArgb(
                            190,
                            Color.Black)))
                {
                    g.FillRectangle(
                        b,
                        numberRect);
                }

                using (Font font =
                    new Font(
                        "Segoe UI",
                        11,
                        FontStyle.Bold))
                using (Brush brush =
                    new SolidBrush(Color.White))
                {
                    string text =
                        number.ToString();

                    g.DrawString(
                        text,
                        font,
                        brush,
                        numberRect);
                }
            }

            return bitmap;
        }


        // =========================================================
        // 이미지 Zoom rectangle
        // =========================================================

        private Rectangle CalculateZoomRectangle(
            int imageWidth,
            int imageHeight,
            int boxWidth,
            int boxHeight)
        {
            double imageRatio =
                (double)imageWidth / imageHeight;

            double boxRatio =
                (double)boxWidth / boxHeight;

            int width;
            int height;

            if (imageRatio > boxRatio)
            {
                width = boxWidth;

                height =
                    (int)(boxWidth / imageRatio);
            }
            else
            {
                height = boxHeight;

                width =
                    (int)(boxHeight * imageRatio);
            }

            int x =
                (boxWidth - width) / 2;

            int y =
                (boxHeight - height) / 2;

            return new Rectangle(
                x,
                y,
                width,
                height);
        }


        // =========================================================
        // 파일 잠금 없이 Image 로드
        // =========================================================

        public static Image LoadImageWithoutLock(
            string path)
        {
            using (FileStream fs =
                new FileStream(
                    path,
                    FileMode.Open,
                    FileAccess.Read,
                    FileShare.ReadWrite))
            {
                using (Image temp =
                    Image.FromStream(fs))
                {
                    return new Bitmap(temp);
                }
            }
        }


        // =========================================================
        // 썸네일 메모리 해제
        // =========================================================

        private void DisposeThumbnailImages()
        {
            foreach (Image image
                in imageList.Images)
            {
                try
                {
                    image.Dispose();
                }
                catch
                {
                }
            }
        }


        // =========================================================
        // 선택 변경 → 큰 Preview
        // =========================================================

        private void LvImages_SelectedIndexChanged(
            object sender,
            EventArgs e)
        {
            if (lvImages.SelectedIndices.Count == 0)
                return;

            int index =
                lvImages.SelectedIndices[0];

            ShowPreview(index);
        }


        // =========================================================
        // Preview 표시
        // =========================================================

        private void ShowPreview(int index)
        {
            if (index < 0 ||
                index >= _images.Count)
                return;

            ImageItem item =
                _images[index];

            if (pbPreview.Image != null)
            {
                Image old =
                    pbPreview.Image;

                pbPreview.Image = null;

                old.Dispose();
            }

            try
            {
                pbPreview.Image =
                    LoadImageWithoutLock(
                        item.FullPath);

                lbFileName.Text =
                    $"{item.Order:D3}  {item.FileName}";

                lbInfo.Text =
                    $"{item.Order} / {_images.Count}";
            }
            catch (Exception ex)
            {
                lbFileName.Text =
                    ex.Message;
            }
        }


        // =========================================================
        // Preview 초기화
        // =========================================================

        private void ClearPreview()
        {
            if (pbPreview.Image != null)
            {
                Image old =
                    pbPreview.Image;

                pbPreview.Image = null;

                old.Dispose();
            }

            lbFileName.Text =
                "이미지를 선택하세요.";

            lbInfo.Text = "";
        }


        // =========================================================
        // 더블클릭 → 큰 보기
        // =========================================================

        private void LvImages_DoubleClick(
            object sender,
            EventArgs e)
        {
            OpenLargePreview();
        }


        private void PbPreview_DoubleClick(
            object sender,
            EventArgs e)
        {
            OpenLargePreview();
        }


        private void OpenLargePreview()
        {
            if (lvImages.SelectedIndices.Count == 0)
                return;

            int index =
                lvImages.SelectedIndices[0];

            using (var form =
                new PreviewForm(
                    _images,
                    index))
            {
                form.ShowDialog(this);
            }
        }


        // =========================================================
        // 이전
        // =========================================================

        private void BtnPrev_Click(
            object sender,
            EventArgs e)
        {
            MoveSelection(-1);
        }


        // =========================================================
        // 다음
        // =========================================================

        private void BtnNext_Click(
            object sender,
            EventArgs e)
        {
            MoveSelection(1);
        }


        private void MoveSelection(int step)
        {
            if (_images.Count == 0)
                return;

            int index = 0;

            if (lvImages.SelectedIndices.Count > 0)
            {
                index =
                    lvImages.SelectedIndices[0] + step;
            }

            if (index < 0)
                index = 0;

            if (index >= _images.Count)
                index = _images.Count - 1;

            lvImages.SelectedItems.Clear();

            lvImages.Items[index].Selected = true;

            lvImages.Items[index].Focused = true;

            lvImages.EnsureVisible(index);

            ShowPreview(index);
        }


        // =========================================================
        // Drag 시작
        // =========================================================

        private void LvImages_ItemDrag(
            object sender,
            ItemDragEventArgs e)
        {
            if (lvImages.SelectedIndices.Count == 0)
                return;

            _dragging = true;

            lvImages.DoDragDrop(
                "IMAGE_REORDER",
                DragDropEffects.Move);

            _dragging = false;
        }


        // =========================================================
        // DragEnter
        // =========================================================

        private void LvImages_DragEnter(
            object sender,
            DragEventArgs e)
        {
            if (_dragging)
            {
                e.Effect =
                    DragDropEffects.Move;
            }
        }


        // =========================================================
        // DragOver
        // =========================================================

        private void LvImages_DragOver(
            object sender,
            DragEventArgs e)
        {
            if (_dragging)
            {
                e.Effect =
                    DragDropEffects.Move;
            }
        }


        // =========================================================
        // Drag Drop 순서 변경
        // =========================================================

        private void LvImages_DragDrop(
            object sender,
            DragEventArgs e)
        {
            if (!_dragging)
                return;

            Point point =
                lvImages.PointToClient(
                    new Point(e.X, e.Y));

            ListViewItem target =
                lvImages.GetItemAt(
                    point.X,
                    point.Y);


            int targetIndex;

            if (target == null)
            {
                targetIndex =
                    _images.Count;
            }
            else
            {
                targetIndex =
                    target.Index;
            }


            // 선택된 항목 index
            List<int> selectedIndices =
                lvImages.SelectedIndices
                .Cast<int>()
                .OrderBy(x => x)
                .ToList();


            if (selectedIndices.Count == 0)
                return;


            // 이동 대상들
            List<ImageItem> moving =
                selectedIndices
                .Select(x => _images[x])
                .ToList();


            // 제거되면서 target index가 당겨지는 만큼 보정
            int removeBefore =
                selectedIndices
                .Count(x => x < targetIndex);

            int insertIndex =
                targetIndex - removeBefore;


            // 뒤에서부터 삭제
            for (int i =
                selectedIndices.Count - 1;
                i >= 0;
                i--)
            {
                _images.RemoveAt(
                    selectedIndices[i]);
            }


            if (insertIndex < 0)
                insertIndex = 0;

            if (insertIndex > _images.Count)
                insertIndex = _images.Count;


            _images.InsertRange(
                insertIndex,
                moving);


            Reorder();


            // 이동한 것 다시 선택
            lvImages.SelectedItems.Clear();

            for (int i = 0;
                i < moving.Count;
                i++)
            {
                int index =
                    insertIndex + i;

                if (index >= 0 &&
                    index < lvImages.Items.Count)
                {
                    lvImages.Items[index].Selected = true;
                }
            }


            if (insertIndex <
                lvImages.Items.Count)
            {
                lvImages.EnsureVisible(
                    insertIndex);
            }


            cbSort.SelectedIndex = 0;
        }


        // =========================================================
        // 선택 항목 위로
        // =========================================================

        private void BtnUp_Click(
            object sender,
            EventArgs e)
        {
            var indices =
                lvImages.SelectedIndices
                .Cast<int>()
                .OrderBy(x => x)
                .ToList();

            if (indices.Count == 0)
                return;

            if (indices.First() == 0)
                return;


            foreach (int index in indices)
            {
                ImageItem temp =
                    _images[index - 1];

                _images[index - 1] =
                    _images[index];

                _images[index] =
                    temp;
            }


            Reorder();


            lvImages.SelectedItems.Clear();

            foreach (int index in indices)
            {
                lvImages.Items[index - 1]
                    .Selected = true;
            }


            cbSort.SelectedIndex = 0;
        }


        // =========================================================
        // 선택 항목 아래로
        // =========================================================

        private void BtnDown_Click(
            object sender,
            EventArgs e)
        {
            var indices =
                lvImages.SelectedIndices
                .Cast<int>()
                .OrderByDescending(x => x)
                .ToList();

            if (indices.Count == 0)
                return;

            if (indices.First() ==
                _images.Count - 1)
                return;


            foreach (int index in indices)
            {
                ImageItem temp =
                    _images[index + 1];

                _images[index + 1] =
                    _images[index];

                _images[index] =
                    temp;
            }


            Reorder();


            lvImages.SelectedItems.Clear();

            foreach (int index in indices)
            {
                lvImages.Items[index + 1]
                    .Selected = true;
            }


            cbSort.SelectedIndex = 0;
        }


        // =========================================================
        // 목록에서 제외
        // 원본 파일 삭제 아님
        // =========================================================

        private void BtnDelete_Click(
            object sender,
            EventArgs e)
        {
            var indices =
                lvImages.SelectedIndices
                .Cast<int>()
                .OrderByDescending(x => x)
                .ToList();


            if (indices.Count == 0)
                return;


            foreach (int index in indices)
            {
                _images.RemoveAt(index);
            }


            ClearPreview();

            Reorder();

            cbSort.SelectedIndex = 0;
        }


        // =========================================================
        // 정렬
        // =========================================================

        private void CbSort_SelectedIndexChanged(
            object sender,
            EventArgs e)
        {
            if (_images.Count == 0)
                return;


            switch (cbSort.SelectedIndex)
            {
                // 사용자 지정
                case 0:
                    return;


                // 파일명 ↑
                case 1:

                    _images.Sort(
                        (a, b) =>
                        string.Compare(
                            a.FileName,
                            b.FileName,
                            StringComparison.CurrentCultureIgnoreCase));

                    break;


                // 파일명 ↓
                case 2:

                    _images.Sort(
                        (a, b) =>
                        string.Compare(
                            b.FileName,
                            a.FileName,
                            StringComparison.CurrentCultureIgnoreCase));

                    break;


                // 수정일 ↑
                case 3:

                    _images.Sort(
                        (a, b) =>
                        a.LastWriteTime.CompareTo(
                            b.LastWriteTime));

                    break;


                // 수정일 ↓
                case 4:

                    _images.Sort(
                        (a, b) =>
                        b.LastWriteTime.CompareTo(
                            a.LastWriteTime));

                    break;
            }


            Reorder();
        }


        // =========================================================
        // 썸네일 크기 변경
        // =========================================================

        private void TbThumbSize_ValueChanged(
            object sender,
            EventArgs e)
        {
            int width =
                tbThumbSize.Value;

            int height =
                (int)(width * 0.73);


            if (height < 50)
                height = 50;

            if (height > 256)
                height = 256;

            if (width > 256)
                width = 256;


            imageList.ImageSize =
                new Size(
                    width,
                    height);


            RefreshListView();
        }


        // =========================================================
        // 새로고침
        // =========================================================

        private void BtnRefresh_Click(
            object sender,
            EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(
                    _currentFolder))
                return;

            LoadImages(
                _currentFolder);
        }


        // =========================================================
        // 정렬본 만들기
        // =========================================================

        private void BtnMakeFolder_Click(
            object sender,
            EventArgs e)
        {
            if (_images.Count == 0)
            {
                MessageBox.Show(
                    "이미지가 없습니다.");

                return;
            }


            using (var dialog =
                new FolderBrowserDialog())
            {
                dialog.Description =
                    "정렬된 이미지를 저장할 폴더를 선택하세요.";

                dialog.SelectedPath =
                    _currentFolder;


                if (dialog.ShowDialog() !=
                    DialogResult.OK)
                    return;


                string outputFolder =
                    Path.Combine(
                        dialog.SelectedPath,
                        "정렬");


                // 이미 있으면 다른 이름 사용
                outputFolder =
                    GetAvailableFolderName(
                        outputFolder);


                Directory.CreateDirectory(
                    outputFolder);


                try
                {
                    for (int i = 0;
                        i < _images.Count;
                        i++)
                    {
                        ImageItem item =
                            _images[i];


                        string extension =
                            Path.GetExtension(
                                item.FileName);


                        string newFileName;


                        if (chkKeepOriginalName.Checked)
                        {
                            newFileName =
                                $"{i + 1:D3}_{item.FileName}";
                        }
                        else
                        {
                            newFileName =
                                $"{i + 1:D3}{extension}";
                        }


                        string dest =
                            Path.Combine(
                                outputFolder,
                                newFileName);


                        File.Copy(
                            item.FullPath,
                            dest,
                            false);
                    }


                    MessageBox.Show(
                        $"완료했습니다.\r\n\r\n{outputFolder}",
                        "완료",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);


                    System.Diagnostics.Process.Start(
                        new System.Diagnostics.ProcessStartInfo
                        {
                            FileName = outputFolder,
                            UseShellExecute = true
                        });
                }
                catch (Exception ex)
                {
                    MessageBox.Show(
                        ex.Message,
                        "복사 오류",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                }
            }
        }


        // =========================================================
        // 같은 폴더가 이미 있으면 정렬_2 등
        // =========================================================

        private string GetAvailableFolderName(
            string path)
        {
            if (!Directory.Exists(path))
                return path;


            int number = 2;


            while (true)
            {
                string newPath =
                    path + "_" + number;


                if (!Directory.Exists(
                        newPath))
                {
                    return newPath;
                }


                number++;
            }
        }


        // =========================================================
        // Form 종료
        // =========================================================

        protected override void OnFormClosed(
            FormClosedEventArgs e)
        {
            ClearPreview();

            DisposeThumbnailImages();

            base.OnFormClosed(e);
        }

        private void FormSortImage_Load(object sender, EventArgs e)
        {
            //    bottomPanel.Resize += (s, ev) =>
            //    {
            //        btnMakeFolder.Left =
            //            bottomPanel.ClientSize.Width -
            //            btnMakeFolder.Width - 15;
            //    };
        }

        private void bottomPanel_Resize(object sender, EventArgs e)
        {
            btnMakeFolder.Left =
                    bottomPanel.ClientSize.Width -
                    btnMakeFolder.Width - 15;

            btnMakeFolder.Left =
                bottomPanel.Width - 160;
        }
    }
}
