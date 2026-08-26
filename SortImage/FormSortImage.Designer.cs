using System.Windows.Forms;
using System.Drawing;

namespace WindowsFormsApp1.SortImage
{
    partial class FormSortImage
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        private TextBox tbFolder;
        private Button btnFolder;
        private Button btnRefresh;

        private TreeView tvFolders;

        private ListView lvImages;
        private ImageList imageList;

        private PictureBox pbPreview;

        private Label lbFileName;
        private Label lbInfo;

        private ComboBox cbSort;

        private Button btnPrev;
        private Button btnNext;

        private Button btnUp;
        private Button btnDown;
        private Button btnDelete;

        private Button btnMakeFolder;

        private CheckBox chkKeepOriginalName;

        private TrackBar tbThumbSize;

        private Panel bottomPanel;
        private Panel topPanel;
        private TableLayoutPanel root;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            //this.components = new System.ComponentModel.Container();
            //this.root = new System.Windows.Forms.TableLayoutPanel();
            //this.topPanel = new System.Windows.Forms.Panel();
            //this.lblFolder = new System.Windows.Forms.Label();
            //this.tbFolder = new System.Windows.Forms.TextBox();
            //this.btnFolder = new System.Windows.Forms.Button();
            //this.btnRefresh = new System.Windows.Forms.Button();
            //this.lblSort = new System.Windows.Forms.Label();
            //this.cbSort = new System.Windows.Forms.ComboBox();
            //this.mainSplit = new System.Windows.Forms.SplitContainer();
            //this.tvFolders = new System.Windows.Forms.TreeView();
            //this.imageSplit = new System.Windows.Forms.SplitContainer();
            //this.lvImages = new System.Windows.Forms.ListView();
            //this.imageList = new System.Windows.Forms.ImageList(this.components);
            //this.previewPanel = new System.Windows.Forms.TableLayoutPanel();
            //this.lbFileName = new System.Windows.Forms.Label();
            //this.pbPreview = new System.Windows.Forms.PictureBox();
            //this.previewButtonPanel = new System.Windows.Forms.Panel();
            //this.btnPrev = new System.Windows.Forms.Button();
            //this.btnNext = new System.Windows.Forms.Button();
            //this.lbInfo = new System.Windows.Forms.Label();
            //this.bottomPanel = new System.Windows.Forms.Panel();
            //this.btnUp = new System.Windows.Forms.Button();
            //this.btnDown = new System.Windows.Forms.Button();
            //this.btnDelete = new System.Windows.Forms.Button();
            //this.lblThumb = new System.Windows.Forms.Label();
            //this.tbThumbSize = new System.Windows.Forms.TrackBar();
            //this.chkKeepOriginalName = new System.Windows.Forms.CheckBox();
            //this.btnMakeFolder = new System.Windows.Forms.Button();
            //this.root.SuspendLayout();
            //this.topPanel.SuspendLayout();
            //((System.ComponentModel.ISupportInitialize)(this.mainSplit)).BeginInit();
            //this.mainSplit.Panel1.SuspendLayout();
            //this.mainSplit.Panel2.SuspendLayout();
            //this.mainSplit.SuspendLayout();
            //((System.ComponentModel.ISupportInitialize)(this.imageSplit)).BeginInit();
            //this.imageSplit.Panel1.SuspendLayout();
            //this.imageSplit.Panel2.SuspendLayout();
            //this.imageSplit.SuspendLayout();
            //this.previewPanel.SuspendLayout();
            //((System.ComponentModel.ISupportInitialize)(this.pbPreview)).BeginInit();
            //this.previewButtonPanel.SuspendLayout();
            //this.bottomPanel.SuspendLayout();
            //((System.ComponentModel.ISupportInitialize)(this.tbThumbSize)).BeginInit();
            //this.SuspendLayout();
            //// 
            //// root
            //// 
            //this.root.ColumnCount = 1;
            //this.root.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 20F));
            //this.root.Controls.Add(this.topPanel, 0, 0);
            //this.root.Controls.Add(this.mainSplit, 0, 1);
            //this.root.Controls.Add(this.bottomPanel, 0, 2);
            //this.root.Dock = System.Windows.Forms.DockStyle.Fill;
            //this.root.Location = new System.Drawing.Point(0, 0);
            //this.root.Name = "root";
            //this.root.RowCount = 3;
            //this.root.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 55F));
            //this.root.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            //this.root.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 65F));
            //this.root.Size = new System.Drawing.Size(1079, 450);
            //this.root.TabIndex = 0;
            //// 
            //// topPanel
            //// 
            //this.topPanel.Controls.Add(this.lblFolder);
            //this.topPanel.Controls.Add(this.tbFolder);
            //this.topPanel.Controls.Add(this.btnFolder);
            //this.topPanel.Controls.Add(this.btnRefresh);
            //this.topPanel.Controls.Add(this.lblSort);
            //this.topPanel.Controls.Add(this.cbSort);
            //this.topPanel.Dock = System.Windows.Forms.DockStyle.Fill;
            //this.topPanel.Location = new System.Drawing.Point(3, 3);
            //this.topPanel.Name = "topPanel";
            //this.topPanel.Size = new System.Drawing.Size(1073, 49);
            //this.topPanel.TabIndex = 0;
            //// 
            //// lblFolder
            //// 
            //this.lblFolder.AutoSize = true;
            //this.lblFolder.Location = new System.Drawing.Point(10, 19);
            //this.lblFolder.Name = "lblFolder";
            //this.lblFolder.Size = new System.Drawing.Size(44, 18);
            //this.lblFolder.TabIndex = 0;
            //this.lblFolder.Text = "폴더";
            //// 
            //// tbFolder
            //// 
            //this.tbFolder.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            //| System.Windows.Forms.AnchorStyles.Right)));
            //this.tbFolder.Location = new System.Drawing.Point(55, 13);
            //this.tbFolder.Name = "tbFolder";
            //this.tbFolder.Size = new System.Drawing.Size(554, 28);
            //this.tbFolder.TabIndex = 1;
            //// 
            //// btnFolder
            //// 
            //this.btnFolder.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            //this.btnFolder.Location = new System.Drawing.Point(615, 11);
            //this.btnFolder.Name = "btnFolder";
            //this.btnFolder.Size = new System.Drawing.Size(100, 32);
            //this.btnFolder.TabIndex = 2;
            //this.btnFolder.Text = "폴더 선택";
            //// 
            //// btnRefresh
            //// 
            //this.btnRefresh.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            //this.btnRefresh.Location = new System.Drawing.Point(725, 11);
            //this.btnRefresh.Name = "btnRefresh";
            //this.btnRefresh.Size = new System.Drawing.Size(90, 32);
            //this.btnRefresh.TabIndex = 3;
            //this.btnRefresh.Text = "새로고침";
            //// 
            //// lblSort
            //// 
            //this.lblSort.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            //this.lblSort.AutoSize = true;
            //this.lblSort.Location = new System.Drawing.Point(840, 19);
            //this.lblSort.Name = "lblSort";
            //this.lblSort.Size = new System.Drawing.Size(44, 18);
            //this.lblSort.TabIndex = 4;
            //this.lblSort.Text = "정렬";
            //// 
            //// cbSort
            //// 
            //this.cbSort.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            //this.cbSort.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            //this.cbSort.Items.AddRange(new object[] {
            //"사용자 지정",
            //"파일명 ↑",
            //"파일명 ↓",
            //"수정일 ↑",
            //"수정일 ↓"});
            //this.cbSort.Location = new System.Drawing.Point(880, 13);
            //this.cbSort.Name = "cbSort";
            //this.cbSort.Size = new System.Drawing.Size(150, 26);
            //this.cbSort.TabIndex = 5;
            //// 
            //// mainSplit
            //// 
            //this.mainSplit.Dock = System.Windows.Forms.DockStyle.Fill;
            //this.mainSplit.FixedPanel = System.Windows.Forms.FixedPanel.Panel1;
            //this.mainSplit.Location = new System.Drawing.Point(3, 58);
            //this.mainSplit.Name = "mainSplit";
            //// 
            //// mainSplit.Panel1
            //// 
            //this.mainSplit.Panel1.Controls.Add(this.tvFolders);
            //// 
            //// mainSplit.Panel2
            //// 
            //this.mainSplit.Panel2.Controls.Add(this.imageSplit);
            //this.mainSplit.Size = new System.Drawing.Size(1073, 324);
            //this.mainSplit.SplitterDistance = 121;
            //this.mainSplit.TabIndex = 1;
            //// 
            //// tvFolders
            //// 
            //this.tvFolders.Dock = System.Windows.Forms.DockStyle.Fill;
            //this.tvFolders.HideSelection = false;
            //this.tvFolders.Location = new System.Drawing.Point(0, 0);
            //this.tvFolders.Name = "tvFolders";
            //this.tvFolders.Size = new System.Drawing.Size(121, 324);
            //this.tvFolders.TabIndex = 0;
            //// 
            //// imageSplit
            //// 
            //this.imageSplit.Dock = System.Windows.Forms.DockStyle.Fill;
            //this.imageSplit.Location = new System.Drawing.Point(0, 0);
            //this.imageSplit.Name = "imageSplit";
            //// 
            //// imageSplit.Panel1
            //// 
            //this.imageSplit.Panel1.Controls.Add(this.lvImages);
            //// 
            //// imageSplit.Panel2
            //// 
            //this.imageSplit.Panel2.Controls.Add(this.previewPanel);
            //this.imageSplit.Size = new System.Drawing.Size(948, 324);
            //this.imageSplit.SplitterDistance = 764;
            //this.imageSplit.TabIndex = 0;
            //// 
            //// lvImages
            //// 
            //this.lvImages.AllowDrop = true;
            //this.lvImages.BackColor = System.Drawing.Color.White;
            //this.lvImages.Dock = System.Windows.Forms.DockStyle.Fill;
            //this.lvImages.HideSelection = false;
            //this.lvImages.LargeImageList = this.imageList;
            //this.lvImages.Location = new System.Drawing.Point(0, 0);
            //this.lvImages.Name = "lvImages";
            //this.lvImages.Size = new System.Drawing.Size(764, 324);
            //this.lvImages.TabIndex = 0;
            //this.lvImages.UseCompatibleStateImageBehavior = false;
            //// 
            //// imageList
            //// 
            //this.imageList.ColorDepth = System.Windows.Forms.ColorDepth.Depth32Bit;
            //this.imageList.ImageSize = new System.Drawing.Size(150, 110);
            //this.imageList.TransparentColor = System.Drawing.Color.Transparent;
            //// 
            //// previewPanel
            //// 
            //this.previewPanel.ColumnCount = 1;
            //this.previewPanel.Controls.Add(this.lbFileName, 0, 0);
            //this.previewPanel.Controls.Add(this.pbPreview, 0, 1);
            //this.previewPanel.Controls.Add(this.previewButtonPanel, 0, 2);
            //this.previewPanel.Controls.Add(this.lbInfo, 0, 3);
            //this.previewPanel.Dock = System.Windows.Forms.DockStyle.Fill;
            //this.previewPanel.Location = new System.Drawing.Point(0, 0);
            //this.previewPanel.Name = "previewPanel";
            //this.previewPanel.RowCount = 4;
            //this.previewPanel.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 35F));
            //this.previewPanel.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            //this.previewPanel.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 50F));
            //this.previewPanel.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 45F));
            //this.previewPanel.Size = new System.Drawing.Size(180, 324);
            //this.previewPanel.TabIndex = 0;
            //// 
            //// lbFileName
            //// 
            //this.lbFileName.Dock = System.Windows.Forms.DockStyle.Fill;
            //this.lbFileName.Location = new System.Drawing.Point(3, 0);
            //this.lbFileName.Name = "lbFileName";
            //this.lbFileName.Size = new System.Drawing.Size(200, 35);
            //this.lbFileName.TabIndex = 0;
            //this.lbFileName.Text = "이미지를 선택하세요.";
            //this.lbFileName.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            //// 
            //// pbPreview
            //// 
            //this.pbPreview.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(40)))), ((int)(((byte)(40)))), ((int)(((byte)(40)))));
            //this.pbPreview.Dock = System.Windows.Forms.DockStyle.Fill;
            //this.pbPreview.Location = new System.Drawing.Point(3, 38);
            //this.pbPreview.Name = "pbPreview";
            //this.pbPreview.Size = new System.Drawing.Size(200, 188);
            //this.pbPreview.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            //this.pbPreview.TabIndex = 1;
            //this.pbPreview.TabStop = false;
            //// 
            //// previewButtonPanel
            //// 
            //this.previewButtonPanel.Controls.Add(this.btnPrev);
            //this.previewButtonPanel.Controls.Add(this.btnNext);
            //this.previewButtonPanel.Dock = System.Windows.Forms.DockStyle.Fill;
            //this.previewButtonPanel.Location = new System.Drawing.Point(3, 232);
            //this.previewButtonPanel.Name = "previewButtonPanel";
            //this.previewButtonPanel.Size = new System.Drawing.Size(200, 44);
            //this.previewButtonPanel.TabIndex = 2;
            //// 
            //// btnPrev
            //// 
            //this.btnPrev.Location = new System.Drawing.Point(30, 7);
            //this.btnPrev.Name = "btnPrev";
            //this.btnPrev.Size = new System.Drawing.Size(100, 34);
            //this.btnPrev.TabIndex = 0;
            //this.btnPrev.Text = "◀ 이전";
            //// 
            //// btnNext
            //// 
            //this.btnNext.Location = new System.Drawing.Point(140, 7);
            //this.btnNext.Name = "btnNext";
            //this.btnNext.Size = new System.Drawing.Size(100, 34);
            //this.btnNext.TabIndex = 1;
            //this.btnNext.Text = "다음 ▶";
            //// 
            //// lbInfo
            //// 
            //this.lbInfo.Dock = System.Windows.Forms.DockStyle.Fill;
            //this.lbInfo.Location = new System.Drawing.Point(3, 279);
            //this.lbInfo.Name = "lbInfo";
            //this.lbInfo.Size = new System.Drawing.Size(200, 45);
            //this.lbInfo.TabIndex = 3;
            //this.lbInfo.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            //// 
            //// bottomPanel
            //// 
            //this.bottomPanel.Controls.Add(this.btnUp);
            //this.bottomPanel.Controls.Add(this.btnDown);
            //this.bottomPanel.Controls.Add(this.btnDelete);
            //this.bottomPanel.Controls.Add(this.lblThumb);
            //this.bottomPanel.Controls.Add(this.tbThumbSize);
            //this.bottomPanel.Controls.Add(this.chkKeepOriginalName);
            //this.bottomPanel.Controls.Add(this.btnMakeFolder);
            //this.bottomPanel.Dock = System.Windows.Forms.DockStyle.Fill;
            //this.bottomPanel.Location = new System.Drawing.Point(3, 388);
            //this.bottomPanel.Name = "bottomPanel";
            //this.bottomPanel.Size = new System.Drawing.Size(1073, 59);
            //this.bottomPanel.TabIndex = 2;
            //this.bottomPanel.Resize += new System.EventHandler(this.bottomPanel_Resize);
            //// 
            //// btnUp
            //// 
            //this.btnUp.Location = new System.Drawing.Point(10, 14);
            //this.btnUp.Name = "btnUp";
            //this.btnUp.Size = new System.Drawing.Size(80, 35);
            //this.btnUp.TabIndex = 0;
            //this.btnUp.Text = "▲ 위로";
            //// 
            //// btnDown
            //// 
            //this.btnDown.Location = new System.Drawing.Point(100, 14);
            //this.btnDown.Name = "btnDown";
            //this.btnDown.Size = new System.Drawing.Size(80, 35);
            //this.btnDown.TabIndex = 1;
            //this.btnDown.Text = "▼ 아래로";
            //// 
            //// btnDelete
            //// 
            //this.btnDelete.Location = new System.Drawing.Point(190, 14);
            //this.btnDelete.Name = "btnDelete";
            //this.btnDelete.Size = new System.Drawing.Size(90, 35);
            //this.btnDelete.TabIndex = 2;
            //this.btnDelete.Text = "목록 제외";
            //// 
            //// lblThumb
            //// 
            //this.lblThumb.AutoSize = true;
            //this.lblThumb.Location = new System.Drawing.Point(310, 24);
            //this.lblThumb.Name = "lblThumb";
            //this.lblThumb.Size = new System.Drawing.Size(62, 18);
            //this.lblThumb.TabIndex = 3;
            //this.lblThumb.Text = "썸네일";
            //// 
            //// tbThumbSize
            //// 
            //this.tbThumbSize.Location = new System.Drawing.Point(365, 5);
            //this.tbThumbSize.Maximum = 250;
            //this.tbThumbSize.Minimum = 80;
            //this.tbThumbSize.Name = "tbThumbSize";
            //this.tbThumbSize.Size = new System.Drawing.Size(180, 69);
            //this.tbThumbSize.TabIndex = 4;
            //this.tbThumbSize.TickFrequency = 20;
            //this.tbThumbSize.Value = 150;
            //// 
            //// chkKeepOriginalName
            //// 
            //this.chkKeepOriginalName.AutoSize = true;
            //this.chkKeepOriginalName.Checked = true;
            //this.chkKeepOriginalName.CheckState = System.Windows.Forms.CheckState.Checked;
            //this.chkKeepOriginalName.Location = new System.Drawing.Point(560, 22);
            //this.chkKeepOriginalName.Name = "chkKeepOriginalName";
            //this.chkKeepOriginalName.Size = new System.Drawing.Size(256, 22);
            //this.chkKeepOriginalName.TabIndex = 5;
            //this.chkKeepOriginalName.Text = "순번 뒤에 원본 파일명 유지";
            //// 
            //// btnMakeFolder
            //// 
            //this.btnMakeFolder.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            //this.btnMakeFolder.Location = new System.Drawing.Point(873, 11);
            //this.btnMakeFolder.Name = "btnMakeFolder";
            //this.btnMakeFolder.Size = new System.Drawing.Size(140, 40);
            //this.btnMakeFolder.TabIndex = 6;
            //this.btnMakeFolder.Text = "정렬본 만들기";
            // 
            // FormSortImage
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(10F, 18F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1079, 450);
            this.Controls.Add(this.root);
            this.Name = "FormSortImage";
            this.Text = "Form1";
            this.Load += new System.EventHandler(this.FormSortImage_Load);
            this.root.ResumeLayout(false);
            this.topPanel.ResumeLayout(false);
            this.topPanel.PerformLayout();
            this.mainSplit.Panel1.ResumeLayout(false);
            this.mainSplit.Panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.mainSplit)).EndInit();
            this.mainSplit.ResumeLayout(false);
            this.imageSplit.Panel1.ResumeLayout(false);
            this.imageSplit.Panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.imageSplit)).EndInit();
            this.imageSplit.ResumeLayout(false);
            this.previewPanel.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.pbPreview)).EndInit();
            this.previewButtonPanel.ResumeLayout(false);
            this.bottomPanel.ResumeLayout(false);
            this.bottomPanel.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.tbThumbSize)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private Label lblFolder;
        private Label lblSort;
        private SplitContainer mainSplit;
        private SplitContainer imageSplit;
        private TableLayoutPanel previewPanel;
        private Panel previewButtonPanel;
        private Label lblThumb;
    }
}