using System.Drawing;
using System.Windows.Forms;

namespace SmartReport
{
    partial class Form1
    {
        /// <summary>
        /// 필수 디자이너 변수입니다.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 사용 중인 모든 리소스를 정리합니다.
        /// </summary>
        /// <param name="disposing">관리되는 리소스를 삭제해야 하면 true이고, 그렇지 않으면 false입니다.</param>
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
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.pnlMain = new System.Windows.Forms.Panel();
            this.statusStrip1 = new System.Windows.Forms.StatusStrip();
            this.toolStripStatusLabel1 = new System.Windows.Forms.ToolStripStatusLabel();
            this.splitContainer1 = new System.Windows.Forms.SplitContainer();
            this.tabMain = new System.Windows.Forms.TabControl();
            this.tpFunction = new System.Windows.Forms.TabPage();
            this.gbReport = new System.Windows.Forms.GroupBox();
            this.btnQuantityFile = new System.Windows.Forms.Button();
            this.tbQuantityFile = new System.Windows.Forms.TextBox();
            this.lbQuantityFile = new System.Windows.Forms.Label();
            this.btnPageNumber = new System.Windows.Forms.Button();
            this.btnExportForPdf = new System.Windows.Forms.Button();
            this.gbConvertValue = new System.Windows.Forms.GroupBox();
            this.checkBox1 = new System.Windows.Forms.CheckBox();
            this.checkBoxDate = new System.Windows.Forms.CheckBox();
            this.checkBoxChecker = new System.Windows.Forms.CheckBox();
            this.gbImage = new System.Windows.Forms.GroupBox();
            this.checkBoxOcr = new System.Windows.Forms.CheckBox();
            this.comboBoxSignalGraph = new System.Windows.Forms.ComboBox();
            this.comboBoxTestGraph = new System.Windows.Forms.ComboBox();
            this.comboBoxHwaveGraph = new System.Windows.Forms.ComboBox();
            this.btnGapjiPictureRelocate = new System.Windows.Forms.Button();
            this.textBoxPictureFolder = new System.Windows.Forms.TextBox();
            this.textBoxCoronaFolder = new System.Windows.Forms.TextBox();
            this.textBoxQuntatyFolder = new System.Windows.Forms.TextBox();
            this.btErrorPageUpdate = new System.Windows.Forms.Button();
            this.textBoxFeverImageFolder = new System.Windows.Forms.TextBox();
            this.checkBoxFeverPicture = new System.Windows.Forms.CheckBox();
            this.checkBoxPicture = new System.Windows.Forms.CheckBox();
            this.checkBoxCorona = new System.Windows.Forms.CheckBox();
            this.checkBoxQuantity = new System.Windows.Forms.CheckBox();
            this.btnQuntityFileRun = new System.Windows.Forms.Button();
            this.tpGap = new System.Windows.Forms.TabPage();
            this.panel2 = new System.Windows.Forms.Panel();
            this.btnDownload = new System.Windows.Forms.Button();
            this.tbCompany = new System.Windows.Forms.TextBox();
            this.lbCompany = new System.Windows.Forms.Label();
            this.panel1 = new System.Windows.Forms.Panel();
            this.dgvFiles = new System.Windows.Forms.DataGridView();
            this.tapLog = new System.Windows.Forms.TabControl();
            this.all = new System.Windows.Forms.TabPage();
            this.richTextBox1 = new System.Windows.Forms.RichTextBox();
            this.info = new System.Windows.Forms.TabPage();
            this.tabError = new System.Windows.Forms.TabPage();
            this.pnlTop = new System.Windows.Forms.Panel();
            this.btnDownloadWork = new System.Windows.Forms.Button();
            this.textBoxKeyword = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.checkBoxAnnual = new System.Windows.Forms.CheckBox();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.lbSample = new System.Windows.Forms.Label();
            this.cbAutoGenerateExcel = new System.Windows.Forms.CheckBox();
            this.comboBoxSeason = new System.Windows.Forms.ComboBox();
            this.label2 = new System.Windows.Forms.Label();
            this.dateTimePicker1 = new System.Windows.Forms.DateTimePicker();
            this.label1 = new System.Windows.Forms.Label();
            this.tbPerson = new System.Windows.Forms.TextBox();
            this.panel3 = new System.Windows.Forms.Panel();
            this.btFolder = new System.Windows.Forms.Button();
            this.tbFolder = new System.Windows.Forms.TextBox();
            this.lbFolder = new System.Windows.Forms.Label();
            this.cbName = new System.Windows.Forms.ComboBox();
            this.lbName = new System.Windows.Forms.Label();
            this.pnlMain.SuspendLayout();
            this.statusStrip1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).BeginInit();
            this.splitContainer1.Panel1.SuspendLayout();
            this.splitContainer1.Panel2.SuspendLayout();
            this.splitContainer1.SuspendLayout();
            this.tabMain.SuspendLayout();
            this.tpFunction.SuspendLayout();
            this.gbReport.SuspendLayout();
            this.gbConvertValue.SuspendLayout();
            this.gbImage.SuspendLayout();
            this.tpGap.SuspendLayout();
            this.panel2.SuspendLayout();
            this.panel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvFiles)).BeginInit();
            this.tapLog.SuspendLayout();
            this.all.SuspendLayout();
            this.pnlTop.SuspendLayout();
            this.panel3.SuspendLayout();
            this.SuspendLayout();
            // 
            // pnlMain
            // 
            this.pnlMain.BackColor = System.Drawing.SystemColors.Control;
            this.pnlMain.Controls.Add(this.statusStrip1);
            this.pnlMain.Controls.Add(this.splitContainer1);
            this.pnlMain.Controls.Add(this.pnlTop);
            this.pnlMain.Dock = System.Windows.Forms.DockStyle.Fill;
            this.pnlMain.Location = new System.Drawing.Point(0, 0);
            this.pnlMain.Margin = new System.Windows.Forms.Padding(4);
            this.pnlMain.Name = "pnlMain";
            this.pnlMain.Padding = new System.Windows.Forms.Padding(4);
            this.pnlMain.Size = new System.Drawing.Size(1160, 925);
            this.pnlMain.TabIndex = 4;
            // 
            // statusStrip1
            // 
            this.statusStrip1.ImageScalingSize = new System.Drawing.Size(24, 24);
            this.statusStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripStatusLabel1});
            this.statusStrip1.Location = new System.Drawing.Point(4, 889);
            this.statusStrip1.Name = "statusStrip1";
            this.statusStrip1.Size = new System.Drawing.Size(1152, 32);
            this.statusStrip1.TabIndex = 5;
            this.statusStrip1.Text = "statusStrip1";
            // 
            // toolStripStatusLabel1
            // 
            this.toolStripStatusLabel1.Name = "toolStripStatusLabel1";
            this.toolStripStatusLabel1.Size = new System.Drawing.Size(183, 25);
            this.toolStripStatusLabel1.Text = "toolStripStatusLabel1";
            // 
            // splitContainer1
            // 
            this.splitContainer1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.splitContainer1.Location = new System.Drawing.Point(4, 186);
            this.splitContainer1.Name = "splitContainer1";
            this.splitContainer1.Orientation = System.Windows.Forms.Orientation.Horizontal;
            // 
            // splitContainer1.Panel1
            // 
            this.splitContainer1.Panel1.Controls.Add(this.tabMain);
            this.splitContainer1.Panel1.Padding = new System.Windows.Forms.Padding(5);
            // 
            // splitContainer1.Panel2
            // 
            this.splitContainer1.Panel2.Controls.Add(this.tapLog);
            this.splitContainer1.Panel2.Padding = new System.Windows.Forms.Padding(5);
            this.splitContainer1.Size = new System.Drawing.Size(1152, 710);
            this.splitContainer1.SplitterDistance = 444;
            this.splitContainer1.TabIndex = 4;
            // 
            // tabMain
            // 
            this.tabMain.Controls.Add(this.tpFunction);
            this.tabMain.Controls.Add(this.tpGap);
            this.tabMain.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tabMain.Location = new System.Drawing.Point(5, 5);
            this.tabMain.Margin = new System.Windows.Forms.Padding(4);
            this.tabMain.Name = "tabMain";
            this.tabMain.SelectedIndex = 0;
            this.tabMain.Size = new System.Drawing.Size(1142, 434);
            this.tabMain.TabIndex = 0;
            // 
            // tpFunction
            // 
            this.tpFunction.Controls.Add(this.gbReport);
            this.tpFunction.Controls.Add(this.gbConvertValue);
            this.tpFunction.Controls.Add(this.gbImage);
            this.tpFunction.Location = new System.Drawing.Point(4, 39);
            this.tpFunction.Margin = new System.Windows.Forms.Padding(4);
            this.tpFunction.Name = "tpFunction";
            this.tpFunction.Padding = new System.Windows.Forms.Padding(4);
            this.tpFunction.Size = new System.Drawing.Size(1134, 391);
            this.tpFunction.TabIndex = 1;
            this.tpFunction.Text = "기능";
            this.tpFunction.UseVisualStyleBackColor = true;
            // 
            // gbReport
            // 
            this.gbReport.Controls.Add(this.btnQuantityFile);
            this.gbReport.Controls.Add(this.tbQuantityFile);
            this.gbReport.Controls.Add(this.lbQuantityFile);
            this.gbReport.Controls.Add(this.btnPageNumber);
            this.gbReport.Controls.Add(this.btnExportForPdf);
            this.gbReport.Location = new System.Drawing.Point(5, 0);
            this.gbReport.Name = "gbReport";
            this.gbReport.Size = new System.Drawing.Size(1127, 109);
            this.gbReport.TabIndex = 26;
            this.gbReport.TabStop = false;
            this.gbReport.Text = "보고서";
            // 
            // btnQuantityFile
            // 
            this.btnQuantityFile.Location = new System.Drawing.Point(985, 30);
            this.btnQuantityFile.Name = "btnQuantityFile";
            this.btnQuantityFile.Size = new System.Drawing.Size(139, 69);
            this.btnQuantityFile.TabIndex = 26;
            this.btnQuantityFile.Text = "찾기";
            this.btnQuantityFile.UseVisualStyleBackColor = true;
            this.btnQuantityFile.Click += new System.EventHandler(this.btnQuantityFile_Click);
            // 
            // tbQuantityFile
            // 
            this.tbQuantityFile.Location = new System.Drawing.Point(154, 30);
            this.tbQuantityFile.Margin = new System.Windows.Forms.Padding(4);
            this.tbQuantityFile.Name = "tbQuantityFile";
            this.tbQuantityFile.Size = new System.Drawing.Size(769, 37);
            this.tbQuantityFile.TabIndex = 25;
            // 
            // lbQuantityFile
            // 
            this.lbQuantityFile.AutoSize = true;
            this.lbQuantityFile.Location = new System.Drawing.Point(10, 33);
            this.lbQuantityFile.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.lbQuantityFile.Name = "lbQuantityFile";
            this.lbQuantityFile.Size = new System.Drawing.Size(198, 30);
            this.lbQuantityFile.TabIndex = 23;
            this.lbQuantityFile.Text = "보고서 파일명    : ";
            // 
            // btnPageNumber
            // 
            this.btnPageNumber.Location = new System.Drawing.Point(6, 64);
            this.btnPageNumber.Name = "btnPageNumber";
            this.btnPageNumber.Size = new System.Drawing.Size(172, 32);
            this.btnPageNumber.TabIndex = 0;
            this.btnPageNumber.Text = "페이지번호 업데이트";
            this.btnPageNumber.UseVisualStyleBackColor = true;
            this.btnPageNumber.Click += new System.EventHandler(this.btnPageNumber_Click);
            // 
            // btnExportForPdf
            // 
            this.btnExportForPdf.Location = new System.Drawing.Point(185, 66);
            this.btnExportForPdf.Name = "btnExportForPdf";
            this.btnExportForPdf.Size = new System.Drawing.Size(129, 30);
            this.btnExportForPdf.TabIndex = 20;
            this.btnExportForPdf.Text = "PDF로 내보내기";
            this.btnExportForPdf.UseVisualStyleBackColor = true;
            this.btnExportForPdf.MouseUp += new System.Windows.Forms.MouseEventHandler(this.btnExportForPdf_MouseUp);
            // 
            // gbConvertValue
            // 
            this.gbConvertValue.Controls.Add(this.checkBox1);
            this.gbConvertValue.Controls.Add(this.checkBoxDate);
            this.gbConvertValue.Controls.Add(this.checkBoxChecker);
            this.gbConvertValue.Location = new System.Drawing.Point(9, 115);
            this.gbConvertValue.Name = "gbConvertValue";
            this.gbConvertValue.Size = new System.Drawing.Size(1117, 100);
            this.gbConvertValue.TabIndex = 25;
            this.gbConvertValue.TabStop = false;
            this.gbConvertValue.Text = "셀 값 적용";
            // 
            // checkBox1
            // 
            this.checkBox1.AutoSize = true;
            this.checkBox1.Checked = true;
            this.checkBox1.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBox1.Location = new System.Drawing.Point(154, 32);
            this.checkBox1.Name = "checkBox1";
            this.checkBox1.Size = new System.Drawing.Size(291, 34);
            this.checkBox1.TabIndex = 2;
            this.checkBox1.Text = "갑지 사진 중앙 정렬 적용";
            this.checkBox1.UseVisualStyleBackColor = true;
            // 
            // checkBoxDate
            // 
            this.checkBoxDate.AutoSize = true;
            this.checkBoxDate.Checked = true;
            this.checkBoxDate.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBoxDate.Location = new System.Drawing.Point(17, 62);
            this.checkBoxDate.Name = "checkBoxDate";
            this.checkBoxDate.Size = new System.Drawing.Size(157, 34);
            this.checkBoxDate.TabIndex = 1;
            this.checkBoxDate.Text = "측정일 적용";
            this.checkBoxDate.UseVisualStyleBackColor = true;
            // 
            // checkBoxChecker
            // 
            this.checkBoxChecker.AutoSize = true;
            this.checkBoxChecker.Checked = true;
            this.checkBoxChecker.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBoxChecker.Location = new System.Drawing.Point(17, 32);
            this.checkBoxChecker.Name = "checkBoxChecker";
            this.checkBoxChecker.Size = new System.Drawing.Size(157, 34);
            this.checkBoxChecker.TabIndex = 0;
            this.checkBoxChecker.Text = "측정자 적용";
            this.checkBoxChecker.UseVisualStyleBackColor = true;
            // 
            // gbImage
            // 
            this.gbImage.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.gbImage.Controls.Add(this.checkBoxOcr);
            this.gbImage.Controls.Add(this.comboBoxSignalGraph);
            this.gbImage.Controls.Add(this.comboBoxTestGraph);
            this.gbImage.Controls.Add(this.comboBoxHwaveGraph);
            this.gbImage.Controls.Add(this.btnGapjiPictureRelocate);
            this.gbImage.Controls.Add(this.textBoxPictureFolder);
            this.gbImage.Controls.Add(this.textBoxCoronaFolder);
            this.gbImage.Controls.Add(this.textBoxQuntatyFolder);
            this.gbImage.Controls.Add(this.btErrorPageUpdate);
            this.gbImage.Controls.Add(this.textBoxFeverImageFolder);
            this.gbImage.Controls.Add(this.checkBoxFeverPicture);
            this.gbImage.Controls.Add(this.checkBoxPicture);
            this.gbImage.Controls.Add(this.checkBoxCorona);
            this.gbImage.Controls.Add(this.checkBoxQuantity);
            this.gbImage.Controls.Add(this.btnQuntityFileRun);
            this.gbImage.Location = new System.Drawing.Point(5, 218);
            this.gbImage.Name = "gbImage";
            this.gbImage.Size = new System.Drawing.Size(1126, 203);
            this.gbImage.TabIndex = 24;
            this.gbImage.TabStop = false;
            this.gbImage.Text = "이미지 삽입";
            // 
            // checkBoxOcr
            // 
            this.checkBoxOcr.AutoSize = true;
            this.checkBoxOcr.Checked = true;
            this.checkBoxOcr.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBoxOcr.Location = new System.Drawing.Point(551, 27);
            this.checkBoxOcr.Name = "checkBoxOcr";
            this.checkBoxOcr.Size = new System.Drawing.Size(261, 34);
            this.checkBoxOcr.TabIndex = 45;
            this.checkBoxOcr.Text = "온도 인식값 자동 입력";
            this.checkBoxOcr.UseVisualStyleBackColor = true;
            // 
            // comboBoxSignalGraph
            // 
            this.comboBoxSignalGraph.FormattingEnabled = true;
            this.comboBoxSignalGraph.Items.AddRange(new object[] {
            "시보.pdf",
            "시_보.pdf"});
            this.comboBoxSignalGraph.Location = new System.Drawing.Point(803, 52);
            this.comboBoxSignalGraph.Name = "comboBoxSignalGraph";
            this.comboBoxSignalGraph.Size = new System.Drawing.Size(98, 38);
            this.comboBoxSignalGraph.TabIndex = 44;
            this.comboBoxSignalGraph.Text = "시보.pdf";
            // 
            // comboBoxTestGraph
            // 
            this.comboBoxTestGraph.FormattingEnabled = true;
            this.comboBoxTestGraph.Items.AddRange(new object[] {
            "시그.pdf",
            "시_그.pdf"});
            this.comboBoxTestGraph.Location = new System.Drawing.Point(678, 52);
            this.comboBoxTestGraph.Name = "comboBoxTestGraph";
            this.comboBoxTestGraph.Size = new System.Drawing.Size(98, 38);
            this.comboBoxTestGraph.TabIndex = 43;
            this.comboBoxTestGraph.Text = "시그.pdf";
            // 
            // comboBoxHwaveGraph
            // 
            this.comboBoxHwaveGraph.FormattingEnabled = true;
            this.comboBoxHwaveGraph.Items.AddRange(new object[] {
            "고그.pdf",
            "고_그.pdf"});
            this.comboBoxHwaveGraph.Location = new System.Drawing.Point(551, 52);
            this.comboBoxHwaveGraph.Name = "comboBoxHwaveGraph";
            this.comboBoxHwaveGraph.Size = new System.Drawing.Size(98, 38);
            this.comboBoxHwaveGraph.TabIndex = 42;
            this.comboBoxHwaveGraph.Text = "고그.pdf";
            // 
            // btnGapjiPictureRelocate
            // 
            this.btnGapjiPictureRelocate.Location = new System.Drawing.Point(16, 145);
            this.btnGapjiPictureRelocate.Name = "btnGapjiPictureRelocate";
            this.btnGapjiPictureRelocate.Size = new System.Drawing.Size(157, 32);
            this.btnGapjiPictureRelocate.TabIndex = 41;
            this.btnGapjiPictureRelocate.Text = "갑지 사진 중앙 정렬";
            this.btnGapjiPictureRelocate.UseVisualStyleBackColor = true;
            this.btnGapjiPictureRelocate.Click += new System.EventHandler(this.btnGapjiPictureRelocate_Click);
            // 
            // textBoxPictureFolder
            // 
            this.textBoxPictureFolder.Location = new System.Drawing.Point(292, 113);
            this.textBoxPictureFolder.Margin = new System.Windows.Forms.Padding(4);
            this.textBoxPictureFolder.Name = "textBoxPictureFolder";
            this.textBoxPictureFolder.Size = new System.Drawing.Size(217, 37);
            this.textBoxPictureFolder.TabIndex = 40;
            this.textBoxPictureFolder.Text = "03 점검사진";
            // 
            // textBoxCoronaFolder
            // 
            this.textBoxCoronaFolder.Location = new System.Drawing.Point(292, 83);
            this.textBoxCoronaFolder.Margin = new System.Windows.Forms.Padding(4);
            this.textBoxCoronaFolder.Name = "textBoxCoronaFolder";
            this.textBoxCoronaFolder.Size = new System.Drawing.Size(217, 37);
            this.textBoxCoronaFolder.TabIndex = 38;
            this.textBoxCoronaFolder.Text = "05 영상코로나 또는 부분방전";
            // 
            // textBoxQuntatyFolder
            // 
            this.textBoxQuntatyFolder.Location = new System.Drawing.Point(292, 53);
            this.textBoxQuntatyFolder.Margin = new System.Windows.Forms.Padding(4);
            this.textBoxQuntatyFolder.Name = "textBoxQuntatyFolder";
            this.textBoxQuntatyFolder.Size = new System.Drawing.Size(217, 37);
            this.textBoxQuntatyFolder.TabIndex = 34;
            this.textBoxQuntatyFolder.Text = "02 전원품질";
            // 
            // btErrorPageUpdate
            // 
            this.btErrorPageUpdate.Enabled = false;
            this.btErrorPageUpdate.Location = new System.Drawing.Point(185, 146);
            this.btErrorPageUpdate.Name = "btErrorPageUpdate";
            this.btErrorPageUpdate.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.btErrorPageUpdate.Size = new System.Drawing.Size(222, 30);
            this.btErrorPageUpdate.TabIndex = 22;
            this.btErrorPageUpdate.Text = "의견 부적합 페이지 업데이트";
            this.btErrorPageUpdate.UseVisualStyleBackColor = true;
            this.btErrorPageUpdate.Click += new System.EventHandler(this.btErrorPageUpdate_Click);
            // 
            // textBoxFeverImageFolder
            // 
            this.textBoxFeverImageFolder.Location = new System.Drawing.Point(292, 24);
            this.textBoxFeverImageFolder.Margin = new System.Windows.Forms.Padding(4);
            this.textBoxFeverImageFolder.Name = "textBoxFeverImageFolder";
            this.textBoxFeverImageFolder.Size = new System.Drawing.Size(217, 37);
            this.textBoxFeverImageFolder.TabIndex = 30;
            this.textBoxFeverImageFolder.Text = "01 열화상";
            // 
            // checkBoxFeverPicture
            // 
            this.checkBoxFeverPicture.AutoSize = true;
            this.checkBoxFeverPicture.Checked = true;
            this.checkBoxFeverPicture.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBoxFeverPicture.Location = new System.Drawing.Point(17, 29);
            this.checkBoxFeverPicture.Name = "checkBoxFeverPicture";
            this.checkBoxFeverPicture.Size = new System.Drawing.Size(163, 34);
            this.checkBoxFeverPicture.TabIndex = 28;
            this.checkBoxFeverPicture.Text = "열화상(분기)";
            this.checkBoxFeverPicture.UseVisualStyleBackColor = true;
            // 
            // checkBoxPicture
            // 
            this.checkBoxPicture.AutoSize = true;
            this.checkBoxPicture.Checked = true;
            this.checkBoxPicture.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBoxPicture.Location = new System.Drawing.Point(18, 117);
            this.checkBoxPicture.Name = "checkBoxPicture";
            this.checkBoxPicture.Size = new System.Drawing.Size(83, 34);
            this.checkBoxPicture.TabIndex = 26;
            this.checkBoxPicture.Text = "사진";
            this.checkBoxPicture.UseVisualStyleBackColor = true;
            // 
            // checkBoxCorona
            // 
            this.checkBoxCorona.AutoSize = true;
            this.checkBoxCorona.Checked = true;
            this.checkBoxCorona.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBoxCorona.Location = new System.Drawing.Point(18, 87);
            this.checkBoxCorona.Name = "checkBoxCorona";
            this.checkBoxCorona.Size = new System.Drawing.Size(105, 34);
            this.checkBoxCorona.TabIndex = 25;
            this.checkBoxCorona.Text = "코로나";
            this.checkBoxCorona.UseVisualStyleBackColor = true;
            // 
            // checkBoxQuantity
            // 
            this.checkBoxQuantity.AutoSize = true;
            this.checkBoxQuantity.Checked = true;
            this.checkBoxQuantity.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBoxQuantity.Location = new System.Drawing.Point(18, 59);
            this.checkBoxQuantity.Name = "checkBoxQuantity";
            this.checkBoxQuantity.Size = new System.Drawing.Size(83, 34);
            this.checkBoxQuantity.TabIndex = 24;
            this.checkBoxQuantity.Text = "품질";
            this.checkBoxQuantity.UseVisualStyleBackColor = true;
            // 
            // btnQuntityFileRun
            // 
            this.btnQuntityFileRun.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnQuntityFileRun.Location = new System.Drawing.Point(990, 67);
            this.btnQuntityFileRun.Name = "btnQuntityFileRun";
            this.btnQuntityFileRun.Size = new System.Drawing.Size(134, 116);
            this.btnQuntityFileRun.TabIndex = 23;
            this.btnQuntityFileRun.Text = "실행";
            this.btnQuntityFileRun.UseVisualStyleBackColor = true;
            this.btnQuntityFileRun.Click += new System.EventHandler(this.btnQuntityFileRun_Click);
            // 
            // tpGap
            // 
            this.tpGap.Controls.Add(this.panel2);
            this.tpGap.Controls.Add(this.panel1);
            this.tpGap.Location = new System.Drawing.Point(4, 25);
            this.tpGap.Margin = new System.Windows.Forms.Padding(4);
            this.tpGap.Name = "tpGap";
            this.tpGap.Padding = new System.Windows.Forms.Padding(4);
            this.tpGap.Size = new System.Drawing.Size(1134, 405);
            this.tpGap.TabIndex = 0;
            this.tpGap.Text = "파일다운로드";
            this.tpGap.UseVisualStyleBackColor = true;
            // 
            // panel2
            // 
            this.panel2.Controls.Add(this.btnDownload);
            this.panel2.Controls.Add(this.tbCompany);
            this.panel2.Controls.Add(this.lbCompany);
            this.panel2.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel2.Location = new System.Drawing.Point(4, 4);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(1126, 44);
            this.panel2.TabIndex = 17;
            this.panel2.Paint += new System.Windows.Forms.PaintEventHandler(this.panel2_Paint);
            // 
            // btnDownload
            // 
            this.btnDownload.Location = new System.Drawing.Point(819, 4);
            this.btnDownload.Margin = new System.Windows.Forms.Padding(4);
            this.btnDownload.Name = "btnDownload";
            this.btnDownload.Size = new System.Drawing.Size(134, 40);
            this.btnDownload.TabIndex = 11;
            this.btnDownload.Text = "다운로드";
            this.btnDownload.UseVisualStyleBackColor = true;
            this.btnDownload.Click += new System.EventHandler(this.btnDownload_Click);
            // 
            // tbCompany
            // 
            this.tbCompany.Location = new System.Drawing.Point(162, 15);
            this.tbCompany.Margin = new System.Windows.Forms.Padding(4);
            this.tbCompany.Name = "tbCompany";
            this.tbCompany.Size = new System.Drawing.Size(649, 37);
            this.tbCompany.TabIndex = 16;
            this.tbCompany.Enter += new System.EventHandler(this.tbCompany_Enter);
            this.tbCompany.KeyDown += new System.Windows.Forms.KeyEventHandler(this.tbCompany_KeyDown);
            // 
            // lbCompany
            // 
            this.lbCompany.AutoSize = true;
            this.lbCompany.Location = new System.Drawing.Point(18, 18);
            this.lbCompany.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.lbCompany.Name = "lbCompany";
            this.lbCompany.Size = new System.Drawing.Size(124, 30);
            this.lbCompany.TabIndex = 15;
            this.lbCompany.Text = "업체명    : ";
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.dgvFiles);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel1.Location = new System.Drawing.Point(4, 4);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(1126, 397);
            this.panel1.TabIndex = 16;
            // 
            // dgvFiles
            // 
            this.dgvFiles.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvFiles.Location = new System.Drawing.Point(0, 50);
            this.dgvFiles.Name = "dgvFiles";
            this.dgvFiles.RowHeadersWidth = 62;
            this.dgvFiles.RowTemplate.Height = 30;
            this.dgvFiles.Size = new System.Drawing.Size(1611, 628);
            this.dgvFiles.TabIndex = 15;
            this.dgvFiles.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView1_CellContentClick);
            // 
            // tapLog
            // 
            this.tapLog.Controls.Add(this.all);
            this.tapLog.Controls.Add(this.info);
            this.tapLog.Controls.Add(this.tabError);
            this.tapLog.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tapLog.Location = new System.Drawing.Point(5, 5);
            this.tapLog.Name = "tapLog";
            this.tapLog.SelectedIndex = 0;
            this.tapLog.Size = new System.Drawing.Size(1142, 252);
            this.tapLog.TabIndex = 0;
            // 
            // all
            // 
            this.all.Controls.Add(this.richTextBox1);
            this.all.Location = new System.Drawing.Point(4, 39);
            this.all.Name = "all";
            this.all.Padding = new System.Windows.Forms.Padding(5);
            this.all.Size = new System.Drawing.Size(1134, 209);
            this.all.TabIndex = 0;
            this.all.Text = "All log";
            this.all.UseVisualStyleBackColor = true;
            // 
            // richTextBox1
            // 
            this.richTextBox1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.richTextBox1.DetectUrls = false;
            this.richTextBox1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.richTextBox1.Location = new System.Drawing.Point(5, 5);
            this.richTextBox1.Name = "richTextBox1";
            this.richTextBox1.ReadOnly = true;
            this.richTextBox1.Size = new System.Drawing.Size(1124, 199);
            this.richTextBox1.TabIndex = 1;
            this.richTextBox1.TabStop = false;
            this.richTextBox1.Text = "";
            this.richTextBox1.WordWrap = false;
            // 
            // info
            // 
            this.info.Location = new System.Drawing.Point(4, 25);
            this.info.Name = "info";
            this.info.Padding = new System.Windows.Forms.Padding(3);
            this.info.Size = new System.Drawing.Size(1134, 223);
            this.info.TabIndex = 1;
            this.info.Text = "Info";
            this.info.UseVisualStyleBackColor = true;
            // 
            // tabError
            // 
            this.tabError.Location = new System.Drawing.Point(4, 25);
            this.tabError.Name = "tabError";
            this.tabError.Padding = new System.Windows.Forms.Padding(5);
            this.tabError.Size = new System.Drawing.Size(1134, 223);
            this.tabError.TabIndex = 2;
            this.tabError.Text = "Error";
            this.tabError.UseVisualStyleBackColor = true;
            // 
            // pnlTop
            // 
            this.pnlTop.Controls.Add(this.btnDownloadWork);
            this.pnlTop.Controls.Add(this.textBoxKeyword);
            this.pnlTop.Controls.Add(this.label3);
            this.pnlTop.Controls.Add(this.checkBoxAnnual);
            this.pnlTop.Controls.Add(this.textBox1);
            this.pnlTop.Controls.Add(this.lbSample);
            this.pnlTop.Controls.Add(this.cbAutoGenerateExcel);
            this.pnlTop.Controls.Add(this.comboBoxSeason);
            this.pnlTop.Controls.Add(this.label2);
            this.pnlTop.Controls.Add(this.dateTimePicker1);
            this.pnlTop.Controls.Add(this.label1);
            this.pnlTop.Controls.Add(this.tbPerson);
            this.pnlTop.Controls.Add(this.panel3);
            this.pnlTop.Controls.Add(this.tbFolder);
            this.pnlTop.Controls.Add(this.lbFolder);
            this.pnlTop.Controls.Add(this.cbName);
            this.pnlTop.Controls.Add(this.lbName);
            this.pnlTop.Dock = System.Windows.Forms.DockStyle.Top;
            this.pnlTop.Location = new System.Drawing.Point(4, 4);
            this.pnlTop.Margin = new System.Windows.Forms.Padding(4);
            this.pnlTop.Name = "pnlTop";
            this.pnlTop.Padding = new System.Windows.Forms.Padding(3);
            this.pnlTop.Size = new System.Drawing.Size(1152, 175);
            this.pnlTop.TabIndex = 3;
            // 
            // btnDownloadWork
            // 
            this.btnDownloadWork.Location = new System.Drawing.Point(972, 50);
            this.btnDownloadWork.Name = "btnDownloadWork";
            this.btnDownloadWork.Size = new System.Drawing.Size(169, 32);
            this.btnDownloadWork.TabIndex = 51;
            this.btnDownloadWork.Text = "내 폴더 다운로드";
            this.btnDownloadWork.UseVisualStyleBackColor = true;
            this.btnDownloadWork.Click += new System.EventHandler(this.btnDownloadWork_Click);
            // 
            // textBoxKeyword
            // 
            this.textBoxKeyword.Location = new System.Drawing.Point(785, 47);
            this.textBoxKeyword.Name = "textBoxKeyword";
            this.textBoxKeyword.Size = new System.Drawing.Size(167, 37);
            this.textBoxKeyword.TabIndex = 50;
            this.textBoxKeyword.Text = "유랑진";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(612, 49);
            this.label3.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(204, 30);
            this.label3.TabIndex = 49;
            this.label3.Text = "서버 폴더 키워드 : ";
            // 
            // checkBoxAnnual
            // 
            this.checkBoxAnnual.AutoSize = true;
            this.checkBoxAnnual.Location = new System.Drawing.Point(488, 85);
            this.checkBoxAnnual.Name = "checkBoxAnnual";
            this.checkBoxAnnual.Size = new System.Drawing.Size(135, 34);
            this.checkBoxAnnual.TabIndex = 3;
            this.checkBoxAnnual.Text = "연차 여부";
            this.checkBoxAnnual.UseVisualStyleBackColor = true;
            // 
            // textBox1
            // 
            this.textBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.textBox1.Location = new System.Drawing.Point(120, 124);
            this.textBox1.Margin = new System.Windows.Forms.Padding(4);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(838, 37);
            this.textBox1.TabIndex = 48;
            this.textBox1.Text = "D:\\work\\한경이엔지\\0_org\\한경이엔지2본부_26년연차보고서(샘플)_rang.xlsx";
            // 
            // lbSample
            // 
            this.lbSample.AutoSize = true;
            this.lbSample.Location = new System.Drawing.Point(22, 125);
            this.lbSample.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.lbSample.Name = "lbSample";
            this.lbSample.Size = new System.Drawing.Size(130, 30);
            this.lbSample.TabIndex = 47;
            this.lbSample.Text = "샘플 위치 : ";
            // 
            // cbAutoGenerateExcel
            // 
            this.cbAutoGenerateExcel.AutoSize = true;
            this.cbAutoGenerateExcel.Checked = true;
            this.cbAutoGenerateExcel.CheckState = System.Windows.Forms.CheckState.Checked;
            this.cbAutoGenerateExcel.Location = new System.Drawing.Point(809, 85);
            this.cbAutoGenerateExcel.Name = "cbAutoGenerateExcel";
            this.cbAutoGenerateExcel.Size = new System.Drawing.Size(187, 34);
            this.cbAutoGenerateExcel.TabIndex = 46;
            this.cbAutoGenerateExcel.Text = "파일 자동 생성";
            this.cbAutoGenerateExcel.UseVisualStyleBackColor = true;
            // 
            // comboBoxSeason
            // 
            this.comboBoxSeason.FormattingEnabled = true;
            this.comboBoxSeason.Items.AddRange(new object[] {
            "1",
            "2",
            "3",
            "4"});
            this.comboBoxSeason.Location = new System.Drawing.Point(384, 82);
            this.comboBoxSeason.Name = "comboBoxSeason";
            this.comboBoxSeason.Size = new System.Drawing.Size(53, 38);
            this.comboBoxSeason.TabIndex = 45;
            this.comboBoxSeason.Text = "1";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(435, 87);
            this.label2.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(57, 30);
            this.label2.TabIndex = 15;
            this.label2.Text = "분기";
            // 
            // dateTimePicker1
            // 
            this.dateTimePicker1.Location = new System.Drawing.Point(120, 82);
            this.dateTimePicker1.Name = "dateTimePicker1";
            this.dateTimePicker1.Size = new System.Drawing.Size(253, 37);
            this.dateTimePicker1.TabIndex = 14;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(22, 87);
            this.label1.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(132, 30);
            this.label1.TabIndex = 13;
            this.label1.Text = "측정일     : ";
            // 
            // tbPerson
            // 
            this.tbPerson.Location = new System.Drawing.Point(120, 44);
            this.tbPerson.Name = "tbPerson";
            this.tbPerson.Size = new System.Drawing.Size(317, 37);
            this.tbPerson.TabIndex = 12;
            this.tbPerson.Text = "김희철 이사 외 1명";
            // 
            // panel3
            // 
            this.panel3.Anchor = System.Windows.Forms.AnchorStyles.Right;
            this.panel3.Controls.Add(this.btFolder);
            this.panel3.Location = new System.Drawing.Point(972, 6);
            this.panel3.Name = "panel3";
            this.panel3.Padding = new System.Windows.Forms.Padding(3);
            this.panel3.Size = new System.Drawing.Size(172, 38);
            this.panel3.TabIndex = 11;
            // 
            // btFolder
            // 
            this.btFolder.Dock = System.Windows.Forms.DockStyle.Fill;
            this.btFolder.Location = new System.Drawing.Point(3, 3);
            this.btFolder.Margin = new System.Windows.Forms.Padding(4);
            this.btFolder.Name = "btFolder";
            this.btFolder.Size = new System.Drawing.Size(166, 32);
            this.btFolder.TabIndex = 10;
            this.btFolder.Text = "폴더 선택";
            this.btFolder.UseVisualStyleBackColor = true;
            this.btFolder.Click += new System.EventHandler(this.btFolder_Click);
            // 
            // tbFolder
            // 
            this.tbFolder.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.tbFolder.Location = new System.Drawing.Point(120, 8);
            this.tbFolder.Margin = new System.Windows.Forms.Padding(4);
            this.tbFolder.Name = "tbFolder";
            this.tbFolder.Size = new System.Drawing.Size(838, 37);
            this.tbFolder.TabIndex = 9;
            // 
            // lbFolder
            // 
            this.lbFolder.AutoSize = true;
            this.lbFolder.Location = new System.Drawing.Point(22, 11);
            this.lbFolder.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.lbFolder.Name = "lbFolder";
            this.lbFolder.Size = new System.Drawing.Size(130, 30);
            this.lbFolder.TabIndex = 8;
            this.lbFolder.Text = "폴더 위치 : ";
            // 
            // cbName
            // 
            this.cbName.FormattingEnabled = true;
            this.cbName.Items.AddRange(new object[] {
            "김희철이사",
            "김영철과장",
            "심재현과장",
            "김희용대리",
            "서원진대리",
            "문동민대리"});
            this.cbName.Location = new System.Drawing.Point(1438, 4);
            this.cbName.Margin = new System.Windows.Forms.Padding(4);
            this.cbName.Name = "cbName";
            this.cbName.Size = new System.Drawing.Size(306, 38);
            this.cbName.TabIndex = 6;
            this.cbName.Text = "김희철 이사";
            this.cbName.Visible = false;
            // 
            // lbName
            // 
            this.lbName.AutoSize = true;
            this.lbName.Location = new System.Drawing.Point(22, 49);
            this.lbName.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.lbName.Name = "lbName";
            this.lbName.Size = new System.Drawing.Size(132, 30);
            this.lbName.TabIndex = 5;
            this.lbName.Text = "측정자     : ";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(12F, 30F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1160, 925);
            this.Controls.Add(this.pnlMain);
            this.Font = new System.Drawing.Font("맑은 고딕", 11F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(129)));
            this.Margin = new System.Windows.Forms.Padding(4);
            this.Name = "Form1";
            this.Text = "한경이엔지 보고서";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.pnlMain.ResumeLayout(false);
            this.pnlMain.PerformLayout();
            this.statusStrip1.ResumeLayout(false);
            this.statusStrip1.PerformLayout();
            this.splitContainer1.Panel1.ResumeLayout(false);
            this.splitContainer1.Panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).EndInit();
            this.splitContainer1.ResumeLayout(false);
            this.tabMain.ResumeLayout(false);
            this.tpFunction.ResumeLayout(false);
            this.gbReport.ResumeLayout(false);
            this.gbReport.PerformLayout();
            this.gbConvertValue.ResumeLayout(false);
            this.gbConvertValue.PerformLayout();
            this.gbImage.ResumeLayout(false);
            this.gbImage.PerformLayout();
            this.tpGap.ResumeLayout(false);
            this.panel2.ResumeLayout(false);
            this.panel2.PerformLayout();
            this.panel1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dgvFiles)).EndInit();
            this.tapLog.ResumeLayout(false);
            this.all.ResumeLayout(false);
            this.pnlTop.ResumeLayout(false);
            this.pnlTop.PerformLayout();
            this.panel3.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion
        private Panel pnlMain;
        private TabControl tabMain;
        private TabPage tpGap;
        private TabPage tpFunction;
        private Panel panel1;
        private DataGridView dgvFiles;
        private Label lbName;
        private ComboBox cbName;
        private Label lbFolder;
        private Button btFolder;
        private TextBox tbFolder;
        private Panel pnlTop;
        private Panel panel2;
        private TextBox tbCompany;
        private Label lbCompany;
        private Button btnDownload;
        private Button btnPageNumber;
        private Button btnExportForPdf;
        private Button btErrorPageUpdate;
        private GroupBox gbImage;
        private Button btnQuntityFileRun;
        private CheckBox checkBoxQuantity;
        private CheckBox checkBoxCorona;
        private CheckBox checkBoxPicture;
        private CheckBox checkBoxFeverPicture;
        private TextBox textBoxFeverImageFolder;
        private SplitContainer splitContainer1;
        private TabControl tapLog;
        private TabPage all;
        private TabPage info;
        private Panel panel3;
        private TextBox tbPerson;
        private TabPage tabError;
        private TextBox textBoxQuntatyFolder;
        private TextBox textBoxPictureFolder;
        private TextBox textBoxCoronaFolder;
        private Button btnGapjiPictureRelocate;
        private ComboBox comboBoxSignalGraph;
        private ComboBox comboBoxTestGraph;
        private ComboBox comboBoxHwaveGraph;
        private StatusStrip statusStrip1;
        private ToolStripStatusLabel toolStripStatusLabel1;
        private RichTextBox richTextBox1;
        private DateTimePicker dateTimePicker1;
        private Label label1;
        private GroupBox gbConvertValue;
        private CheckBox checkBox1;
        private CheckBox checkBoxDate;
        private CheckBox checkBoxChecker;
        private ComboBox comboBoxSeason;
        private Label label2;
        private GroupBox gbReport;
        private Button btnQuantityFile;
        private TextBox tbQuantityFile;
        private Label lbQuantityFile;
        private CheckBox checkBoxOcr;
        private TextBox textBox1;
        private Label lbSample;
        private CheckBox cbAutoGenerateExcel;
        private CheckBox checkBoxAnnual;
        private TextBox textBoxKeyword;
        private Label label3;
        private Button btnDownloadWork;
    }
}

