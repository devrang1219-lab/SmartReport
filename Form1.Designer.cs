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
            this.tabMain = new System.Windows.Forms.TabControl();
            this.tpGap = new System.Windows.Forms.TabPage();
            this.panel2 = new System.Windows.Forms.Panel();
            this.btnDownload = new System.Windows.Forms.Button();
            this.tbCompany = new System.Windows.Forms.TextBox();
            this.lbCompany = new System.Windows.Forms.Label();
            this.panel1 = new System.Windows.Forms.Panel();
            this.dgvFiles = new System.Windows.Forms.DataGridView();
            this.tpFunction = new System.Windows.Forms.TabPage();
            this.panel4 = new System.Windows.Forms.Panel();
            this.tbDefultDirectory = new System.Windows.Forms.TextBox();
            this.tbGapFeverPicture = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.lblGapFeverPicture = new System.Windows.Forms.Label();
            this.tbWidthFeverPicture = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.tbHeightFeverPicture = new System.Windows.Forms.TextBox();
            this.lblWidthFeverPicture = new System.Windows.Forms.Label();
            this.lblHeightFeverPicture = new System.Windows.Forms.Label();
            this.button3 = new System.Windows.Forms.Button();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.panel3 = new System.Windows.Forms.Panel();
            this.btErrorPageUpdate = new System.Windows.Forms.Button();
            this.cbAutoExportPdf = new System.Windows.Forms.CheckBox();
            this.btnExportForPdf = new System.Windows.Forms.Button();
            this.btnFindFileForFunction = new System.Windows.Forms.Button();
            this.tbFileNameForFunction = new System.Windows.Forms.TextBox();
            this.lblFileNmaeForFunction = new System.Windows.Forms.Label();
            this.btnPageNumber = new System.Windows.Forms.Button();
            this.lbName = new System.Windows.Forms.Label();
            this.cbName = new System.Windows.Forms.ComboBox();
            this.comboBox1 = new System.Windows.Forms.ComboBox();
            this.lbFolder = new System.Windows.Forms.Label();
            this.btFolder = new System.Windows.Forms.Button();
            this.tbFolder = new System.Windows.Forms.TextBox();
            this.pnlTop = new System.Windows.Forms.Panel();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.btnQuantityFile = new System.Windows.Forms.Button();
            this.tbQuantityFile = new System.Windows.Forms.TextBox();
            this.lbQuantityFile = new System.Windows.Forms.Label();
            this.btnQuntityFileRun = new System.Windows.Forms.Button();
            this.pnlMain.SuspendLayout();
            this.tabMain.SuspendLayout();
            this.tpGap.SuspendLayout();
            this.panel2.SuspendLayout();
            this.panel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvFiles)).BeginInit();
            this.tpFunction.SuspendLayout();
            this.panel4.SuspendLayout();
            this.panel3.SuspendLayout();
            this.pnlTop.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // pnlMain
            // 
            this.pnlMain.Controls.Add(this.tabMain);
            this.pnlMain.Dock = System.Windows.Forms.DockStyle.Fill;
            this.pnlMain.Location = new System.Drawing.Point(0, 186);
            this.pnlMain.Margin = new System.Windows.Forms.Padding(4);
            this.pnlMain.Name = "pnlMain";
            this.pnlMain.Size = new System.Drawing.Size(1633, 730);
            this.pnlMain.TabIndex = 4;
            // 
            // tabMain
            // 
            this.tabMain.Controls.Add(this.tpGap);
            this.tabMain.Controls.Add(this.tpFunction);
            this.tabMain.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tabMain.Location = new System.Drawing.Point(0, 0);
            this.tabMain.Margin = new System.Windows.Forms.Padding(4);
            this.tabMain.Name = "tabMain";
            this.tabMain.SelectedIndex = 0;
            this.tabMain.Size = new System.Drawing.Size(1633, 730);
            this.tabMain.TabIndex = 0;
            // 
            // tpGap
            // 
            this.tpGap.Controls.Add(this.panel2);
            this.tpGap.Controls.Add(this.panel1);
            this.tpGap.Location = new System.Drawing.Point(4, 39);
            this.tpGap.Margin = new System.Windows.Forms.Padding(4);
            this.tpGap.Name = "tpGap";
            this.tpGap.Padding = new System.Windows.Forms.Padding(4);
            this.tpGap.Size = new System.Drawing.Size(1625, 687);
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
            this.panel2.Size = new System.Drawing.Size(1617, 44);
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
            this.panel1.Size = new System.Drawing.Size(1617, 679);
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
            // tpFunction
            // 
            this.tpFunction.Controls.Add(this.groupBox1);
            this.tpFunction.Controls.Add(this.panel4);
            this.tpFunction.Controls.Add(this.panel3);
            this.tpFunction.Location = new System.Drawing.Point(4, 39);
            this.tpFunction.Margin = new System.Windows.Forms.Padding(4);
            this.tpFunction.Name = "tpFunction";
            this.tpFunction.Padding = new System.Windows.Forms.Padding(4);
            this.tpFunction.Size = new System.Drawing.Size(1625, 687);
            this.tpFunction.TabIndex = 1;
            this.tpFunction.Text = "기능";
            this.tpFunction.UseVisualStyleBackColor = true;
            // 
            // panel4
            // 
            this.panel4.Controls.Add(this.tbDefultDirectory);
            this.panel4.Controls.Add(this.tbGapFeverPicture);
            this.panel4.Controls.Add(this.label4);
            this.panel4.Controls.Add(this.lblGapFeverPicture);
            this.panel4.Controls.Add(this.tbWidthFeverPicture);
            this.panel4.Controls.Add(this.label3);
            this.panel4.Controls.Add(this.label2);
            this.panel4.Controls.Add(this.tbHeightFeverPicture);
            this.panel4.Controls.Add(this.lblWidthFeverPicture);
            this.panel4.Controls.Add(this.lblHeightFeverPicture);
            this.panel4.Controls.Add(this.button3);
            this.panel4.Controls.Add(this.textBox1);
            this.panel4.Controls.Add(this.label1);
            this.panel4.Location = new System.Drawing.Point(8, 239);
            this.panel4.Name = "panel4";
            this.panel4.Size = new System.Drawing.Size(1610, 213);
            this.panel4.TabIndex = 23;
            // 
            // tbDefultDirectory
            // 
            this.tbDefultDirectory.Location = new System.Drawing.Point(959, 93);
            this.tbDefultDirectory.Name = "tbDefultDirectory";
            this.tbDefultDirectory.Size = new System.Drawing.Size(483, 37);
            this.tbDefultDirectory.TabIndex = 29;
            this.tbDefultDirectory.Text = "C:\\_D\\work\\한경이엔지\\NAS다운로드";
            // 
            // tbGapFeverPicture
            // 
            this.tbGapFeverPicture.Location = new System.Drawing.Point(724, 43);
            this.tbGapFeverPicture.Margin = new System.Windows.Forms.Padding(4);
            this.tbGapFeverPicture.Name = "tbGapFeverPicture";
            this.tbGapFeverPicture.Size = new System.Drawing.Size(121, 37);
            this.tbGapFeverPicture.TabIndex = 28;
            this.tbGapFeverPicture.Text = "0.2";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(853, 46);
            this.label4.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label4.Name = "label4";
            this.label4.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.label4.Size = new System.Drawing.Size(42, 30);
            this.label4.TabIndex = 27;
            this.label4.Text = "cm";
            // 
            // lblGapFeverPicture
            // 
            this.lblGapFeverPicture.AutoSize = true;
            this.lblGapFeverPicture.Location = new System.Drawing.Point(627, 45);
            this.lblGapFeverPicture.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.lblGapFeverPicture.Name = "lblGapFeverPicture";
            this.lblGapFeverPicture.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.lblGapFeverPicture.Size = new System.Drawing.Size(78, 30);
            this.lblGapFeverPicture.TabIndex = 26;
            this.lblGapFeverPicture.Text = "간격 : ";
            // 
            // tbWidthFeverPicture
            // 
            this.tbWidthFeverPicture.Location = new System.Drawing.Point(398, 38);
            this.tbWidthFeverPicture.Margin = new System.Windows.Forms.Padding(4);
            this.tbWidthFeverPicture.Name = "tbWidthFeverPicture";
            this.tbWidthFeverPicture.Size = new System.Drawing.Size(121, 37);
            this.tbWidthFeverPicture.TabIndex = 25;
            this.tbWidthFeverPicture.Text = "8.9";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(527, 41);
            this.label3.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label3.Name = "label3";
            this.label3.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.label3.Size = new System.Drawing.Size(42, 30);
            this.label3.TabIndex = 24;
            this.label3.Text = "cm";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(226, 45);
            this.label2.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label2.Name = "label2";
            this.label2.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.label2.Size = new System.Drawing.Size(42, 30);
            this.label2.TabIndex = 23;
            this.label2.Text = "cm";
            // 
            // tbHeightFeverPicture
            // 
            this.tbHeightFeverPicture.Location = new System.Drawing.Point(97, 38);
            this.tbHeightFeverPicture.Margin = new System.Windows.Forms.Padding(4);
            this.tbHeightFeverPicture.Name = "tbHeightFeverPicture";
            this.tbHeightFeverPicture.Size = new System.Drawing.Size(121, 37);
            this.tbHeightFeverPicture.TabIndex = 22;
            this.tbHeightFeverPicture.Text = "7.63";
            // 
            // lblWidthFeverPicture
            // 
            this.lblWidthFeverPicture.AutoSize = true;
            this.lblWidthFeverPicture.Location = new System.Drawing.Point(321, 38);
            this.lblWidthFeverPicture.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.lblWidthFeverPicture.Name = "lblWidthFeverPicture";
            this.lblWidthFeverPicture.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.lblWidthFeverPicture.Size = new System.Drawing.Size(78, 30);
            this.lblWidthFeverPicture.TabIndex = 21;
            this.lblWidthFeverPicture.Text = "너비 : ";
            // 
            // lblHeightFeverPicture
            // 
            this.lblHeightFeverPicture.AutoSize = true;
            this.lblHeightFeverPicture.Location = new System.Drawing.Point(14, 38);
            this.lblHeightFeverPicture.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.lblHeightFeverPicture.Name = "lblHeightFeverPicture";
            this.lblHeightFeverPicture.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.lblHeightFeverPicture.Size = new System.Drawing.Size(78, 30);
            this.lblHeightFeverPicture.TabIndex = 20;
            this.lblHeightFeverPicture.Text = "높이 : ";
            // 
            // button3
            // 
            this.button3.Location = new System.Drawing.Point(745, 87);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(182, 41);
            this.button3.TabIndex = 19;
            this.button3.Text = "찾기";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(158, 90);
            this.textBox1.Margin = new System.Windows.Forms.Padding(4);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(580, 37);
            this.textBox1.TabIndex = 18;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(14, 93);
            this.label1.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(124, 30);
            this.label1.TabIndex = 17;
            this.label1.Text = "폴더명    : ";
            // 
            // panel3
            // 
            this.panel3.Controls.Add(this.btErrorPageUpdate);
            this.panel3.Controls.Add(this.cbAutoExportPdf);
            this.panel3.Controls.Add(this.btnExportForPdf);
            this.panel3.Controls.Add(this.btnFindFileForFunction);
            this.panel3.Controls.Add(this.tbFileNameForFunction);
            this.panel3.Controls.Add(this.lblFileNmaeForFunction);
            this.panel3.Controls.Add(this.btnPageNumber);
            this.panel3.Location = new System.Drawing.Point(8, 7);
            this.panel3.Name = "panel3";
            this.panel3.Size = new System.Drawing.Size(1610, 213);
            this.panel3.TabIndex = 0;
            // 
            // btErrorPageUpdate
            // 
            this.btErrorPageUpdate.Location = new System.Drawing.Point(295, 84);
            this.btErrorPageUpdate.Name = "btErrorPageUpdate";
            this.btErrorPageUpdate.Size = new System.Drawing.Size(199, 95);
            this.btErrorPageUpdate.TabIndex = 22;
            this.btErrorPageUpdate.Text = "의견 부적합 페이지 업데이트";
            this.btErrorPageUpdate.UseVisualStyleBackColor = true;
            this.btErrorPageUpdate.Click += new System.EventHandler(this.btErrorPageUpdate_Click);
            // 
            // cbAutoExportPdf
            // 
            this.cbAutoExportPdf.AutoSize = true;
            this.cbAutoExportPdf.Location = new System.Drawing.Point(1053, 20);
            this.cbAutoExportPdf.Name = "cbAutoExportPdf";
            this.cbAutoExportPdf.Size = new System.Drawing.Size(471, 34);
            this.cbAutoExportPdf.TabIndex = 21;
            this.cbAutoExportPdf.Text = "자동으로 페이지 번호 매기고 PDF 내보내기";
            this.cbAutoExportPdf.UseVisualStyleBackColor = true;
            // 
            // btnExportForPdf
            // 
            this.btnExportForPdf.Location = new System.Drawing.Point(509, 84);
            this.btnExportForPdf.Name = "btnExportForPdf";
            this.btnExportForPdf.Size = new System.Drawing.Size(270, 95);
            this.btnExportForPdf.TabIndex = 20;
            this.btnExportForPdf.Text = "PDF로 내보내기";
            this.btnExportForPdf.UseVisualStyleBackColor = true;
            this.btnExportForPdf.MouseUp += new System.Windows.Forms.MouseEventHandler(this.btnExportForPdf_MouseUp);
            // 
            // btnFindFileForFunction
            // 
            this.btnFindFileForFunction.Location = new System.Drawing.Point(814, 16);
            this.btnFindFileForFunction.Name = "btnFindFileForFunction";
            this.btnFindFileForFunction.Size = new System.Drawing.Size(182, 41);
            this.btnFindFileForFunction.TabIndex = 19;
            this.btnFindFileForFunction.Text = "찾기";
            this.btnFindFileForFunction.UseVisualStyleBackColor = true;
            this.btnFindFileForFunction.Click += new System.EventHandler(this.btnFindFileForFunction_Click);
            // 
            // tbFileNameForFunction
            // 
            this.tbFileNameForFunction.Location = new System.Drawing.Point(158, 20);
            this.tbFileNameForFunction.Margin = new System.Windows.Forms.Padding(4);
            this.tbFileNameForFunction.Name = "tbFileNameForFunction";
            this.tbFileNameForFunction.Size = new System.Drawing.Size(649, 37);
            this.tbFileNameForFunction.TabIndex = 18;
            // 
            // lblFileNmaeForFunction
            // 
            this.lblFileNmaeForFunction.AutoSize = true;
            this.lblFileNmaeForFunction.Location = new System.Drawing.Point(14, 23);
            this.lblFileNmaeForFunction.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.lblFileNmaeForFunction.Name = "lblFileNmaeForFunction";
            this.lblFileNmaeForFunction.Size = new System.Drawing.Size(124, 30);
            this.lblFileNmaeForFunction.TabIndex = 17;
            this.lblFileNmaeForFunction.Text = "파일명    : ";
            // 
            // btnPageNumber
            // 
            this.btnPageNumber.Location = new System.Drawing.Point(19, 84);
            this.btnPageNumber.Name = "btnPageNumber";
            this.btnPageNumber.Size = new System.Drawing.Size(270, 95);
            this.btnPageNumber.TabIndex = 0;
            this.btnPageNumber.Text = "페이지번호 업데이트";
            this.btnPageNumber.UseVisualStyleBackColor = true;
            this.btnPageNumber.Click += new System.EventHandler(this.btnPageNumber_Click);
            // 
            // lbName
            // 
            this.lbName.AutoSize = true;
            this.lbName.Location = new System.Drawing.Point(18, 22);
            this.lbName.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.lbName.Name = "lbName";
            this.lbName.Size = new System.Drawing.Size(132, 30);
            this.lbName.TabIndex = 5;
            this.lbName.Text = "측정자     : ";
            // 
            // cbName
            // 
            this.cbName.FormattingEnabled = true;
            this.cbName.Location = new System.Drawing.Point(164, 14);
            this.cbName.Margin = new System.Windows.Forms.Padding(4);
            this.cbName.Name = "cbName";
            this.cbName.Size = new System.Drawing.Size(306, 38);
            this.cbName.TabIndex = 6;
            this.cbName.Text = "김희철 이사";
            // 
            // comboBox1
            // 
            this.comboBox1.FormattingEnabled = true;
            this.comboBox1.Location = new System.Drawing.Point(506, 14);
            this.comboBox1.Margin = new System.Windows.Forms.Padding(4);
            this.comboBox1.Name = "comboBox1";
            this.comboBox1.Size = new System.Drawing.Size(306, 38);
            this.comboBox1.TabIndex = 7;
            this.comboBox1.Text = "김희철 이사";
            // 
            // lbFolder
            // 
            this.lbFolder.AutoSize = true;
            this.lbFolder.Location = new System.Drawing.Point(19, 73);
            this.lbFolder.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.lbFolder.Name = "lbFolder";
            this.lbFolder.Size = new System.Drawing.Size(130, 30);
            this.lbFolder.TabIndex = 8;
            this.lbFolder.Text = "폴더 위치 : ";
            // 
            // btFolder
            // 
            this.btFolder.Location = new System.Drawing.Point(845, 4);
            this.btFolder.Margin = new System.Windows.Forms.Padding(4);
            this.btFolder.Name = "btFolder";
            this.btFolder.Size = new System.Drawing.Size(134, 97);
            this.btFolder.TabIndex = 10;
            this.btFolder.Text = "폴더 선택";
            this.btFolder.UseVisualStyleBackColor = true;
            this.btFolder.Click += new System.EventHandler(this.btFolder_Click);
            // 
            // tbFolder
            // 
            this.tbFolder.Location = new System.Drawing.Point(163, 64);
            this.tbFolder.Margin = new System.Windows.Forms.Padding(4);
            this.tbFolder.Name = "tbFolder";
            this.tbFolder.Size = new System.Drawing.Size(649, 37);
            this.tbFolder.TabIndex = 9;
            // 
            // pnlTop
            // 
            this.pnlTop.Controls.Add(this.tbFolder);
            this.pnlTop.Controls.Add(this.btFolder);
            this.pnlTop.Controls.Add(this.lbFolder);
            this.pnlTop.Controls.Add(this.comboBox1);
            this.pnlTop.Controls.Add(this.cbName);
            this.pnlTop.Controls.Add(this.lbName);
            this.pnlTop.Dock = System.Windows.Forms.DockStyle.Top;
            this.pnlTop.Location = new System.Drawing.Point(0, 0);
            this.pnlTop.Margin = new System.Windows.Forms.Padding(4);
            this.pnlTop.Name = "pnlTop";
            this.pnlTop.Size = new System.Drawing.Size(1633, 186);
            this.pnlTop.TabIndex = 3;
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.btnQuntityFileRun);
            this.groupBox1.Controls.Add(this.btnQuantityFile);
            this.groupBox1.Controls.Add(this.tbQuantityFile);
            this.groupBox1.Controls.Add(this.lbQuantityFile);
            this.groupBox1.Location = new System.Drawing.Point(8, 476);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(1398, 152);
            this.groupBox1.TabIndex = 24;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "품질";
            // 
            // btnQuantityFile
            // 
            this.btnQuantityFile.Location = new System.Drawing.Point(814, 37);
            this.btnQuantityFile.Name = "btnQuantityFile";
            this.btnQuantityFile.Size = new System.Drawing.Size(182, 41);
            this.btnQuantityFile.TabIndex = 22;
            this.btnQuantityFile.Text = "찾기";
            this.btnQuantityFile.UseVisualStyleBackColor = true;
            this.btnQuantityFile.Click += new System.EventHandler(this.btnQuantityFile_Click);
            // 
            // tbQuantityFile
            // 
            this.tbQuantityFile.Location = new System.Drawing.Point(158, 41);
            this.tbQuantityFile.Margin = new System.Windows.Forms.Padding(4);
            this.tbQuantityFile.Name = "tbQuantityFile";
            this.tbQuantityFile.Size = new System.Drawing.Size(649, 37);
            this.tbQuantityFile.TabIndex = 21;
            // 
            // lbQuantityFile
            // 
            this.lbQuantityFile.AutoSize = true;
            this.lbQuantityFile.Location = new System.Drawing.Point(14, 44);
            this.lbQuantityFile.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.lbQuantityFile.Name = "lbQuantityFile";
            this.lbQuantityFile.Size = new System.Drawing.Size(124, 30);
            this.lbQuantityFile.TabIndex = 20;
            this.lbQuantityFile.Text = "파일명    : ";
            // 
            // btnQuntityFileRun
            // 
            this.btnQuntityFileRun.Location = new System.Drawing.Point(125, 107);
            this.btnQuntityFileRun.Name = "btnQuntityFileRun";
            this.btnQuntityFileRun.Size = new System.Drawing.Size(218, 39);
            this.btnQuntityFileRun.TabIndex = 23;
            this.btnQuntityFileRun.Text = "실행";
            this.btnQuntityFileRun.UseVisualStyleBackColor = true;
            this.btnQuntityFileRun.Click += new System.EventHandler(this.btnQuntityFileRun_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(12F, 30F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1633, 916);
            this.Controls.Add(this.pnlMain);
            this.Controls.Add(this.pnlTop);
            this.Font = new System.Drawing.Font("맑은 고딕", 11F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(129)));
            this.Margin = new System.Windows.Forms.Padding(4);
            this.Name = "Form1";
            this.Text = "한경이엔지 보고서";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.pnlMain.ResumeLayout(false);
            this.tabMain.ResumeLayout(false);
            this.tpGap.ResumeLayout(false);
            this.panel2.ResumeLayout(false);
            this.panel2.PerformLayout();
            this.panel1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dgvFiles)).EndInit();
            this.tpFunction.ResumeLayout(false);
            this.panel4.ResumeLayout(false);
            this.panel4.PerformLayout();
            this.panel3.ResumeLayout(false);
            this.panel3.PerformLayout();
            this.pnlTop.ResumeLayout(false);
            this.pnlTop.PerformLayout();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
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
        private ComboBox comboBox1;
        private Label lbFolder;
        private Button btFolder;
        private TextBox tbFolder;
        private Panel pnlTop;
        private Panel panel2;
        private TextBox tbCompany;
        private Label lbCompany;
        private Button btnDownload;
        private Panel panel3;
        private Button btnPageNumber;
        private Button btnFindFileForFunction;
        private TextBox tbFileNameForFunction;
        private Label lblFileNmaeForFunction;
        private CheckBox cbAutoExportPdf;
        private Button btnExportForPdf;
        private Button btErrorPageUpdate;
        private Panel panel4;
        private Button button3;
        private TextBox textBox1;
        private Label label1;
        private Label lblHeightFeverPicture;
        private TextBox tbGapFeverPicture;
        private Label label4;
        private Label lblGapFeverPicture;
        private TextBox tbWidthFeverPicture;
        private Label label3;
        private Label label2;
        private TextBox tbHeightFeverPicture;
        private Label lblWidthFeverPicture;
        private TextBox tbDefultDirectory;
        private GroupBox groupBox1;
        private Button btnQuantityFile;
        private TextBox tbQuantityFile;
        private Label lbQuantityFile;
        private Button btnQuntityFileRun;
    }
}

