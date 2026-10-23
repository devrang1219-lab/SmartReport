namespace WindowsFormsApp1
{
    partial class Form2
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

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
            this.pnlTopNav = new System.Windows.Forms.Panel();
            this.btnNavDataMerge = new System.Windows.Forms.Button();
            this.btnNavReportWork = new System.Windows.Forms.Button();
            this.btnNavFileSettings = new System.Windows.Forms.Button();

            this.grpFileTools = new System.Windows.Forms.GroupBox();
            this.btnSelectFile = new System.Windows.Forms.Button();
            this.txtReportFileName = new System.Windows.Forms.TextBox();
            this.lblReportFileName = new System.Windows.Forms.Label();
            this.btnCoverAlign = new System.Windows.Forms.Button();
            this.btnPageNumUpdate = new System.Windows.Forms.Button();
            this.btnFooterChange = new System.Windows.Forms.Button();
            this.btnDateUpdate = new System.Windows.Forms.Button();
            this.btnPdfExport = new System.Windows.Forms.Button();
            this.btnAutoDivider = new System.Windows.Forms.Button();
            this.btnAddSignature = new System.Windows.Forms.Button();
            this.btnChangeInspector = new System.Windows.Forms.Button();

            this.grpImageSettings = new System.Windows.Forms.GroupBox();
            this.lblSheetName = new System.Windows.Forms.Label();
            this.cmbSheetName = new System.Windows.Forms.ComboBox();
            this.btnRemoveSheetImg = new System.Windows.Forms.Button();
            this.btnOptimizeImg = new System.Windows.Forms.Button();
            this.btnFillSheetImg = new System.Windows.Forms.Button();

            this.grpCellPadding = new System.Windows.Forms.GroupBox();
            this.txtTop = new System.Windows.Forms.TextBox();
            this.txtLeft = new System.Windows.Forms.TextBox();
            this.txtBottom = new System.Windows.Forms.TextBox();
            this.txtRight = new System.Windows.Forms.TextBox();
            this.lblTop = new System.Windows.Forms.Label();
            this.lblLeft = new System.Windows.Forms.Label();
            this.lblBottom = new System.Windows.Forms.Label();
            this.lblRight = new System.Windows.Forms.Label();

            this.grpConfigMgmt = new System.Windows.Forms.GroupBox();
            this.chkCorrectionActive = new System.Windows.Forms.CheckBox();
            this.btnUpdateList = new System.Windows.Forms.Button();
            this.chkBackgroundYellow = new System.Windows.Forms.CheckBox();
            this.btnUpdateInvalidPage = new System.Windows.Forms.Button();
            this.btnExecute = new System.Windows.Forms.Button();

            this.grpDataMerge = new System.Windows.Forms.GroupBox();
            this.chkAutoTemp = new System.Windows.Forms.CheckBox();
            // 데이터 병합 내부 행 컨트롤들
            this.chkRow1 = new System.Windows.Forms.CheckBox();
            this.txtRow1Label = new System.Windows.Forms.TextBox();
            this.txtRow1Folder = new System.Windows.Forms.TextBox();
            this.txtRow1Sheet = new System.Windows.Forms.TextBox();

            this.chkRow2 = new System.Windows.Forms.CheckBox();
            this.txtRow2Label = new System.Windows.Forms.TextBox();
            this.txtRow2Folder = new System.Windows.Forms.TextBox();
            this.txtRow2Sheet = new System.Windows.Forms.TextBox();

            this.chkRow3 = new System.Windows.Forms.CheckBox();
            this.txtRow3Label = new System.Windows.Forms.TextBox();
            this.txtRow3Folder = new System.Windows.Forms.TextBox();
            this.txtRow3Sheet = new System.Windows.Forms.TextBox();

            this.grpExecutionLog = new System.Windows.Forms.GroupBox();
            this.pnlSpinner = new System.Windows.Forms.Panel();

            this.grpLogViewer = new System.Windows.Forms.GroupBox();
            this.pnlLogNav = new System.Windows.Forms.Panel();
            this.btnLogAll = new System.Windows.Forms.Button();
            this.btnLogInfo = new System.Windows.Forms.Button();
            this.btnLogError = new System.Windows.Forms.Button();
            this.txtLogContent = new System.Windows.Forms.TextBox();

            this.pnlTopNav.SuspendLayout();
            this.grpFileTools.SuspendLayout();
            this.grpImageSettings.SuspendLayout();
            this.grpCellPadding.SuspendLayout();
            this.grpConfigMgmt.SuspendLayout();
            this.grpDataMerge.SuspendLayout();
            this.grpLogViewer.SuspendLayout();
            this.pnlLogNav.SuspendLayout();
            this.SuspendLayout();

            // 
            // pnlTopNav
            // 
            this.pnlTopNav.BackColor = System.Drawing.Color.FromArgb(240, 243, 246);
            this.pnlTopNav.Controls.Add(this.btnNavDataMerge);
            this.pnlTopNav.Controls.Add(this.btnNavReportWork);
            this.pnlTopNav.Controls.Add(this.btnNavFileSettings);
            this.pnlTopNav.Dock = System.Windows.Forms.DockStyle.Top;
            this.pnlTopNav.Location = new System.Drawing.Point(0, 0);
            this.pnlTopNav.Name = "pnlTopNav";
            this.pnlTopNav.Size = new System.Drawing.Size(1008, 45);

            this.btnNavFileSettings.Location = new System.Drawing.Point(12, 8);
            this.btnNavFileSettings.Size = new System.Drawing.Size(90, 30);
            this.btnNavFileSettings.Text = "파일/설정";
            this.btnNavFileSettings.UseVisualStyleBackColor = true;

            this.btnNavReportWork.Location = new System.Drawing.Point(108, 8);
            this.btnNavReportWork.Size = new System.Drawing.Size(90, 30);
            this.btnNavReportWork.Text = "보고서 작업";
            this.btnNavReportWork.UseVisualStyleBackColor = true;

            this.btnNavDataMerge.Location = new System.Drawing.Point(204, 8);
            this.btnNavDataMerge.Size = new System.Drawing.Size(90, 30);
            this.btnNavDataMerge.Text = "데이터 병합";
            this.btnNavDataMerge.UseVisualStyleBackColor = true;

            // 
            // grpFileTools
            // 
            this.grpFileTools.Controls.Add(this.btnSelectFile);
            this.grpFileTools.Controls.Add(this.txtReportFileName);
            this.grpFileTools.Controls.Add(this.lblReportFileName);
            this.grpFileTools.Location = new System.Drawing.Point(12, 55);
            this.grpFileTools.Size = new System.Drawing.Size(984, 115);
            this.grpFileTools.Text = "파일 및 주요 도구";

            this.lblReportFileName.AutoSize = true;
            this.lblReportFileName.Location = new System.Drawing.Point(15, 28);
            this.lblReportFileName.Text = "보고서 파일명 :";

            this.txtReportFileName.Location = new System.Drawing.Point(105, 25);
            this.txtReportFileName.Size = new System.Drawing.Size(750, 23);

            this.btnSelectFile.Location = new System.Drawing.Point(865, 24);
            this.btnSelectFile.Size = new System.Drawing.Size(105, 25);
            this.btnSelectFile.Text = "📁 파일 선택";
            this.btnSelectFile.UseVisualStyleBackColor = true;

            int btnX = 15, btnY = 58, btnW = 108, btnH = 45, gap = 10;
            System.Windows.Forms.Button[] tools = { btnCoverAlign, btnPageNumUpdate, btnFooterChange, btnDateUpdate, btnPdfExport, btnAutoDivider, btnAddSignature, btnChangeInspector };
            string[] toolTexts = { "커버 사진 정렬", "페이지 번호 갱신", "바닥글 변경", "날짜 갱신", "PDF 내보내기", "자동 구분선", "서명 추가", "측정자 변경" };

            for (int i = 0; i < tools.Length; i++)
            {
                tools[i].Location = new System.Drawing.Point(btnX + (i * (btnW + gap)), btnY);
                tools[i].Size = new System.Drawing.Size(btnW, btnH);
                tools[i].Text = toolTexts[i];
                tools[i].UseVisualStyleBackColor = true;
                this.grpFileTools.Controls.Add(tools[i]);
            }

            // 
            // grpImageSettings
            // 
            this.grpImageSettings.Controls.Add(this.lblSheetName);
            this.grpImageSettings.Controls.Add(this.cmbSheetName);
            this.grpImageSettings.Controls.Add(this.btnRemoveSheetImg);
            this.grpImageSettings.Controls.Add(this.btnOptimizeImg);
            this.grpImageSettings.Controls.Add(this.btnFillSheetImg);
            this.grpImageSettings.Controls.Add(this.grpCellPadding);
            this.grpImageSettings.Location = new System.Drawing.Point(12, 178);
            this.grpImageSettings.Size = new System.Drawing.Size(490, 150);
            this.grpImageSettings.Text = "이미지 설정";

            this.lblSheetName.AutoSize = true;
            this.lblSheetName.Location = new System.Drawing.Point(15, 28);
            this.lblSheetName.Text = "Sheet name";

            this.cmbSheetName.Location = new System.Drawing.Point(95, 25);
            this.cmbSheetName.Size = new System.Drawing.Size(120, 23);

            this.btnRemoveSheetImg.Location = new System.Drawing.Point(15, 60);
            this.btnRemoveSheetImg.Size = new System.Drawing.Size(65, 75);
            this.btnRemoveSheetImg.Text = "시트 이미지 제거";
            this.btnRemoveSheetImg.UseVisualStyleBackColor = true;

            this.btnOptimizeImg.Location = new System.Drawing.Point(88, 60);
            this.btnOptimizeImg.Size = new System.Drawing.Size(65, 75);
            this.btnOptimizeImg.Text = "시트 이미지 용량 최적화";
            this.btnOptimizeImg.UseVisualStyleBackColor = true;

            this.btnFillSheetImg.Location = new System.Drawing.Point(161, 60);
            this.btnFillSheetImg.Size = new System.Drawing.Size(65, 75);
            this.btnFillSheetImg.Text = "시트 이미지 셀에 채우기";
            this.btnFillSheetImg.UseVisualStyleBackColor = true;

            // grpCellPadding
            this.grpCellPadding.Controls.Add(this.txtTop);
            this.grpCellPadding.Controls.Add(this.txtLeft);
            this.grpCellPadding.Controls.Add(this.txtBottom);
            this.grpCellPadding.Controls.Add(this.txtRight);
            this.grpCellPadding.Controls.Add(this.lblTop);
            this.grpCellPadding.Controls.Add(this.lblLeft);
            this.grpCellPadding.Controls.Add(this.lblBottom);
            this.grpCellPadding.Controls.Add(this.lblRight);
            this.grpCellPadding.Location = new System.Drawing.Point(235, 20);
            this.grpCellPadding.Size = new System.Drawing.Size(245, 115);
            this.grpCellPadding.Text = "셀 채우기 갭";

            this.lblLeft.AutoSize = true; this.lblLeft.Location = new System.Drawing.Point(12, 30); this.lblLeft.Text = "left";
            this.txtLeft.Location = new System.Drawing.Point(50, 27); this.txtLeft.Size = new System.Drawing.Size(60, 23); this.txtLeft.Text = "1.5";

            this.lblTop.AutoSize = true; this.lblTop.Location = new System.Drawing.Point(125, 30); this.lblTop.Text = "top";
            this.txtTop.Location = new System.Drawing.Point(165, 27); this.txtTop.Size = new System.Drawing.Size(60, 23); this.txtTop.Text = "1.5";

            this.lblRight.AutoSize = true; this.lblRight.Location = new System.Drawing.Point(12, 70); this.lblRight.Text = "right";
            this.txtRight.Location = new System.Drawing.Point(50, 67); this.txtRight.Size = new System.Drawing.Size(60, 23); this.txtRight.Text = "0";

            this.lblBottom.AutoSize = true; this.lblBottom.Location = new System.Drawing.Point(125, 70); this.lblBottom.Text = "bottom";
            this.txtBottom.Location = new System.Drawing.Point(165, 67); this.txtBottom.Size = new System.Drawing.Size(60, 23); this.txtBottom.Text = "0.5";

            // 
            // grpConfigMgmt
            // 
            this.grpConfigMgmt.Controls.Add(this.chkCorrectionActive);
            this.grpConfigMgmt.Controls.Add(this.btnUpdateList);
            this.grpConfigMgmt.Controls.Add(this.chkBackgroundYellow);
            this.grpConfigMgmt.Controls.Add(this.btnUpdateInvalidPage);
            this.grpConfigMgmt.Controls.Add(this.btnExecute);
            this.grpConfigMgmt.Location = new System.Drawing.Point(512, 178);
            this.grpConfigMgmt.Size = new System.Drawing.Size(484, 150);
            this.grpConfigMgmt.Text = "설정 관리";

            this.chkCorrectionActive.AutoSize = true;
            this.chkCorrectionActive.Checked = true;
            this.chkCorrectionActive.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkCorrectionActive.Location = new System.Drawing.Point(15, 25);
            this.chkCorrectionActive.Text = "보정 활성화";

            this.btnUpdateList.Location = new System.Drawing.Point(15, 55);
            this.btnUpdateList.Size = new System.Drawing.Size(210, 38);
            this.btnUpdateList.Text = "리스트 업데이트";
            this.btnUpdateList.UseVisualStyleBackColor = true;

            this.btnUpdateInvalidPage.Location = new System.Drawing.Point(15, 100);
            this.btnUpdateInvalidPage.Size = new System.Drawing.Size(210, 38);
            this.btnUpdateInvalidPage.Text = "부적합 페이지 업데이트";
            this.btnUpdateInvalidPage.UseVisualStyleBackColor = true;

            this.chkBackgroundYellow.AutoSize = true;
            this.chkBackgroundYellow.Location = new System.Drawing.Point(235, 25);
            this.chkBackgroundYellow.Text = "교정 (unchecked 시 배경 노랑)";

            this.btnExecute.Location = new System.Drawing.Point(235, 55);
            this.btnExecute.Size = new System.Drawing.Size(234, 83);
            this.btnExecute.Text = "실행";
            this.btnExecute.UseVisualStyleBackColor = true;

            // 
            // grpDataMerge
            // 
            this.grpDataMerge.Controls.Add(this.chkAutoTemp);
            this.grpDataMerge.Controls.Add(this.chkRow1); this.grpDataMerge.Controls.Add(this.txtRow1Label); this.grpDataMerge.Controls.Add(this.txtRow1Folder); this.grpDataMerge.Controls.Add(this.txtRow1Sheet);
            this.grpDataMerge.Controls.Add(this.chkRow2); this.grpDataMerge.Controls.Add(this.txtRow2Label); this.grpDataMerge.Controls.Add(this.txtRow2Folder); this.grpDataMerge.Controls.Add(this.txtRow2Sheet);
            this.grpDataMerge.Controls.Add(this.chkRow3); this.grpDataMerge.Controls.Add(this.txtRow3Label); this.grpDataMerge.Controls.Add(this.txtRow3Folder); this.grpDataMerge.Controls.Add(this.txtRow3Sheet);
            this.grpDataMerge.Location = new System.Drawing.Point(12, 334);
            this.grpDataMerge.Size = new System.Drawing.Size(610, 160);
            this.grpDataMerge.Text = "데이터 병합 및 삽입";

            this.chkAutoTemp.AutoSize = true;
            this.chkAutoTemp.Location = new System.Drawing.Point(450, 20);
            this.chkAutoTemp.Text = "온도 인식값 자동 입력";

            // Row 1
            this.chkRow1.Location = new System.Drawing.Point(15, 45); this.chkRow1.Size = new System.Drawing.Size(80, 20); this.chkRow1.Checked = true; this.chkRow1.Text = "열화상";
            this.txtRow1Label.Location = new System.Drawing.Point(100, 43); this.txtRow1Label.Size = new System.Drawing.Size(150, 23); this.txtRow1Label.Text = "01_Folder";
            this.txtRow1Sheet.Location = new System.Drawing.Point(260, 43); this.txtRow1Sheet.Size = new System.Drawing.Size(120, 23); this.txtRow1Sheet.Text = "7호";

            // Row 2
            this.chkRow2.Location = new System.Drawing.Point(15, 78); this.chkRow2.Size = new System.Drawing.Size(80, 20); this.chkRow2.Checked = true; this.chkRow2.Text = "Sheet2";
            this.txtRow2Label.Location = new System.Drawing.Point(100, 76); this.txtRow2Label.Size = new System.Drawing.Size(150, 23); this.txtRow2Label.Text = "02_Folder";
            this.txtRow2Sheet.Location = new System.Drawing.Point(260, 76); this.txtRow2Sheet.Size = new System.Drawing.Size(120, 23); this.txtRow2Sheet.Text = "Sheet2";

            // Row 3
            this.chkRow3.Location = new System.Drawing.Point(15, 111); this.chkRow3.Size = new System.Drawing.Size(80, 20); this.chkRow3.Text = "Sheet3";
            this.txtRow3Label.Location = new System.Drawing.Point(100, 109); this.txtRow3Label.Size = new System.Drawing.Size(150, 23); this.txtRow3Label.Text = "03_Folder";
            this.txtRow3Sheet.Location = new System.Drawing.Point(260, 109); this.txtRow3Sheet.Size = new System.Drawing.Size(120, 23); this.txtRow3Sheet.Text = "Sheet3";

            // 
            // grpExecutionLog (우측 하단 스피너 영역)
            // 
            this.grpExecutionLog.Controls.Add(this.pnlSpinner);
            this.grpExecutionLog.Location = new System.Drawing.Point(630, 334);
            this.grpExecutionLog.Size = new System.Drawing.Size(366, 160);
            this.grpExecutionLog.Text = "실행 로그";

            this.pnlSpinner.Location = new System.Drawing.Point(150, 45);
            this.pnlSpinner.Size = new System.Drawing.Size(64, 64);

            // 
            // grpLogViewer
            // 
            this.grpLogViewer.Controls.Add(this.pnlLogNav);
            this.grpLogViewer.Controls.Add(this.txtLogContent);
            this.grpLogViewer.Location = new System.Drawing.Point(12, 500);
            this.grpLogViewer.Size = new System.Drawing.Size(984, 218);
            this.grpLogViewer.Text = "로그 뷰어";

            this.pnlLogNav.Controls.Add(this.btnLogAll);
            this.pnlLogNav.Controls.Add(this.btnLogInfo);
            this.pnlLogNav.Controls.Add(this.btnLogError);
            this.pnlLogNav.Location = new System.Drawing.Point(15, 22);
            this.pnlLogNav.Size = new System.Drawing.Size(250, 28);

            this.btnLogAll.Location = new System.Drawing.Point(0, 0); this.btnLogAll.Size = new System.Drawing.Size(75, 25); this.btnLogAll.Text = "전체"; this.btnLogAll.UseVisualStyleBackColor = true;
            this.btnLogInfo.Location = new System.Drawing.Point(80, 0); this.btnLogInfo.Size = new System.Drawing.Size(75, 25); this.btnLogInfo.Text = "정보"; this.btnLogInfo.UseVisualStyleBackColor = true;
            this.btnLogError.Location = new System.Drawing.Point(160, 0); this.btnLogError.Size = new System.Drawing.Size(75, 25); this.btnLogError.Text = "오류"; this.btnLogError.UseVisualStyleBackColor = true;

            this.txtLogContent.Location = new System.Drawing.Point(15, 55);
            this.txtLogContent.Multiline = true;
            this.txtLogContent.Name = "txtLogContent";
            this.txtLogContent.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.txtLogContent.Size = new System.Drawing.Size(954, 150);

            // 
            // Form2
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(248, 249, 250);
            this.ClientSize = new System.Drawing.Size(1008, 729);
            this.Controls.Add(this.grpLogViewer);
            this.Controls.Add(this.grpExecutionLog);
            this.Controls.Add(this.grpDataMerge);
            this.Controls.Add(this.grpConfigMgmt);
            this.Controls.Add(this.grpImageSettings);
            this.Controls.Add(this.grpFileTools);
            this.Controls.Add(this.pnlTopNav);
            this.Font = new System.Drawing.Font("맑은 고딕", 9F);
            this.Name = "Form2";
            this.Text = "소보르소프트 엑셀보고서 도우미";

            this.pnlTopNav.ResumeLayout(false);
            this.grpFileTools.ResumeLayout(false);
            this.grpFileTools.PerformLayout();
            this.grpImageSettings.ResumeLayout(false);
            this.grpImageSettings.PerformLayout();
            this.grpCellPadding.ResumeLayout(false);
            this.grpCellPadding.PerformLayout();
            this.grpConfigMgmt.ResumeLayout(false);
            this.grpConfigMgmt.PerformLayout();
            this.grpDataMerge.ResumeLayout(false);
            this.grpDataMerge.PerformLayout();
            this.grpExecutionLog.ResumeLayout(false);
            this.grpLogViewer.ResumeLayout(false);
            this.pnlLogNav.ResumeLayout(false);
            this.ResumeLayout(false);
        }

        #region 컨트롤 변수 선언
        private System.Windows.Forms.Panel pnlTopNav;
        private System.Windows.Forms.Button btnNavFileSettings;
        private System.Windows.Forms.Button btnNavReportWork;
        private System.Windows.Forms.Button btnNavDataMerge;
        private System.Windows.Forms.GroupBox grpFileTools;
        private System.Windows.Forms.Label lblReportFileName;
        private System.Windows.Forms.TextBox txtReportFileName;
        private System.Windows.Forms.Button btnSelectFile;
        private System.Windows.Forms.Button btnCoverAlign;
        private System.Windows.Forms.Button btnPageNumUpdate;
        private System.Windows.Forms.Button btnFooterChange;
        private System.Windows.Forms.Button btnDateUpdate;
        private System.Windows.Forms.Button btnPdfExport;
        private System.Windows.Forms.Button btnAutoDivider;
        private System.Windows.Forms.Button btnAddSignature;
        private System.Windows.Forms.Button btnChangeInspector;
        private System.Windows.Forms.GroupBox grpImageSettings;
        private System.Windows.Forms.ComboBox cmbSheetName;
        private System.Windows.Forms.Label lblSheetName;
        private System.Windows.Forms.Button btnRemoveSheetImg;
        private System.Windows.Forms.Button btnOptimizeImg;
        private System.Windows.Forms.Button btnFillSheetImg;
        private System.Windows.Forms.GroupBox grpCellPadding;
        private System.Windows.Forms.TextBox txtLeft;
        private System.Windows.Forms.TextBox txtTop;
        private System.Windows.Forms.TextBox txtRight;
        private System.Windows.Forms.TextBox txtBottom;
        private System.Windows.Forms.Label lblLeft;
        private System.Windows.Forms.Label lblTop;
        private System.Windows.Forms.Label lblRight;
        private System.Windows.Forms.Label lblBottom;
        private System.Windows.Forms.GroupBox grpConfigMgmt;
        private System.Windows.Forms.CheckBox chkCorrectionActive;
        private System.Windows.Forms.Button btnUpdateList;
        private System.Windows.Forms.Button btnUpdateInvalidPage;
        private System.Windows.Forms.CheckBox chkBackgroundYellow;
        private System.Windows.Forms.Button btnExecute;
        private System.Windows.Forms.GroupBox grpDataMerge;
        private System.Windows.Forms.CheckBox chkAutoTemp;
        private System.Windows.Forms.CheckBox chkRow1;
        private System.Windows.Forms.TextBox txtRow1Label;
        private System.Windows.Forms.TextBox txtRow1Folder;
        private System.Windows.Forms.TextBox txtRow1Sheet;
        private System.Windows.Forms.CheckBox chkRow2;
        private System.Windows.Forms.TextBox txtRow2Label;
        private System.Windows.Forms.TextBox txtRow2Folder;
        private System.Windows.Forms.TextBox txtRow2Sheet;
        private System.Windows.Forms.CheckBox chkRow3;
        private System.Windows.Forms.TextBox txtRow3Label;
        private System.Windows.Forms.TextBox txtRow3Folder;
        private System.Windows.Forms.TextBox txtRow3Sheet;
        private System.Windows.Forms.GroupBox grpExecutionLog;
        private System.Windows.Forms.Panel pnlSpinner;
        private System.Windows.Forms.GroupBox grpLogViewer;
        private System.Windows.Forms.Panel pnlLogNav;
        private System.Windows.Forms.Button btnLogAll;
        private System.Windows.Forms.Button btnLogInfo;
        private System.Windows.Forms.Button btnLogError;
        private System.Windows.Forms.TextBox txtLogContent;
        #endregion
        #endregion
    }
}