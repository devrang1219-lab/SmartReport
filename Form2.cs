using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WindowsFormsApp1
{
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
            ApplyModernStyling();
        }

        private void ApplyModernStyling()
        {
            // 1. 폼 전체 더블 버퍼링 적용 (화면 깜빡임 방지)
            this.SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.UserPaint | ControlStyles.OptimizedDoubleBuffer, true);
            this.UpdateStyles();

            // 2. 메인 실행 버튼 강조 (포인트 컬러 적용)
            if (btnExecute != null)
            {
                btnExecute.FlatStyle = FlatStyle.Flat;
                btnExecute.FlatAppearance.BorderSize = 0;
                btnExecute.BackColor = Color.FromArgb(28, 150, 230); // 모던 블루
                btnExecute.ForeColor = Color.White;
            }

            // 3. 상단 주요 도구 버튼 스타일 일괄 정돈
            if (grpFileTools != null)
            {
                foreach (Control ctrl in grpFileTools.Controls)
                {
                    if (ctrl is System.Windows.Forms.Button btn && btn != btnSelectFile)
                    {
                        btn.FlatStyle = FlatStyle.Flat;
                        btn.FlatAppearance.BorderColor = Color.FromArgb(218, 222, 225);
                        btn.BackColor = Color.White;
                    }
                }
            }
        }

        private void Form2_Load(object sender, EventArgs e)
        {
            // 초기화 로직 및 기본값 설정
            if (txtReportFileName != null) txtReportFileName.ReadOnly = true;
            if (cmbSheetName != null)
            {
                cmbSheetName.Items.AddRange(new object[] { "Sheet2", "Sheet3", "Sheet4", "Sheet5", "7호" });
                if (cmbSheetName.Items.Count > 0) cmbSheetName.SelectedIndex = 1;
            }
        }

        // 상단 네비게이션 탭 전환 이벤트 예시
        private void btnNavFileSettings_Click(object sender, EventArgs e)
        {
            // TODO: 파일/설정 화면 패널 활성화
        }

        private void btnNavReportWork_Click(object sender, EventArgs e)
        {
            // TODO: 보고서 작업 화면 패널 활성화
        }

        private void btnNavDataMerge_Click(object sender, EventArgs e)
        {
            // TODO: 데이터 병합 화면 패널 활성화
        }

        private void btnExecute_Click(object sender, EventArgs e)
        {
            // 실행 버튼 클릭 이벤트
            if (txtLogContent != null)
            {
                txtLogContent.AppendText($"[{DateTime.Now:HH:mm:ss}] 작업이 실행되었습니다.\r\n");
            }
        }
    }
}