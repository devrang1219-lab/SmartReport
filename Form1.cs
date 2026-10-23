using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using WindowsFormsApp1.Comm;

namespace WindowsFormsApp1
{
    public partial class Form1 : Form, IReport
    {
        private SoborLog soborLog = null;
        private Report report = null;


        public Form1()
        {
            InitializeComponent();


            try
            {
                tbQuantityFile.AllowDrop = true;
                tbQuantityFile.DragEnter += TbQuantityFile_DragEnter;
                tbQuantityFile.DragDrop += TbQuantityFile_DragDrop;
            }
            catch { /* 디자이너에서 컨트롤이 아직 없을 수 있음 */ }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
        }


        private void TbQuantityFile_DragEnter(object sender, DragEventArgs e)
        {
            try
            {
                e.Effect = e.Data.GetDataPresent(DataFormats.FileDrop) ? DragDropEffects.Copy : DragDropEffects.None;
            }
            catch
            {
                e.Effect = DragDropEffects.None;
            }
        }

        private void TbQuantityFile_DragDrop(object sender, DragEventArgs e)
        {
            var files = e.Data.GetData(DataFormats.FileDrop) as string[];
            if (files == null || files.Length == 0) return;

            var file = files[0];
            try
            {
                var ext = Path.GetExtension(file).ToLowerInvariant();
                var textExts = new[] { ".txt", ".csv", ".log", ".cs", ".xml", ".json", ".htm", ".html" };
                if (textExts.Contains(ext))
                {
                    // 큰 파일일 수 있으므로 비동기 읽기
                    Task.Run(() =>
                    {
                        string content = File.ReadAllText(file);
                        try { Invoke(new Action(() => tbQuantityFile.Text = content)); } catch { }
                    });
                }
                else
                {
                    // 경로만 표시
                    tbQuantityFile.Text = file;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"파일 처리 중 오류: {ex.Message}", "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnQuantityFile_Click(object sender, EventArgs e)
        {
            using (var dlg = new OpenFileDialog())
            {
                dlg.Title = "엑셀 파일을 선택하세요";
                dlg.Filter = "Excel 파일 (*.xlsx;*.xls;*.xlsm)|*.xlsx;*.xls;*.xlsm|모든 파일 (*.*)|*.*";
                dlg.CheckFileExists = true;
                dlg.CheckPathExists = true;
                dlg.Multiselect = false;

                // 기본 위치 설정 (폴더가 존재하면 InitialDirectory로 설정)
                var defaultPath = @"C:\";
                if (Directory.Exists(defaultPath))
                {
                    dlg.InitialDirectory = defaultPath;
                }

                var result = dlg.ShowDialog();
                if (result == DialogResult.OK)
                {
                    tbQuantityFile.Text = dlg.FileName;
                }
            }
        }

        private void btnExportForPdf_Click(object sender, EventArgs e)
        {
            btnExportForPdf.Enabled = false;
            Cursor = Cursors.WaitCursor;
            string outputFile = string.Empty;

            try
            {
                string filePath = tbQuantityFile.Text?.Trim();

                var result = report?.ExportPdfAsync(filePath, false);

                if (result == null)
                {
                    AddLog("Error", $"파일 형식이 올바르지 않음");
                    return;
                }

                if (!result.Success)
                {
                    MessageBox.Show(result.ErrorMessage);
                    AddLog("Error", $"통합 PDF 내보내기 실패: {result.ErrorMessage}");
                }
                else
                {
                    outputFile = result.OutputFile;
                    Process.Start(new ProcessStartInfo
                    {
                        FileName = result.OutputFile,
                        UseShellExecute = true
                    });
                }
            }
            finally
            {
                AddLog("SUCCESS", $"통합 PDF 내보내기 성공: {outputFile}");
                Cursor = Cursors.Default;
                btnExportForPdf.Enabled = true;
            }
        }

        #region [로그 넣기]
        private void WriteLog(string log)
        {
            if (soborLog == null)
            {
                soborLog = new SoborLog();
            }

            if (log == "")
            {
                return;
            }
            try
            {
                soborLog.Add(log);
            }
            catch
            {
                //throw ex;
            }
        }

        public void AddLog(string level, string message)
        {
            Color color = Color.Black;

            switch ((level ?? "").ToUpperInvariant())
            {
                case "ERROR":
                    color = Color.Red;
                    break;

                case "WARNING":
                    color = Color.DarkOrange;
                    break;

                case "INFO":
                    color = Color.Black;
                    break;

                case "SUCCESS":
                    color = Color.Green;
                    break;
            }
            if (richTextBox1 == null) return;

            Action doAppend = () =>
            {
                try
                {
                    // 날짜+시간 표시 추가
                    string ts = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                    richTextBox1.SelectionStart = richTextBox1.TextLength;
                    richTextBox1.SelectionColor = Color.Gray;
                    richTextBox1.AppendText($"[{ts}] ");

                    // 레벨 색상으로 출력
                    richTextBox1.SelectionStart = richTextBox1.TextLength;
                    richTextBox1.SelectionColor = color;
                    richTextBox1.AppendText($"[{level}] ");

                    // 메시지
                    richTextBox1.SelectionColor = Color.Black;
                    richTextBox1.AppendText(message + Environment.NewLine);

                    // 탭 자동 선택 (All 로그 탭이 있으면 표시)
                    try
                    {
                        if (tapLog != null && all != null)
                        {
                            tapLog.SelectedTab = all;
                        }


                        int start = richTextBox1.GetFirstCharIndexOfCurrentLine();

                        richTextBox1.SelectionStart = start;
                        richTextBox1.SelectionLength = richTextBox1.TextLength - start;

                        richTextBox1.ScrollToCaret();
                    }
                    catch { }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"AddLog failed: {ex.Message}");
                }
            };

            if (richTextBox1.InvokeRequired)
            {
                richTextBox1.BeginInvoke(doAppend);
            }
            else
            {
                doAppend();
            }

            WriteLog(log: $"[{level}] {message}");
        }
        #endregion
        private void tbQuantityFile_TextChanged(object sender, EventArgs e)
        {
            string filePath = tbQuantityFile.Text?.Trim();
            //string xlsSamplePath = txtBxSampleReport.Text?.Trim();

            if (string.IsNullOrEmpty(filePath) || !File.Exists(filePath))
            {
                return;
            }

            Cursor = Cursors.WaitCursor;
            //예 전기점검보고서_26년예시_260711.xlsx

            // Report.ParseReport 호출하여 Report 인스턴스 생성
            try
            {
                this.report = Report.ParseReport(filePath, "", soborLog, this);

            }
            catch (Exception ex)
            {
                AddLog("Error", $"보고서 파싱 중 오류: {ex.Message}");
            }
            finally
            {
                AddLog("SUCCESS", $"보고서 읽어오기 완료: {filePath}");
                Cursor = Cursors.Default;
            }
        }

        private void btnChangeInspector_Click(object sender, EventArgs e)
        {

        }
    }
}
