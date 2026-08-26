using Microsoft.Office.Interop.Excel;
using OpenCvSharp;
using SynologyIntegration;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;
using WindowsFormsApp1;
using WindowsFormsApp1.Comm;
using WindowsFormsApp1.SortImage;
using Action = System.Action;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace SmartReport
{
    public partial class FormMain : Form
    {
        private SoborLog soborLog = null;
        private Report report = null;
        private SynologyFileUploader _uploader = null;
        public FormMain()
        {
            InitializeComponent();
            //TestCopyGapjiSheet();
            // 파일 드래그 앤 드롭을 허용하고 이벤트를 연결
            try
            {
                tbQuantityFile.AllowDrop = true;
                tbQuantityFile.DragEnter += TbQuantityFile_DragEnter;
                tbQuantityFile.DragDrop += TbQuantityFile_DragDrop;
            }
            catch { /* 디자이너에서 컨트롤이 아직 없을 수 있음 */ }

        }

        private void btnExportForPdf_MouseUp(object sender, MouseEventArgs e)
        {
            // 마우스 왼쪽 버튼에서만 처리
            if (e.Button != MouseButtons.Left)
                return;

            // 기존 Click 핸들러 로직 재사용
            try
            {
                btnExportForPdf_Click(sender, EventArgs.Empty);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"PDF 내보내기 처리 중 오류: {ex.Message}", "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void DumpWorkbookSheetInfo(Excel.Workbook wb, string label)
        {
            try
            {
                if (wb == null)
                {
                    System.Diagnostics.Debug.WriteLine($"{label}: workbook is null");
                    return;
                }

                int wc = 0;
                try { wc = wb.Worksheets.Count; } catch { }
                System.Diagnostics.Debug.WriteLine($"{label}: worksheet count = {wc}");

                for (int i = 1; i <= wc; i++)
                {
                    Excel.Worksheet sh = null;
                    try
                    {
                        sh = (Excel.Worksheet)wb.Worksheets[i];
                        string name = "(unknown)";
                        string visible = "?";
                        string used = "(n/a)";
                        string shapes = "(n/a)";
                        string charts = "(n/a)";
                        string protect = "(n/a)";

                        try { name = sh.Name; } catch { }
                        try { visible = sh.Visible.ToString(); } catch { }
                        try
                        {
                            var ur = sh.UsedRange;
                            try
                            {
                                int r = 0, c = 0;
                                try { r = ur.Rows.Count; } catch { }
                                try { c = ur.Columns.Count; } catch { }
                                used = $"UsedRange {r}x{c}";
                            }
                            finally { if (ur != null) try { Marshal.ReleaseComObject(ur); } catch { } }
                        }
                        catch { }

                        try { shapes = sh.Shapes.Count.ToString(); } catch { }
                        try { charts = sh.ChartObjects().Count.ToString(); } catch { }
                        try { protect = sh.ProtectContents ? "protected" : "unprotected"; } catch { }

                        System.Diagnostics.Debug.WriteLine($"{label}: idx={i}, name='{name}', visible={visible}, {used}, shapes={shapes}, charts={charts}, {protect}");
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"{label}: failed to read sheet idx={i}: {ex.Message}");
                    }
                    finally
                    {
                        if (sh != null) try { Marshal.ReleaseComObject(sh); } catch { }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"DumpWorkbookSheetInfo('{label}') failed: {ex.Message}");
            }
        }

        public static void TestCopyGapjiSheet()
        {
            string sourceFile =
                @"C:\_D\work\한경이엔지\0_org\한경이엔지2본부_26년연차보고서(샘플).xlsx";

            string targetFile =
                @"C:\_D\work\한경이엔지\0_org\갑지.xlsx";

            Excel.Application app = null;
            Excel.Workbook wb = null;
            Excel.Worksheet ws = null;

            try
            {
                app = new Excel.Application();
                app.Visible = false;
                app.DisplayAlerts = false;

                wb = app.Workbooks.Open(sourceFile);

                ws = (Excel.Worksheet)wb.Worksheets["갑지"];

                // 갑지 시트만 새 통합문서로 복사
                ws.Copy();

                // 복사 후 활성 통합문서가 새 파일이 됨
                Excel.Workbook tmpWb = app.ActiveWorkbook;

                tmpWb.SaveAs(
                    targetFile,
                    Excel.XlFileFormat.xlOpenXMLWorkbook);

                tmpWb.Close(false);
                Marshal.ReleaseComObject(tmpWb);
            }
            finally
            {
                if (ws != null) Marshal.ReleaseComObject(ws);

                if (wb != null)
                {
                    wb.Close(false);
                    Marshal.ReleaseComObject(wb);
                }

                if (app != null)
                {
                    app.Quit();
                    Marshal.ReleaseComObject(app);
                }

                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        private void btFolder_Click(object sender, EventArgs e)
        {
            // 폴더 선택 대화상자로 변경
            using (var dlg = new FolderBrowserDialog())
            {
                dlg.Description = "보고서가 있는 폴더를 선택하세요";

                // 기본 위치 설정
                var defaultPath = @"C:\_D\work\한경이엔지\2_report";
                if (Directory.Exists(defaultPath))
                {
                    dlg.SelectedPath = defaultPath;
                }

                var result = dlg.ShowDialog();
                if (result == DialogResult.OK || result == DialogResult.Yes)
                {
                    tbFolder.Text = dlg.SelectedPath;
                }
            }

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

        private void FindAnuualReportFile(string folderPath, out string annualReportFile)
        {
            annualReportFile = null;
            try
            {
                var exts = new[] { ".xls", ".xlsx", ".xlsm", ".xlsb" };
                var files = Directory.GetFiles(folderPath)
                    .Where(f => exts.Contains(Path.GetExtension(f), StringComparer.OrdinalIgnoreCase))
                    .ToList();
                foreach (var f in files)
                {
                    if (Path.GetFileName(f).IndexOf("연차", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        annualReportFile = f;
                        return;
                    }
                }
            }
            catch (Exception ex)
            {
                AddLog("ERROR", $"연차 보고서 파일 검색 중 오류: {ex.Message}");
            }
        }

        private string makeNewFileName(string folderName)
        {
            string candidate = "";
            var folderParts = folderName.Split('_');

            if (folderParts.Length < 1)
                throw new InvalidOperationException("폴더명이 예상 형식이 아닙니다. '_'로 구분된 첫번째 부분에 날짜가 있어야 합니다.");
            
            var folderDate = folderParts[0];
            string year = "";
            string month = "";
            string day = "";
            string formattedDate = folderDate;

            if (System.Text.RegularExpressions.Regex.IsMatch(folderDate, "^\\d{6}$"))
            {
                var yy = folderDate.Substring(0, 2);
                year = yy;
                month = folderDate.Substring(2, 2);
                day = folderDate.Substring(4, 2);
                formattedDate = year + month + day;
            }

            int iMonth = 1;
            Int32.TryParse(month, out iMonth);

            string quater = (iMonth < 4) ? "1" : (iMonth < 7) ? "2" : (iMonth < 10) ? "3" : "4";

            candidate = $"{folderName}\\{folderParts[1]}_{year}년{quater}분기{(folderParts[2] == "연차" ? "연차" : "")}보고서_{folderParts[0]}.xlsx";

            return candidate;
        }

        private void SetRecentQuaterSheet(Excel.Workbook wbNew, string dir)
        {
            Excel.Workbook wbOld = null;

            try
            {
                var exts = new[] { ".xls", ".xlsx", ".xlsm", ".xlsb" };

                string latestFile = Directory.GetFiles(dir)
                    .Where(f => exts.Contains(Path.GetExtension(f), StringComparer.OrdinalIgnoreCase))
                    .OrderByDescending(f => Path.GetFileName(f), StringComparer.OrdinalIgnoreCase)
                    .FirstOrDefault();

                if (string.IsNullOrEmpty(latestFile))
                    return;

                Console.WriteLine(latestFile);

                // 파일명에 "연차"가 있으면 기존 유지
                if (Path.GetFileName(latestFile).Contains("연차"))
                    return;

                Excel.Application app = wbNew.Application;

                // 원본 파일 열기
                wbOld = app.Workbooks.Open(latestFile, ReadOnly: true);

                // wbNew의 "분기" 시트 삭제
                for (int i = wbNew.Worksheets.Count; i >= 1; i--)
                {
                    Excel.Worksheet ws = wbNew.Worksheets[i];

                    if (ws.Name.Contains("분기"))
                    {
                        ws.Delete();
                    }

                    Marshal.ReleaseComObject(ws);
                }

                // latestFile의 "분기" 시트 복사
                foreach (Excel.Worksheet wsOld in wbOld.Worksheets)
                {
                    if (wsOld.Name.Contains("분기"))
                    {
                        wsOld.Copy(
                            After: wbNew.Worksheets[wbNew.Worksheets.Count]
                        );
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {
                if (wbOld != null)
                {
                    wbOld.Close(false);
                    Marshal.ReleaseComObject(wbOld);
                }
            }
        }

        private string createNewReportFile(string originalFilePath)
        {
            var dir = Path.GetDirectoryName(originalFilePath) ?? throw new InvalidOperationException("디렉터리 정보를 가져올 수 없습니다.");
            var originalName = Path.GetFileNameWithoutExtension(originalFilePath);
            var ext = Path.GetExtension(originalFilePath);


            var folderName = originalName;// new DirectoryInfo(dir).Name;
            var folderParts = originalName.Split('_');
            if (folderParts.Length < 1)
                throw new InvalidOperationException("폴더명이 예상 형식이 아닙니다. '_'로 구분된 첫번째 부분에 날짜가 있어야 합니다.");

            var xlsReportFile = "";
            var reportPath = Path.Combine(originalFilePath, "04 보고서");
            FindAnuualReportFile(reportPath, out xlsReportFile);

            if (string.IsNullOrEmpty(xlsReportFile) || !File.Exists(xlsReportFile))
                throw new InvalidOperationException("연차 보고서 파일을 찾을 수 없습니다. 폴더 내에 '연차'가 포함된 엑셀 파일이 있어야 합니다.");

            string newfileName = dir + "\\" + makeNewFileName(originalName);
            if (string.IsNullOrEmpty(newfileName))
                throw new InvalidOperationException("새 파일 이름 생성을 하지 못했습니다.");

            File.Copy(xlsReportFile, newfileName);            

            // 템플릿 파일 찾기
            string templatePath = txtBxSampleReport.Text.Trim();
            if (string.IsNullOrEmpty(templatePath))
                throw new InvalidOperationException("탬플릿 파일이 유효하지 않습니다.");

            var conditionalSheets = new[] { "제출문", "의견", "연계획"};
            var alwaysCopySheets = new[] { "전기설비", "장비", "검교정", "부록", "마" };

            Excel.Application xlApp = null;
            Excel.Workbook wbTemplate = null;
            Excel.Workbook wbNew = null;

            try
            {
                xlApp = new Excel.Application
                {
                    Visible = false,
                    DisplayAlerts = false
                };

                wbTemplate = xlApp.Workbooks.Open(templatePath, ReadOnly: false);

                // 원본 파일을 열어 어떤 시트가 있는지 확인
                Microsoft.Office.Interop.Excel.Workbook wbOrig = null;
                var origSheetNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                try
                {
                    wbOrig = xlApp.Workbooks.Open(newfileName, ReadOnly: true);
                    foreach (Microsoft.Office.Interop.Excel.Worksheet sh in wbOrig.Worksheets)
                    {
                        try { origSheetNames.Add(sh.Name); } catch { }
                        finally { if (sh != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(sh); }
                    }
                }
                finally
                {
                    //if (wbOrig != null)
                    //{
                    //    wbOrig.Close(false);
                    //    System.Runtime.InteropServices.Marshal.ReleaseComObject(wbOrig);
                    //    wbOrig = null;
                    //}
                }

                // 새 통합문서 생성: 원본 복사 대신 연차 파일과 템플릿의 조합으로 새 파일을 만든다.
                wbNew = xlApp.Workbooks.Open(newfileName, ReadOnly: true);
                try
                {
                    // 1) 연차 문서를 임시 복사해 외부 링크 제거 및 불필요 시트 삭제 후 사용
                    if (wbNew != null)
                    {
                        // 외부 링크 끊기(있으면)
                        try
                        {
                            object linksObj = null;
                            try { linksObj = wbNew.LinkSources(Excel.XlLinkType.xlLinkTypeExcelLinks); } catch { linksObj = null; }
                            if (linksObj is System.Array linksArr)
                            {
                                foreach (var l in linksArr)
                                {
                                    try { wbNew.BreakLink(l.ToString(), Excel.XlLinkType.xlLinkTypeExcelLinks); } catch (Exception ex) { System.Diagnostics.Debug.WriteLine($"BreakLink failed: {ex.Message}"); }
                                }
                            }
                        }
                        catch { }

                        // 연결 삭제
                        try
                        {
                            dynamic conns = null;
                            try { conns = wbNew.Connections; } catch { conns = null; }
                            if (conns != null)
                            {
                                try
                                {
                                    int cc = conns.Count;
                                    for (int ci = cc; ci >= 1; ci--)
                                    {
                                        try { conns.Item(ci).Delete(); } catch { }
                                    }
                                }
                                catch { }
                            }
                        }
                        catch { }

                        // 불필요 시트 삭제: 연차에서 필요한 패턴(갑지, 제출문, 저압*, 예비*, 분기*)만 남김
                        try
                        {
                            var deletePatterns = new[] { "전기설비", "장비", "검교정", "부록", "마" };
                            var names = new List<string>();
                            try
                            {
                                int yc = wbNew.Worksheets.Count;
                                for (int yi = 1; yi <= yc; yi++)
                                {
                                    Excel.Worksheet tsh = null;
                                    try { tsh = (Excel.Worksheet)wbNew.Worksheets[yi]; names.Add(tsh.Name); }
                                    catch { }
                                    finally { if (tsh != null) try { Marshal.ReleaseComObject(tsh); } catch { } }
                                }
                            }
                            catch { }

                            foreach (var nm in names)
                            {
                                try
                                {
                                    bool delete = deletePatterns.Any(p => nm.IndexOf(p, StringComparison.OrdinalIgnoreCase) >= 0);
                                    if (delete)
                                    {
                                        Excel.Worksheet del = null;
                                        try { del = (Excel.Worksheet)wbNew.Worksheets[nm]; del.Delete(); }
                                        catch (Exception ex) { System.Diagnostics.Debug.WriteLine($"Failed to delete sheet '{nm}': {ex.Message}"); }
                                        finally { if (del != null) try { Marshal.ReleaseComObject(del); } catch { } }
                                    }
                                }
                                catch { }
                            }
                        }
                        catch { }
                    }

                    // 2) 템플릿의 필요한 시트 복사 (전기설비, 장비, 검교정, 부록, 마)
                    var templateNames = new[] { "전기설비", "장비", "검교정", "부록", "마" };
                    Excel.Worksheet last = null;
                    try { last = (Excel.Worksheet)wbNew.Worksheets[wbNew.Worksheets.Count]; } catch { }
                    foreach (var sheetName in templateNames)
                    {
                        Excel.Worksheet src = null;
                        try
                        {
                            try { src = wbTemplate.Worksheets[sheetName] as Excel.Worksheet; } catch { src = null; }
                            if (src != null)
                            {
                                try
                                {
                                    if (last != null)
                                    {
                                        src.Copy(After: last);
                                        // 업데이트 last
                                        try { System.Runtime.InteropServices.Marshal.ReleaseComObject(last); } catch { }
                                        try { last = (Excel.Worksheet)wbNew.Worksheets[wbNew.Worksheets.Count]; } catch { last = null; }
                                    }
                                    else
                                    {
                                        src.Copy(After: wbNew.Worksheets[wbNew.Worksheets.Count]);
                                        try { last = (Excel.Worksheet)wbNew.Worksheets[wbNew.Worksheets.Count]; } catch { last = null; }
                                    }
                                }
                                catch { }
                            }
                            else
                            {
                                // 템플릿에 없으면 빈 시트 추가
                                Excel.Worksheet added = null;
                                try { added = (Excel.Worksheet)wbNew.Worksheets.Add(After: wbNew.Worksheets[wbNew.Worksheets.Count]); added.Name = sheetName; }
                                catch { }
                                finally { if (added != null) { try { if (last != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(last); } catch { } last = added; } }
                            }
                        }
                        catch { }
                        finally { if (src != null) try { Marshal.ReleaseComObject(src); } catch { } }
                    }

                    // 최신 분기 엑셀 파일을 찾아 분기 시트 교체
                    SetRecentQuaterSheet(wbNew, reportPath);

                    // 불필요한 기본 시트 제거 (이름이 Sheet로 시작하는 경우)
                    try
                    {
                        var defaults = wbNew.Worksheets.Cast<Excel.Worksheet>().Where(s => s.Name.StartsWith("Sheet", StringComparison.OrdinalIgnoreCase)).ToList();
                        foreach (var s in defaults)
                        {
                            try { s.Delete(); } catch { }
                            try { Marshal.ReleaseComObject(s); } catch { }
                        }
                    }
                    catch { }

                    // 텍스트 치환
                    //if (!string.IsNullOrEmpty(site))
                    //{
                    //    ReplaceTextInWorkbook(wbNew, "금천구청", site);
                    //}

                    // 저장
                    try { wbNew.Save(); }
                    catch { }
                    finally
                    {
                        try { wbNew.Close(false); } catch { }
                        try { wbTemplate.Close(false); } catch { }
                    }
                }
                catch { }
            }
            finally
            {
                if (wbNew != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(wbNew);
                if (wbTemplate != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(wbTemplate);
                if (xlApp != null)
                {
                    xlApp.Quit();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp);
                }
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }


            return newfileName;
        }

        bool copyFile(string filePath, string newFilePath)
        {
            try
            {
                var tempExt = Path.GetExtension(filePath);
                newFilePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + tempExt);
                try { File.Copy(filePath, newFilePath, true); } 
                catch (Exception ex) { 
                    System.Diagnostics.Debug.WriteLine($"Failed to copy xlsReportFile to temp: {ex.Message}"); 
                    newFilePath = null;
                    return false;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Failed to open workbook '{filePath}': {ex.Message}");
                //throw;
                return false;
            }
            return true;
        }

        private string createNewReportFileUsingOrgFile(string originalFilePath)
        {
            var dir = Path.GetDirectoryName(originalFilePath) ?? throw new InvalidOperationException("디렉터리 정보를 가져올 수 없습니다.");
            var originalName = Path.GetFileNameWithoutExtension(originalFilePath);
            var ext = Path.GetExtension(originalFilePath);


            var folderName = new DirectoryInfo(dir).Name;
            var folderParts = folderName.Split('_');
            if (folderParts.Length < 1)
                throw new InvalidOperationException("폴더명이 예상 형식이 아닙니다. '_'로 구분된 첫번째 부분에 날짜가 있어야 합니다.");

            var xlsReportFile = "";
            var reportPath = Path.Combine(originalFilePath, "04 보고서");
            FindAnuualReportFile(reportPath, out xlsReportFile);

            if (string.IsNullOrEmpty(xlsReportFile) || !File.Exists(xlsReportFile))
                throw new InvalidOperationException("연차 보고서 파일을 찾을 수 없습니다. 폴더 내에 '연차'가 포함된 엑셀 파일이 있어야 합니다.");

            var folderDate = folderParts[0];
            string year;
            string month = "";
            string day = "";
            string formattedDate = folderDate;

            if (System.Text.RegularExpressions.Regex.IsMatch(folderDate, "^\\d{6}$"))
            {
                var yy = folderDate.Substring(0, 2);
                year = yy;
                month = folderDate.Substring(2, 2);
                day = folderDate.Substring(4, 2);
                formattedDate = year + month + day;
            }
            else if (System.Text.RegularExpressions.Regex.IsMatch(folderDate, "^\\d{8}$"))
            {
                year = folderDate.Substring(0, 4);
                month = folderDate.Substring(4, 2);
                day = folderDate.Substring(6, 2);
                formattedDate = year + month + day;
            }
            else
            {
                var m = System.Text.RegularExpressions.Regex.Match(folderDate, "^(\\d{4})");
                if (m.Success)
                    year = m.Groups[1].Value;
                else
                {
                    m = System.Text.RegularExpressions.Regex.Match(folderDate, "^(\\d{2})");
                    year = m.Success ? m.Groups[1].Value : folderDate;
                }
            }

            var fileParts = originalName.Split('_');
            string site = null;

            if (folderParts.Length >= 2 && fileParts.Length >= 1)
            {
                site = folderParts[1];
                foreach (var c in Path.GetInvalidFileNameChars())
                    site = site.Replace(c, '_');
                fileParts[0] = site;
            }

            if (fileParts.Length >= 2)
                fileParts[1] = System.Text.RegularExpressions.Regex.Replace(fileParts[1], "\\d{2,4}년", year + "년");
            if (fileParts.Length >= 3)
                fileParts[2] = folderDate;

            var newBase = string.Join("_", fileParts);
            var newPath = Path.Combine(dir, newBase + ext);

            var idx = 1;
            var candidate = newPath;
            while (File.Exists(candidate))
            {
                candidate = Path.Combine(dir, newBase + $"_{idx}" + ext);
                idx++;
            }

            // 템플릿 파일 찾기
            string templateFileName = "한경이엔지2본부_26년연차보고서(샘플).xlsx";
            string templatePath = FindTemplatePath(templateFileName);
            templatePath = txtBxSampleReport.Text;

            //var conditionalSheets = new[] { "제출문", "의견", "연계획", "검교정" };
            var conditionalSheets = new[] { "제출문", "의견", "연계획", "검교정" };
            var alwaysCopySheets = new[] { "전기설비", "장비", "부록", "마" };

            Excel.Application xlApp = null;
            Excel.Workbook wbTemplate = null;
            Excel.Workbook wbNew = null;

            try
            {
                xlApp = new Excel.Application
                {
                    Visible = false,
                    DisplayAlerts = false
                };

                wbTemplate = xlApp.Workbooks.Open(templatePath, ReadOnly: false);

                // 원본 파일을 열어 어떤 시트가 있는지 확인
                HashSet<string> sheetNames = GetWorkbookSheetNames(wbTemplate);

                // 새 통합문서 생성: 원본 복사 대신 연차 파일과 템플릿의 조합으로 새 파일을 만든다.
                wbNew = xlApp.Workbooks.Add();

                Excel.Workbook wbYear = null;
                try
                {
                    // 연차 파일 찾기
                    string yearFile = null;
                    try
                    {
                        var exts = new[] { ".xls", ".xlsx", ".xlsm", ".xlsb" };
                        yearFile = Directory.GetFiles(dir)
                            .Where(f => exts.Contains(Path.GetExtension(f), StringComparer.OrdinalIgnoreCase))
                            .FirstOrDefault(f => Path.GetFileName(f).IndexOf("연차", StringComparison.OrdinalIgnoreCase) >= 0);
                    }
                    catch { }
                    if (!string.IsNullOrEmpty(yearFile) && File.Exists(yearFile))
                    {
                        wbYear = xlApp.Workbooks.Open(yearFile, ReadOnly: true);
                        try { DumpWorkbookSheetInfo(wbYear, "wbYear (opened)"); } catch { }
                    }


                    // 1) 연차 문서를 임시 복사해 외부 링크 제거 및 불필요 시트 삭제 후 사용
                    //string tempYearPath = null;
                    //copyFile(xlsReportFile, Path.Combine(originalFilePath, ));
                        

                    if (wbYear != null)
                    {
                        // 외부 링크 끊기(있으면)
                        try
                        {
                            object linksObj = null;
                            try { linksObj = wbYear.LinkSources(Excel.XlLinkType.xlLinkTypeExcelLinks); } catch { linksObj = null; }
                            if (linksObj is System.Array linksArr)
                            {
                                foreach (var l in linksArr)
                                {
                                    try { wbYear.BreakLink(l.ToString(), Excel.XlLinkType.xlLinkTypeExcelLinks); } catch (Exception ex) { System.Diagnostics.Debug.WriteLine($"BreakLink failed: {ex.Message}"); }
                                }
                            }
                        }
                        catch { }

                        // 연결 삭제
                        try
                        {
                            dynamic conns = null;
                            try { conns = wbYear.Connections; } catch { conns = null; }
                            if (conns != null)
                            {
                                try
                                {
                                    int cc = conns.Count;
                                    for (int ci = cc; ci >= 1; ci--)
                                    {
                                        try { conns.Item(ci).Delete(); } catch { }
                                    }
                                }
                                catch { }
                            }
                        }
                        catch { }

                        // 불필요 시트 삭제: 연차에서 필요한 패턴(갑지, 제출문, 저압*, 예비*, 분기*)만 남김
                        try
                        {
                            var keepPatterns = new[] { "갑지", "제출문", "저압", "예비", "분기" };
                            var names = new List<string>();
                            try
                            {
                                int yc = wbYear.Worksheets.Count;
                                for (int yi = 1; yi <= yc; yi++)
                                {
                                    Excel.Worksheet tsh = null;
                                    try { tsh = (Excel.Worksheet)wbYear.Worksheets[yi]; names.Add(tsh.Name); }
                                    catch { }
                                    finally { if (tsh != null) try { Marshal.ReleaseComObject(tsh); } catch { } }
                                }
                            }
                            catch { }

                            foreach (var nm in names)
                            {
                                try
                                {
                                    bool keep = keepPatterns.Any(p => nm.IndexOf(p, StringComparison.OrdinalIgnoreCase) >= 0);
                                    if (!keep)
                                    {
                                        Excel.Worksheet del = null;
                                        try { del = (Excel.Worksheet)wbYear.Worksheets[nm]; del.Delete(); }
                                        catch (Exception ex) { System.Diagnostics.Debug.WriteLine($"Failed to delete sheet '{nm}': {ex.Message}"); }
                                        finally { if (del != null) try { Marshal.ReleaseComObject(del); } catch { } }
                                    }
                                }
                                catch { }
                            }
                        }
                        catch { }

                        var exactNames = new[] { "갑지", "제출문", "의견" };
                        Excel.Worksheet firstSheet = null;
                        try { firstSheet = (Excel.Worksheet)wbNew.Worksheets[1]; } catch { }
                        if (firstSheet != null)
                        {
                            for (int i = exactNames.Length - 1; i >= 0; i--)
                            {
                                var name = exactNames[i];
                                try
                                {
                                    foreach (Excel.Worksheet src in wbYear.Worksheets)
                                    {
                                        try
                                        {
                                            if (string.Equals(src.Name, name, StringComparison.OrdinalIgnoreCase))
                                            {
                                                try { src.Copy(Before: firstSheet); }
                                                catch { }
                                            }
                                        }
                                        catch { }
                                        finally { try { Marshal.ReleaseComObject(src); } catch { } }
                                    }
                                }
                                catch { }
                            }

                            try { Marshal.ReleaseComObject(firstSheet); } catch { }
                        }
                    }

                    // 2) 템플릿의 필요한 시트 복사 (전기설비, 장비, 검교정, 부록, 마)
                    var templateNames = new[] { "전기설비", "장비", "검교정", "부록", "마" };
                    Excel.Worksheet last = null;
                    try { last = (Excel.Worksheet)wbNew.Worksheets[wbNew.Worksheets.Count]; } catch { }
                    foreach (var sheetName in templateNames)
                    {
                        Excel.Worksheet src = null;
                        try
                        {
                            try { src = wbTemplate.Worksheets[sheetName] as Excel.Worksheet; } catch { src = null; }
                            if (src != null)
                            {
                                try
                                {
                                    if (last != null)
                                    {
                                        src.Copy(After: last);
                                        // 업데이트 last
                                        try { System.Runtime.InteropServices.Marshal.ReleaseComObject(last); } catch { }
                                        try { last = (Excel.Worksheet)wbNew.Worksheets[wbNew.Worksheets.Count]; } catch { last = null; }
                                    }
                                    else
                                    {
                                        src.Copy(After: wbNew.Worksheets[wbNew.Worksheets.Count]);
                                        try { last = (Excel.Worksheet)wbNew.Worksheets[wbNew.Worksheets.Count]; } catch { last = null; }
                                    }
                                }
                                catch { }
                            }
                            else
                            {
                                // 템플릿에 없으면 빈 시트 추가
                                Excel.Worksheet added = null;
                                try { added = (Excel.Worksheet)wbNew.Worksheets.Add(After: wbNew.Worksheets[wbNew.Worksheets.Count]); added.Name = sheetName; }
                                catch { }
                                finally { if (added != null) { try { if (last != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(last); } catch { } last = added; } }
                            }
                        }
                        catch { }
                        finally { if (src != null) try { Marshal.ReleaseComObject(src); } catch { } }
                    }

                    // 3) 장비 시트 뒤에 연차의 패턴 시트(저압*, 예비*, *분기*) 복사
                    Excel.Worksheet equipmentSheet = ((IEnumerable<Excel.Worksheet>)wbNew.Worksheets.Cast<Excel.Worksheet>())
                        .FirstOrDefault(w => string.Equals(w.Name, "장비", StringComparison.OrdinalIgnoreCase));
                    Excel.Worksheet insertAfter = equipmentSheet ?? (Excel.Worksheet)wbNew.Worksheets[wbNew.Worksheets.Count];

                    var patterns = new[] { "예비", "분기" };
                    foreach (var pattern in patterns)
                    {
                        try
                        {
                            // 기존에 동일 패턴 시트가 있으면 삭제
                            var toDelete = wbNew.Worksheets.Cast<Excel.Worksheet>()
                                .Where(w => w.Name.IndexOf(pattern, StringComparison.OrdinalIgnoreCase) >= 0)
                                .ToList();
                            foreach (var d in toDelete)
                            {
                                try { d.Delete(); } catch { }
                                try { Marshal.ReleaseComObject(d); } catch { }
                            }
                        }
                        catch { }

                        if (wbYear == null) continue;

                        try
                        {
                            var copied = new List<string>();
                            int cnt = wbYear.Worksheets.Count;
                            for (int si = 1; si <= cnt; si++)
                            {
                                Excel.Worksheet src = null;
                                try
                                {
                                    src = (Excel.Worksheet)wbYear.Worksheets[si];
                                    if (src == null) continue;

                                    if (src.Name.IndexOf(pattern, StringComparison.OrdinalIgnoreCase) >= 0)
                                    {
                                        try
                                        {
                                            src.Copy(After: insertAfter);
                                            int newIndex = insertAfter.Index + 1;
                                            Excel.Worksheet newSh = null;
                                            try
                                            {
                                                newSh = (Excel.Worksheet)wbNew.Worksheets[newIndex];
                                                insertAfter = newSh;
                                            }
                                            catch { }
                                            finally { if (newSh != null) Marshal.ReleaseComObject(newSh); }

                                            copied.Add(src.Name);
                                        }
                                        catch (Exception ex)
                                        {
                                            string sheetName = null;
                                            try { sheetName = src?.Name; } catch { sheetName = "(unknown)"; }
                                            System.Diagnostics.Debug.WriteLine($"Copy failed for sheet '{sheetName}': {ex.Message}");

                                            // 만약 RCW 분리 오류라면 wbYear를 다시 열어서 시도
                                            bool reopened = false;
                                            try
                                            {
                                                if (ex is System.Runtime.InteropServices.InvalidComObjectException)
                                                {
                                                    try
                                                    {
                                                        if (wbYear != null)
                                                        {
                                                            try { wbYear.Close(false); } catch { }
                                                            try { Marshal.ReleaseComObject(wbYear); } catch { }
                                                            wbYear = null;
                                                        }
                                                    }
                                                    catch { }

                                                    try
                                                    {
                                                        if (!string.IsNullOrEmpty(yearFile) && File.Exists(yearFile))
                                                        {
                                                            wbYear = xlApp.Workbooks.Open(yearFile, ReadOnly: true);
                                                            reopened = true;
                                                        }
                                                    }
                                                    catch (Exception rex)
                                                    {
                                                        System.Diagnostics.Debug.WriteLine($"Failed to reopen year workbook: {rex.Message}");
                                                    }
                                                }
                                            }
                                            catch { }

                                            // 폴백 로직: 새 시트 추가하고 UsedRange 복사
                                            Excel.Worksheet added = null;
                                            Excel.Range srcUsed = null;
                                            try
                                            {
                                                // 재취득: 동일 이름의 시트를 wbYear에서 찾아 src로 재할당
                                                if (reopened && wbYear != null)
                                                {
                                                    try
                                                    {
                                                        // 시트 이름이 있으면 찾아서 할당
                                                        for (int jj = 1; jj <= wbYear.Worksheets.Count; jj++)
                                                        {
                                                            try
                                                            {
                                                                var tmp = (Excel.Worksheet)wbYear.Worksheets[jj];
                                                                try
                                                                {
                                                                    if (!string.IsNullOrEmpty(sheetName) && tmp.Name == sheetName)
                                                                    {
                                                                        // 교체
                                                                        try { if (src != null) Marshal.ReleaseComObject(src); } catch { }
                                                                        src = tmp;
                                                                        break;
                                                                    }
                                                                }
                                                                catch { try { Marshal.ReleaseComObject(tmp); } catch { } }
                                                            }
                                                            catch { }
                                                        }
                                                    }
                                                    catch { }
                                                }

                                                added = (Excel.Worksheet)wbNew.Worksheets.Add(After: insertAfter);
                                                try { if (!string.IsNullOrEmpty(sheetName)) added.Name = sheetName; } catch { }

                                                try { srcUsed = src?.UsedRange; } catch { srcUsed = null; }
                                                if (srcUsed != null)
                                                {
                                                    try
                                                    {
                                                        int rows = srcUsed.Rows.Count;
                                                        int cols = srcUsed.Columns.Count;
                                                        Excel.Range destStart = added.Range["A1"];
                                                        try { destStart.Resize[rows, cols].Value = srcUsed.Value; } catch { }
                                                        try { destStart.Resize[rows, cols].NumberFormat = srcUsed.NumberFormat; } catch { }
                                                        try
                                                        {
                                                            for (int c = 1; c <= cols; c++)
                                                            {
                                                                try { added.Columns[c].ColumnWidth = src.Columns[c].ColumnWidth; } catch { }
                                                            }
                                                        }
                                                        catch { }
                                                    }
                                                    catch { }
                                                }

                                                try { insertAfter = added; } catch { }
                                                copied.Add((sheetName ?? "(unknown)") + "(fallback)");
                                            }
                                            catch (Exception ex2)
                                            {
                                                System.Diagnostics.Debug.WriteLine($"Fallback copy failed for sheet '{sheetName}': {ex2.Message}");
                                            }
                                            finally
                                            {
                                                if (srcUsed != null) try { Marshal.ReleaseComObject(srcUsed); } catch { }
                                                if (added != null) try { Marshal.ReleaseComObject(added); } catch { }
                                            }
                                        }
                                    }
                                }
                                catch { }
                                finally { if (src != null) try { Marshal.ReleaseComObject(src); } catch { } }
                            }

                            if (copied.Count > 0)
                            {
                                System.Diagnostics.Debug.WriteLine($"Copied pattern '{pattern}': {string.Join(",", copied)}");
                            }
                            else
                            {
                                System.Diagnostics.Debug.WriteLine($"No sheets copied for pattern '{pattern}'");
                            }
                        }
                        catch { }
                    }

                    try { if (equipmentSheet != null) Marshal.ReleaseComObject(equipmentSheet); } catch { }

                    // 불필요한 기본 시트 제거 (이름이 Sheet로 시작하는 경우)
                    try
                    {
                        var defaults = wbNew.Worksheets.Cast<Excel.Worksheet>().Where(s => s.Name.StartsWith("Sheet", StringComparison.OrdinalIgnoreCase)).ToList();
                        foreach (var s in defaults)
                        {
                            try { s.Delete(); } catch { }
                            try { Marshal.ReleaseComObject(s); } catch { }
                        }
                    }
                    catch { }

                    // 텍스트 치환
                    if (!string.IsNullOrEmpty(site))
                    {
                        ReplaceTextInWorkbook(wbNew, "금천구청", site);
                    }

                    // 저장
                    try { wbNew.SaveAs(candidate, Excel.XlFileFormat.xlOpenXMLWorkbook); }
                    catch
                    {
                        // SaveAs 실패하면 일반 Save 시도
                        try { wbNew.Save(); }
                        catch { }
                    }
                    finally
                    {
                        try { wbNew.Close(false); } catch { }
                        try { wbTemplate.Close(false); } catch { }
                    }

                    return candidate;
                }
                finally
                {
                    if (wbYear != null)
                    {
                        try { wbYear.Close(false); } catch { }
                        try { Marshal.ReleaseComObject(wbYear); } catch { }
                    }
                }
            }
            finally
            {
                if (wbNew != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(wbNew);
                if (wbTemplate != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(wbTemplate);
                if (xlApp != null)
                {
                    xlApp.Quit();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp);
                }
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        private void EnsureConditionalAndAlwaysSheets(Excel.Workbook wbTemplate, Excel.Workbook wbNew, HashSet<string> wbNewSheetNames, string[] conditionalSheets, string[] alwaysCopySheets)
        {
            // 조건부 시트: 대상에 없으면 템플릿에서 복사
            foreach (var sheetName in conditionalSheets)
            {
                if (wbNewSheetNames.Contains(sheetName))
                    continue;
                Excel.Worksheet src = null;
                try
                {
                    src = wbTemplate.Worksheets[sheetName] as Excel.Worksheet;
                    if (src != null)
                    {
                        var last = (Excel.Worksheet)wbNew.Worksheets[wbNew.Worksheets.Count];
                        src.Copy(After: last);
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(last);
                    }
                }
                catch { }
                finally { if (src != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(src); }
            }

            // 무조건 복사 시트: 대상에 있으면 삭제 후 템플릿에서 복사(없으면 빈 시트 추가)
            foreach (var sheetName in alwaysCopySheets)
            {
                try
                {
                    var existing = ((IEnumerable<Excel.Worksheet>)wbNew.Worksheets.Cast<Excel.Worksheet>())
                        .FirstOrDefault(w => string.Equals(w.Name, sheetName, StringComparison.OrdinalIgnoreCase));
                    if (existing != null)
                    {
                        existing.Delete();
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(existing);
                    }
                }
                catch { }

                Excel.Worksheet src = null;
                try
                {
                    src = wbTemplate.Worksheets[sheetName] as Excel.Worksheet;
                    if (src != null)
                    {
                        var last = (Excel.Worksheet)wbNew.Worksheets[wbNew.Worksheets.Count];
                        src.Copy(After: last);
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(last);
                    }
                    else
                    {
                        var added = (Excel.Worksheet)wbNew.Worksheets.Add(After: wbNew.Worksheets[wbNew.Worksheets.Count]);
                        try { added.Name = sheetName; } catch { }
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(added);
                    }
                }
                catch
                {
                    try
                    {
                        var added = (Excel.Worksheet)wbNew.Worksheets.Add(After: wbNew.Worksheets[wbNew.Worksheets.Count]);
                        try { added.Name = sheetName; } catch { }
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(added);
                    }
                    catch { }
                }
                finally { if (src != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(src); }
            }
        }


        

        private void InsertYearSheetsAfterEquipment(Excel.Application xlApp, Excel.Workbook wbNew, string dir)
        {
            try
            {
                string yearFile = null;
                try
                {
                    var exts = new[] { ".xls", ".xlsx", ".xlsm", ".xlsb" };
                    var files = Directory.GetFiles(dir)
                        .Where(f => exts.Contains(Path.GetExtension(f), StringComparer.OrdinalIgnoreCase))
                        .ToArray();
                    yearFile = files.FirstOrDefault(f => Path.GetFileName(f).IndexOf("연차", StringComparison.OrdinalIgnoreCase) >= 0);
                }
                catch { }

                if (string.IsNullOrEmpty(yearFile) || !File.Exists(yearFile))
                    return;

                Excel.Workbook wbYear = null;
                string tempYearPath = null;
                try
                {
                    // 복사본 생성 및 외부 링크 업데이트 비활성화해서 열기
                    try
                    {
                        var yearExt = Path.GetExtension(yearFile);
                        tempYearPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + yearExt);
                        try { File.Copy(yearFile, tempYearPath, true); } catch (Exception ex) { System.Diagnostics.Debug.WriteLine($"Failed to copy yearFile to temp: {ex.Message}"); tempYearPath = null; }
                    }
                    catch { tempYearPath = null; }

                    try
                    {
                        if (!string.IsNullOrEmpty(tempYearPath) && File.Exists(tempYearPath))
                        {
                            try { xlApp.AskToUpdateLinks = false; } catch { }
                            try { wbYear = xlApp.Workbooks.Open(tempYearPath, 0, ReadOnly: false); } catch (Exception ex) { System.Diagnostics.Debug.WriteLine($"Failed to open temp year workbook: {ex.Message}"); }
                        }
                        else
                        {
                            try { wbYear = xlApp.Workbooks.Open(yearFile, 0, ReadOnly: true); } catch { wbYear = null; }
                        }
                    }
                    catch { }

                    if (wbYear != null)
                    {
                        // 외부 링크 끊기
                        try
                        {
                            object linksObj = null;
                            try { linksObj = wbYear.LinkSources(Excel.XlLinkType.xlLinkTypeExcelLinks); } catch { linksObj = null; }
                            if (linksObj is System.Array linksArr)
                            {
                                foreach (var l in linksArr)
                                {
                                    try { wbYear.BreakLink(l.ToString(), Excel.XlLinkType.xlLinkTypeExcelLinks); } catch (Exception ex) { System.Diagnostics.Debug.WriteLine($"BreakLink failed: {ex.Message}"); }
                                }
                            }
                        }
                        catch { }

                        // 연결 삭제
                        try
                        {
                            dynamic conns = null;
                            try { conns = wbYear.Connections; } catch { conns = null; }
                            if (conns != null)
                            {
                                try
                                {
                                    int cc = conns.Count;
                                    for (int ci = cc; ci >= 1; ci--)
                                    {
                                        try { conns.Item(ci).Delete(); } catch { }
                                    }
                                }
                                catch { }
                            }
                        }
                        catch { }

                        // wbYear 시트 이름 수집
                        var wbYearSheetNames = new List<string>();
                        try
                        {
                            int yc = wbYear.Worksheets.Count;
                            for (int yi = 1; yi <= yc; yi++)
                            {
                                Excel.Worksheet tsh = null;
                                try { tsh = (Excel.Worksheet)wbYear.Worksheets[yi]; wbYearSheetNames.Add(tsh.Name); }
                                catch { }
                                finally { if (tsh != null) try { Marshal.ReleaseComObject(tsh); } catch { } }
                            }
                        }
                        catch { }

                        // 먼저 연차 파일의 '갑지'와 '제출문'을 통합문서 앞쪽에 복사
                        var exactNames = new[] { "갑지", "제출문" };
                        Excel.Worksheet firstSheet = null;
                        try { firstSheet = (Excel.Worksheet)wbNew.Worksheets[1]; } catch { }

                        if (firstSheet != null)
                        {
                            // 역순으로 Before 복사하여 순서 유지
                            for (int ni = exactNames.Length - 1; ni >= 0; ni--)
                            {
                                var name = exactNames[ni];
                                try
                                {
                                    foreach (var sname in wbYearSheetNames)
                                    {
                                        Excel.Worksheet src = null;
                                        Excel.Worksheet added = null;
                                        try
                                        {
                                            if (sname == null) continue;
                                            if (sname.IndexOf(name, StringComparison.OrdinalIgnoreCase) < 0) continue;

                                            try { src = (Excel.Worksheet)wbYear.Worksheets[sname]; } catch { src = null; }
                                            if (src == null) continue;

                                            // 시트 전체를 복사하는 대신 새 시트를 추가하고 셀값과 서식을 복사
                                            try
                                            {
                                                added = (Excel.Worksheet)wbNew.Worksheets.Add(Before: firstSheet);
                                                try { added.Name = src.Name; } catch { }

                                                Excel.Range srcUsed = null;
                                                Excel.Range destStart = null;
                                                try
                                                {
                                                    srcUsed = src.UsedRange;
                                                    if (srcUsed != null)
                                                    {
                                                        destStart = added.Range["A1"];
                                                        // 값 복사
                                                        try { destStart.Resize[srcUsed.Rows.Count, srcUsed.Columns.Count].Value = srcUsed.Value; } catch { }
                                                        // 숫자/표시 형식 복사
                                                        try { destStart.Resize[srcUsed.Rows.Count, srcUsed.Columns.Count].NumberFormat = srcUsed.NumberFormat; } catch { }
                                                        // 열 너비 복사
                                                        try
                                                        {
                                                            for (int c = 1; c <= srcUsed.Columns.Count; c++)
                                                            {
                                                                try { added.Columns[c].ColumnWidth = src.Columns[c].ColumnWidth; } catch { }
                                                            }
                                                        }
                                                        catch { }
                                                    }
                                                }
                                                catch { }
                                                finally { if (srcUsed != null) try { Marshal.ReleaseComObject(srcUsed); } catch { } }
                                            }
                                            catch { }
                                            finally { if (added != null) try { Marshal.ReleaseComObject(added); } catch { } }

                                            try { if (src != null) Marshal.ReleaseComObject(src); } catch { }
                                            break; // 한 시트만 복사
                                        }
                                        catch { }
                                    }
                                }
                                catch { }
                            }

                            try { Marshal.ReleaseComObject(firstSheet); } catch { }
                        }

                        // 장비 시트 뒤에 패턴 기반 시트(저압*, 예비*, *분기*)를 복사
                        Excel.Worksheet equipmentSheet = ((IEnumerable<Excel.Worksheet>)wbNew.Worksheets.Cast<Excel.Worksheet>())
                            .FirstOrDefault(w => string.Equals(w.Name, "장비", StringComparison.OrdinalIgnoreCase));

                        Excel.Worksheet insertAfter = equipmentSheet;
                        if (insertAfter == null)
                        {
                            insertAfter = (Excel.Worksheet)wbNew.Worksheets[wbNew.Worksheets.Count];
                        }

                        var patterns = new[] { "저압", "예비", "분기" };
                        foreach (var pattern in patterns)
                        {
                            try
                            {
                                var toDelete = wbNew.Worksheets.Cast<Excel.Worksheet>()
                                    .Where(w => w.Name.IndexOf(pattern, StringComparison.OrdinalIgnoreCase) >= 0)
                                    .ToList();
                                foreach (var d in toDelete)
                                {
                                    try { d.Delete(); } catch { }
                                    try { Marshal.ReleaseComObject(d); } catch { }
                                }
                            }
                            catch { }

                            try
                            {
                                // 이름 리스트 기반으로 새 RCW를 얻어 안전하게 복사
                                foreach (var sname in wbYearSheetNames)
                                {
                                    Excel.Worksheet src = null;
                                    try
                                    {
                                        if (sname == null) continue;
                                        if (sname.IndexOf(pattern, StringComparison.OrdinalIgnoreCase) < 0) continue;
                                        try { src = (Excel.Worksheet)wbYear.Worksheets[sname]; } catch { src = null; }
                                        if (src == null) continue;

                                        try
                                        {
                                            src.Copy(After: insertAfter);
                                            int newIndex = insertAfter.Index + 1;
                                            Excel.Worksheet newSh = null;
                                            try { newSh = (Excel.Worksheet)wbNew.Worksheets[newIndex]; insertAfter = newSh; }
                                            catch { }
                                            finally { if (newSh != null) Marshal.ReleaseComObject(newSh); }
                                        }
                                        catch (Exception ex)
                                        {
                                            System.Diagnostics.Debug.WriteLine($"Copy failed for sheet '{sname}': {ex.Message}");
                                            // 폴백 처리(생략 here for brevity)
                                        }
                                    }
                                    catch { }
                                    finally { if (src != null) try { Marshal.ReleaseComObject(src); } catch { } }
                                }
                            }
                            catch { }
                        }

                        try { if (equipmentSheet != null) Marshal.ReleaseComObject(equipmentSheet); } catch { }
                    }
                }
                finally
                {
                    if (wbYear != null)
                    {
                        try { wbYear.Close(false); } catch { }
                        try { Marshal.ReleaseComObject(wbYear); } catch { }
                    }
                    try
                    {
                        if (!string.IsNullOrEmpty(tempYearPath) && File.Exists(tempYearPath))
                        {
                            try { File.Delete(tempYearPath); } catch { }
                        }
                    }
                    catch { }
                }
            }
            catch { }
        }

        private void ReplaceTextInWorkbook(Excel.Workbook wb, string oldText, string newText)
        {
            try
            {
                var sheets = wb.Worksheets;
                int count = sheets.Count;
                for (int i = 1; i <= count; i++)
                {
                    Excel.Worksheet sh = null;
                    Excel.Range cells = null;
                    try
                    {
                        sh = sheets[i] as Excel.Worksheet;
                        if (sh != null)
                        {
                            cells = sh.Cells as Excel.Range;
                            if (cells != null)
                            {
                                cells.Replace(oldText, newText,
                                    Excel.XlLookAt.xlPart,
                                    Excel.XlSearchOrder.xlByRows,
                                    false, Type.Missing, Type.Missing, Type.Missing);
                            }
                        }
                    }
                    catch { }
                    finally
                    {
                        if (cells != null) Marshal.ReleaseComObject(cells);
                        if (sh != null) Marshal.ReleaseComObject(sh);
                    }
                }
            }
            catch { }
        }

        private string FindTemplatePath(string fileName)
        {
            var dir = new DirectoryInfo(AppContext.BaseDirectory);
            for (int i = 0; i < 8; i++)
            {
                var candidate = Path.Combine(dir.FullName, "0_org", fileName);
                if (File.Exists(candidate))
                    return Path.GetFullPath(candidate);

                if (dir.Parent == null)
                    break;
                dir = dir.Parent;
            }

            throw new FileNotFoundException($"템플릿 파일을 찾을 수 없습니다: {fileName}");
        }

        // --- 헬퍼 메서드들 ---
        private HashSet<string> GetWorkbookSheetNames(Excel.Workbook wb)
        {
            var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            try
            {
                foreach (Excel.Worksheet sh in wb.Worksheets)
                {
                    try { names.Add(sh.Name); } catch { }
                    finally { try { System.Runtime.InteropServices.Marshal.ReleaseComObject(sh); } catch { } }
                }
            }
            catch { }
            return names;
        }


        private void EnsureSheetOrder(Excel.Workbook wbNew)
        {
            try
            {
                var order = new[] { "연계획", "전기설비", "장비", "검교정" };
                for (int i = 1; i < order.Length; i++)
                {
                    string prevName = order[i - 1];
                    string curName = order[i];

                    Excel.Worksheet prevSheet = null;
                    Excel.Worksheet curSheet = null;
                    try
                    {
                        prevSheet = ((IEnumerable<Excel.Worksheet>)wbNew.Worksheets.Cast<Excel.Worksheet>())
                            .FirstOrDefault(w => string.Equals(w.Name, prevName, StringComparison.OrdinalIgnoreCase));
                        curSheet = ((IEnumerable<Excel.Worksheet>)wbNew.Worksheets.Cast<Excel.Worksheet>())
                            .FirstOrDefault(w => string.Equals(w.Name, curName, StringComparison.OrdinalIgnoreCase));

                        if (curSheet != null && prevSheet != null)
                        {
                            try
                            {
                                int curIndex = curSheet.Index;
                                int prevIndex = prevSheet.Index;
                                if (curIndex != prevIndex + 1)
                                {
                                    curSheet.Move(After: prevSheet);
                                }
                            }
                            catch { }
                        }
                    }
                    catch { }
                    finally
                    {
                        if (prevSheet != null) Marshal.ReleaseComObject(prevSheet);
                        if (curSheet != null) Marshal.ReleaseComObject(curSheet);
                    }
                }
            }
            catch { }
        }

        private void textBox1_Enter(object sender, EventArgs e)
        {

        }

        private void panel2_Paint(object sender, PaintEventArgs e)
        {

        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        #region [다운로드]
        private SynologyFileDownloader _downloader;
        //private BindingList<SynologyFileGridItem> _gridFiles;

        private BindingSource _fileBindingSource = new BindingSource();
        private List<SynologyFileItem> _files = new List<SynologyFileItem>();

        private BindingList<SynologyFileGridItem> ConvertToGridItems(List<SynologyFileItem> files)
        {
            var list = files
                .Select(x => new SynologyFileGridItem
                {
                    Selected = true,
                    Name = x.Name,
                    Path = x.Path,
                    Size = x.Size,
                    SizeText = FormatFileSize(x.Size),
                    ModifiedTime = x.ModifiedTime,
                    ModifiedTimeText = x.ModifiedTime == DateTime.MinValue
                        ? ""
                        : x.ModifiedTime.ToString("yyyy-MM-dd HH:mm:ss")
                })
                .ToList();

            return new BindingList<SynologyFileGridItem>(list);
        }

        private string FormatFileSize(long bytes)
        {
            if (bytes < 1024) return bytes + " B";
            if (bytes < 1024 * 1024) return (bytes / 1024.0).ToString("0.0") + " KB";
            if (bytes < 1024 * 1024 * 1024) return (bytes / 1024.0 / 1024.0).ToString("0.0") + " MB";
            return (bytes / 1024.0 / 1024.0 / 1024.0).ToString("0.0") + " GB";
        }

        private void ExtractAllZipFiles(string folderPath, bool deleteZip = true)
        {
            if (!Directory.Exists(folderPath))
                return;

            foreach (string zipFile in Directory.GetFiles(folderPath, "*.zip"))
            {
                string extractPath = Path.Combine(
                    folderPath,
                    Path.GetFileNameWithoutExtension(zipFile));

                Directory.CreateDirectory(extractPath);

                if (Directory.Exists(extractPath))
                {
                    Directory.Delete(extractPath, true);
                }
                ZipFile.ExtractToDirectory(zipFile, folderPath);

                if (deleteZip)
                {
                    File.Delete(zipFile);
                }
            }
        }

        private async Task DownloadPreviousReportAsync(string folderPath)
        {
            if (!Directory.Exists(folderPath))
                return;

            await _downloader.DownloadPreviousReportsAsync(folderPath);

            AddLog("INFO", folderPath + "이전 보고서 다운로드 완료");
        }

        public async Task TestSynologyAsync(List<SearchFolderOption> folders = null, string keyword = null)
        {
            Cursor = Cursors.WaitCursor;

            if (folders == null || folders.Count == 0) { }
                string downloadFolderSeasonReport = "/2_1전기직무고시점검보고서/2 본부(김희철)/0001 시흥보고서방/001열화상";
                string downloadFolderAnuualReport = "/2_1전기직무고시점검보고서/2 본부(김희철)/0001 시흥보고서방";


            folders = new System.Collections.Generic.List<SearchFolderOption>
            {
                new SearchFolderOption
                {
                    Folder = downloadFolderSeasonReport,
                    Recursive = false
                },
                new SearchFolderOption
                {
                    Folder = downloadFolderAnuualReport,
                    Recursive = false
                }
            };



            if (string.IsNullOrEmpty(keyword)) keyword = "유랑진";

            try
            {
                var config = new SynologyFileDownloaderConfig
                {
                    BaseUrl = "http://hk2ng.synology.me:5000",
                    Username = "유량진",
                    Password = "HKeng717241",
                    SearchFolders = folders,
                    SearchFolder = "/2_1전기직무고시점검보고서",
                    Keyword = keyword,
                    Extensions = new[] { ".xlsx", ".pdf" },
                    Recursive = true,
                    DownloadLatestOnly = false
                };

                if (_downloader != null)
                {
                    _downloader.Dispose();
                    _downloader = null;
                }


                _downloader = new SynologyFileDownloader(config);

                await _downloader.LoginAsync();

                _uploader = new SynologyFileUploader(
                                config,
                                _downloader.HttpClient,
                                _downloader.Sid);

                //_files = await _downloader.SearchFilesAsync();
                //_fileBindingSource.DataSource = _files;
                //dgvFiles.DataSource = _fileBindingSource;

                //MessageBox.Show("검색 완료: " + _files.Count + "건");

                int count = await _downloader.DownloadFoldersAsync(@"D:\work\Report\0now");

                if (count > 0)
                {
                    ExtractAllZipFiles(@"D:\work\Report\0now");
                    AddLog("INFO", keyword + " 검색 완료: " + count + "건 다운로드 완료");

                    await DownloadPreviousReportAsync(@"D:\work\Report\0now");

                    Process.Start("explorer.exe", @"D:\work\Report\0now");
                }
                else
                {
                    AddLog("INFO", keyword + " 검색 완료: 다운로드할 폴더가 없습니다.");
                    MessageBox.Show("다운로드할 폴더가 없습니다.");
                }
            }
            catch (Exception ex)
            {
                AddLog("ERROR", keyword + " 검색 중 오류 발생: " + ex.Message);
                MessageBox.Show(ex.ToString(), "오류");
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        private void InitGrid()
        {
            dgvFiles.AutoGenerateColumns = false;
            dgvFiles.AllowUserToAddRows = false;
            dgvFiles.AllowUserToDeleteRows = false;
            dgvFiles.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            dgvFiles.MultiSelect = false;
            dgvFiles.RowHeadersVisible = false;

            dgvFiles.Columns.Clear();

            dgvFiles.Columns.Add(new DataGridViewCheckBoxColumn
            {
                Name = "colSelected",
                HeaderText = "선택",
                DataPropertyName = "Selected",
                Width = 50
            });

            dgvFiles.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "colName",
                HeaderText = "파일명",
                DataPropertyName = "Name",
                Width = 220
            });

            dgvFiles.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "colPath",
                HeaderText = "경로",
                DataPropertyName = "Path",
                AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
            });

            dgvFiles.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "colSize",
                HeaderText = "크기",
                DataPropertyName = "SizeText",
                Width = 90
            });

            dgvFiles.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "colModified",
                HeaderText = "수정일",
                DataPropertyName = "ModifiedTimeText",
                Width = 150
            });
        }
        #endregion

        // newWb의 첫 번째 시트를 wb의 '분기'라는 이름을 포함한 시트 앞에 삽입 (복사만)
        private void InsertFirstSheetFromNewWbBeforeBranch(Excel.Workbook wb, Excel.Workbook newWb)
        {
            if (wb == null || newWb == null) return;

            Excel.Worksheet src = null;
            try
            {
                if (newWb.Worksheets.Count < 1) return;

                src = (Excel.Worksheet)newWb.Worksheets[1];

                int targetIndex = -1;
                int wc = 0;
                try { wc = wb.Worksheets.Count; } catch { }

                for (int i = 1; i <= wc; i++)
                {
                    Excel.Worksheet sh = null;
                    try
                    {
                        sh = (Excel.Worksheet)wb.Worksheets[i];
                        string name = null;
                        try { name = sh.Name; } catch { }
                        if (!string.IsNullOrEmpty(name) && name.Contains("분기"))
                        {
                            targetIndex = i;
                            break;
                        }
                    }
                    finally
                    {
                        if (sh != null) try { Marshal.ReleaseComObject(sh); } catch { }
                    }
                }

                if (targetIndex != -1)
                {
                    Excel.Worksheet target = null;
                    try
                    {
                        target = (Excel.Worksheet)wb.Worksheets[targetIndex];
                        src.Copy(Before: target);
                    }
                    finally
                    {
                        if (target != null) try { Marshal.ReleaseComObject(target); } catch { }
                    }
                }
                else
                {
                    // 분기 시트를 찾지 못하면 마지막에 복사
                    src.Copy(After: wb.Worksheets[wb.Worksheets.Count]);
                }
            }
            catch (Exception ex)
            {
                AddLog("Error", $"시트 삽입 중 오류: {ex.Message}");
            }
            finally
            {
                if (src != null) 
                    try {
                        wb.Save();
                        Marshal.ReleaseComObject(src); 
                    } 
                    catch { }
            }
        }


        private void FormMain_Load(object sender, EventArgs e)
        {
            InitGrid();
            AddLog("INFO", "프로그램 시작");



            try
            {
                // 탭에 Form을 임베드할 때는 TopLevel을 false로 설정해야 보입니다.
                var imageOrderForm = new FormSortImage();

                imageOrderForm.TopLevel = false;
                imageOrderForm.TopMost = false;
                imageOrderForm.FormBorderStyle = FormBorderStyle.None;
                imageOrderForm.Dock = DockStyle.Fill;

                // 중복 추가 방지
                this.tabSortImage.Controls.Clear();

                // 부모를 명시적으로 설정
                imageOrderForm.Parent = this.tabSortImage;
                this.tabSortImage.Controls.Add(imageOrderForm);
                imageOrderForm.Visible = true;
                imageOrderForm.Show();
                imageOrderForm.BringToFront();

                // 진단 로그: 탭에 추가된 컨트롤 정보 출력
                try
                {
                    AddLog("DEBUG", $"tabSortImage.Controls.Count={this.tabSortImage.Controls.Count}");
                    foreach (Control c in this.tabSortImage.Controls)
                    {
                        AddLog("DEBUG", $"Control: {c.GetType().FullName}, Name={c.Name}, Visible={c.Visible}");
                    }
                    // 탭 배경과 스크롤 옵션 체크
                    this.tabSortImage.AutoScroll = true;
                }
                catch { }
            }
            catch (Exception ex)
            {
                AddLog("ERROR", $"이미지 정렬 탭 초기화 실패: {ex.Message}");
            }
        }

        private async void tbCompany_Enter(object sender, EventArgs e)
        {

            //if (tbCompany.Text.Trim() == "")
            //{ 
            //    return;
            //}

            //await TestSynologyAsync();
        }

        private async void btnDownload_Click(object sender, EventArgs e)
        {
            try
            {
                Cursor = Cursors.WaitCursor;

                dgvFiles.EndEdit(); // 체크박스 편집 반영

                var files = _fileBindingSource.DataSource as List<SynologyFileItem>;
                if (files == null || files.Count == 0)
                {
                    MessageBox.Show("다운로드할 검색 결과가 없습니다.");
                    return;
                }

                string baseSaveFolder = @"C:\_D\work\한경이엔지\NAS다운로드";

                List<string> downloaded = await _downloader.DownloadSelectedFilesAsync(files, baseSaveFolder);

                MessageBox.Show(downloaded.Count + "개 파일 다운로드 완료");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "오류");
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        private async void tbCompany_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode != Keys.Enter)
                return;

            e.SuppressKeyPress = true;
            e.Handled = true;

            string keyword = tbCompany.Text.Trim();
            if (keyword == "")
                return;

            await TestSynologyAsync(null, keyword);
        }

        private void btnPageNumber_Click(object sender, EventArgs e)
        {
            var filePath = tbQuantityFile.Text?.Trim();
            if (string.IsNullOrEmpty(filePath) || !File.Exists(filePath))
            {
                MessageBox.Show("페이지 번호를 매길 엑셀 파일을 먼저 선택하세요.", "오류", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            Cursor = Cursors.WaitCursor;
            Excel.Application xlApp = null;
            Excel.Workbook wb = null;

            try
            {
                xlApp = new Excel.Application { Visible = false, DisplayAlerts = false };
                // Open for write because we modify PageSetup
                wb = xlApp.Workbooks.Open(filePath, ReadOnly: false);

                int sheetCount = 0;
                try { sheetCount = wb.Worksheets.Count; } catch { }

                //AddLog("Info", $"Worksheet Count = {sheetCount}");
                //AddLog("Info",
                //    $"Sheets={wb.Sheets.Count}, Worksheets={wb.Worksheets.Count}");

                //for (int i = 1; i <= wb.Sheets.Count; i++)
                //{
                //    object obj = null;

                //    try
                //    {
                //        obj = wb.Sheets[i];

                //        if (obj is Excel.Worksheet ws)
                //        {
                //            AddLog("Info", $"[{i}] Worksheet : {ws.Name}");
                //            Marshal.ReleaseComObject(ws);
                //        }
                //        else if (obj is Excel.Chart chart)
                //        {
                //            AddLog("Info", $"[{i}] Chart : {chart.Name}");
                //            Marshal.ReleaseComObject(chart);
                //        }
                //        else
                //        {
                //            AddLog("Info", $"[{i}] 기타 Sheet");
                //        }
                //    }
                //    catch (Exception ex)
                //    {
                //        AddLog("Error", $"Sheet[{i}] : {ex.Message}");
                //    }
                //}

                int currentStartPage = 1;

                for (int i = 1; i <= sheetCount; i++)
                {
                    Excel.Worksheet sh = null;
                    try
                    {
                        sh = (Excel.Worksheet)wb.Worksheets[i];
                        string name = "";
                        try { name = sh.Name; } catch { }

                        AddLog(
                            "Info",
                            $"[{i}] Name={sh.Name}, " +
                            $"Visible={sh.Visible}, " +
                            $"Type={sh.Type}");

                        // '갑지' 시트는 페이지 번호 매기기에서 제외
                        if (string.Equals(name, "갑지", StringComparison.OrdinalIgnoreCase))
                            continue;

                        // 현재 시트를 활성화
                        sh.Activate();
                        //int pages = sh.HPageBreaks.Count + 1; // 페이지 나누기의 개수

                        //if (sh.PageSetup.PrintArea == null)
                        //{
                        //    continue;
                        //}

                        //string printArea = sh.PageSetup.PrintArea;

                        //AddLog("Info", $"{sh.Name} PrintArea=[{printArea}]");

                        //if (string.IsNullOrWhiteSpace(printArea))
                        //{
                        //    AddLog("WARN", $"{sh.Name} : 인쇄영역 없음");
                        //    continue;
                        //}



                        //Excel.Range printRange = sh.Range[sh.PageSetup.PrintArea];
                        int lastRow = 0;// = printRange.Row + printRange.Rows.Count - 1;

                        string printAreaAddress = sh.PageSetup.PrintArea;

                        if (!string.IsNullOrWhiteSpace(printAreaAddress))
                        {
                            // 인쇄영역이 설정되어 있는 경우
                            Excel.Range printArea = sh.Range[printAreaAddress];

                            lastRow = printArea.Row + printArea.Rows.Count - 1;

                            Marshal.ReleaseComObject(printArea);
                        }
                        else
                        {
                            // 인쇄영역이 없는 경우 UsedRange 기준
                            Excel.Range usedRange = sh.UsedRange;

                            lastRow = usedRange.Row + usedRange.Rows.Count - 1;

                            Marshal.ReleaseComObject(usedRange);
                        }

                        int pages = 1;
                        int preRow = 0;

                        foreach (Excel.HPageBreak pb in sh.HPageBreaks)
                        {
                            if ((pb.Type == Excel.XlPageBreak.xlPageBreakManual &&
                                pb.Location.Row >= lastRow + 1)
                                || preRow >= pb.Location.Row)
                            {
                                // 인쇄영역 바로 다음 행에 있는 수동 페이지 나누기는 무시
                                continue;
                            }
                            Debug.WriteLine($"page : {pages}, row: {pb.Location.Row}");
                            pages++;
                            preRow = pb.Location.Row;
                        }
                        sh.PageSetup.FirstPageNumber = currentStartPage;

                        System.Diagnostics.Debug.WriteLine(
                            $"{name} : 시작={currentStartPage}, 페이지수={pages}");

                        Debug.WriteLine(sh.PageSetup.PrintArea);
                        Debug.WriteLine(sh.UsedRange.Address);
                        Debug.WriteLine(sh.HPageBreaks.Count);
                        Debug.WriteLine(sh.DisplayPageBreaks);

                        Debug.WriteLine($"시트 : {sh.Name}");

                        foreach (Excel.HPageBreak pb in sh.HPageBreaks)
                        {
                            Debug.WriteLine($"Break : {pb.Location.Address}");
                            Debug.WriteLine($"{pb.Location.Address}  {pb.Type}");
                        }

                        currentStartPage += pages;

                        //if (sh.Name == "마")
                        //{
                        //    break;
                        //}
                    }
                    catch (Exception ex)
                    {
                        string sheetName = "(알 수 없음)";

                        try
                        {
                            if (sh != null)
                                sheetName = sh.Name;
                        }
                        catch { }

                        AddLog(
                            "Error",
                            $"페이지 번호 처리 중 오류(idx={i}, name={sheetName}): {ex.Message}");
                    }
                    finally
                    {
                        try { wb.Save(); } catch { }
                        if (sh != null)
                            Marshal.ReleaseComObject(sh);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"페이지 번호 매기기 중 오류가 발생했습니다:\r\n{ex.Message}", "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);

                AddLog("Error", $"페이지 번호 매기기 중 오류가 발생했습니다: {ex.Message}");
            }
            finally
            {
                try
                {
                    if (wb != null)
                    {
                        wb.Save();
                        wb.Close(false);
                        Marshal.ReleaseComObject(wb);
                    }
                }
                catch { }

                try
                {
                    if (xlApp != null)
                    {
                        xlApp.Quit();
                        Marshal.ReleaseComObject(xlApp);
                    }
                }
                catch { }

                GC.Collect();
                GC.WaitForPendingFinalizers();
                Cursor = Cursors.Default;
            }
        }

        public class ExportResult
        {
            public bool Success { get; set; }
            public string OutputFile { get; set; }
            public string ErrorMessage { get; set; }
        }

        private void btnExportForPdf_Click(object sender, EventArgs e)
        {
            btnExportForPdf.Enabled = false;
            Cursor = Cursors.WaitCursor;

            try
            {
                string filePath = tbQuantityFile.Text?.Trim();

                var result = ExportPdfAsync(filePath);

                if (!result.Success)
                {
                    MessageBox.Show(result.ErrorMessage);
                    AddLog("Error", $"통합 PDF 내보내기 실패: {result.ErrorMessage}");
                }
                else
                {
                    Process.Start(new ProcessStartInfo
                    {
                        FileName = result.OutputFile,
                        UseShellExecute = true
                    });
                }
            }
            finally
            {
                Cursor = Cursors.Default;
                btnExportForPdf.Enabled = true;
            }
        }

        private ExportResult ExportPdfAsync(string filePath)
        {
            ExportResult returnValue = new ExportResult
            {
                Success = true
            };

            //var filePath = tbFileNameForFunction.Text?.Trim();
            if (string.IsNullOrEmpty(filePath) || !File.Exists(filePath))
            {
                returnValue = new ExportResult
                {
                    Success = false,
                    ErrorMessage = "내보낼 엑셀 파일을 먼저 선택하세요."
                };
                return returnValue;
            }

            Excel.Application xlApp = null;
            Excel.Workbook wb = null;
            string outFile = null;

            try
            {
                xlApp = new Excel.Application { Visible = false, DisplayAlerts = false };
                wb = xlApp.Workbooks.Open(filePath, ReadOnly: true);

                var dir = Path.GetDirectoryName(filePath) ?? Environment.CurrentDirectory;
                string baseName = Path.GetFileNameWithoutExtension(filePath);
                outFile = Path.Combine(dir, baseName + ".pdf");
                //int idx = 1;
                //var candidate = outFile;
                //while (File.Exists(candidate))
                //{
                //    candidate = Path.Combine(dir, baseName + "_merged_" + idx + ".pdf");
                //    idx++;
                //}

                // Export entire workbook as a single PDF (모든 시트를 하나의 PDF로)
                try
                {
                    wb.ExportAsFixedFormat(
                        Excel.XlFixedFormatType.xlTypePDF,
                        outFile,
                        Excel.XlFixedFormatQuality.xlQualityStandard,
                        IncludeDocProperties: true,
                        IgnorePrintAreas: false,
                        OpenAfterPublish: false); 
                }
                catch (Exception ex)
                {
                    //throw new Exception("통합 PDF 내보내기 실패: " + ex.Message, ex);
                    AddLog("Error", $"통합 PDF 내보내기 실패: {ex.Message}");
                }
            }
            catch (Exception ex)
            {
                returnValue = new ExportResult
                {
                    Success = false,
                    ErrorMessage = $"PDF 내보내기 중 오류가 발생했습니다:\r\n{ex.Message}"
                };
            }
            finally
            {
                try
                {
                    if (wb != null)
                    {
                        wb.Close(false);
                        Marshal.ReleaseComObject(wb);
                    }
                }
                catch { returnValue.Success = false; }

                try
                {
                    if (xlApp != null)
                    {
                        xlApp.Quit();
                        Marshal.ReleaseComObject(xlApp);
                    }
                }
                catch { returnValue.Success = false; }

                GC.Collect();
                GC.WaitForPendingFinalizers();

                returnValue.OutputFile = outFile;
            }
            return returnValue;
        }

        private void btnFindFileForFunction_Click(object sender, EventArgs e)
        {
            using (var dlg = new OpenFileDialog())
            {
                dlg.Title = "엑셀 파일을 선택하세요";
                dlg.Filter = "Excel 파일 (*.xlsx;*.xls;*.xlsm)|*.xlsx;*.xls;*.xlsm|모든 파일 (*.*)|*.*";
                dlg.CheckFileExists = true;
                dlg.CheckPathExists = true;
                dlg.Multiselect = false;

                // 초기 디렉터리 설정: 먼저 폼의 tbFolder에 입력된 값을 사용
                string initial = tbFolder.Text;
                if (string.IsNullOrWhiteSpace(initial))
                    initial = @"C:\_D\work\한경이엔지\2_report";

                try
                {
                    if (Directory.Exists(initial))
                        dlg.InitialDirectory = initial;
                    else
                    {
                        var dir = Path.GetDirectoryName(initial);
                        if (!string.IsNullOrEmpty(dir) && Directory.Exists(dir))
                            dlg.InitialDirectory = dir;
                    }
                }
                catch { }

                if (dlg.ShowDialog() == DialogResult.OK)
                {
                    tbQuantityFile.Text = dlg.FileName;

                    //if (cbAutoExportPdf.Checked)
                    //{


                    //    Cursor = Cursors.WaitCursor;
                    //    btnPageNumber_Click(sender, e);

                    //    btnExportForPdf_Click(sender, e);
                    //    Cursor = Cursors.Default;
                    //}
                }
            }
        }

        private void btErrorPageUpdate_Click(object sender, EventArgs e)
        {
            var filePath = tbQuantityFile.Text?.Trim();
            if (string.IsNullOrEmpty(filePath) || !File.Exists(filePath))
            {
                MessageBox.Show("페이지 번호를 매길 엑셀 파일을 먼저 선택하세요.", "오류", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            Excel.Application xlApp = null;
            Excel.Workbook wb = null;

            try
            {
                xlApp = new Excel.Application { Visible = false, DisplayAlerts = false };
                // Open for write because we modify PageSetup
                wb = xlApp.Workbooks.Open(filePath, ReadOnly: false);

                OpinionMaker.FillOpinionFromInsulation(wb);

                try { wb.Save(); } catch { }
            }
            catch (Exception ex)
            {
                //MessageBox.Show($"페이지 번호 매기기 중 오류가 발생했습니다:\r\n{ex.Message}", "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                AddLog("Error", $"페이지 번호 매기기 중 오류가 발생했습니다: {ex.Message}");
            }
            finally
            {
                try
                {
                    if (wb != null)
                    {
                        wb.Close(false);
                        Marshal.ReleaseComObject(wb);
                    }
                }
                catch { }

                try
                {
                    if (xlApp != null)
                    {
                        xlApp.Quit();
                        Marshal.ReleaseComObject(xlApp);
                    }
                }
                catch { }

                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        
        private void ClearAllShapes(Excel.Worksheet ws)
        {
            if (ws == null) return;

            // 뒤에서부터 삭제해야 안전함
            for (int i = ws.Shapes.Count; i >= 1; i--)
            {
                Excel.Shape shp = null;
                try
                {
                    shp = ws.Shapes.Item(i);
                    shp.Delete();
                }
                finally
                {
                    ReleaseObject(shp);
                }
            }
        }

        private void CreateExcelWithImages(
            string[] imageFiles,
            string savePath,
            double imageWidthCm,
            double imageHeightCm,
            double gapCm)
            {
            Excel.Application xlApp = null;
            Excel.Workbook wb = null;
            Excel.Worksheet ws = null;

            bool success = false;

            try
            {
                xlApp = new Excel.Application();
                xlApp.Visible = false;
                xlApp.DisplayAlerts = false;

                wb = xlApp.Workbooks.Add();
                ws = (Excel.Worksheet)wb.Sheets[1];
                ws.Name = "Images";

                // cm -> point 변환
                double imgWidthPt = CmToPoint(imageWidthCm);
                double imgHeightPt = CmToPoint(imageHeightCm);
                double gapPt = CmToPoint(gapCm);

                // 시작 위치
                double startLeft = 10;
                double startTop = 10;

                // 세로 간격
                double rowStep = imgHeightPt + 10;

                // ★ 2개씩 한 줄 처리
                for (int i = 0; i < imageFiles.Length; i += 2)
                {
                    int rowIndex = i / 2;
                    double top = startTop + (rowIndex * rowStep);

                    Excel.Shape leftShape = null;
                    Excel.Shape rightShape = null;

                    try
                    {
                        // 왼쪽 이미지
                        leftShape = AddImage(
                            ws,
                            imageFiles[i],
                            startLeft,
                            top,
                            imgWidthPt,
                            imgHeightPt);

                        // 오른쪽 이미지
                        if (i + 1 < imageFiles.Length && leftShape != null)
                        {
                            double rightLeft = leftShape.Left + leftShape.Width + gapPt;

                            rightShape = AddImage(
                                ws,
                                imageFiles[i + 1],
                                rightLeft,
                                top,
                                imgWidthPt,
                                imgHeightPt);
                        }
                    }
                    finally
                    {
                        ReleaseObject(rightShape);
                        ReleaseObject(leftShape);
                    }
                }

                wb.SaveAs(savePath, Excel.XlFileFormat.xlOpenXMLWorkbook);

                ws.Activate();
                xlApp.Visible = true;
                xlApp.WindowState = Excel.XlWindowState.xlMaximized;

                success = true;
            }
            finally
            {
                ReleaseObject(ws);
                ReleaseObject(wb);

                if (!success && xlApp != null)
                {
                    try { xlApp.Quit(); } catch { }
                    ReleaseObject(xlApp);
                }

                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        private Excel.Shape AddImage(
            Excel.Worksheet ws,
            string imagePath,
            double left,
            double top,
            double widthPt,
            double heightPt)
        {
            Excel.Shape shape = null;

            // 먼저 원본 크기로 넣기(-1, -1)
            shape = ws.Shapes.AddPicture(
                imagePath,
                Office.MsoTriState.msoFalse,   // LinkToFile
                Office.MsoTriState.msoTrue,    // SaveWithDocument
                (float)left,
                (float)top,
                -1,   // 원본 너비
                -1);  // 원본 높이

            // 원본 크기 기억
            float originalWidth = shape.Width;
            float originalHeight = shape.Height;

            // 셀 변화 영향 최소화
            shape.Placement = Excel.XlPlacement.xlFreeFloating;


            // 비율 잠금 해제
            shape.LockAspectRatio = Office.MsoTriState.msoFalse;

            // ★ 원본 대비 배율도 강제로 갱신
            float widthPercent = (float)(widthPt / originalWidth * 100.0);
            float heightPercent = (float)(heightPt / originalHeight * 100.0);

            shape.ScaleWidth(widthPercent, Office.MsoTriState.msoFalse,
                Microsoft.Office.Core.MsoScaleFrom.msoScaleFromTopLeft);

            shape.ScaleHeight(heightPercent, Office.MsoTriState.msoFalse,
                Microsoft.Office.Core.MsoScaleFrom.msoScaleFromTopLeft);

            // 원하는 크기로 강제
            shape.Width = (float)widthPt;
            shape.Height = (float)heightPt;

            //shape.LockAspectRatio = Office.MsoTriState.msoFalse;

            // 위치도 다시 한번 고정
            shape.Left = (float)left;
            shape.Top = (float)top;

            return shape;
        }

        //private Excel.Shape AddImage(
        //    Excel.Worksheet ws,
        //    string imagePath,
        //    double left,
        //    double top,
        //    double widthPt,
        //    double heightPt)
        //{
        //    object pictures = null;
        //    object picture = null;
        //    Excel.Shape shape = null;

        //    try
        //    {
        //        pictures = ws.Pictures();
        //        picture = pictures.GetType().InvokeMember(
        //            "Insert",
        //            System.Reflection.BindingFlags.InvokeMethod,
        //            null,
        //            pictures,
        //            new object[] { imagePath });

        //        shape = picture as Excel.Shape;
        //        if (shape == null)
        //        {
        //            shape = ws.Shapes.Item(ws.Shapes.Count);
        //        }

        //        // ★ 비율/배치 관련 속성
        //        shape.LockAspectRatio = Office.MsoTriState.msoFalse;
        //        shape.Placement = Excel.XlPlacement.xlFreeFloating;

        //        // 위치/크기 강제 지정
        //        shape.Left = (float)left;
        //        shape.Top = (float)top;
        //        shape.Width = (float)widthPt;
        //        shape.Height = (float)heightPt;

        //        return shape;
        //    }
        //    finally
        //    {
        //        ReleaseObject(picture);
        //        ReleaseObject(pictures);
        //    }
        //}

        private void InsertImagesToExistingWorkbook(
            string excelPath,
            string[] imageFiles,
            double imageWidthCm,
            double imageHeightCm,
            double gapCm)
        {
            Excel.Application xlApp = null;
            Excel.Workbook wb = null;
            Excel.Worksheet ws = null;

            bool success = false;

            try
            {
                xlApp = new Excel.Application();
                xlApp.Visible = false;
                xlApp.DisplayAlerts = false;

                wb = xlApp.Workbooks.Open(excelPath);

                // 넣을 시트
                ws = (Excel.Worksheet)wb.Sheets[1];
                // 시트명으로 하고 싶으면:
                // ws = (Excel.Worksheet)wb.Sheets["Images"];

                // ★ 기존 객체(이미지/도형) 삭제
                ClearAllShapes(ws);

                double imgWidthPt = CmToPoint(imageWidthCm);
                double imgHeightPt = CmToPoint(imageHeightCm);
                double gapPt = CmToPoint(gapCm);

                double startLeft = 10;
                double startTop = 10;

                // 세로 간격은 일단 이미지 높이 + 10pt
                double rowStep = imgHeightPt + 10;

                for (int i = 0; i < imageFiles.Length; i += 2)
                {
                    int rowIndex = i / 2;
                    double top = startTop + (rowIndex * rowStep);

                    Excel.Shape leftShape = null;
                    Excel.Shape rightShape = null;

                    try
                    {
                        // 왼쪽 이미지
                        leftShape = AddImage(ws, imageFiles[i], startLeft, top, imgWidthPt, imgHeightPt);

                        // 오른쪽 이미지
                        if (i + 1 < imageFiles.Length && leftShape != null)
                        {
                            double rightLeft = leftShape.Left + leftShape.Width + gapPt;

                            rightShape = AddImage(ws, imageFiles[i + 1], rightLeft, top, imgWidthPt, imgHeightPt);
                        }
                    }
                    finally
                    {
                        ReleaseObject(rightShape);
                        ReleaseObject(leftShape);
                    }
                }

                wb.Save();

                ws.Activate();
                xlApp.Visible = true;
                xlApp.WindowState = Excel.XlWindowState.xlMaximized;

                success = true;
            }
            finally
            {
                ReleaseObject(ws);
                ReleaseObject(wb);

                if (!success && xlApp != null)
                {
                    try { xlApp.Quit(); } catch { }
                    ReleaseObject(xlApp);
                }

                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        private double CmToPoint(double cm)
        {
            return cm * 28.3464567;
        }

        private void ReleaseObject(object obj)
        {
            if (obj != null)
            {
                try
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                }
                catch
                {
                }
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
                var defaultPath = @"C:\_D\work\한경이엔지\NAS다운로드";
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


        private void btnQuntityFileRun_Click(object sender, EventArgs e)
        {
            ProcImagesSheet();
        }

        private void ProcImagesSheet()
        {
            var filePath = tbQuantityFile.Text?.Trim();
            if (string.IsNullOrEmpty(filePath) || !File.Exists(filePath))
            {
                MessageBox.Show("엑셀 파일을 먼저 선택하세요.", "오류", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            Cursor = Cursors.WaitCursor;
            Excel.Application xlApp = null;
            Excel.Workbook wb = null;
            Excel.Worksheet ws = null;

            try
            {
                xlApp = new Excel.Application { Visible = false, DisplayAlerts = false };
                //wb = xlApp.Workbooks.Open(filePath, ReadOnly: false);\
                wb = xlApp.Workbooks.Open(filePath);


                string baseFolder = Path.GetDirectoryName(filePath);
                string pdfPath = Path.Combine(baseFolder, "02 전원품질", "K.pdf");

                // 열화상
                if (checkBoxFeverPicture.Checked)
                {
                    ProcFeverPicture(xlApp, wb, baseFolder, textBoxFeverImageFolder.Text);
                }

                // 품질
                if (checkBoxQuantity.Checked)
                {
                    ProcQuantitySheet(xlApp, wb, baseFolder, textBoxQuntatyFolder.Text, comboBoxTestReport.Text,
                         comboBoxTimeGraph.Text, comboBoxHwaveGraph.Text);
                }


                // 코로나
                if (checkBoxCorona.Checked)
                {
                    ProcCoronaSheet(xlApp, wb, baseFolder, textBoxCoronaFolder.Text);
                }

                // 점검사진
                if (checkBoxPicture.Checked)
                {
                    ProcCheckPicture(xlApp, wb, baseFolder, textBoxPictureFolder.Text);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"이미지 삽입 실패:\r\n{ex.Message}", "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                //return;
            }
            finally
            {
                try
                {
                    if (ws != null) Marshal.ReleaseComObject(ws);
                    if (wb != null)
                    {
                        wb.Close(false);
                        Marshal.ReleaseComObject(wb);
                    }
                }
                catch { }

                try
                {
                    if (xlApp != null)
                    {
                        xlApp.Quit();
                        Marshal.ReleaseComObject(xlApp);
                    }
                }
                catch { }

                GC.Collect();
                GC.WaitForPendingFinalizers();
                Cursor = Cursors.Default;
            }
        }

        private Excel.Worksheet GetWorksheetByName(Excel.Workbook wb, string sheetName)
        {
            foreach (Excel.Worksheet sheet in wb.Worksheets)
            {
                if (sheet.Name.Trim().IndexOf(sheetName, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    return sheet;
                }
                //Marshal.ReleaseComObject(sheet);
            }
            return null;
        }

        private Excel.Worksheet GetWorksheetByLastName(Excel.Workbook wb, string sheetName)
        {
            foreach (Excel.Worksheet sheet in wb.Worksheets)
            {
                if (sheet.Name.Trim().EndsWith(
                        sheetName.Trim(),
                        StringComparison.OrdinalIgnoreCase))
                {
                    return sheet;
                }
            }

            return null;
        }

        private void ProcCoronaSheet(Excel.Application xlApp, Excel.Workbook wb, string baseFolder, string text)
        {
            string tmpFolder = Path.Combine(baseFolder, text);

            string xlsPath = Directory.GetFiles(tmpFolder, "*.xls").FirstOrDefault();

            // PD 부분방전 결과지가 있으면 PD부분방전 처리, 없으면 영코 처리
            if (xlsPath != null)
            {
                // PD부분방전
                ProcPdCoronaSheet(xlApp, wb, baseFolder, text, xlsPath);
                return;
            }

            // 영코
            ProcVideoCoronaSheet(xlApp, wb, baseFolder, text);
        }

        #region [점검사진 시트 처리]
        private void ProcCheckPicture(Excel.Application xlApp, Excel.Workbook wb, string baseFolder, object text)
        {

            Excel.Worksheet ws = null;

            try
            {
                ws = GetWorksheetByName(wb, "사진");

                if (ws == null)
                {
                    throw new Exception("사진 시트를 찾을 수 없습니다.");
                }

                RemovePictures(ws);


                string[] files = Directory.GetFiles(baseFolder + $"\\{text}", "*.jpg")
                .OrderBy(f =>
                {
                    return int.TryParse(Path.GetFileNameWithoutExtension(f), out int n)
                        ? n
                        : int.MaxValue;
                })
                .ToArray();

                int imageIndex = 0;

                for (int page = 0; imageIndex < files.Length; page++)
                {
                    int pageOffset = page * 41;

                    for (int row = 0; row < 2; row++)
                    {
                        int rowOffset = row * 17;

                        for (int col = 0; col < 2; col++)
                        {
                            if (imageIndex >= files.Length)
                                break;

                            string fromCol = (col == 0) ? "A" : "O";
                            string toCol = (col == 0) ? "M" : "AA";

                            int startRow = 7 + pageOffset + rowOffset;
                            int endRow = 20 + pageOffset + rowOffset;

                            string cellFrom = $"{fromCol}{startRow}";
                            string cellTo = $"{toCol}{endRow}";

                            //여기

                            using (var inserter = new ImageInserter(ws, files[imageIndex]))
                            {
                                inserter.InsertFit(
                                    cellFrom,
                                    cellTo,
                                    new ImageInsertOptions
                                    {
                                        KeepAspectRatio = false
                                    });

                                wb.Save();
                            }

                            imageIndex++;
                        }
                    }
                }
                AddLog("Info", $"사진 이미지 삽입 완료");

            }

            catch (Exception ex)
            {
                AddLog("Error", $"사진 이미지 삽입 실패: {ex.Message}");

            }
            finally
            {
                try
                {
                    if (ws != null) Marshal.ReleaseComObject(ws);
                }
                catch { }
            }
        }
        #endregion

        #region [PD부분방전 시트 처리]
        private void ProcPdCoronaSheet(object xla, Excel.Workbook wb, string baseFolder, string pdFolder, string xlsPath)
        {
            Excel.Worksheet ws = null;

            try
            {

                string tmpFolder = Path.Combine(baseFolder, pdFolder);

                // PD 부분방전 결과지가 있으므로 PD 부분 방전 처리를 했다는 뜻으로 변환 과정이 실패하더라도 모두 true를 반환하도록 함

                ws = GetWorksheetByName(wb, "PD부분방전");

                if (ws == null)
                {
                    throw new Exception("PD부분방전 시트를 찾을 수 없습니다.");
                }
                
                RemovePictures(ws, 1);


                SetPrintAreaForPDCorona(xlsPath);

                ExportExcelToPdf(xlsPath);

                // xls를 pdf로 변환한 파일을 찾음
                string pdfPath = Directory.GetFiles(tmpFolder, "*.pdf").FirstOrDefault();

                if (pdfPath == null)
                {
                    throw new Exception($"PD 부분방전 엑셀 파일을 pdf로 변환하지 못했습니다.: {tmpFolder}");
                }

                using (ImageInserter inserter = new ImageInserter(ws, pdfPath))
                {
                    ImageInsertOptions option = new ImageInsertOptions
                    {
                        CropLeft = 265,
                        CropTop = 350,
                        CropRight = 260,
                        CropBottom = 300,
                        GapRight = 5,
                        GapLeft = 3
                    };

                    int lastRow = 49;
                    int rowPerPage = 25;
                    int startRow = 25;

                    for (int i = 0; i < inserter.ImageCount; i++)
                    {
                        int index = startRow + i * 25;
                        string cellFrom = $"A{index}";
                        string cellTo = $"I{(rowPerPage-1) + index}";
                        inserter.InsertFit(i,
                            cellFrom,
                            cellTo,
                            option
                            );
                        lastRow = (rowPerPage - 1) + index;
                    }

                    ws.PageSetup.PrintArea = $"$A$1:$I${lastRow}";

                    wb.Save();
                }
            }
            catch (Exception ex)
            {
                AddLog("Error", $"PD 부분방전 이미지 삽입 실패: {ex.Message}");
            }
            finally
            {
                try
                {
                    if (ws != null) Marshal.ReleaseComObject(ws);
                }
                catch { }
            }
        }
        #endregion

        #region [영코 시트 처리]

        private void ProcVideoCoronaSheet(Excel.Application xlApp, Excel.Workbook wb, string baseFolder, object text)
        {
            Excel.Worksheet ws = null;

            try
            {
                string tmpFolder = Path.Combine(baseFolder, "05 영상코로나 또는 부분방전");
                string pdfPath = Directory.GetFiles(tmpFolder, "*.pdf").FirstOrDefault();

                if (pdfPath == null)
                {
                    throw new FileNotFoundException("PDF 파일을 찾을 수 없습니다.", tmpFolder);
                }

                ws = GetWorksheetByName(wb, "영코");

                if (ws == null)
                {
                    throw new Exception("영코 시트를 찾을 수 없습니다.");
                }

                Excel.Range rng = ws.Range["A1:I1"];
                RemovePicturesInRange(ws, rng, false);

                using (ImageInserter inserter = new ImageInserter(ws, pdfPath))
                {
                    ImageInsertOptions option = new ImageInsertOptions
                    {
                        CropLeft = 175,
                        CropTop = 195,
                        CropRight = 1200,
                        CropBottom = 195
                    };

                    inserter.InsertFit(0,
                        "A5",
                        "I24"
                        );

                    for (int i = 1; i < inserter.ImageCount; i++)
                    {
                        string cellFrom = $"A{25 + (i - 1) * 25}";
                        string cellTo = $"I{49 + (i - 1) * 25}";
                        inserter.InsertFit(i,
                            cellFrom,
                            cellTo
                            );
                    }

                    ws.PageSetup.LeftMargin = xlApp.CentimetersToPoints(2.06);
                    wb.Save();
                }
            }
            catch (Exception ex)
            {
                AddLog("Error", $"영상코로나 이미지 삽입 실패: {ex.Message}");
            }
            finally
            {
                try
                {
                    if (ws != null) Marshal.ReleaseComObject(ws);
                }
                catch { }
            }
        }
        #endregion

        #region [품질 시트 처리]
        private void ProcQuantitySheet(Excel.Application xlApp, Excel.Workbook wb, string baseFolder, string quantityFolder, 
            string testReport, string timeGraph, string HighGraph)
        {
            Excel.Worksheet ws = null;
            string pdfPath = null;

            try
            {
                AddLog("Info", $"품질 시트 처리 시작");
                ws = GetWorksheetByName(wb, "품질");

                if (ws == null)
                {
                    throw new Exception("품질 시트를 찾을 수 없습니다.");
                }

                RemovePictures(ws);

                pdfPath = Path.Combine(baseFolder, quantityFolder, "K.pdf");
                using (ImageInserter inserter = new ImageInserter(ws, pdfPath))
                {
                    inserter.Insert("U6", 0.8);

                    //wb.Save();
                }
                AddLog("Info", $"K.pdf 처리 완료");

                pdfPath = Path.Combine(baseFolder, quantityFolder, testReport);
                using (ImageInserter inserter = new ImageInserter(ws, pdfPath))
                {
                    inserter.InsertFit(
                        "A36",
                        "R67",
                        new ImageInsertOptions
                        {
                            CropLeft = 195,
                            CropTop = 170,
                            CropRight = 210,
                            CropBottom = 1000,
                            GapTop = 0,
                            GapBottom = 10.0f
                        });

                    //wb.Save();
                }

                AddLog("Info", $"시험보고서 삽입 완료");

                pdfPath = Path.Combine(baseFolder, quantityFolder, "E.pdf");
                using (ImageInserter inserter = new ImageInserter(ws, pdfPath))
                {

                    ImageInsertOptions option = new ImageInsertOptions
                    {
                        //CropLeft = 135,
                        //CropTop = 55,
                        //CropRight = 105,
                        //CropBottom = 146
                        GapTop = 0f,
                        GapBottom = 10.0f
                    };

                    // 1페이지
                    inserter.InsertFit(0, "A68", "R100", option);

                    // 2~5페이지
                    string[,] group1 =
                    {
                                { "A101", "G116" },   // 2
                                { "H101", "R116" },   // 3
                                { "A117", "G132" },   // 4
                                { "H117", "R132" }    // 5
                            };

                    // 6~9페이지
                    string[,] group2 =
                              {
                                { "A133", "G148" },   // 6
                                { "H133", "R148" },   // 7
                                { "A149", "G164" },   // 8
                                { "H149", "R164" }    // 9
                            };

                    // 10~13페이지
                    string[,] group3 =
                    {
                                { "A165", "G180" },   //10
                                { "H165", "R180" },   //11
                                { "A181", "G196" },   //12
                                { "H181", "R196" }    //13
                            };

                    // 14~17페이지
                    string[,] group4 =
                    {
                                { "A197", "G212" },   //14
                                { "H197", "R212" },   //15
                                { "A213", "G228" },   //16
                                { "H213", "R228" }    //17
                            };

                    int page = 1;

                    option.GapLeft = 2.8f;
                    option.GapTop = 2.8f;
                    option.GapRight = 2.8f;
                    option.GapBottom = 2.8f;

                    foreach (var group in new[] { group1, group2, group3, group4 })
                    {
                        for (int i = 0; i < 4 && page < 17; i++, page++)
                        {
                            inserter.InsertFit(
                                page,
                                group[i, 0],
                                group[i, 1],
                                option);
                        }
                    }

                    //wb.Save();
                }


                AddLog("Info", $"E.pdf 처리 완료");


                pdfPath = Path.Combine(baseFolder, quantityFolder, timeGraph);

                using (ImageInserter inserter = new ImageInserter(ws, pdfPath))
                {

                    ImageInsertOptions option = new ImageInsertOptions
                    {
                        CropLeft = 150,
                        CropTop = 145,
                        CropRight = 1330,
                        CropBottom = 150,
                        GapLeft = 5
                    };

                    // 1페이지
                    inserter.InsertFit("A230", "R261", option);
                    //inserter.Insert("A228", 1.0);

                    //wb.Save();
                }

                AddLog("Info", $"시계열그래프 처리 완료");

                pdfPath = Path.Combine(baseFolder, quantityFolder, HighGraph);

                using (ImageInserter inserter = new ImageInserter(ws, pdfPath))
                {

                    ImageInsertOptions option = new ImageInsertOptions
                    {
                        CropLeft = 150,
                        CropTop = 145,
                        CropRight = 1130,
                        CropBottom = 145,
                        GapLeft = 5
                    };

                    // 1페이지
                    inserter.InsertFit("A263", "R294", option);
                    //inserter.Insert("A262", 1.0);

                    // wb.Save();
                }

                AddLog("Info", $"고조파그래프 처리 완료");

                string[] files = Directory.GetFiles(baseFolder + $"\\{quantityFolder}", "*.bmp")
                    .OrderBy(f =>
                    {
                        return int.TryParse(Path.GetFileNameWithoutExtension(f), out int n)
                            ? n
                            : int.MaxValue;
                    })
                    .ToArray();

                int startRow = 295;
                int blockHeight = 8;   // 293~300 = 8행

                List<int> rows = new List<int> { 0, 2, 1, 7, 3, 5, 4, 6 };

                for (int i = 0; i < files.Length && i < 8; i++)
                {
                    int row = i / 2;    // 0~3
                    int col = i % 2;    // 0=좌, 1=우

                    int fromRow = startRow + row * blockHeight;
                    int toRow = fromRow + blockHeight - 1;

                    string fromCell = col == 0
                        ? $"A{fromRow}"
                        : $"H{fromRow}";

                    string toCell = col == 0
                        ? $"G{toRow}"
                        : $"R{toRow}";

                    using (var inserter = new ImageInserter(ws, files[rows[i]]))
                    {
                        inserter.InsertFit(
                            fromCell,
                            toCell,
                            new ImageInsertOptions
                            {
                                KeepAspectRatio = false
                            });
                    }
                }

                AddLog("Info", $"캡처이미지 처리 완료");

                wb.Save();
            }

            catch (Exception ex)
            {
                AddLog("Error", $"품질 이미지 삽입 실패: {ex.Message}");
            }
            finally
            {
                try
                {
                    AddLog("Info", $"품질 이미지 삽입 완료");
                    if (ws != null) Marshal.ReleaseComObject(ws);
                    //if (wb != null)
                    //{
                    //    wb.Close(false);
                    //    Marshal.ReleaseComObject(wb);
                    //}
                }
                catch
                {
                    AddLog("Error", $"분기 파일 정리 실패");
                }
            }
        }

        #region [열화상 이미지 분기 시트 삽입]

        private void ProcFeverPicture(Excel.Application xlApp, Excel.Workbook wb, string baseFolder, string pictureFolder)
        {
            // 새 통합문서 생성
            Excel.Workbook newWb = null;
            Excel.Worksheet sourceWs = null;
            Excel.Worksheet ws = null;

            try
            {
                sourceWs = GetWorksheetByLastName(wb, "분기");

                if (sourceWs == null)
                {
                    throw new Exception("분기 시트를 찾을 수 없습니다.");
                }

                // 새 통합문서 생성
                newWb = xlApp.Workbooks.Add();

                try
                {
                    // 기본 시트 삭제
                    while (newWb.Worksheets.Count > 1)
                    {
                        ((Excel.Worksheet)newWb.Worksheets[2]).Delete();
                    }

                    // 분기 시트 복사
                    sourceWs.Copy(Before: newWb.Worksheets[1]);

                    // 기본 Sheet1 삭제
                    ((Excel.Worksheet)newWb.Worksheets[newWb.Worksheets.Count]).Delete();

                    // 복사된 시트
                    ws = (Excel.Worksheet)newWb.Worksheets[1];

                    string tempFile = Path.Combine(baseFolder, "images.xlsx");
                    newWb.SaveAs(tempFile);
                }
                catch { }


                RemovePictures(ws);
                RemovePictures(sourceWs);


                string[] files = Directory.GetFiles(baseFolder + "\\" + pictureFolder, "*.jpg")
                .OrderBy(f =>
                {
                    return int.TryParse(Path.GetFileNameWithoutExtension(f), out int n)
                        ? n
                        : int.MaxValue;
                })
                .ToArray();

                int imageIndex = 0;

                for (int page = 0; imageIndex < files.Length; page++)
                {
                    int startRow = 27 + (page * 56);
                    int endRow = 43 + (page * 56);

                    string[] fromCols = { "A", "R" };
                    string[] toCols = { "Q", "AC" };   // 실제 병합 끝 열에 맞게 수정

                    for (int i = 0; i < fromCols.Length; i++)
                    {
                        if (imageIndex >= files.Length)
                            break;

                        string fromCell = $"{fromCols[i]}{startRow}";
                        string toCell = $"{toCols[i]}{endRow}";

                        float gapRight = (i % 2 == 0) ? 2.3f : 2.9f;
                        float gapLeft = (i % 2 == 0) ? 4.2f : 3.6f;

                        float gapBottom = 2.8f;
                        float gapTop = 4.8f;

                        using (var inserter = new ImageInserter(ws, files[imageIndex]))
                        {
                            inserter.InsertFit(
                                fromCell,
                                toCell,
                                new ImageInsertOptions
                                {
                                    KeepAspectRatio = false,
                                    //GapRight = 2.7f,
                                    //GapRight = gapRight,
                                    ////GapBottom = 2.6f,
                                    //GapBottom = gapBottom,
                                    ////GapLeft = 4.0f,
                                    //GapLeft = gapLeft,
                                    //GapTop = gapTop
                                });
                        }

                        using (var inserter = new ImageInserter(sourceWs, files[imageIndex]))
                        {
                            inserter.InsertFit(
                                fromCell,
                                toCell,
                                new ImageInsertOptions
                                {
                                    KeepAspectRatio = false,
                                    //GapRight = 2.7f,
                                    //GapRight = gapRight,
                                    //GapBottom = gapBottom,
                                    ////GapLeft = 4.0f,
                                    //GapLeft = gapLeft,
                                    //GapTop = gapTop
                                });
                        }

                        imageIndex++;
                    }
                }

                if (checkBoxOcr.Checked)
                {

                    using (var reader = new FlirOcrReader())
                    {
                        OcrExcelMap map = new OcrExcelMap
                        {
                            ValueCells = new[]
                            {
                                            "H49",
                                            "P49",
                                            "W49",
                                            "H52",
                                            "P52",
                                            "W52"
                                        },
                            MinTemperatureCell = "AD6",
                            RowOffset = 56
                        };

                        string[] evenFiles = files
                            .Where((file, index) => index % 2 == 0)
                            .ToArray();

                        OcrDataToExcel.ProcessAll(
                            sourceWs,
                            evenFiles,
                            reader,
                            map);

                        OcrDataToExcel.ProcessAll(
                            ws,
                            evenFiles,
                            reader,
                            map);
                    }
                }

                wb.Save();
                newWb.Save();

                // newWb의 첫번째 시트를 wb의 '분기'라는 이름을 포함한 시트 앞에 복사해서 넣음 (복사만, newWb에 남김)
                //InsertFirstSheetFromNewWbBeforeBranch(wb, newWb);
            }

            catch (Exception ex)
            {
                //throw new Exception("분기 이미지 삽입 실패: " + ex.Message, ex);
                AddLog("Error", $"분기 이미지 삽입 실패: {ex.Message}");
            }
            finally
            {
                try
                {
                    AddLog("Info", $"분기 이미지 삽입 완료");
                    if (ws != null) Marshal.ReleaseComObject(ws);
                    if (sourceWs != null) Marshal.ReleaseComObject(sourceWs);   
                    if (newWb != null)
                    {
                        newWb.Close(false);
                        Marshal.ReleaseComObject(newWb);
                    }
                    //if (wb != null)
                    //{
                    //    wb.Close(false);
                    //    Marshal.ReleaseComObject(wb);
                    //}
                }
                catch {
                    AddLog("Error", $"분기 파일 정리 실패");
                }
            }
        }
        #endregion

        #region [엑셀 시트에서 그림 삭제]

        public void RemovePictures(Excel.Worksheet ws)
        {
            for (int i = ws.Shapes.Count; i >= 1; i--)
            {
                Excel.Shape shape = ws.Shapes.Item(i);

                if (shape.Type == Office.MsoShapeType.msoPicture ||
                    shape.Type == Office.MsoShapeType.msoLinkedPicture)
                {
                    shape.Delete();
                }
            }
        }

        public void RemovePictures(Excel.Worksheet ws, int? keepPictureIndex = null)
        {
            var pictureNames = new List<string>();

            // 이미지 이름 수집
            for (int i = 1; i <= ws.Shapes.Count; i++)
            {
                Excel.Shape shape = ws.Shapes.Item(i);

                try
                {
                    if (shape.Type == Office.MsoShapeType.msoPicture ||
                        shape.Type == Office.MsoShapeType.msoLinkedPicture)
                    {
                        pictureNames.Add(shape.Name);
                    }
                }
                finally
                {
                    Marshal.ReleaseComObject(shape);
                }
            }

            string keepName = null;

            if (keepPictureIndex.HasValue &&
                keepPictureIndex.Value >= 1 &&
                keepPictureIndex.Value <= pictureNames.Count)
            {
                keepName = pictureNames[keepPictureIndex.Value - 1];
            }

            // 삭제
            for (int i = ws.Shapes.Count; i >= 1; i--)
            {
                Excel.Shape shape = ws.Shapes.Item(i);

                try
                {
                    if ((shape.Type == Office.MsoShapeType.msoPicture ||
                         shape.Type == Office.MsoShapeType.msoLinkedPicture) &&
                        shape.Name != keepName)
                    {
                        shape.Delete();
                    }
                }
                finally
                {
                    Marshal.ReleaseComObject(shape);
                }
            }
        }

        public void RemovePicturesInRange(
            Excel.Worksheet ws,
            Excel.Range targetRange,
            bool removeInside = true)
        {
            double left = (double)targetRange.Left;
            double top = (double)targetRange.Top;
            double right = left + (double)targetRange.Width;
            double bottom = top + (double)targetRange.Height;

            for (int i = ws.Shapes.Count; i >= 1; i--)
            {
                Excel.Shape shape = ws.Shapes.Item(i);

                if (shape.Type != Office.MsoShapeType.msoPicture &&
                    shape.Type != Office.MsoShapeType.msoLinkedPicture)
                    continue;

                double sLeft = shape.Left;
                double sTop = shape.Top;
                double sRight = sLeft + shape.Width;
                double sBottom = sTop + shape.Height;

                bool overlap =
                    sLeft < right &&
                    sRight > left &&
                    sTop < bottom &&
                    sBottom > top;

                if ((removeInside && overlap) ||
                    (!removeInside && !overlap))
                {
                    shape.Delete();
                }
            }
        }
        #endregion
        #endregion

        #region 갑지 이미지 중앙 정렬
        public void relocatePictures()
        {

            var filePath = tbQuantityFile.Text?.Trim();
            if (string.IsNullOrEmpty(filePath) || !File.Exists(filePath))
            {
                MessageBox.Show("페이지 번호를 매길 엑셀 파일을 먼저 선택하세요.", "오류", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            Cursor = Cursors.WaitCursor;
            Excel.Application xlApp = null;
            Excel.Workbook wb = null;
            Excel.Worksheet ws = null;

            try
            {
                xlApp = new Excel.Application { Visible = false, DisplayAlerts = false };
                // Open for write because we modify PageSetup
                wb = xlApp.Workbooks.Open(filePath, ReadOnly: false);

                ws = GetWorksheetByName(wb, "갑지");

                if (ws == null)
                {
                    AddLog("Error", "갑지 시트를 찾을 수 없습니다.");
                }

                //int quarterCount = GetQuarterCount(filePath);

                Match match = Regex.Match(filePath, @"(\d{2})년(\d)분기");

                if (match.Success)
                {

                    string title =
                    $"{2000 + int.Parse(match.Groups[1].Value)}년 {match.Groups[2].Value}분기" +
                    (filePath.Contains("연차") ? " 연차" : "");

                    ws.Cells[11, 1].Value = title;
                }


                ws.PageSetup.LeftMargin = 28.35;   // 약 1cm
                ws.PageSetup.RightMargin = 28.35;  // 약 1cm
                ws.PageSetup.BottomMargin = 28.35;   // 약 1cm
                ws.PageSetup.TopMargin = 28.35;  // 약 1cm
                ws.PageSetup.CenterHorizontally = true;   // 좌우 가운데
                ws.PageSetup.CenterVertically = true;     // 상하 가운데

                // 인쇄 영역 기준
                Excel.Range printRange = ws.Range[ws.PageSetup.PrintArea];
                if (string.IsNullOrWhiteSpace(ws.PageSetup.PrintArea))
                    printRange = ws.UsedRange;
                else
                    printRange = ws.Range[ws.PageSetup.PrintArea];

                double pageLeft = (double)printRange.Left;
                double pageWidth = (double)printRange.Width;

                // 페이지 중앙
                double centerX = pageLeft + pageWidth / 2;

                double centerY = (double)printRange.Top + (double)printRange.Height / 2;

                Excel.Shape picture = ws.Shapes.Cast<Excel.Shape>()
                    .Where(s => s.Type == Microsoft.Office.Core.MsoShapeType.msoPicture)
                    .OrderBy(s =>
                    {
                        double shapeCenterY = s.Top + s.Height / 2.0;
                        return Math.Abs(shapeCenterY - centerY);
                    })
                    .FirstOrDefault();

                if (picture != null)
                {
                    picture.Left = (float)(centerX - picture.Width / 2);
                }

                // 모서리 둥근 사각형 1개
                Excel.Shape roundRect = ws.Shapes.Cast<Excel.Shape>()
                    .FirstOrDefault(s =>
                        s.Type == Microsoft.Office.Core.MsoShapeType.msoAutoShape &&
                        s.AutoShapeType == Microsoft.Office.Core.MsoAutoShapeType.msoShapeRoundedRectangle);

                if (roundRect != null)
                {
                    //roundRect.Left = (float)(centerX - roundRect.Width / 2);



                    Debug.WriteLine($"PrintArea={ws.PageSetup.PrintArea}");
                    Debug.WriteLine($"Print Left={printRange.Left}");
                    Debug.WriteLine($"Print Width={printRange.Width}");
                    Debug.WriteLine($"CenterX={centerX}");

                    Debug.WriteLine($"Before={roundRect.Left}");
                    roundRect.Left = (float)(centerX - roundRect.Width / 2);
                    Debug.WriteLine($"After={roundRect.Left}");
                }


                if (printRange != null)
                    Marshal.ReleaseComObject(printRange);



                try { wb.Save(); } catch { }
            }
            catch (Exception ex)
            {
                //MessageBox.Show($"갑지 시트 위치 조정 중 오류가 발생했습니다:\r\n{ex.Message}", "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                AddLog("Error", $"갑지 시트 위치 조정 중 오류가 발생했습니다: {ex.Message}");
            }
            finally
            {

                if (ws != null)
                    Marshal.ReleaseComObject(ws);
                try
                {
                    if (wb != null)
                    {
                        wb.Close(true);
                        Marshal.ReleaseComObject(wb);
                    }
                }
                catch { }

                try
                {
                    if (xlApp != null)
                    {
                        xlApp.Quit();
                        Marshal.ReleaseComObject(xlApp);
                    }
                }
                catch { }

                GC.Collect();
                GC.WaitForPendingFinalizers();
                Cursor = Cursors.Default;
            }
        }


        private void btnGapjiPictureRelocate_Click(object sender, EventArgs e)
        {
            relocatePictures();
        }
        #endregion

        #region 엑셀 pdf 내보내기
        private void SetPrintAreaForPDCorona(string xlsFile)
        {
            Excel.Application xlApp = null;
            Excel.Workbook wb = null;
            Excel.Worksheet ws = null;

            try
            {
                xlApp = new Excel.Application
                {
                    Visible = false,
                    DisplayAlerts = false
                };

                wb = xlApp.Workbooks.Open(xlsFile);

                // 첫 번째 시트
                ws = (Excel.Worksheet)wb.Worksheets[1];

                // 상·하·좌·우 여백 0cm
                ws.PageSetup.TopMargin = xlApp.CentimetersToPoints(0);
                ws.PageSetup.BottomMargin = xlApp.CentimetersToPoints(0);
                ws.PageSetup.LeftMargin = xlApp.CentimetersToPoints(0);
                ws.PageSetup.RightMargin = xlApp.CentimetersToPoints(0);

                ws.PageSetup.CenterHorizontally = true;
                ws.PageSetup.CenterVertically = true;

                // 실제 사용 중인 마지막 행
                Excel.Range usedRange = ws.UsedRange;

                int lastUsedRow =
                    usedRange.Row + usedRange.Rows.Count - 1;

                Marshal.ReleaseComObject(usedRange);

                // 50행 단위로 올림
                // 349 → 350
                // 300 → 300
                // 301 → 350
                int lastPrintRow =
                    ((lastUsedRow + 49) / 50) * 50;

                // 인쇄 영역 설정
                ws.PageSetup.PrintArea =
                    $"$A$1:$K${lastPrintRow}";

                // 기존 수동 페이지 나누기 삭제
                for (int i = ws.HPageBreaks.Count; i >= 1; i--)
                {
                    Excel.HPageBreak pb =
                        (Excel.HPageBreak)ws.HPageBreaks[i];

                    try
                    {
                        if (pb.Type == Excel.XlPageBreak.xlPageBreakManual)
                            pb.Delete();
                    }
                    finally
                    {
                        Marshal.ReleaseComObject(pb);
                    }
                }

                // 50행마다 페이지 나누기
                // 350행까지라면 51, 101, 151, 201, 251, 301
                for (int row = 51; row <= lastPrintRow; row += 50)
                {
                    Excel.Range cell = null;

                    try
                    {
                        cell = (Excel.Range)ws.Cells[row, 1];
                        ws.HPageBreaks.Add(cell);
                    }
                    finally
                    {
                        if (cell != null)
                            Marshal.ReleaseComObject(cell);
                    }
                }

                wb.Save();
            }
            finally
            {
                if (ws != null)
                    Marshal.ReleaseComObject(ws);

                if (wb != null)
                {
                    wb.Close(true);
                    Marshal.ReleaseComObject(wb);
                }

                if (xlApp != null)
                {
                    xlApp.Quit();
                    Marshal.ReleaseComObject(xlApp);
                }

                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }
        //private void SetPrintAreaForPDCorona(string xlsFile)
        //{
        //    Excel.Application xlApp = null;
        //    Excel.Workbook wb = null;
        //    Excel.Worksheet ws = null;

        //    try
        //    {
        //        xlApp = new Excel.Application
        //        {
        //            Visible = false,
        //            DisplayAlerts = false
        //        };

        //        wb = xlApp.Workbooks.Open(xlsFile);

        //        // 첫 번째 시트
        //        ws = (Excel.Worksheet)wb.Worksheets[1];

        //        // 상·하 여백 0cm
        //        ws.PageSetup.TopMargin = xlApp.CentimetersToPoints(0);
        //        ws.PageSetup.BottomMargin = xlApp.CentimetersToPoints(0);
        //        ws.PageSetup.LeftMargin = xlApp.CentimetersToPoints(0);
        //        ws.PageSetup.RightMargin = xlApp.CentimetersToPoints(0);
        //        ws.PageSetup.CenterHorizontally = true;   // 좌우 가운데
        //        ws.PageSetup.CenterVertically = true;     // 상하 가운데

        //        // 인쇄 영역 설정
        //        ws.PageSetup.PrintArea = "$A$1:$K$300";

        //        // 기존 수동 페이지 나누기 삭제
        //        for (int i = ws.HPageBreaks.Count; i >= 1; i--)
        //        {
        //            Excel.HPageBreak pb = (Excel.HPageBreak)ws.HPageBreaks[i];

        //            if (pb.Type == Excel.XlPageBreak.xlPageBreakManual)
        //            {
        //                pb.Delete();
        //            }

        //            Marshal.ReleaseComObject(pb);
        //        }

        //        // 50행마다 페이지 나누기 추가
        //        for (int row = 51; row <= 251; row += 50)
        //        {
        //            ws.HPageBreaks.Add((Excel.Range)ws.Cells[row, 1]);
        //        }

        //        wb.Save();
        //    }
        //    finally
        //    {
        //        if (ws != null) Marshal.ReleaseComObject(ws);

        //        if (wb != null)
        //        {
        //            wb.Close(true);
        //            Marshal.ReleaseComObject(wb);
        //        }

        //        if (xlApp != null)
        //        {
        //            xlApp.Quit();
        //            Marshal.ReleaseComObject(xlApp);
        //        }

        //        GC.Collect();
        //        GC.WaitForPendingFinalizers();
        //    }
        //}

        //private string ExportExcelToPdf(string xlsPath)
        //{
        //    var pdfPath = Path.ChangeExtension(xlsPath, ".pdf");

        //    Excel.Application app = null;
        //    Excel.Workbook wb = null;

        //    try
        //    {
        //        app = new Excel.Application();
        //        app.DisplayAlerts = false;

        //        wb = app.Workbooks.Open(xlsPath);

        //        foreach (Excel.Worksheet ws in wb.Worksheets)
        //        {
        //            try
        //            {
        //                // 기존 인쇄영역 제거
        //                ws.PageSetup.PrintArea = "";
        //            }
        //            finally
        //            {
        //                Marshal.ReleaseComObject(ws);
        //            }
        //        }

        //        wb.ExportAsFixedFormat(
        //            Excel.XlFixedFormatType.xlTypePDF,
        //            pdfPath);

        //        return pdfPath;
        //    }
        //    finally
        //    {
        //        if (wb != null)
        //        {
        //            wb.Close(false);
        //            Marshal.ReleaseComObject(wb);
        //        }

        //        if (app != null)
        //        {
        //            app.Quit();
        //            Marshal.ReleaseComObject(app);
        //        }

        //        GC.Collect();
        //        GC.WaitForPendingFinalizers();
        //    }
        //}

        private string ExportExcelToPdf(string xlsPath)
        {
            var pdfPath = Path.ChangeExtension(xlsPath, ".pdf");

            Excel.Application app = null;
            Excel.Workbook wb = null;

            try
            {
                app = new Excel.Application();
                wb = app.Workbooks.Open(xlsPath);

                wb.ExportAsFixedFormat(
                    Excel.XlFixedFormatType.xlTypePDF,
                    pdfPath);

                return pdfPath;
            }
            finally
            {
                if (wb != null)
                {
                    wb.Close(false);
                    Marshal.ReleaseComObject(wb);
                }

                if (app != null)
                {
                    app.Quit();
                    Marshal.ReleaseComObject(app);
                }

                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }
        #endregion

        #region [절연페이지 20000을 2000으로 변경]
        #endregion

        #region [측정자 앞의 공백이 측정일의 공백 개수와 동일하도록 변경]
        #endregion

        #region [절연의 점검요 항목 의견에 옮기기]
        #endregion

        #region [점검 분기에 따른 시트 복사 및 정리]
        #endregion

        #region [서버에서 자동으로 파일 다운로드하기]
        #endregion

        #region [로그 넣기]
        private void WriteLog(string log)
        {
            if (soborLog == null) {
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
                    }
                    catch { }

                    richTextBox1.ScrollToCaret();
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

        #region [서버에서 자동으로 파일 다운로드하기]
        private void MoveFoldersToRoot()
        {
            string sourceFolder = @"D:\work\Report\0now";
            string targetFolder = @"D:\work\Report";

            try
            {
                // 하위 폴더들
                string[] folders = Directory.GetDirectories(sourceFolder);

                foreach (string folder in folders)
                {
                    string folderName = Path.GetFileName(folder);
                    string targetPath = Path.Combine(targetFolder, folderName);

                    if (Directory.Exists(targetPath))
                    {
                        AddLog("WARN", $"이미 존재하는 폴더: {targetPath}");
                        continue;
                    }

                    // 폴더 전체 이동
                    Directory.Move(folder, targetPath);

                    AddLog("Info", $"폴더 이동: {folder} → {targetPath}");
                }
            }
            catch (Exception ex)
            {
                AddLog("Error", $"폴더 이동 중 오류: {ex.Message}");
            }
        }

        private async void btnDownloadWork_Click(object sender, EventArgs e)
        {
            string keyword = textBoxKeyword.Text;
            List<SearchFolderOption> SearchFolders = new List<SearchFolderOption>();

            MoveFoldersToRoot();

            string downloadFolderSeasonReport = "/2_1전기직무고시점검보고서/2 본부(김희철)/0001 시흥보고서방/001열화상";
            string downloadFolderAnuualReport = "/2_1전기직무고시점검보고서/2 본부(김희철)/0001 시흥보고서방";

            SearchFolders = new List<SearchFolderOption>
            {
                new SearchFolderOption
                {
                    Folder = downloadFolderSeasonReport,
                    Recursive = false
                },
                new SearchFolderOption
                {
                    Folder = downloadFolderAnuualReport,
                    Recursive = false
                }
            };

            if (string.IsNullOrWhiteSpace(keyword))
            {
                AddLog("Error", "검색어를 입력하세요.");
                return;
            }

            await TestSynologyAsync(SearchFolders, keyword);
        }
        #endregion



        // 할 일
        // 제출문 시트의 A20 셀을 2025년 7월 ==> 26년 6월로 변경하는 예시
        // 제출문 시트 B18 셀 "첨부 별지서식 : 2~8,코로나방전,축전지"
        // AD17 "2~8,코로나방전,축전지 "
        // AD18 "7"
        // AD19 "2접지,6,7"

        // 연계획 시트의 A2 2025년을 2026으로 변경
        // 측정일 바꾸기
        #region [측정일 바꾸기]
        #endregion

        // 분기 바꾸기
        // 연계획 시트의 D24부터 열을 하나씩 올라가며 해당월까지 "●" 이 문자의 개수가 분기 수
        #region [분기 바꾸기]
        // 연계획 시트의 D24부터 해당 월 컬럼까지 "●" 개수 = 분기 수

        private int GetMonthFromFileName(string filePath)
        {
            string fileName = Path.GetFileNameWithoutExtension(filePath);

            Match m = Regex.Match(fileName, @"_(\d{6})$");

            if (!m.Success)
                throw new Exception("파일명에서 날짜를 찾을 수 없습니다.");

            string yymmdd = m.Groups[1].Value;

            return int.Parse(yymmdd.Substring(2, 2)); // MM
        }

        int GetQuarterCount(string filePath)
        {
            if (string.IsNullOrEmpty(filePath) || !File.Exists(filePath))
            {
                MessageBox.Show("페이지 번호를 매길 엑셀 파일을 먼저 선택하세요.", "오류", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return 0;
            }

            int targetMonth = GetMonthFromFileName(filePath);

            Cursor = Cursors.WaitCursor;
            Excel.Application xlApp = null;
            Excel.Workbook wb = null;
            Excel.Worksheet ws = null;

            try
            {
                xlApp = new Excel.Application { Visible = false, DisplayAlerts = false };
                // Open for write because we modify PageSetup
                wb = xlApp.Workbooks.Open(filePath, ReadOnly: false);

                ws = GetWorksheetByName(wb, "연계획");

                int count = 0;

                // D열 = 4, 1월이라고 가정
                for (int month = 1; month <= targetMonth; month++)
                {
                    int col = 3 + month; // D=4

                    var value = ws.Cells[24, col].Text.Trim();

                    if (value == "●")
                        count++;
                }

                return count;
            }
            catch (Exception ex)
            {
                AddLog("Error", $"분기 수 계산 중 오류 발생: {ex.Message}");
                return 0;
            }
            finally
            {
                if (ws != null) Marshal.ReleaseComObject(ws);
                if (wb != null)
                {
                    wb.Close(false);
                    Marshal.ReleaseComObject(wb);
                }
                if (xlApp != null)
                {
                    xlApp.Quit();
                    Marshal.ReleaseComObject(xlApp);
                }
                GC.Collect();
                GC.WaitForPendingFinalizers();
                Cursor = Cursors.Default;
            }
        }
        #endregion

        // Sheet1 제거 (기본 생성)
        #region [측정자 셀 변경]
        private void UpdateInspectorCell(Excel.Workbook wb, string inspectorName)
        {
            // 기존 단일 셀 호출을 유지하도록 wrapper로 구현
            var list = new List<KeyValuePair<string, string>>()
            {
                // sheetName이 null 또는 빈 문자열이면 ActiveSheet를 의미합니다.
                new KeyValuePair<string,string>("절연", "A4"),
                new KeyValuePair<string,string>("저압", "A3"),
                new KeyValuePair<string,string>("예비", "A3"),
                new KeyValuePair<string,string>("분기", "A3")
            };

            UpdateInspectorCells(wb, inspectorName, list);
            UpdateMachine(wb);
        }

        private void UpdateMachine(Excel.Workbook wb)
        {
            if (wb == null) return;

            foreach (Excel.Worksheet ws in wb.Worksheets)
            {
                Excel.Range found = ws.Cells.Find(
                    What: "▣ 측정장비",
                    LookAt: Excel.XlLookAt.xlPart,
                    LookIn: Excel.XlFindLookIn.xlValues);

                if (found != null && found.Column == 1)
                {
                    string text = found.Value2?.ToString() ?? "";
                    found.Value2 = " " + text.TrimStart();
                }
            }
        }

        // 시트 이름과 셀 주소 리스트로 여러 시트의 동일한 셀을 한 번에 업데이트합니다.
        // sheetCellList: (sheetName, cellAddress) 쌍의 열거. sheetName이 null/빈 문자열이면 ActiveSheet를 사용합니다.
        private void UpdateInspectorCells(Excel.Workbook wb, string inspectorName, IEnumerable<KeyValuePair<string,string>> sheetCellList)
        {
            if (!checkBoxChecker.Checked) return;
            if (string.IsNullOrEmpty(inspectorName))
            {
                MessageBox.Show("측정자 이름이 null입니다. 측정자 이름을 입력하세요.",
                    "오류", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (wb == null) return;

            var pairs = sheetCellList.ToList();

            // 절연 시트가 있으면 참조용 시트 이름과 셀 주소를 가져온다
            bool isJeolyeonSheet = false;

            foreach (var pair in pairs)
            {
                string sheetName = pair.Key;
                string cellAddress = pair.Value;

                Excel.Worksheet ws = null;
                try
                {
                    if (string.IsNullOrEmpty(sheetName))
                    {
                        // ActiveSheet 사용
                        ws = (Excel.Worksheet)wb.ActiveSheet;
                    }
                    else
                    {
                        // 시트가 존재하면 가져오고, 없으면 건너뜀
                        try
                        {
                            ws = GetWorksheetByName(wb, sheetName);
                        }
                        catch
                        {
                            // 시트 없음
                            ws = null;
                        }
                    }

                    if (ws == null) continue;

                    if (ws.Name == "절연") isJeolyeonSheet = true;

                    if (!string.IsNullOrEmpty(cellAddress))
                    {
                        try
                        {
                            // 절연 시트가 존재하고 현재 타겟이 저압 또는 예비이면 절연 시트의 셀을 참조하는 수식으로 설정
                            if (isJeolyeonSheet && ws.Name != "절연")
                            {
                                string formula = "='" + "절연" + "'!" + "A4";
                                ws.Range[cellAddress].Formula = formula;
                            }
                            else
                            {
                                ws.Range[cellAddress].Value = inspectorName;
                            }
                        }
                        catch (Exception ex)
                        {
                            Debug.WriteLine($"UpdateInspectorCells: failed to set {sheetName ?? "(Active)"}!{cellAddress}: {ex.Message}");
                        }
                    }
                }
                finally
                {
                    if (ws != null) try { Marshal.ReleaseComObject(ws); } catch { }
                }
            }
        }

        private void tbQuantityFile_TextChanged(object sender, EventArgs e)
        {
            string filePath = tbQuantityFile.Text?.Trim();

            if (string.IsNullOrEmpty(filePath) || !File.Exists(filePath))
            {
                return;
            }

            //세정물류센터_26년3분기연차보고서_260711.xlsx

            // Report.ParseReport 호출하여 Report 인스턴스 생성
            try
            {
                this.report = Report.ParseReport(filePath, soborLog);
                
            }
            catch (Exception ex)
            {
                AddLog("Error", $"보고서 파싱 중 오류: {ex.Message}");
            }


        }

        #region [보고서 이미지 셀 안에 정렬]
        private void btnSnapImage_Click(object sender, EventArgs e)
        {
            if (report == null)
            {
                AddLog("Error", "보고서 정보가 없습니다. 엑셀 파일을 먼저 선택하세요.");
                return;
            }
            if (string.IsNullOrEmpty(textBoxSheetForSnapImage.Text))
            {
                AddLog("Error", "이미지 셀에 정렬할 시트 이름을 입력하세요.");
                return;
            }

            Cursor = Cursors.WaitCursor;
            float gapLeft = 1.5f, gapRight = 1.5f, gapTop = 0f, gapBottom = 0.5f;
            float.TryParse(textBoxImageAlignLeftGap.Text.Trim(), out gapLeft);
            float.TryParse(textBoxImageAlignTopGap.Text.Trim(), out gapTop);
            float.TryParse(textBoxImageAlignRightGap.Text.Trim(), out gapRight);
            float.TryParse(textBoxImageAlignBottomGap.Text.Trim(), out gapBottom);
            report.CopySheetToXlsxAndProcess(textBoxSheetForSnapImage.Text, gapLeft, gapTop, gapRight, gapBottom);
            Cursor = Cursors.Default;
        }
        #endregion

        private void btnChangeInspector_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(tbInspector.Text.Trim()))
            {
                AddLog("Error", "담당자 이름을 입력하세요.");
                return;
            }

            Cursor = Cursors.WaitCursor;
            var filePath = tbQuantityFile.Text?.Trim();
            if (string.IsNullOrEmpty(filePath) || !File.Exists(filePath))
            {
                MessageBox.Show("측정자를 입력할 엑셀 파일을 먼저 선택하세요.", "오류", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            Cursor = Cursors.WaitCursor;
            Excel.Application xlApp = null;
            Excel.Workbook wb = null;

            try
            {
                xlApp = new Excel.Application { Visible = false, DisplayAlerts = false };
                // Open for write because we modify PageSetup
                wb = xlApp.Workbooks.Open(filePath, ReadOnly: false);
                UpdateInspectorCell(wb, " ▣ 측정자 : ㈜한경이엔지 - " + tbInspector.Text.Trim());

                wb.Save();
            } catch (Exception ex) {
                AddLog("Error", $"측정자 변경 중 오류: {ex.Message}");
            }
            finally
            {
                if (wb != null)
                {
                    wb.Close(false);
                    Marshal.ReleaseComObject(wb);
                }
                if (xlApp != null)
                {
                    xlApp.Quit();
                    Marshal.ReleaseComObject(xlApp);
                }
                GC.Collect();
                GC.WaitForPendingFinalizers();
                Cursor = Cursors.Default;
            }
        }
        #endregion

        // 검교정 P1 (현장명) 바꾸기
        // 의견 시트 별지서식 변경
        // A5 셀 "  ○ 전기안전관리자 직무고시 점검 : 별지서식 [2,3,4,5,6,7,8,코로나방전,축전지] 점검."



        // 갑지 시트의 A11 셀 "2026년 1분기"

        // 검교정 P1 (현장명) 바꾸기
        // 의견 시트 별지서식 변경
        // A5 셀 "  ○ 전기안전관리자 직무고시 점검 : 별지서식 [2,3,4,5,6,7,8,코로나방전,축전지] 점검."


        // 서버에 유랑진 있는 폴더 가져오기

        #region [보고서 업로드]
        // 서버에 보고서 파일 업로드하기
        // 현재 조건은 해당 폴더에서 오늘 날짜의 보고서 xls와 pdf를 올림

        private async Task UploadReportsToSynoAsync()
        {
            string downloadFolderSeasonReport = "/2_1전기직무고시점검보고서/2 본부(김희철)/0001 시흥보고서방/001열화상";
            string downloadFolderAnuualReport = "/2_1전기직무고시점검보고서/2 본부(김희철)/0001 시흥보고서방";

            // 구성: DSM 접속 정보
            var config = new SynologyFileDownloaderConfig
            {
                BaseUrl = "http://hk2ng.synology.me:5000",
                Username = "유량진",
                Password = "HKeng717241",
                // 검색 시작 폴더 (상위 루트). 필요시 변경하세요.

                SearchFolders = new List<SearchFolderOption>
                {
                    new SearchFolderOption
                    {
                        Folder = downloadFolderSeasonReport,
                        Recursive = false
                    },
                    new SearchFolderOption
                    {
                        Folder = downloadFolderAnuualReport,
                        Recursive = false
                    }
                }
            };

            string localRoot = @"D:\work\Report\0now";

            using (var downloader = new SynologyIntegration.SynologyFileDownloader(config))
            {
                await downloader.LoginAsync();
                Debug.WriteLine($"login sid = {downloader.Sid}");
                using (var uploader = new SynologyFileUploader(
                        config,
                        downloader.HttpClient,
                        downloader.Sid))
                {
                    //await uploader.LoginAsync();

                    // 원격 폴더 목록 조회
                    var remoteFolders = await downloader.SearchFoldersAsync();

                    foreach (string localFolder in Directory.GetDirectories(localRoot))
                    {
                        string folderName = Path.GetFileName(localFolder);
                        AddLog("INFO", "처리 폴더: " + folderName);
                        folderName = folderName.Trim();

                        bool ok = await uploader.CreateFolderAsync(folderName, "test_upload");

                        Debug.WriteLine($"CreateFolder={ok}");

                        // 원격에서 동일한 이름의 폴더 찾기
                        var match = remoteFolders.FirstOrDefault(f => string.Equals(f.Name, folderName, StringComparison.OrdinalIgnoreCase));
                        if (match == null)
                        {
                            match = remoteFolders.FirstOrDefault(f =>
                                    f.Name.Trim().StartsWith(folderName, StringComparison.OrdinalIgnoreCase));
                            // 대체 매칭: 포함 또는 끝부분
                        }


                        if (match == null)
                        {
                            AddLog("WARN", $"원격에서 폴더를 찾지 못함: {folderName} (스킵)");
                            continue;
                        }

                        string remoteBase = match.Path; // 예: /.../folderName

                        bool isAnnual = folderName.IndexOf("연차", StringComparison.OrdinalIgnoreCase) >= 0;

                        if (isAnnual)
                        {
                            string reportSub = Path.Combine(localFolder, "04 보고서");
                            if (!Directory.Exists(reportSub))
                            {
                                AddLog("WARN", $"연차 폴더지만 '04 보고서'를 찾을 수 없음: {reportSub} (스킵)");
                                continue;
                            }

                            // 원격에 '04 보고서' 폴더 생성(없으면)
                            string remoteReportPath = remoteBase.TrimEnd('/') + "/04 보고서";
                            try
                            {
                                await uploader.CreateFolderAsync(remoteBase, "04 보고서");
                            }
                            catch { }

                            var files = Directory.GetFiles(reportSub);
                            foreach (var f in files)
                            {
                                try
                                {
                                    AddLog("INFO", $"업로드: {f} -> {remoteReportPath}");
                                    await uploader.UploadFileAsync(f, remoteReportPath);
                                    AddLog("INFO", $"업로드 성공: {Path.GetFileName(f)}");
                                }
                                catch (Exception ex)
                                {
                                    AddLog("ERROR", $"업로드 실패: {f} -> {remoteReportPath} : {ex.Message}");
                                }
                            }
                        }
                        else
                        {
                            var files = Directory.GetFiles(localFolder)
                                .Where(p => p.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase) || p.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase));

                            foreach (var f in files)
                            {
                                string nameNoExt = Path.GetFileNameWithoutExtension(f);
                                if (nameNoExt.Length < 7) continue;

                                string last6 = nameNoExt.Substring(nameNoExt.Length - 6);
                                if (!System.Text.RegularExpressions.Regex.IsMatch(last6, "^\\d{6}$"))
                                    continue;

                                if (!DateTime.TryParseExact(last6, "yyMMdd", null, System.Globalization.DateTimeStyles.None, out DateTime dt))
                                    continue;

                                if (dt.Year != DateTime.Now.Year)
                                    continue; // 올해 파일만 업로드

                                try
                                {
                                    AddLog("INFO", $"업로드: {f} -> {remoteBase}");
                                    await uploader.UploadFileAsync(f, remoteBase);
                                    AddLog("INFO", $"업로드 성공: {Path.GetFileName(f)}");
                                }
                                catch (Exception ex)
                                {
                                    AddLog("ERROR", $"업로드 실패: {f} -> {remoteBase} : {ex.Message}");
                                }
                            }
                        }
                    }

                    try { await downloader.LogoutAsync(); } catch { }
                }
            }
        }

        private async void btnUploadReport_Click(object sender, EventArgs e)
        {
            // DSM을 통해 업로드 실행
            try
            {
                Cursor = Cursors.WaitCursor;
                await UploadReportsToSynoAsync();
                MessageBox.Show("업로드 작업 완료");
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }
        #endregion

        #region [이미지 용량 줄이기]
        private void btnCompressImages_Click(object sender, EventArgs e)
        {
            var filePath = tbQuantityFile.Text?.Trim();
            if (string.IsNullOrEmpty(filePath) || !File.Exists(filePath))
            {
                MessageBox.Show("엑셀 파일을 먼저 선택하세요.", "오류", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (textBoxSheetForSnapImage.Text.Trim() == "")
            {
                MessageBox.Show("이미지 용량을 줄일 시트 이름을 입력하세요.", "오류", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            Cursor = Cursors.WaitCursor;
            Excel.Application xlApp = null;
            Excel.Workbook wb = null;
            Excel.Worksheet ws = null;

            try
            {
                xlApp = new Excel.Application { Visible = false, DisplayAlerts = false };
                //wb = xlApp.Workbooks.Open(filePath, ReadOnly: false);
                wb = xlApp.Workbooks.Open(filePath);


                string baseFolder = Path.GetDirectoryName(filePath);

                ws = GetWorksheetByName(wb, textBoxSheetForSnapImage.Text.Trim());

                if (ws == null)
                {
                    throw new Exception($"{textBoxSheetForSnapImage.Text.Trim()} 시트를 찾을 수 없습니다.");
                }

                CompressSheetImages(ws, 1);
                //CompressMediaImages(filePath);

            }

            catch (Exception ex)
            {
                AddLog("Error", $"사진 용량 줄이기 실패: {ex.Message}");

            }
            finally
            {
                try
                {
                    if (ws != null) Marshal.ReleaseComObject(ws);
                    if (wb != null)
                    {
                        wb.Save();
                        wb.Close(false);
                        Marshal.ReleaseComObject(wb);
                    }

                    if (xlApp != null)
                    {
                        xlApp.Quit();
                        Marshal.ReleaseComObject(xlApp);
                    }

                    GC.Collect();
                    GC.WaitForPendingFinalizers();

                    Cursor = Cursors.Default;
                }
                catch { }
            }
        }

        private void CompressMediaImages(
            string xlsxPath,
            double scale = 0.5)
        {
            string tempXlsx = xlsxPath + ".tmp";


            using (var archive =
                ZipFile.Open(
                    xlsxPath,
                    ZipArchiveMode.Read))
            {
                using (var newArchive =
                    ZipFile.Open(
                        tempXlsx,
                        ZipArchiveMode.Create))
                {
                    foreach (var entry in archive.Entries)
                    {
                        // 이미지 파일만 처리
                        if (entry.FullName.StartsWith("xl/media/") &&
                           (entry.Name.EndsWith(".png") ||
                            entry.Name.EndsWith(".jpg") ||
                            entry.Name.EndsWith(".jpeg")))
                        {
                            using (var stream = entry.Open())
                            using (var img = Image.FromStream(stream))
                            {
                                int width =
                                    (int)(img.Width * scale);

                                int height =
                                    (int)(img.Height * scale);


                                using (Bitmap bmp =
                                    ResizeBitmap(
                                        new Bitmap(img),
                                        width,
                                        height))
                                {
                                    var newEntry =
                                        newArchive.CreateEntry(
                                            entry.FullName,
                                            CompressionLevel.Optimal);


                                    using (var outStream =
                                        newEntry.Open())
                                    {
                                        bmp.Save(
                                            outStream,
                                            ImageFormat.Jpeg);
                                    }
                                }
                            }
                        }
                        else
                        {
                            // 나머지 xml 등은 그대로 복사
                            var newEntry =
                                newArchive.CreateEntry(
                                    entry.FullName);

                            using (var input =
                                entry.Open())
                            using (var output =
                                newEntry.Open())
                            {
                                input.CopyTo(output);
                            }
                        }
                    }
                }
            }


            File.Delete(xlsxPath);
            File.Move(tempXlsx, xlsxPath);
        }

        private static void ApplyExifOrientation(Image image)
        {
            const int ExifOrientationId = 0x0112;

            if (!image.PropertyIdList.Contains(ExifOrientationId))
                return;

            var prop = image.GetPropertyItem(ExifOrientationId);
            ushort orientation = BitConverter.ToUInt16(prop.Value, 0);

            switch (orientation)
            {
                case 2:
                    image.RotateFlip(RotateFlipType.RotateNoneFlipX);
                    break;
                case 3:
                    image.RotateFlip(RotateFlipType.Rotate180FlipNone);
                    break;
                case 4:
                    image.RotateFlip(RotateFlipType.Rotate180FlipX);
                    break;
                case 5:
                    image.RotateFlip(RotateFlipType.Rotate90FlipX);
                    break;
                case 6:
                    image.RotateFlip(RotateFlipType.Rotate90FlipNone);
                    break;
                case 7:
                    image.RotateFlip(RotateFlipType.Rotate270FlipX);
                    break;
                case 8:
                    image.RotateFlip(RotateFlipType.Rotate270FlipNone);
                    break;
            }

            image.RemovePropertyItem(ExifOrientationId);
        }

        public void CompressSheetImages(
                Excel.Worksheet ws,
                double scale = 1.5)
        {
            var pictures = new System.Collections.Generic.List<Excel.Shape>();

            // Shape 목록 복사 (삭제하면서 순회하면 오류 발생)
            foreach (Excel.Shape shape in ws.Shapes)
            {
                if (shape.Type == Office.MsoShapeType.msoPicture ||
                    shape.Type == Office.MsoShapeType.msoLinkedPicture)
                {
                    pictures.Add(shape);
                }
            }


            foreach (Excel.Shape shape in pictures)
            {
                string tempFile = null;

                try
                {
                    float left = shape.Left;
                    float top = shape.Top;
                    float width = shape.Width;
                    float height = shape.Height;

                    float rotation = shape.Rotation;

                    // Excel point -> pixel
                    int targetWidth = (int)(width * 96 / 72 * scale);
                    int targetHeight = (int)(height * 96 / 72 * scale);

                    var shadowVisible = shape.Shadow.Visible;
                    var placement = shape.Placement;
                    var lockAspect = shape.LockAspectRatio;
                    var z = shape.ZOrderPosition;

                    shape.Shadow.Visible =
                        Office.MsoTriState.msoFalse;

                    // 이미지 복사
                    shape.Copy();

                    Bitmap source = null;

                    // Clipboard 대기
                    for (int i = 0; i < 10; i++)
                    {
                        if (Clipboard.ContainsImage())
                        {
                            source = Clipboard.GetImage() as Bitmap;
                            break;
                        }

                        System.Threading.Thread.Sleep(100);
                    }


                    if (source == null)
                        continue;

                    Console.WriteLine(
                             $"{shape.Name}  {source.Width}x{source.Height}  Rotation={shape.Rotation}");
                    Console.WriteLine("leftTop: " + left + " , " + top);
                    Console.WriteLine(source.Width + " x " + source.Height);
                    Console.WriteLine(rotation);

                    using (source)
                    {
                        ApplyExifOrientation(source);
                        using (Bitmap resized = ResizeBitmap(
                            source,
                            targetWidth,
                            targetHeight))
                        {
                            tempFile = Path.Combine(
                                Path.GetTempPath(),
                                Guid.NewGuid() + ".jpg");

                            var encoder = ImageCodecInfo.GetImageEncoders()
                                .First(x => x.FormatID == ImageFormat.Jpeg.Guid);

                            var param = new EncoderParameters(1);

                            param.Param[0] = new EncoderParameter(
                                Encoder.Quality,
                                80L);


                            resized.Save(
                                tempFile,
                                encoder,
                                param);
                        }
                    }


                    // 기존 그림 삭제
                    shape.Delete();

                    if (rotation == 90 || rotation == 270)
                    {
                        float oldWidth = width;
                        float oldHeight = height;

                        width = oldHeight;
                        height = oldWidth;

                        left += (oldWidth - width) / 2f;
                        top += (oldHeight - height) / 2f;

                    }
                        // 다시 삽입
                    Excel.Shape newShape = ws.Shapes.AddPicture(
                        tempFile,
                        Office.MsoTriState.msoFalse,
                        Office.MsoTriState.msoTrue,
                        left,
                        top,
                        width,
                        height);


                    //newShape.Rotation = rotation;
                    if (rotation == 90 || rotation == 270)
                        newShape.Rotation = 0;
                    else
                        newShape.Rotation = rotation;
                    newShape.Shadow.Visible = shadowVisible;
                    newShape.Placement = placement;
                    newShape.LockAspectRatio = lockAspect;

                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
                finally
                {
                    if (tempFile != null && File.Exists(tempFile))
                    {
                        try
                        {
                            File.Delete(tempFile);
                        }
                        catch { }
                    }
                }
            }
        }

        private Bitmap ResizeBitmap(
            Bitmap src,
            int maxWidth,
            int maxHeight)
        {
            //double ratio = Math.Min(
            //    (double)maxWidth / src.Width,
            //    (double)maxHeight / src.Height);


            //if (ratio >= 1)
            //    return new Bitmap(src);


            //int width = (int)(src.Width * ratio);
            //int height = (int)(src.Height * ratio);


            Bitmap bmp = new Bitmap(maxWidth, maxHeight);

            using (Graphics g = Graphics.FromImage(bmp))
            {
                g.InterpolationMode =
                    System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;

                g.DrawImage(
                    src,
                    0,
                    0,
                    maxWidth,
                    maxHeight);
            }

            return bmp;
        }
        #endregion

        private void tbFolder_TextChanged(object sender, EventArgs e)
        {


            //260703_동탄시범월드반도아파트_연차_유량진
            //260703_당진시네마타워_김희철_분기_유량진

            // 폴더명에 연차가 있으면 새 엑셀 파일 생성
            // 연차가 없으면 

            //if (string.IsNullOrEmpty(tbFolder.Text) || !Directory.Exists(tbFolder.Text))
            //{
            //    AddLog("WARN", "폴더 경로가 유효하지 않습니다.");
            //    return;
            //}

            //try
            //{
            //    var newPath = createNewReportFile(tbFolder.Text);
            //    AddLog("Info", $"새 엑셀 파일을 생성했습니다:\r\n{newPath}");
            //    MessageBox.Show($"새 엑셀 파일을 생성했습니다:\r\n{newPath}", "완료", MessageBoxButtons.OK, MessageBoxIcon.Information);
            //}
            //catch (Exception ex)
            //{
            //    AddLog("ERROR", $"파일 생성 중 오류: {ex.Message}");
            //    MessageBox.Show($"파일 생성 중 오류가 발생했습니다:\r\n{ex.Message}", "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
            //}
        }

        private void tbFolder_Enter(object sender, EventArgs e)
        {
            //if (string.IsNullOrEmpty(tbFolder.Text) || !Directory.Exists(tbFolder.Text))
            //{
            //    AddLog("WARN", "폴더 경로가 유효하지 않습니다.");
            //    return;
            //}

            //try
            //{
            //    var newPath = createNewReportFile(tbFolder.Text);
            //    AddLog("Info", $"새 엑셀 파일을 생성했습니다:\r\n{newPath}");
            //    MessageBox.Show($"새 엑셀 파일을 생성했습니다:\r\n{newPath}", "완료", MessageBoxButtons.OK, MessageBoxIcon.Information);
            //}
            //catch (Exception ex)
            //{
            //    AddLog("ERROR", $"파일 생성 중 오류: {ex.Message}");
            //    MessageBox.Show($"파일 생성 중 오류가 발생했습니다:\r\n{ex.Message}", "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
            //}
        }

        private void tbFolder_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode != Keys.Enter)
                return;

            if (string.IsNullOrEmpty(tbFolder.Text) || !Directory.Exists(tbFolder.Text))
            {
                AddLog("WARN", "폴더 경로가 유효하지 않습니다.");
                return;
            }

            try
            {
                var newPath = createNewReportFile(tbFolder.Text);
                AddLog("Info", $"새 엑셀 파일을 생성했습니다:\r\n{newPath}");
                MessageBox.Show($"새 엑셀 파일을 생성했습니다:\r\n{newPath}", "완료", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                AddLog("ERROR", $"파일 생성 중 오류: {ex.Message}");
                MessageBox.Show($"파일 생성 중 오류가 발생했습니다:\r\n{ex.Message}", "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        #region [저압 (접지저항) 목록 업데이트]
        private void UpdateEquipmentList(Excel.Workbook wb, string filePath)
        {
            Excel.Application app = null;
            bool openedHere = false;
            Excel.Worksheet wsSrc = null;
            Excel.Worksheet wsDst = null;

            try
            {
                if (wb == null)
                {
                    if (!File.Exists(filePath))
                    {
                        AddLog("WARN", "파일이 존재하지 않습니다.");
                        return;
                    }

                    app = new Excel.Application();
                    app.Visible = false;
                    app.DisplayAlerts = false;

                    wb = app.Workbooks.Open(filePath);
                    openedHere = true;
                }

                wsSrc = wb.Worksheets["절연"];
                wsDst = GetWorksheetByName(wb, "저압");

                if (wsSrc == null || wsDst == null)
                    throw new Exception("sheet is not available");

                int lastRow = wsSrc.Cells[wsSrc.Rows.Count, "U"]
                                   .End(Excel.XlDirection.xlUp).Row;

                int row = 7;
                bool isLeft = true;

                int pageStartRow = 7;
                int pageEndRow = 27;

                const int pageHeight = 21; // 7~27
                const int pageGap = 6;     // 28~33
                const int pageStep = pageHeight + pageGap; // 27

                for (int r = 1; r <= lastRow; r++)
                {
                    Excel.Range srcCell = null;

                    try
                    {
                        srcCell = wsSrc.Cells[r, "U"] as Excel.Range;

                        string value = Convert.ToString(srcCell?.Value2)?.Trim();

                        if (string.IsNullOrEmpty(value))
                            continue;

                        string col = isLeft ? "A" : "G";

                        wsDst.Range[$"{col}{row}"].Value2 = value;

                        if (!isLeft)
                        {
                            row++;

                            // 현재 페이지 끝을 넘어감
                            if (row > pageEndRow)
                            {
                                int nextPageStart = pageStartRow + pageStep;
                                int nextPageEnd = nextPageStart + pageHeight - 1;

                                // 다음 페이지가 없으면 현재 페이지를 복사
                                if (!IsPageExists(wsDst, nextPageStart, nextPageEnd))
                                {
                                    CopyLowVoltagePage(
                                        wsDst,
                                        pageStartRow,
                                        pageEndRow,
                                        nextPageStart
                                    );
                                }

                                pageStartRow = nextPageStart;
                                pageEndRow = nextPageEnd;
                                row = pageStartRow;
                            }
                        }

                        isLeft = !isLeft;
                    }
                    finally
                    {
                        if (srcCell != null)
                            Marshal.ReleaseComObject(srcCell);
                    }
                }

                if (openedHere)
                    wb.Save();
            }
            finally
            {
                if (wsSrc != null)
                    Marshal.ReleaseComObject(wsSrc);

                if (wsDst != null)
                    Marshal.ReleaseComObject(wsDst);

                if (openedHere)
                {
                    wb.Close(false);
                    app.Quit();

                    Marshal.ReleaseComObject(wb);
                    Marshal.ReleaseComObject(app);
                }
            }
        }

        private bool IsPageExists(
            Excel.Worksheet ws,
            int startRow,
            int endRow)
        {
            Excel.Range range = null;

            try
            {
                range = ws.Range[$"A{startRow}:L{endRow}"];

                // 영역 안에 값이 하나라도 있으면 페이지가 있다고 판단
                object count = ws.Application.WorksheetFunction.CountA(range);

                return Convert.ToDouble(count) > 0;
            }
            finally
            {
                if (range != null)
                    Marshal.ReleaseComObject(range);
            }
        }

        private void CopyLowVoltagePage(
            Excel.Worksheet ws,
            int sourceStartRow,
            int sourceEndRow,
            int destStartRow)
        {
            Excel.Range src = null;
            Excel.Range dst = null;

            try
            {
                int rowCount = sourceEndRow - sourceStartRow + 1;
                int destEndRow = destStartRow + rowCount - 1;

                src = ws.Range[$"A{sourceStartRow}:L{sourceEndRow}"];
                dst = ws.Range[$"A{destStartRow}:L{destEndRow}"];

                src.Copy(dst);

                AddLog(
                    "Info",
                    $"저압 페이지 추가: {sourceStartRow}~{sourceEndRow} → " +
                    $"{destStartRow}~{destEndRow}"
                );
            }
            finally
            {
                if (src != null)
                    Marshal.ReleaseComObject(src);

                if (dst != null)
                    Marshal.ReleaseComObject(dst);
            }
        }

        private void btnUpdateJuapList_Click(object sender, EventArgs e)
        {
            string filePath= tbQuantityFile.Text.Trim();
            if (string.IsNullOrWhiteSpace(filePath) || !File.Exists(filePath))
            {
                AddLog("WARN", "파일 경로가 올바르지 않습니다.");
                return;
            }
            UpdateEquipmentList(null, filePath);
        }
        #endregion

        #region [절연의 저항과 전류, 결과값 유효성 확인]

        private void SetIncorrectCell(Excel.Range cell, string value, string message)
        {
            if (cell == null) return;
            if (value == null) return;

            if (Convert.ToString(cell.Value2) == value) return;

            if (checkBoxCorrect.Checked)
            {
                if (string.IsNullOrEmpty(value))
                {
                    cell.ClearContents();
                }
                else
                {
                    cell.Value = value;
                }
            }
            else
            {
                cell.Interior.Color = ColorTranslator.ToOle(Color.Yellow);
            }
            AddLog("WARN", message);
        }

        private void CheckResistanceAndCurrentForJulyeon(Excel.Workbook wb, string filePath)
        {
            Excel.Application app = null;
            bool openedHere = false;
            bool bCorrect = checkBoxCorrect.Checked;

            try
            {
                if (wb == null)
                {
                    if (!File.Exists(filePath))
                    {
                        AddLog("WARN", "파일이 존재하지 않습니다.");
                        return;
                    }

                    app = new Excel.Application();
                    app.Visible = false;
                    app.DisplayAlerts = false;

                    wb = app.Workbooks.Open(filePath);
                    openedHere = true;
                }

                Excel.Worksheet wsSrc = wb.Worksheets["절연"];

                int lastRow = wsSrc.Cells[wsSrc.Rows.Count, "U"]
                                   .End(Excel.XlDirection.xlUp).Row;

                ProcessSide("A", "B", "E", "G", "H", "I", "J");   // 좌
                ProcessSide("K", "L", "O", "Q", "R", "S", "T");   // 우

                if (openedHere)
                    wb.Save();

                // -----------------------------
                void ProcessSide(string sideCol,
                                 string noCol,
                                 string limitCol,
                                 string resistCol,
                                 string currentCol,
                                 string resultCol,
                                 string remarkCol)
                {
                    // 1. 측정번호 1번 여부 확인
                    for (int row = 2; row <= lastRow; row++)
                    {
                        Excel.Range rNo = wsSrc.Range[$"{noCol}{row}"];
                        if (Convert.ToBoolean(rNo.MergeCells))
                            continue;

                        string no = Convert.ToString(rNo.Value2)?.Trim();


                        if (no == "1")
                        {
                            if (sideCol == "A")
                            {
                                string rightNo = Convert.ToString(wsSrc.Range[$"L{row}"].Value2)?.Trim();
                                string leftNo = Convert.ToString(wsSrc.Range[$"B{row}"].Value2)?.Trim();

                                bool bRight = !string.IsNullOrEmpty(rightNo);
                                SetIncorrectCell(wsSrc.Range[$"A{row}"], (bRight)?"좌":"",
                                        $"좌측: A{row}에 '{Convert.ToString(wsSrc.Range[$"A{row}"].Value2)}' 대신 '좌' 채움");
                                SetIncorrectCell(wsSrc.Range[$"K{row}"], (bRight) ? "우" : "",
                                        $"우측: K{row}에 '{Convert.ToString(wsSrc.Range[$"K{row}"].Value2)}' 대신 '우' 채움");
                                
                            }
                        }



                        Excel.Range rResist = wsSrc.Range[$"{resistCol}{row}"];
                        Excel.Range rCurrent = wsSrc.Range[$"{currentCol}{row}"];
                        Excel.Range rResult = wsSrc.Range[$"{resultCol}{row}"];
                        Excel.Range rRemark = wsSrc.Range[$"{remarkCol}{row}"];
                        Excel.Range rLimit = wsSrc.Range[$"{limitCol}{row}"];

                        if (string.IsNullOrWhiteSpace(no))
                        {
                            if (Convert.ToString(rResist.Value2) != null
                                || Convert.ToString(rCurrent.Value2) != null
                                || Convert.ToString(rResult.Value2) != null
                                || Convert.ToString(rRemark.Value2) != null
                                || Convert.ToString(rLimit.Value2) != null)    
                            {
                                if (bCorrect)
                                {
                                    rResist?.ClearContents();
                                    rRemark?.ClearContents();
                                    rLimit?.ClearContents();
                                    rResult?.ClearContents();
                                    rLimit?.ClearContents();
                                }
                                else
                                {
                                    rResist.Interior.Color = ColorTranslator.ToOle(Color.Yellow);
                                    rCurrent.Interior.Color = ColorTranslator.ToOle(Color.Yellow);
                                    rRemark.Interior.Color = ColorTranslator.ToOle(Color.Yellow);
                                    rLimit.Interior.Color = ColorTranslator.ToOle(Color.Yellow);
                                    rResult.Interior.Color = ColorTranslator.ToOle(Color.Yellow);
                                }

                                AddLog("WARN", $"측정번호가 없는데 값이 있음");
                            }
                            continue;
                        }

                        bool hasResist = double.TryParse(Convert.ToString(rResist.Value2), out double resist);
                        bool hasCurrent = double.TryParse(Convert.ToString(rCurrent.Value2), out double current);
                        double.TryParse(Convert.ToString(rLimit.Value2), out double limit);

                        // 흰색으로 초기화
                        rResist.Interior.Pattern = Excel.XlPattern.xlPatternSolid;
                        rCurrent.Interior.Pattern = Excel.XlPattern.xlPatternSolid;
                        rResist.Interior.ColorIndex = Excel.XlColorIndex.xlColorIndexNone;
                        rCurrent.Interior.ColorIndex = Excel.XlColorIndex.xlColorIndexNone;

                        // 2. 둘 다 없으면
                        if (!hasResist && !hasCurrent)
                        {
                            if (rResult.Value2 != "ᅳ")
                            {
                                AddLog("Info", $"{resultCol}{row}에 'ᅳ' 채움");
                                rResult.Value2 = "'ᅳ";
                            }

                            if (string.IsNullOrWhiteSpace(Convert.ToString(rRemark.Value2)))
                            {

                                AddLog("Info", $"{remarkCol}{row}에 'SP' 채움");
                                rRemark.Value = "SP";
                            }

                            continue;
                        }

                        // 4. 누설전류 판정
                        if (hasCurrent && limit > 0)
                        {

                            if (rResult.Value2 == null && rResult.Value == null)
                            {
                                //AddLog("Info", $"{resultCol}{row} is null");
                                continue;
                            }
                            string result = Convert.ToString(rResult.Value)?.Trim();

                            if (current >= limit / 20.0)
                            {
                                if (result != "주의")
                                {
                                    AddLog("Info", $"{resultCol}{row}에 누설전류 '점검요' 채움");
                                    rResult.Value = "주의";
                                }
                            }
                            else
                            {
                                if (result != "양호")
                                {
                                    AddLog("Info", $"{resultCol}{row}에 누설전류 '양호' 채움");
                                    rResult.Value = "양호";
                                    rRemark.ClearContents();
                                }
                            }
                        }

                        // 5. 절연저항 판정 (누설전류보다 우선)
                        if (hasResist)
                        {
                            if (resist < 0.2)
                            {
                                if (Convert.ToString(rResult.Value2) != "점검요")
                                {
                                    AddLog("Info", $"{rResult}{row}에 절연저항 '점검요' 채움");
                                    rResult.Value = "점검요";
                                }
                            }
                            else if (resist < 1.0)
                            {
                                if (Convert.ToString(rResult.Value2) != "주의")
                                {
                                    AddLog("Info", $"{rResult}{row}에 절연저항 '주의' 채움");

                                    rResult.Value = "주의";
                                }
                            }
                            else
                            {
                                if (Convert.ToString(rResult.Value2) != "양호")
                                {
                                    AddLog("Info", $"{rResult}{row}에 절연저항 '양호' 채움");

                                    rResult.Value = "양호";
                                    rRemark.ClearContents();
                                }
                            }
                        }

                        // 3. 둘 다 있으면 노란색
                        if (hasResist && hasCurrent)
                        {
                            rResist.Interior.Color = ColorTranslator.ToOle(Color.Yellow);
                            rCurrent.Interior.Color = ColorTranslator.ToOle(Color.Yellow);
                            AddLog("Info", $"{rResist}{row}, {rCurrent}{row}에 저항과 전류가 모두 표시되어 있음");
                        }
                    }
                }
            }
            finally
            {
                if (openedHere)
                {
                    wb.Close(false);
                    app.Quit();

                    System.Runtime.InteropServices.Marshal.ReleaseComObject(wb);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(app);
                }
            }
        }

        private void btnConfirmJulyeon_Click(object sender, EventArgs e)
        {
            string filePath = tbQuantityFile.Text.Trim();
            if (string.IsNullOrWhiteSpace(filePath) || !File.Exists(filePath))
            {
                AddLog("WARN", "파일 경로가 올바르지 않습니다.");
                return;
            }
            CheckResistanceAndCurrentForJulyeon(null, filePath);
        }
        #endregion

        #region [날짜로 반기 분기 판정, 적용]
        private void SetDateForJechulmoon(Excel.Workbook wb, string filePath)
        {
            Excel.Application app = null;
            bool openedHere = false;
            Excel.Worksheet wsSrc = null;
            Excel.Range usedRange = null;

            try
            {
                if (wb == null)
                {
                    if (!File.Exists(filePath))
                    {
                        AddLog("WARN", "파일이 존재하지 않습니다.");
                        return;
                    }

                    app = new Excel.Application();
                    app.Visible = false;
                    app.DisplayAlerts = false;

                    wb = app.Workbooks.Open(filePath);
                    openedHere = true;
                }

                wsSrc = wb.Worksheets["제출문"];

                if (wsSrc == null)
                {
                    AddLog("Error", "제출문 시트를 찾을 수 없습니다.");
                    return;
                }

                if (report != null)
                {
                    int month = report.nDay > 15
                        ? report.nMonth + 1
                        : report.nMonth;

                    int year = report.nYear;

                    // 12월 15일 이후라서 다음 달이 13월이 되는 경우 처리
                    if (month > 12)
                    {
                        month = 1;
                        year++;
                    }

                    string newDate = $"{year}년 {month}월";

                    string papers =
                        report.isAnnual
                            ? "2~8,코로나방전,축전지"
                            : report.isHalfYear
                                ? "2접지,6,7"
                                : "7";

                    string newPaperText = $"첨부 별지서식 : {papers}";

                    usedRange = wsSrc.UsedRange;

                    bool dateChanged = false;
                    bool paperChanged = false;

                    foreach (Excel.Range cell in usedRange.Cells)
                    {
                        try
                        {
                            string text = Convert.ToString(cell.Value2)?.Trim();

                            if (string.IsNullOrEmpty(text))
                                continue;

                            // "첨부 별지서식"이 들어있는 셀
                            if (!paperChanged &&
                                text.Contains("첨부 별지서식"))
                            {
                                AddLog(
                                    "Info",
                                    $"제출문 {cell.Address} '{text}' → '{newPaperText}'");

                                cell.Value2 = newPaperText;
                                paperChanged = true;
                            }

                            // "XXXX년 XX월" 형태의 셀
                            if (!dateChanged &&
                                Regex.IsMatch(
                                    text,
                                    @"^\s*\d{4}년\s*\d{1,2}월\s*$"))
                            {
                                AddLog(
                                    "Info",
                                    $"제출문 {cell.Address} '{text}' → '{newDate}'");

                                cell.Value2 = newDate;
                                dateChanged = true;
                            }

                            // 둘 다 찾았으면 더 이상 검색할 필요 없음
                            if (dateChanged && paperChanged)
                                break;
                        }
                        finally
                        {
                            Marshal.ReleaseComObject(cell);
                        }
                    }

                    if (!dateChanged)
                        AddLog("WARN", "제출문에서 'XXXX년 XX월' 형식의 셀을 찾지 못했습니다.");

                    if (!paperChanged)
                        AddLog("WARN", "제출문에서 '첨부 별지서식'이 포함된 셀을 찾지 못했습니다.");
                }

                if (openedHere)
                    wb.Save();
            }
            catch (Exception ex)
            {
                AddLog(
                    "Error",
                    $"SetDateForJechulmoon에서 오류가 발생했습니다: {ex.Message}");
            }
            finally
            {
                if (usedRange != null)
                    Marshal.ReleaseComObject(usedRange);

                if (wsSrc != null)
                    Marshal.ReleaseComObject(wsSrc);

                if (openedHere)
                {
                    wb.Close(false);
                    app.Quit();

                    Marshal.ReleaseComObject(wb);
                    Marshal.ReleaseComObject(app);
                }
            }
        }

        private void SetDateForOpinion(Excel.Workbook wb, string filePath)
        {
            Excel.Application app = null;
            Excel.Worksheet wsSrc = null;
            bool openedHere = false;

            try
            {
                if (wb == null)
                {
                    if (!File.Exists(filePath))
                    {
                        AddLog("WARN", "파일이 존재하지 않습니다.");
                        return;
                    }

                    app = new Excel.Application();
                    app.Visible = false;
                    app.DisplayAlerts = false;

                    wb = app.Workbooks.Open(filePath);
                    openedHere = true;
                }

                wsSrc = wb.Worksheets["의견"];

                if (report != null)
                {
                    string papers =
                        report.isAnnual
                            ? "2,3,4,5,6,7,8,코로나방전,축전지"
                            : report.isHalfYear
                                ? "2접지,6,7"
                                : "7";


                    wsSrc.Range["A5"].Value2 = $"  ○ 전기안전관리자 직무고시 점검 : 별지서식 [{papers}] 점검.";
                    Marshal.ReleaseComObject(wsSrc);
                }

                wsSrc = wb.Worksheets["연계획"];
                if (report != null)
                {
                    //A2 셀 예시 "원흥퍼스트푸르지오시티 2026년 전기 연간점검 계획표"
                    wsSrc.Range["A2"].Value2 = $"{report.strSite} {report.nYear}년 전기 연간점검 계획표";
                    Marshal.ReleaseComObject(wsSrc);
                }

                if (openedHere)
                    wb.Save();

            }
            catch (Exception ex)
            {
                AddLog("Error", $"SetDateForOpinion 오류가 발생했습니다: {ex.Message}");
            }
            finally
            {
                if (wsSrc != null) Marshal.ReleaseComObject(wsSrc);
                if (openedHere)
                {
                    wb.Close(false);
                    app.Quit();

                    Marshal.ReleaseComObject(wb);
                    Marshal.ReleaseComObject(app);
                }
            }
        }

        private void SetDateJeoap(Excel.Workbook wb, string filePath)
        {
            Excel.Application app = null;
            Excel.Worksheet wsSrc = null;
            bool openedHere = false;

            try
            {
                if (wb == null)
                {
                    if (!File.Exists(filePath))
                    {
                        AddLog("WARN", "파일이 존재하지 않습니다.");
                        return;
                    }

                    app = new Excel.Application();
                    app.Visible = false;
                    app.DisplayAlerts = false;

                    wb = app.Workbooks.Open(filePath);
                    openedHere = true;
                }

                wsSrc = GetWorksheetByName(wb, "저압");

                if (wsSrc == null)
                {
                    AddLog("Error", "저압 시트를 찾을 수 없습니다.");
                    return;
                }

                if (report != null)
                {
                    var half = (report.nMonth>6) ? "하" : "상";
                    wsSrc.Range["A2"].Value2 = $"◈ 접지저항 측정기록표({half}반기)";

                    wsSrc.Name = $"저압({half})";
                }

                if (openedHere)
                    wb.Save();

            }
            catch (Exception ex)
            {
                AddLog("Error", $"SetDateJeoap 오류가 발생했습니다: {ex.Message}");
            }
            finally
            {
                if (wsSrc != null) Marshal.ReleaseComObject(wsSrc);
                if (openedHere)
                {
                    wb.Close(false);
                    app.Quit();

                    Marshal.ReleaseComObject(wb);
                    Marshal.ReleaseComObject(app);
                }
            }
        }

        private void SetDateYaeby(Excel.Workbook wb, string filePath)
        {
            Excel.Application app = null;
            Excel.Worksheet wsSrc = null;
            bool openedHere = false;

            try
            {
                if (wb == null)
                {
                    if (!File.Exists(filePath))
                    {
                        AddLog("WARN", "파일이 존재하지 않습니다.");
                        return;
                    }

                    app = new Excel.Application();
                    app.Visible = false;
                    app.DisplayAlerts = false;

                    wb = app.Workbooks.Open(filePath);
                    openedHere = true;
                }

                wsSrc = GetWorksheetByName(wb, "예비");

                if (wsSrc == null)
                {
                    AddLog("Error", "예비 시트를 찾을 수 없습니다.");
                    return;
                }

                if (report != null)
                {
                    var half = (report.nMonth > 6) ? "하" : "상";
                    wsSrc.Range["A2"].Value2 = $"발전설비 점검기록표({half}반기)";

                    wsSrc.Name = $"예비({half})";
                }

                if (openedHere)
                    wb.Save();

            }
            catch (Exception ex)
            {
                AddLog("Error", $"SetDateYaeby 오류가 발생했습니다: {ex.Message}");
            }
            finally
            {
                if (wsSrc != null) Marshal.ReleaseComObject(wsSrc);
                if (openedHere)
                {
                    wb.Close(false);
                    app.Quit();

                    Marshal.ReleaseComObject(wb);
                    Marshal.ReleaseComObject(app);
                }
            }
        }

        private void SetDateBungy(Excel.Workbook wb, string filePath)
        {

            if (report.isOnlyAnnual) return;

            Excel.Application app = null;
            Excel.Worksheet wsSrc = null;
            bool openedHere = false;

            try
            {
                if (wb == null)
                {
                    if (!File.Exists(filePath))
                    {
                        AddLog("WARN", "파일이 존재하지 않습니다.");
                        return;
                    }

                    app = new Excel.Application();
                    app.Visible = false;
                    app.DisplayAlerts = false;

                    wb = app.Workbooks.Open(filePath);
                    openedHere = true;
                }

                wsSrc = GetWorksheetByName(wb, "분기");

                if (wsSrc == null)
                {
                    AddLog("Error", "분기 시트를 찾을 수 없습니다.");
                    return;
                }

                if (report != null)
                {
                    wsSrc.Range["E5"].Value2 = $"{report.nYear}-{report.nMonth}-{report.nDay}"; 
                    wsSrc.Range["L5"].Value2 = $"[{report.nQuater}분기]";
                    wsSrc.Name = $"{report.nQuater}분기";
                }

                if (wsSrc != null) Marshal.ReleaseComObject(wsSrc);
                wsSrc = null;

                wsSrc = GetWorksheetByName(wb, "절연");

                if (wsSrc == null)
                {
                    AddLog("Error", "절연 시트를 찾을 수 없습니다.");
                    return;
                }

                if (report != null)
                {
                    wsSrc.Range["Q5"].Value2 = $"{report.nYear}-{report.nMonth}-{report.nDay}";
                }

                if (openedHere)
                    wb.Save();

            }
            catch (Exception ex)
            {
                AddLog("Error", $"SetDateBungy 오류가 발생했습니다: {ex.Message}");
            }
            finally
            {
                if (wsSrc != null) Marshal.ReleaseComObject(wsSrc);
                if (openedHere)
                {
                    wb.Close(false);
                    app.Quit();

                    Marshal.ReleaseComObject(wb);
                    Marshal.ReleaseComObject(app);
                }
            }
        }
        private void btnSeasonCorrect_Click(object sender, EventArgs e)
        {
            string filePath = tbQuantityFile.Text.Trim();
            if (string.IsNullOrWhiteSpace(filePath) || !File.Exists(filePath))
            {
                AddLog("WARN", "파일 경로가 올바르지 않습니다.");
                return;
            }

            Cursor = Cursors.WaitCursor;
            Excel.Application app = null;
            app = new Excel.Application();
            app.Visible = false;
            app.DisplayAlerts = false; 
            Excel.Workbook wb;

            wb = app.Workbooks.Open(filePath);

            // 제출문 시트의 A21 셀 년 월 예시 "2026년 8월" 날짜가 16일 이상이면 현재달 +1로 표시
            SetDateForJechulmoon(wb, filePath);
            // 연계획 시트의 A2 셀 예시 "원흥퍼스트푸르지오시티 2026년 전기 연간점검 계획표" 사이트 + 연 + 전지 연간점검 계획표
            SetDateForOpinion(wb, filePath);
            // 저압, 예비가 포함된 시트 이름을 저압(상)/저압(하), 예비(상)/예비(하)로 변경
            // 저압 A2셀: ◈ 접지저항 측정기록표(하반기), 예비 A2셀: 발전설비 점검기록표(하반기)
            SetDateYaeby(wb, filePath);
            SetDateJeoap(wb, filePath);
            // 연차 점검만 있는 보고서가 아닐 경우 예시 [3분기], 연차 점검만 할 경우는 E5~O5가 머지되지 않았으면 머지
            SetDateBungy(wb, filePath);

            wb.Save();

            wb.Close(false);
            app.Quit();

            Marshal.ReleaseComObject(wb);
            Marshal.ReleaseComObject(app);

            GC.Collect();
            GC.WaitForPendingFinalizers();

            Cursor = Cursors.Default;

        }
        #endregion

        private void btnChangeFooterLogo_Click(object sender, EventArgs e)
        {
            string filePath = tbQuantityFile.Text?.Trim();
            string logoPath = @"D:\Logo.png";

            Excel.Application app = null;
            Excel.Workbook wb = null;

            try
            {
                if (string.IsNullOrWhiteSpace(filePath) || !File.Exists(filePath))
                {
                    MessageBox.Show(
                        "엑셀 파일 경로가 올바르지 않습니다.",
                        "확인",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                    return;
                }

                if (!File.Exists(logoPath))
                {
                    MessageBox.Show(
                        $"로고 파일이 없습니다.\n{logoPath}",
                        "확인",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                    return;
                }

                Cursor = Cursors.WaitCursor;

                app = new Excel.Application();
                app.Visible = false;
                app.DisplayAlerts = false;

                wb = app.Workbooks.Open(filePath, ReadOnly: false);

                int sheetCount = wb.Worksheets.Count;

                for (int i = 1; i <= sheetCount; i++)
                {
                    Excel.Worksheet ws = null;
                    Excel.PageSetup pageSetup = null;
                    Excel.Graphic graphic = null;

                    try
                    {
                        ws = (Excel.Worksheet)wb.Worksheets[i];
                        pageSetup = ws.PageSetup;

                        graphic = pageSetup.RightFooterPicture;

                        // 바닥글 이미지 교체
                        graphic.Filename = logoPath;

                        // 정확한 크기 지정
                        graphic.LockAspectRatio =
                            Office.MsoTriState.msoFalse;

                        // 2.68cm × 0.53cm
                        graphic.Width = 75.97f;
                        graphic.Height = 15.02f;

                        // 가로세로 비율 고정
                        graphic.LockAspectRatio =
                            Office.MsoTriState.msoTrue;

                        // 오른쪽 바닥글에 그림 표시
                        pageSetup.RightFooter = "&G";

                        AddLog(
                            "Info",
                            $"{ws.Name} 시트 바닥글 로고 변경 완료");
                    }
                    catch (Exception ex)
                    {
                        AddLog(
                            "Error",
                            $"{ws?.Name ?? i.ToString()} 시트 로고 변경 실패: {ex.Message}");
                    }
                    finally
                    {
                        if (graphic != null)
                            Marshal.ReleaseComObject(graphic);

                        if (pageSetup != null)
                            Marshal.ReleaseComObject(pageSetup);

                        if (ws != null)
                            Marshal.ReleaseComObject(ws);
                    }
                }

                wb.Save();

                MessageBox.Show(
                    "모든 시트의 바닥글 로고를 변경했습니다.",
                    "완료",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                AddLog("Error", $"바닥글 로고 변경 실패: {ex.Message}");

                MessageBox.Show(
                    ex.Message,
                    "오류",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
            finally
            {
                if (wb != null)
                {
                    try
                    {
                        wb.Close(SaveChanges: false);
                    }
                    catch { }

                    Marshal.ReleaseComObject(wb);
                }

                if (app != null)
                {
                    try
                    {
                        app.Quit();
                    }
                    catch { }

                    Marshal.ReleaseComObject(app);
                }

                Cursor = Cursors.Default;

                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }
    }

}


