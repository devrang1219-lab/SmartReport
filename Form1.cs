using SynologyIntegration;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace SmartReport
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            //TestCopyGapjiSheet();
            // 파일 드래그 앤 드롭을 허용하고 이벤트를 연결
            try
            {
                tbFileNameForFunction.AllowDrop = true;
                tbFileNameForFunction.DragEnter += TbFileNameForFunction_DragEnter;
                tbFileNameForFunction.DragDrop += TbFileNameForFunction_DragDrop;
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
                Excel.Workbook newWb = app.ActiveWorkbook;

                newWb.SaveAs(
                    targetFile,
                    Excel.XlFileFormat.xlOpenXMLWorkbook);

                newWb.Close(false);
                Marshal.ReleaseComObject(newWb);
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
            using (var dlg = new OpenFileDialog())
            {
                dlg.Title = "엑셀 파일을 선택하세요";
                dlg.Filter = "Excel 파일 (*.xlsx;*.xls;*.xlsm)|*.xlsx;*.xls;*.xlsm|모든 파일 (*.*)|*.*";
                dlg.CheckFileExists = true;
                dlg.CheckPathExists = true;
                dlg.Multiselect = false;

                // 기본 위치 설정 (폴더가 존재하면 InitialDirectory로 설정)
                var defaultPath = @"C:\_D\work\한경이엔지\2_report";
                if (Directory.Exists(defaultPath))
                {
                    dlg.InitialDirectory = defaultPath;
                }

                var result = dlg.ShowDialog();
                if (result == DialogResult.OK)
                {
                    tbFolder.Text = dlg.FileName;

                    try
                    {
                        //var newPath = CreateNewExcelFromTemplate(dlg.FileName);
                        var newPath = creatfileTemp(tbFolder.Text);
                        MessageBox.Show($"새 엑셀 파일을 생성했습니다:\r\n{newPath}", "완료", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"파일 생성 중 오류가 발생했습니다:\r\n{ex.Message}", "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }

        }


        private void TbFileNameForFunction_DragEnter(object sender, DragEventArgs e)
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

        private void TbFileNameForFunction_DragDrop(object sender, DragEventArgs e)
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
                        try { Invoke(new Action(() => tbFileNameForFunction.Text = content)); } catch { }
                    });
                }
                else
                {
                    // 경로만 표시
                    tbFileNameForFunction.Text = file;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"파일 처리 중 오류: {ex.Message}", "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // 할 일
        // 제출문 시트의 A20 셀을 2025년 7월 ==> 26년 6월로 변경하는 예시
        // 연계획 시트의 A2 2025년을 2026으로 변경
        // Sheet1 제거 (기본 생성)
        // 측정자 셀 변경하기
        // 페이지 번호 매기기



        private string creatfileTemp(string originalFilePath)
        {
            var dir = Path.GetDirectoryName(originalFilePath) ?? throw new InvalidOperationException("디렉터리 정보를 가져올 수 없습니다.");
            var originalName = Path.GetFileNameWithoutExtension(originalFilePath);
            var ext = Path.GetExtension(originalFilePath);

            var folderName = new DirectoryInfo(dir).Name;
            var folderParts = folderName.Split('_');
            if (folderParts.Length < 1)
                throw new InvalidOperationException("폴더명이 예상 형식이 아닙니다. '_'로 구분된 첫번째 부분에 날짜가 있어야 합니다.");

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

            var conditionalSheets = new[] { "갑지", "제출문", "의견", "연계획", "검교정" };
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
                Microsoft.Office.Interop.Excel.Workbook wbOrig = null;
                var origSheetNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                try
                {
                    if (File.Exists(originalFilePath))
                    {
                        wbOrig = xlApp.Workbooks.Open(originalFilePath, ReadOnly: true);
                        foreach (Microsoft.Office.Interop.Excel.Worksheet sh in wbOrig.Worksheets)
                        {
                            try { origSheetNames.Add(sh.Name); } catch { }
                            finally { if (sh != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(sh); }
                        }
                    }
                }
                finally
                {
                    if (wbOrig != null)
                    {
                        wbOrig.Close(false);
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(wbOrig);
                        wbOrig = null;
                    }
                }

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
                    string tempYearPath = null;
                    if (wbYear != null)
                    {
                        try
                        {
                            // temp 파일 생성
                            var tempExt = Path.GetExtension(yearFile);
                            tempYearPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + tempExt);
                            try { File.Copy(yearFile, tempYearPath, true); } catch (Exception ex) { System.Diagnostics.Debug.WriteLine($"Failed to copy yearFile to temp: {ex.Message}"); tempYearPath = null; }

                            if (!string.IsNullOrEmpty(tempYearPath) && File.Exists(tempYearPath))
                            {
                                try
                                {
                                    // 닫고 원본 대신 임시 파일을 열기
                                    try { wbYear.Close(false); } catch { }
                                    try { Marshal.ReleaseComObject(wbYear); } catch { }
                                    wbYear = null;
                                }
                                catch { }

                                try
                                {
                                    xlApp.AskToUpdateLinks = false;
                                }
                                catch { }

                                try
                                {
                                    wbYear = xlApp.Workbooks.Open(tempYearPath, 0, ReadOnly: false);
                                }
                                catch (Exception ex)
                                {
                                    System.Diagnostics.Debug.WriteLine($"Failed to open temp year workbook: {ex.Message}");
                                }
                            }
                        }
                        catch { }
                    }

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

        private SynologyFileDownloader _downloader;
        private BindingList<SynologyFileGridItem> _gridFiles;

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

        public async Task TestSynologyAsync()
        {
            Cursor = Cursors.WaitCursor;

            try
            {

                var config = new SynologyFileDownloaderConfig
                {
                    BaseUrl = "http://hk2ng.synology.me:5000",
                    Username = "유량진",
                    Password = "HKeng717241",
                    SearchFolder = "/2_1전기직무고시점검보고서/2 본부(김희철)/0001 시흥보고서방/001열화상",
                    Keyword = tbCompany.Text,
                    Extensions = new[] { ".xlsx", ".pdf" },
                    Recursive = true,
                    DownloadLatestOnly = false
                };

                //MessageBox.Show("1. config 생성 완료");

                if (_downloader != null)
                {
                    _downloader.Dispose();
                    _downloader = null;
                }


                _downloader = new SynologyFileDownloader(config);
                //MessageBox.Show("2. downloader 생성 완료");

                await _downloader.LoginAsync();
                //MessageBox.Show("3. 로그인 완료");

                //List<SynologyFileItem> files = await _downloader.SearchFilesAsync();
                //MessageBox.Show("4. 검색 완료");

                //_gridFiles = ConvertToGridItems(files);
                //dgvFiles.DataSource = _gridFiles;

                _files = await _downloader.SearchFilesAsync();
                _fileBindingSource.DataSource = _files;
                dgvFiles.DataSource = _fileBindingSource;

                MessageBox.Show("검색 완료: " + _files.Count + "건");
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

        private void Form1_Load(object sender, EventArgs e)
        {
            InitGrid();
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

            await TestSynologyAsync();
        }

        private void btnPageNumber_Click(object sender, EventArgs e)
        {
            var filePath = tbFileNameForFunction.Text?.Trim();
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

                int currentStartPage = 1;

                for (int i = 1; i <= sheetCount; i++)
                {
                    Excel.Worksheet sh = null;
                    try
                    {
                        sh = (Excel.Worksheet)wb.Worksheets[i];
                        string name = "";
                        try { name = sh.Name; } catch { }

                        // '갑지' 시트는 페이지 번호 매기기에서 제외
                        if (string.Equals(name, "갑지", StringComparison.OrdinalIgnoreCase))
                        {
                            // leave as automatic
                            continue;
                        }

                        // 계산: 가로/세로 페이지 구분점 수 -> 페이지수 = (H+1)*(V+1)
                        int h = 0, v = 0;
                        try { h = sh.HPageBreaks.Count; } catch { }
                        try { v = sh.VPageBreaks.Count; } catch { }
                        int pages = (h + 1) * (v + 1);
                        if (pages <= 0) pages = 1;

                        try
                        {
                            // 시작 페이지 번호 설정
                            sh.PageSetup.FirstPageNumber = currentStartPage;
                        }
                        catch (Exception ex)
                        {
                            System.Diagnostics.Debug.WriteLine($"Failed to set FirstPageNumber for sheet '{name}': {ex.Message}");
                        }

                        // 형식은 변경하지 않음: FirstPageNumber만 설정

                        currentStartPage += pages;
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"페이지 번호 처리 중 오류(시트 idx={i}): {ex.Message}");
                    }
                    finally
                    {
                        if (sh != null) try { Marshal.ReleaseComObject(sh); } catch { }
                    }
                }

                try { wb.Save(); } catch { }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"페이지 번호 매기기 중 오류가 발생했습니다:\r\n{ex.Message}", "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
                Cursor = Cursors.Default;
            }
        }
        private void btnExportForPdf_Click(object sender, EventArgs e)
        {
            var filePath = tbFileNameForFunction.Text?.Trim();
            if (string.IsNullOrEmpty(filePath) || !File.Exists(filePath))
            {
                MessageBox.Show("내보낼 엑셀 파일을 먼저 선택하세요.", "오류", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            Cursor = Cursors.WaitCursor;
            Excel.Application xlApp = null;
            Excel.Workbook wb = null;
            string outFile = null;

            try
            {
                xlApp = new Excel.Application { Visible = false, DisplayAlerts = false };
                wb = xlApp.Workbooks.Open(filePath, ReadOnly: true);

                var dir = Path.GetDirectoryName(filePath) ?? Environment.CurrentDirectory;
                string baseName = Path.GetFileNameWithoutExtension(filePath);
                outFile = Path.Combine(dir, baseName + "_merged.pdf");
                int idx = 1;
                var candidate = outFile;
                while (File.Exists(candidate))
                {
                    candidate = Path.Combine(dir, baseName + "_merged_" + idx + ".pdf");
                    idx++;
                }

                // Export entire workbook as a single PDF (모든 시트를 하나의 PDF로)
                try
                {
                    wb.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF, candidate,
                        Excel.XlFixedFormatQuality.xlQualityStandard, IncludeDocProperties: true, IgnorePrintAreas: false, OpenAfterPublish: false);
                }
                catch (Exception ex)
                {
                    throw new Exception("통합 PDF 내보내기 실패: " + ex.Message, ex);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"PDF 내보내기 중 오류가 발생했습니다:\r\n{ex.Message}", "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
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
                Cursor = Cursors.Default;
            }

            if (!string.IsNullOrEmpty(outFile) && File.Exists(outFile))
            {
                MessageBox.Show($"통합 PDF로 내보냈습니다:\r\n{outFile}", "완료", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show("PDF 파일이 생성되지 않았습니다.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
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
                    tbFileNameForFunction.Text = dlg.FileName;
                }
            }
        }

        private void btErrorPageUpdate_Click(object sender, EventArgs e)
        {
            var filePath = tbFileNameForFunction.Text?.Trim();
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

                OpinionMaker.FillOpinionFromInsulation(wb);

                try { wb.Save(); } catch { }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"페이지 번호 매기기 중 오류가 발생했습니다:\r\n{ex.Message}", "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
                Cursor = Cursors.Default;
            }
        }

        private void btnMakeExcel_Click(object sender, EventArgs e)
        {
            try
            {
                if (!double.TryParse(tbWidthFeverPicture.Text.Trim(), out double imageWidthCm) || imageWidthCm <= 0)
                {
                    MessageBox.Show("이미지 가로(cm)를 올바르게 입력하세요.");
                    return;
                }

                if (!double.TryParse(tbHeightFeverPicture.Text.Trim(), out double imageHeightCm) || imageHeightCm <= 0)
                {
                    MessageBox.Show("이미지 세로(cm)를 올바르게 입력하세요.");
                    return;
                }

                if (!double.TryParse(tbGapFeverPicture.Text.Trim(), out double gapCm) || gapCm < 0)
                {
                    MessageBox.Show("이미지 가로 간격(cm)을 올바르게 입력하세요.");
                    return;
                }

                // 이미지 폴더 선택
                string selectedFolder = "";
                using (FolderBrowserDialog fbd = new FolderBrowserDialog())
                {
                    fbd.Description = "이미지 폴더를 선택하세요.";

                    // 기본 폴더 지정
                    string defaultDir = tbDefultDirectory.Text.Trim();
                    if (!string.IsNullOrEmpty(defaultDir) && Directory.Exists(defaultDir))
                    {
                        fbd.SelectedPath = defaultDir;
                    }

                    if (fbd.ShowDialog() != DialogResult.OK)
                        return;

                    selectedFolder = fbd.SelectedPath;
                }

                // 폴더 안 이미지 목록
                string[] imageExtensions = new[] { ".jpg", ".jpeg", ".png", ".bmp", ".gif", ".tif", ".tiff" };

                var imageFiles = Directory.GetFiles(selectedFolder)
                    .Where(x => imageExtensions.Contains(Path.GetExtension(x).ToLower()))
                    .OrderBy(x => x, StringComparer.CurrentCultureIgnoreCase)
                    .ToArray();

                if (imageFiles.Length == 0)
                {
                    MessageBox.Show("선택한 폴더에 이미지 파일이 없습니다.");
                    return;
                }

                // 상위 폴더의 Images.xlsx 경로
                DirectoryInfo parent = Directory.GetParent(selectedFolder);
                if (parent == null)
                {
                    MessageBox.Show("상위 폴더를 찾을 수 없습니다.");
                    return;
                }

                string excelPath = Path.Combine(parent.FullName, "Images.xlsx");
                if (!File.Exists(excelPath))
                {
                    MessageBox.Show("상위 폴더에 Images.xlsx 파일이 없습니다.\r\n" + excelPath);
                    return;
                }


                InsertImagesToExistingWorkbook(
                    excelPath,
                    imageFiles,
                    imageWidthCm,
                    imageHeightCm,
                    gapCm
                );

                MessageBox.Show("완료되었습니다.");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
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

        private void button3_Click(object sender, EventArgs e)
        {
            btnMakeExcel_Click(sender, e);
        }


        #region 품질관리 관련 기능
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
            ProcQuantitySheet();
        }

        private void ProcQuantitySheet()
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
                wb = xlApp.Workbooks.Open(filePath, ReadOnly: false);

                foreach (Excel.Worksheet sheet in wb.Worksheets)
                {
                    if (sheet.Name.Trim() == "품질")
                    {
                        ws = sheet;
                        break;
                    }
                }

                if (ws == null)
                    throw new Exception("품질 시트를 찾을 수 없습니다.");

                //ws = (Excel.Worksheet)wb.Worksheets["품질"];

                // Export entire workbook as a single PDF (모든 시트를 하나의 PDF로)
                try
                {
                    string baseFolder = Path.GetDirectoryName(filePath);
                    string pdfPath = Path.Combine(baseFolder, "02 전원품질", "K.pdf");
                    using (ImageInserter inserter = new ImageInserter(ws, pdfPath))
                    {
                        inserter.Insert("U6", 0.8);
                        wb.Save();
                    }
                }
                catch (Exception ex)
                {
                    throw new Exception("이미지 삽입 실패: " + ex.Message, ex);
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
        #endregion
    }

}


