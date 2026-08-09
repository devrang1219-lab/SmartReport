using OpenCvSharp;
using System;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using WindowsFormsApp1.Comm;
using Excel = Microsoft.Office.Interop.Excel;

namespace WindowsFormsApp1
{
    public class Report
    {
        public bool isAnnual { get; set; } = false;
        public bool isHalfYear { get; set; } = false;
        public bool isOnlyAnnual { get; set; } = false;
        public bool isUpperOfHalfYear { get; set; } = false;
        public int nQuater { get; set; } = 1;
        public int nYear { get; set; } = 2026;
        public int nMonth { get; set; } = 7;
        public int nDay { get; set; } = 1;
        public int totalCount { get; set; } = 0;
        public int quaterCount { get; set; } = 0;
        public string strSite { get; set; }
        public string strInspector { get; set; }
        public string xlsFilePath { get; set; }

        public static Report ParseReport(string filePath, SoborLog soborLog)
        {
            string fileName = Path.GetFileNameWithoutExtension(filePath);

            string[] parts = fileName.Split('_');

            Report report = new Report();
            report.xlsFilePath = filePath;

            report.strSite = parts[0];

            string reportName = parts[1];

            // 연차 여부
            report.isAnnual = reportName.Contains("연차");


            // 연도
            Match yearMatch = Regex.Match(reportName, @"(\d{2})년");
            if (yearMatch.Success)
                report.nYear = 2000 + int.Parse(yearMatch.Groups[1].Value);

            // 분기
            //Match quarterMatch = Regex.Match(reportName, @"(\d)분기");
            //if (quarterMatch.Success)
            //    report.nQuater = int.Parse(quarterMatch.Groups[1].Value);

            // 날짜
            if (parts.Length >= 3 && parts[2].Length == 6)
            {
                report.nYear = 2000 + int.Parse(parts[2].Substring(0, 2));
                report.nMonth = int.Parse(parts[2].Substring(2, 2));
                report.nDay = int.Parse(parts[2].Substring(4, 2));
            }

            // 상반기 여부
            report.isUpperOfHalfYear = report.nQuater <= 2;

            report.quaterCount = report.GetQuarterCount(filePath);

            return report;
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

        private int GetMonthFromFileName(string filePath)
        {
            string fileName = Path.GetFileNameWithoutExtension(filePath);

            Match m = Regex.Match(fileName, @"_(\d{6})$");

            if (!m.Success)
                throw new Exception("파일명에서 날짜를 찾을 수 없습니다.");

            string yymmdd = m.Groups[1].Value;

            return int.Parse(yymmdd.Substring(2, 2)); // MM
        }

        public int GetQuarterCount(string filePath)
        {
            totalCount = 0;
            quaterCount = 0;


            if (string.IsNullOrEmpty(filePath) || !File.Exists(filePath))
            {
                MessageBox.Show("페이지 번호를 매길 엑셀 파일을 먼저 선택하세요.", "오류", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return 0;
            }

            int targetMonth = GetMonthFromFileName(filePath);
            Excel.Application xlApp = null;
            Excel.Workbook wb = null;
            Excel.Worksheet ws = null;

            try
            {
                xlApp = new Excel.Application { Visible = false, DisplayAlerts = false };
                // Open for write because we modify PageSetup
                wb = xlApp.Workbooks.Open(filePath, ReadOnly: false);

                ws = GetWorksheetByName(wb, "연계획");

                // D열 = 4, 1월이라고 가정
                for (int month = 1; month <= 12; month++)
                {
                    int col = 3 + month; // D=4

                    var value = ws.Cells[24, col].Text.Trim();

                    if (value == "●")
                    {
                        if (month <= targetMonth)
                            quaterCount++;
                        totalCount++;
                    }
                }

                // 절연 점검 포함은 연차, 절연점검 미포함, 접지저항측정은 반기
                var val1 = ws.Cells[13, 3 + targetMonth].Text.Trim();
                var val2 = ws.Cells[10, 3 + targetMonth].Text.Trim();
                if (val1 == "●" && val2 != "●")
                {
                    isHalfYear = true;
                }

                return quaterCount;
            }
            catch (Exception ex)
            {
                // FormMain 인스턴스가 있으면 AddLog 호출, 없으면 MessageBox 표시
                try
                {
                    var mainForm = Application.OpenForms.OfType<SmartReport.FormMain>().FirstOrDefault();
                    if (mainForm != null)
                        mainForm.AddLog("Error", $"분기 수 계산 중 오류 발생: {ex.Message}");
                    else
                        MessageBox.Show($"분기 수 계산 중 오류 발생: {ex.Message}", "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                catch
                {
                    // 추가 안전 장치: 예외가 발생해도 무시
                    MessageBox.Show($"분기 수 계산 중 오류 발생: {ex.Message}", "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                return 0;
            }
            finally
            {

                // 연차 보고서만 있는 경우 추출
                if (totalCount == 0)
                {
                    MessageBox.Show("연계획 시트에서 분기 수를 찾을 수 없습니다.", "오류", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    isOnlyAnnual = false;
                }
                else if (quaterCount == 1)
                {
                    isOnlyAnnual = true;
                }
                else
                {
                    isOnlyAnnual = false;
                }

                // 분기
                nQuater = quaterCount;

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
            }


        }

        public void CopySheetToXlsxAndProcess(string sheetName, float gapLeft, float gapTop, float gapRight, float gapBottom)
        {
            Excel.Application xlApp = null;
            Excel.Workbook srcWb = null;
            Excel.Workbook newWb = null;

            try
            {
                xlApp = new Excel.Application { Visible = false, DisplayAlerts = false };
                srcWb = xlApp.Workbooks.Open(xlsFilePath, ReadOnly: true);

                Excel.Worksheet srcWs = GetWorksheetByName(srcWb, sheetName);
                if (srcWs == null)
                    throw new InvalidOperationException($"원본 파일에 '{sheetName}' 시트가 없습니다.");

                // 시트 복사(새 워크북으로)
                srcWs.Copy(Type.Missing, Type.Missing);
                newWb = xlApp.ActiveWorkbook;

                // destPath를 xlsFilePath와 같은 폴더로 설정 (xlsFilePath 필드가 있으면 우선 사용)
                string destFolder = Path.GetDirectoryName(this.xlsFilePath);

                //string destFileName = Path.GetFileNameWithoutExtension(xlsFilePath) + "_" + sheetName + ".xlsx";
                string destFileName = sheetName + ".xlsx";
                string destPath = Path.Combine(destFolder, destFileName);

                if (File.Exists(destPath))
                    File.Delete(destPath); // 덮어쓰기

                // .xlsx 형식으로 저장
                newWb.SaveAs(destPath, Excel.XlFileFormat.xlOpenXMLWorkbook);

                // 새 워크북의 첫 시트에서 처리
                Excel.Worksheet newWs = (Excel.Worksheet)newWb.Sheets[1];
                SnapImageMergedCell(srcWs, sheetName);
                SnapImageMergedCell(newWs, sheetName);
            }
            finally
            {
                if (srcWb != null) { srcWb.Save(); srcWb.Close(false); Marshal.ReleaseComObject(srcWb); }
                if (newWb != null) { newWb.Save(); newWb.Close(false); Marshal.ReleaseComObject(newWb); }
                if (xlApp != null) { xlApp.Quit(); Marshal.ReleaseComObject(xlApp); }
                GC.Collect();
                GC.WaitForPendingFinalizers();

                Cursor.Current = Cursors.Default;
            }
        }

        float GetGap(Excel.Border border)
        {
            float gap = 1f;

            if ((Excel.XlBorderWeight)border.Weight == Excel.XlBorderWeight.xlMedium)
                gap = 2f;

            if ((Excel.XlBorderWeight)border.Weight == Excel.XlBorderWeight.xlThick)
                gap = 3f;

            if ((Excel.XlLineStyle)border.LineStyle == Excel.XlLineStyle.xlDouble)
                gap += 2f;

            return gap;
        }

        public void SnapImageMergedCell(Excel.Worksheet ws, string sheetName, 
            float gapLeft = 1.5f, float gapTop = 1.5f, float gapRight = 0f, float gapBottom = 0.5f)
        {
            Excel.Application xlApp = null;
            Excel.Workbook wb = null;
            bool openedHere = false;
            int i = 0;

            try
            {
                // ws가 null이면 xlsFilePath로 파일을 열고 "장비" 시트를 가져옴
                if (ws == null)
                {
                    if (string.IsNullOrEmpty(xlsFilePath) || !File.Exists(xlsFilePath))
                    {
                        MessageBox.Show("엑셀 파일 경로가 없습니다. xlsFilePath를 설정하세요.", "오류", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }

                    xlApp = new Excel.Application { Visible = false, DisplayAlerts = false };
                    wb = xlApp.Workbooks.Open(xlsFilePath, ReadOnly: false);
                    ws = GetWorksheetByName(wb, sheetName);
                    if (ws == null)
                    {
                        MessageBox.Show($"'{sheetName}' 시트를 찾을 수 없습니다.", "오류", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        // 열었던 워크북 닫기
                        wb.Close(false);
                        Marshal.ReleaseComObject(wb);
                        xlApp.Quit();
                        Marshal.ReleaseComObject(xlApp);
                        return;
                    }
                    openedHere = true;
                }

                foreach (Excel.Shape shape in ws.Shapes)
                {
                    // 도형이 걸치는 셀 범위(왼쪽상단~오른쪽하단)를 사용
                    Excel.Range topLeft = shape.TopLeftCell;
                    Excel.Range bottomRight = shape.BottomRightCell;
                    Excel.Range cellRange = ws.Range[topLeft, bottomRight];

                    // 병합 셀이면 병합영역 사용 (COM에서 DBNull 반환 가능하므로 방어적 검사)
                    object mergeObj = null;
                    try { mergeObj = cellRange.MergeCells; } catch { mergeObj = null; }
                    bool isMerged = (mergeObj is bool b && b);


                    //Excel.Range area = shape.TopLeftCell;
                    //Excel.Range area = ws.Range[
                    //                        shape.TopLeftCell,
                    //                        shape.BottomRightCell];

                    //if ((bool)area.MergeCells)
                    //{
                    //    area = area.MergeArea;
                    //}

                    Excel.Range area = shape.TopLeftCell;

                    try
                    {
                        if ((bool)area.MergeCells)
                            area = area.MergeArea;
                    }
                    catch
                    {
                        area = ws.Range[
                                            shape.TopLeftCell,
                                            shape.BottomRightCell];
                    }

                    double rot = shape.Rotation % 360;
                    if (rot < 0) rot += 360;

                    bool rotated90 =
                        Math.Abs(rot - 90) < 1 ||
                        Math.Abs(rot - 270) < 1;

                    //float gapLeft = 1f;
                    //float gapTop = 1f;
                    //float gapRight = 0f;
                    //float gapBottom = 0f;

                    //var borders = area.Borders;

                    //gapLeft =
                    //    GetGap(borders[Excel.XlBordersIndex.xlEdgeLeft]);

                    //gapTop =
                    //    GetGap(borders[Excel.XlBordersIndex.xlEdgeTop]);

                    //gapRight =
                    //    GetGap(borders[Excel.XlBordersIndex.xlEdgeRight]);

                    //gapRight =
                    //    GetGap(borders[Excel.XlBordersIndex.xlEdgeBottom]);

                    if (rotated90)
                    {
                        float targetW = (float)area.Height;
                        float targetH = (float)area.Width;

                        shape.Width = targetW - gapLeft - gapRight;
                        shape.Height = targetH - gapTop - gapBottom;

                        shape.Left = (float)area.Left +
                                     ((float)area.Width - targetW) / 2f + gapLeft;

                        shape.Top = (float)area.Top +
                                    ((float)area.Height - targetH) / 2f + gapTop;
                    }
                    else
                    {
                        shape.Left = (float)area.Left + gapLeft;
                        shape.Top = (float)area.Top + gapTop;
                        shape.Width = (float)area.Width - gapLeft - gapRight;
                        shape.Height = (float)area.Height - gapTop - gapRight;
                    }
                    i++;

                    //shape.Left = (float)area.Left;
                    //shape.Top = (float)area.Top;

                    // 90/270도 회전된 도형은 너비/높이 값을 교체
                    //double rot = 0;
                    //try { rot = shape.Rotation; } catch { rot = 0; }
                    //rot = rot % 360;
                    //if (rot < 0) rot += 360;
                    //bool rotated90 = Math.Abs(rot - 90) < 1.0 || Math.Abs(rot - 270) < 1.0;

                    //if (rotated90)
                    //{
                    //    shape.Width = (float)area.Height;
                    //    shape.Height = (float)area.Width;
                    //}
                    //else
                    //{
                    //    shape.Width = (float)area.Width;
                    //    shape.Height = (float)area.Height;
                    //}
                    shape.LockAspectRatio =
                        Microsoft.Office.Core.MsoTriState.msoFalse;
                    shape.Placement = Excel.XlPlacement.xlMoveAndSize;
                }
            }
            catch (Exception ex)
            {
                try
                {
                    var mainForm = Application.OpenForms.OfType<SmartReport.FormMain>().FirstOrDefault();
                    if (mainForm != null)
                        mainForm.AddLog("Error", $"이미지 스냅 처리 중 오류 발생: {ex.Message}");
                    else
                        MessageBox.Show($"이미지 스냅 처리 중 오류 발생: {ex.Message}", "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                catch
                {
                    MessageBox.Show($"이미지 스냅 처리 중 오류 발생: {ex.Message}", "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            finally
            {
                if (openedHere)
                {
                    if (wb != null)
                    {
                        wb.Save();
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
        }
    }
}
