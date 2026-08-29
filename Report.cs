using OpenCvSharp;
using SmartReport;
using System;
using System.ComponentModel.Composition.Primitives;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using WindowsFormsApp1.Comm;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

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

        static public FormMain mainForm { get; set; }

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

            mainForm = Application.OpenForms.OfType<SmartReport.FormMain>().FirstOrDefault();
            return report;
        }



        public Excel.Worksheet GetWorksheetByName(
            Excel.Workbook wb,
            string sheetName)
        {
            Excel.Sheets sheets = null;

            try
            {
                sheets = wb.Worksheets;

                int count = sheets.Count;

                for (int i = 1; i <= count; i++)
                {
                    Excel.Worksheet sheet = null;

                    try
                    {
                        sheet = (Excel.Worksheet)sheets[i];

                        bool matched =
                            sheet.Name.Trim()
                                .IndexOf(
                                    sheetName,
                                    StringComparison.OrdinalIgnoreCase) >= 0;

                        if (matched)
                        {
                            // 소유권을 호출한 쪽으로 넘김
                            Excel.Worksheet result = sheet;
                            sheet = null;

                            return result;
                        }
                    }
                    finally
                    {
                        if (sheet != null)
                            Marshal.ReleaseComObject(sheet);
                    }
                }

                return null;
            }
            finally
            {
                if (sheets != null)
                    Marshal.ReleaseComObject(sheets);
            }
        }

        private int GetMonthFromFileName(string filePath)
        {
            string fileName = Path.GetFileNameWithoutExtension(filePath);

            Match m = Regex.Match(fileName, @"(\d{6})$");

            if (!m.Success)
                throw new Exception("파일명에서 날짜를 찾을 수 없습니다.");

            string yymmdd = m.Groups[1].Value;

            return int.Parse(yymmdd.Substring(2, 2)); // MM
        }

        public int GetQuarterCount()
        {
            var quarterCount = 1;
            quarterCount = this.nMonth / 3 + 1;
            return quarterCount;
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

                wb = ExcelComHelper.OpenWorkbook(xlApp, filePath, false);
                ws = ExcelComHelper.GetWorksheet(wb, "연계획");

                if (ws == null)
                {
                    throw new Exception("연계획 시트를 찾을 수 없습니다.");
                }

                // D열 = 4, 1월이라고 가정
                for (int month = 1; month <= 12; month++)
                {
                    int col = 3 + month; // D=4

                    string value = ExcelComHelper.GetCellText(ws, 24, col);

                    if (value == "●")
                    {
                        if (month <= targetMonth)
                            quaterCount++;
                        totalCount++;
                    }
                }

                // 절연 점검 포함은 연차, 절연점검 미포함, 접지저항측정은 반기
                string val1 = ExcelComHelper.GetCellText(ws, 13, 3 + targetMonth);
                string val2 = ExcelComHelper.GetCellText(ws, 10, 3 + targetMonth);

                if (val1 == "●" && val2 != "●")
                {
                    isHalfYear = true;
                }

                return quaterCount;
            }
            catch (Exception ex)
            {
                AddLog($"분기 수 계산 중 오류 발생: {ex.Message}");
                return 0;
            }
            finally
            {
                if (totalCount == 0)
                {
                    MessageBox.Show(
                        "연계획 시트에서 분기 수를 찾을 수 없습니다.",
                        "오류",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);

                    isOnlyAnnual = false;
                }
                else
                {
                    isOnlyAnnual = quaterCount == 1;
                }

                nQuater = quaterCount;

                ExcelComHelper.Release(ws);
                ws = null;

                ExcelComHelper.CloseWorkbook(
                    ref wb,
                    false);

                ExcelComHelper.QuitApplication(
                    ref xlApp);

                ExcelComHelper.Cleanup();
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
                AddLog($"이미지 스냅 처리 중 오류 발생: {ex.Message}");
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


        #region 바닥글 이미지 교체
        private void SetRightFooterLogo(
            Excel.Worksheet ws,
            string logoPath)
        {
            if (!File.Exists(logoPath))
                throw new FileNotFoundException(
                    "로고 파일을 찾을 수 없습니다.", logoPath);

            Excel.PageSetup setup = ws.PageSetup;
            Excel.Graphic graphic = setup.RightFooterPicture;

            // 이미지 교체
            graphic.Filename = logoPath;

            // 정확한 크기를 먼저 지정하기 위해 비율 고정 해제
            graphic.LockAspectRatio = Office.MsoTriState.msoFalse;

            // Excel은 point 단위
            // 1cm = 28.3464567 point
            graphic.Width = (float)(2.68 * 28.3464567);
            graphic.Height = (float)(0.53 * 28.3464567);

            // 가로 세로 비율 고정
            graphic.LockAspectRatio = Office.MsoTriState.msoTrue;

            // 오른쪽 바닥글에 그림 표시
            setup.RightFooter = "&G";
        }

        private void ReplaceRightFooterLogo(
            Excel.Worksheet ws,
            string logoPath)
        {
            Excel.PageSetup setup = ws.PageSetup;

            setup.RightFooter = "&G";

            Excel.Graphic graphic = setup.RightFooterPicture;

            // 기존 크기/배율 설정은 그대로 두고 이미지만 교체
            graphic.Filename = logoPath;
        }
        #endregion

        #region 갑지 이미지 중앙 정렬

        public ProcResult relocatePictures(string filePath)
        {

            Excel.Application xlApp = null;
            Excel.Workbook wb = null;
            Excel.Worksheet ws = null;
            Excel.PageSetup pageSetup = null;
            Excel.Range printRange = null;
            Excel.Shapes shapes = null;
            Excel.Shape picture = null;
            Excel.Shape roundRect = null;

            try
            {
                // =====================================================
                // Excel 열기
                // =====================================================

                xlApp = new Excel.Application
                {
                    Visible = false,
                    DisplayAlerts = false
                };

                wb = ExcelComHelper.OpenWorkbook(
                    xlApp,
                    filePath,
                    false);

                ws = ExcelComHelper.GetWorksheet(
                    wb,
                    "갑지");

                if (ws == null)
                    return ProcResult.Fail("갑지 시트를 찾을 수 없습니다.");


                // =====================================================
                // 제목 설정
                // =====================================================

                Match match =
                    Regex.Match(filePath, @"(\d{2})년(\d)분기");

                if (match.Success)
                {
                    string title =
                        $"{2000 + int.Parse(match.Groups[1].Value)}년 " +
                        $"{match.Groups[2].Value}분기" +
                        (filePath.Contains("연차") ? " 연차" : "");

                    ExcelComHelper.SetCellValue(
                        ws,
                        11,
                        1,
                        title);
                }


                // =====================================================
                // PageSetup
                // =====================================================

                pageSetup = ws.PageSetup;

                pageSetup.LeftMargin = 28.35;
                pageSetup.RightMargin = 28.35;
                pageSetup.BottomMargin = 28.35;
                pageSetup.TopMargin = 28.35;

                pageSetup.CenterHorizontally = true;
                pageSetup.CenterVertically = true;


                // =====================================================
                // 인쇄 영역
                // =====================================================

                string printArea = pageSetup.PrintArea;

                if (string.IsNullOrWhiteSpace(printArea))
                {
                    printRange = ws.UsedRange;
                }
                else
                {
                    printRange = ws.Range[printArea];
                }


                double pageLeft = printRange.Left;
                double pageWidth = printRange.Width;

                double centerX =
                    pageLeft + pageWidth / 2.0;

                double centerY =
                    printRange.Top +
                    printRange.Height / 2.0;


                // =====================================================
                // Shape 검색
                // =====================================================

                shapes = ws.Shapes;

                int shapeCount = shapes.Count;

                double minDistance = double.MaxValue;


                // =====================================================
                // 중앙에 가장 가까운 Picture 찾기
                // =====================================================

                for (int i = 1; i <= shapeCount; i++)
                {
                    Excel.Shape shape = null;

                    try
                    {
                        shape = shapes.Item(i);

                        if (shape.Type ==
                            Microsoft.Office.Core.MsoShapeType.msoPicture)
                        {
                            double shapeCenterY =
                                shape.Top +
                                shape.Height / 2.0;

                            double distance =
                                Math.Abs(shapeCenterY - centerY);

                            if (distance < minDistance)
                            {
                                minDistance = distance;

                                // 이전 picture 해제
                                if (picture != null)
                                {
                                    ExcelComHelper.Release(picture);
                                    picture = null;
                                }

                                // 현재 shape를 picture로 넘김
                                picture = shape;
                                shape = null;
                            }
                        }
                    }
                    finally
                    {
                        ExcelComHelper.Release(shape);
                    }
                }


                // =====================================================
                // Picture 좌우 중앙 정렬
                // =====================================================

                if (picture != null)
                {
                    picture.Left =
                        (float)(
                            centerX -
                            picture.Width / 2.0);
                }


                // =====================================================
                // 둥근 사각형 찾기
                // =====================================================

                for (int i = 1; i <= shapeCount; i++)
                {
                    Excel.Shape shape = null;

                    try
                    {
                        shape = shapes.Item(i);

                        if (shape.Type ==
                                Microsoft.Office.Core.MsoShapeType.msoAutoShape &&
                            shape.AutoShapeType ==
                                Microsoft.Office.Core.MsoAutoShapeType
                                    .msoShapeRoundedRectangle)
                        {
                            // 호출자 소유로 넘김
                            roundRect = shape;
                            shape = null;

                            break;
                        }
                    }
                    finally
                    {
                        ExcelComHelper.Release(shape);
                    }
                }


                // =====================================================
                // 둥근 사각형 중앙 정렬
                // =====================================================

                if (roundRect != null)
                {
                    Debug.WriteLine(
                        $"PrintArea={printArea}");

                    Debug.WriteLine(
                        $"Print Left={printRange.Left}");

                    Debug.WriteLine(
                        $"Print Width={printRange.Width}");

                    Debug.WriteLine(
                        $"CenterX={centerX}");

                    Debug.WriteLine(
                        $"Before={roundRect.Left}");

                    roundRect.Left =
                        (float)(
                            centerX -
                            roundRect.Width / 2.0);

                    Debug.WriteLine(
                        $"After={roundRect.Left}");
                }


                // =====================================================
                // 저장
                // =====================================================

                wb.Save();

                return ProcResult.Ok("갑지 이미지 위치 조정이 완료되었습니다.");
            }
            catch (Exception ex)
            {
                return ProcResult.Fail(
                    $"갑지 이미지 위치 조정 중 오류가 발생했습니다.\r\n{ex.Message}");
            }
            finally
            {
                // =====================================================
                // COM 해제
                //
                // 생성한 역순으로 해제
                // =====================================================

                ExcelComHelper.Release(roundRect);
                roundRect = null;

                ExcelComHelper.Release(picture);
                picture = null;

                ExcelComHelper.Release(shapes);
                shapes = null;

                ExcelComHelper.Release(printRange);
                printRange = null;

                ExcelComHelper.Release(pageSetup);
                pageSetup = null;

                ExcelComHelper.Release(ws);
                ws = null;

                ExcelComHelper.CloseWorkbook(
                    ref wb,
                    false);

                ExcelComHelper.QuitApplication(
                    ref xlApp);

                ExcelComHelper.Cleanup();
            }
        }
        #endregion

        #region [페이지 번호 매기기]
        private int GetLastPrintRow(
            Excel.Worksheet ws,
            Excel.PageSetup pageSetup)
        {
            Excel.Range range = null;
            Excel.Range rows = null;

            try
            {
                string printArea =
                    pageSetup.PrintArea;

                if (!string.IsNullOrWhiteSpace(printArea))
                {
                    range = ws.Range[printArea];
                }
                else
                {
                    range = ws.UsedRange;
                }

                int firstRow = range.Row;

                rows = range.Rows;

                int rowCount = rows.Count;

                return firstRow + rowCount - 1;
            }
            finally
            {
                ExcelComHelper.Release(rows);
                ExcelComHelper.Release(range);
            }
        }

        private int GetHorizontalPageCount(
            Excel.Worksheet ws,
            int lastRow)
        {
            Excel.HPageBreaks pageBreaks = null;

            try
            {
                pageBreaks = ws.HPageBreaks;

                int count = pageBreaks.Count;

                int pages = 1;
                int preRow = 0;

                for (int i = 1; i <= count; i++)
                {
                    Excel.HPageBreak pb = null;
                    Excel.Range location = null;

                    try
                    {
                        pb = pageBreaks.Item[i];

                        location = pb.Location;

                        int breakRow =
                            location.Row;

                        Excel.XlPageBreak breakType =
                            pb.Type;

                        // 인쇄영역 바로 다음에 존재하는
                        // 수동 페이지 나누기 무시
                        if ((breakType ==
                                 Excel.XlPageBreak.xlPageBreakManual &&
                             breakRow >= lastRow + 1)
                            ||
                            preRow >= breakRow)
                        {
                            continue;
                        }

                        Debug.WriteLine(
                            $"page : {pages}, row : {breakRow}");

                        pages++;

                        preRow = breakRow;
                    }
                    finally
                    {
                        ExcelComHelper.Release(location);
                        ExcelComHelper.Release(pb);
                    }
                }

                return pages;
            }
            finally
            {
                ExcelComHelper.Release(pageBreaks);
            }
        }

        private void DebugPageBreaks(
            Excel.Worksheet ws)
        {
            Excel.HPageBreaks pageBreaks = null;

            try
            {
                pageBreaks = ws.HPageBreaks;

                int count = pageBreaks.Count;

                for (int i = 1; i <= count; i++)
                {
                    Excel.HPageBreak pb = null;
                    Excel.Range location = null;

                    try
                    {
                        pb = pageBreaks.Item[i];
                        location = pb.Location;

                        Debug.WriteLine(
                            $"Break : {location.Address} / {pb.Type}");
                    }
                    finally
                    {
                        ExcelComHelper.Release(location);
                        ExcelComHelper.Release(pb);
                    }
                }
            }
            finally
            {
                ExcelComHelper.Release(pageBreaks);
            }
        }

        public ProcResult SetPageNumbers(string filePath)
        {
            Excel.Application xlApp = null;
            Excel.Workbook wb = null;
            Excel.Sheets sheets = null;

            try
            {
                if (string.IsNullOrEmpty(filePath) ||
                    !File.Exists(filePath))
                {
                    return ProcResult.Fail("엑셀 파일이 존재하지 않습니다.");
                }

                xlApp = new Excel.Application
                {
                    Visible = false,
                    DisplayAlerts = false
                };

                wb = ExcelComHelper.OpenWorkbook(
                    xlApp,
                    filePath,
                    false);

                sheets = wb.Worksheets;

                int sheetCount = sheets.Count;
                int currentStartPage = 1;

                for (int i = 1; i <= sheetCount; i++)
                {
                    Excel.Worksheet sh = null;
                    Excel.PageSetup pageSetup = null;

                    try
                    {
                        sh = (Excel.Worksheet)sheets[i];

                        string name = sh.Name;

                        // 갑지는 제외
                        if (string.Equals(
                            name,
                            "갑지",
                            StringComparison.OrdinalIgnoreCase))
                        {
                            continue;
                        }

                        sh.Activate();

                        pageSetup = sh.PageSetup;

                        // ============================================
                        // 마지막 인쇄 행
                        // ============================================

                        int lastRow =
                            GetLastPrintRow(
                                sh,
                                pageSetup);

                        // ============================================
                        // 페이지 수 계산
                        // ============================================

                        int pages =
                            GetHorizontalPageCount(
                                sh,
                                lastRow);

                        // ============================================
                        // 시작 페이지 설정
                        // ============================================

                        pageSetup.FirstPageNumber =
                            currentStartPage;

                        Debug.WriteLine(
                            $"{name} : 시작={currentStartPage}, " +
                            $"페이지수={pages}");

                        currentStartPage += pages;
                    }
                    catch (Exception ex)
                    {
                        string sheetName = "(알 수 없음)";

                        try
                        {
                            if (sh != null)
                                sheetName = sh.Name;
                        }
                        catch
                        {
                        }

                        Debug.WriteLine(
                            $"페이지 번호 처리 오류 " +
                            $"idx={i}, " +
                            $"name={sheetName}: " +
                            ex.Message);

                        // 한 시트가 실패해도 다음 시트 계속
                    }
                    finally
                    {
                        ExcelComHelper.Release(pageSetup);
                        pageSetup = null;

                        ExcelComHelper.Release(sh);
                        sh = null;
                    }
                }

                wb.Save();

                return ProcResult.Ok(
                    "페이지 번호 매기기가 완료되었습니다.");
            }
            catch (Exception ex)
            {
                return ProcResult.Fail(
                    $"페이지 번호 매기기 중 오류가 발생했습니다.\r\n{ex.Message}");
            }
            finally
            {
                ExcelComHelper.Release(sheets);
                sheets = null;

                ExcelComHelper.CloseWorkbook(
                    ref wb,
                    false);

                ExcelComHelper.QuitApplication(
                    ref xlApp);

                ExcelComHelper.Cleanup();
            }
        }
        #endregion

        #region [log]
        private void AddLog(string msg)
        {
            try
            {
                if (mainForm != null)
                    mainForm.AddLog("Error", msg);
                else
                    MessageBox.Show(msg, "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (Exception ex)
            {
                MessageBox.Show(msg, "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        #endregion
    }
}
