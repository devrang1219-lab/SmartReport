using Microsoft.Office.Interop.Excel;
using OpenCvSharp;
using SmartReport;
using System;
using System.Collections.Generic;
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

            mainForm = System.Windows.Forms.Application.OpenForms.OfType<SmartReport.FormMain>().FirstOrDefault();
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

        //public void CopySheetToXlsxAndProcess(string sheetName, float gapLeft, float gapTop, float gapRight, float gapBottom)
        //{
        //    Excel.Application xlApp = null;
        //    Excel.Workbook srcWb = null;
        //    Excel.Workbook newWb = null;

        //    try
        //    {
        //        xlApp = new Excel.Application { Visible = false, DisplayAlerts = false };
        //        srcWb = xlApp.Workbooks.Open(xlsFilePath, ReadOnly: false);

        //        Excel.Worksheet srcWs = GetWorksheetByName(srcWb, sheetName);
        //        if (srcWs == null)
        //            throw new InvalidOperationException($"원본 파일에 '{sheetName}' 시트가 없습니다.");

        //        // 시트 복사(새 워크북으로)
        //        srcWs.Copy(Type.Missing, Type.Missing);
        //        newWb = xlApp.ActiveWorkbook;

        //        // destPath를 xlsFilePath와 같은 폴더로 설정 (xlsFilePath 필드가 있으면 우선 사용)
        //        string destFolder = Path.GetDirectoryName(this.xlsFilePath);

        //        //string destFileName = Path.GetFileNameWithoutExtension(xlsFilePath) + "_" + sheetName + ".xlsx";
        //        string destFileName = sheetName + ".xlsx";
        //        string destPath = Path.Combine(destFolder, destFileName);

        //        if (File.Exists(destPath))
        //            File.Delete(destPath); // 덮어쓰기

        //        // .xlsx 형식으로 저장
        //        newWb.SaveAs(destPath, Excel.XlFileFormat.xlOpenXMLWorkbook);

        //        // 새 워크북의 첫 시트에서 처리
        //        Excel.Worksheet newWs = (Excel.Worksheet)newWb.Sheets[1];
        //        SnapImageMergedCell(srcWs, sheetName);
        //        SnapImageMergedCell(newWs, sheetName);
        //    }
        //    finally
        //    {
        //        if (srcWb != null) { srcWb.Save(); srcWb.Close(false); Marshal.ReleaseComObject(srcWb); }
        //        if (newWb != null) { newWb.Save(); newWb.Close(false); Marshal.ReleaseComObject(newWb); }
        //        if (xlApp != null) { xlApp.Quit(); Marshal.ReleaseComObject(xlApp); }
        //        GC.Collect();
        //        GC.WaitForPendingFinalizers();

        //        Cursor.Current = Cursors.Default;
        //    }
        //}
        public void CopySheetToXlsxAndProcess(
            string sheetName,
            float gapLeft,
            float gapTop,
            float gapRight,
            float gapBottom)
        {
            Excel.Application xlApp = null;
            Excel.Workbooks workbooks = null;
            Excel.Workbook srcWb = null;
            Excel.Workbook newWb = null;

            Excel.Worksheet srcWs = null;
            Excel.Worksheet newWs = null;

            Excel.Sheets newSheets = null;

            try
            {
                // =====================================================
                // 1. Excel Application
                // =====================================================
                xlApp = new Excel.Application
                {
                    Visible = false,
                    DisplayAlerts = false
                };

                // =====================================================
                // 2. Workbooks COM 객체를 명시적으로 보관
                // =====================================================
                workbooks = xlApp.Workbooks;


                // =====================================================
                // 2. 원본 Workbook 열기
                // =====================================================
                srcWb = workbooks.Open(
                    xlsFilePath,
                    ReadOnly: false);


                // =====================================================
                // 3. 원본 Sheet 가져오기
                // =====================================================
                srcWs = GetWorksheetByName(
                    srcWb,
                    sheetName);

                if (srcWs == null)
                {
                    throw new InvalidOperationException(
                        $"원본 파일에 '{sheetName}' 시트가 없습니다.");
                }


                // =====================================================
                // 4. Sheet를 새로운 Workbook으로 복사
                // =====================================================
                srcWs.Copy(
                    Type.Missing,
                    Type.Missing);


                // Copy 후 ActiveWorkbook = 새 Workbook
                newWb = xlApp.ActiveWorkbook;

                if (newWb == null)
                {
                    throw new InvalidOperationException(
                        "시트 복사 후 새 Workbook을 가져오지 못했습니다.");
                }


                //// =====================================================
                //// 5. 저장 경로
                //// =====================================================
                string destFolder =
                    Path.GetDirectoryName(xlsFilePath);

                string destFileName =
                    sheetName + ".xlsx";

                string destPath =
                    Path.Combine(
                        destFolder,
                        destFileName);


                //// =====================================================
                //// 6. 기존 파일 삭제
                //// =====================================================
                if (File.Exists(destPath))
                {
                    File.Delete(destPath);
                }


                //// =====================================================
                //// 7. 새 Workbook을 XLSX로 저장
                //// =====================================================
                newWb.SaveAs(
                    destPath,
                    Excel.XlFileFormat.xlOpenXMLWorkbook);

                // =====================================================
                // 10. Sheets COM 객체를 별도로 관리
                // =====================================================
                newSheets = newWb.Sheets;


                //// =====================================================
                //// 8. 새 Workbook의 첫 번째 Sheet 가져오기
                //// =====================================================
                newWs =
                 (Excel.Worksheet)newSheets[1];


                if (newWs == null)
                {
                    throw new InvalidOperationException(
                        "새 Workbook의 Sheet를 가져오지 못했습니다.");
                }


                // =====================================================
                // 9. 새 Sheet의 이미지 위치/크기 조정
                // =====================================================
                SnapImageMergedCell(
                    newWs,
                    sheetName,
                    gapLeft,
                    gapTop,
                    gapRight,
                    gapBottom);
                SnapImageMergedCell(
                    srcWs,
                    sheetName,
                    gapLeft,
                    gapTop,
                    gapRight,
                    gapBottom);


                // =====================================================
                // 10. 처리 결과 저장
                // =====================================================

                srcWb.Save();
                newWb.Save();
            }
            catch (Exception ex)
            {
                AddLog(
                    $"시트 복사 및 이미지 처리 중 오류 발생: {ex.Message}");

                MessageBox.Show(
                    $"시트 복사 및 처리 중 오류가 발생했습니다.\r\n\r\n{ex.Message}",
                    "오류",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
            finally
            {
                // =====================================================
                // 11. Worksheet COM 해제
                // =====================================================
                if (newWs != null)
                {
                    try
                    {
                        Marshal.ReleaseComObject(newWs);
                    }
                    catch
                    {
                    }

                    newWs = null;
                }

                if (srcWs != null)
                {
                    try
                    {
                        Marshal.ReleaseComObject(srcWs);
                    }
                    catch
                    {
                    }

                    srcWs = null;
                }

                if (newSheets != null)
                {
                    try
                    {
                        Marshal.FinalReleaseComObject(newSheets);
                    }
                    catch
                    {
                    }

                    newSheets = null;
                }



                // =====================================================
                // 12. 새 Workbook 저장 및 종료
                // =====================================================
                if (newWb != null)
                {
                    try
                    {
                        newWb.Save();
                    }
                    catch
                    {
                    }

                    try
                    {
                        newWb.Close(
                            SaveChanges: false);
                    }
                    catch
                    {
                    }

                    try
                    {
                        Marshal.ReleaseComObject(newWb);
                    }
                    catch
                    {
                    }

                    newWb = null;
                }


                // =====================================================
                // 13. 원본 Workbook 종료
                // =====================================================
                if (srcWb != null)
                {
                    try
                    {
                        srcWb.Close(
                            SaveChanges: false);
                    }
                    catch
                    {
                    }

                    try
                    {
                        Marshal.ReleaseComObject(srcWb);
                    }
                    catch
                    {
                    }

                    srcWb = null;
                }

                // =====================================================
                // 2. Workbooks 컬렉션 해제
                // =====================================================
                if (workbooks != null)
                {
                    try
                    {
                        Marshal.FinalReleaseComObject(workbooks);
                    }
                    catch
                    {
                    }

                    workbooks = null;
                }


                // =====================================================
                // 14. Excel 종료
                // =====================================================
                if (xlApp != null)
                {
                    try
                    {
                        xlApp.Quit();
                    }
                    catch
                    {
                    }

                    try
                    {
                        Marshal.ReleaseComObject(xlApp);
                    }
                    catch
                    {
                    }

                    xlApp = null;
                }


                // =====================================================
                // 15. 남아 있는 COM RCW 정리
                // =====================================================
                GC.Collect();
                GC.WaitForPendingFinalizers();

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

        public void SnapImageMergedCell(
            Excel.Worksheet ws,
            string sheetName,
            float gapLeft = 1.5f,
            float gapTop = 1.5f,
            float gapRight = 0f,
            float gapBottom = 0.5f)
        {
            Excel.Application xlApp = null;
            Excel.Workbook wb = null;
            Excel.Worksheet targetWs = null;
            Excel.Shapes shapes = null;

            bool openedHere = false;

            try
            {
                // =====================================================
                // 1. Worksheet 준비
                // =====================================================
                if (ws == null)
                {
                    if (string.IsNullOrEmpty(xlsFilePath) ||
                        !File.Exists(xlsFilePath))
                    {
                        MessageBox.Show(
                            "엑셀 파일 경로가 없습니다. xlsFilePath를 설정하세요.",
                            "오류",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Warning);

                        return;
                    }

                    xlApp = new Excel.Application
                    {
                        Visible = false,
                        DisplayAlerts = false
                    };

                    wb = xlApp.Workbooks.Open(
                        xlsFilePath,
                        ReadOnly: false);

                    targetWs = GetWorksheetByName(
                        wb,
                        sheetName);

                    if (targetWs == null)
                    {
                        MessageBox.Show(
                            $"'{sheetName}' 시트를 찾을 수 없습니다.",
                            "오류",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Warning);

                        return;
                    }

                    openedHere = true;
                }
                else
                {
                    targetWs = ws;
                }


                // =====================================================
                // 2. Shapes 가져오기
                // =====================================================
                shapes = targetWs.Shapes;

                int shapeCount = shapes.Count;


                // =====================================================
                // 3. 모든 도형 처리
                // =====================================================
                for (int i = 1; i <= shapeCount; i++)
                {
                    Excel.Shape shape = null;
                    Excel.Range topLeft = null;
                    Excel.Range bottomRight = null;
                    Excel.Range area = null;

                    try
                    {
                        // -------------------------------------------------
                        // Shape
                        // -------------------------------------------------
                        shape = shapes.Item(i);

                        // -------------------------------------------------
                        // 왼쪽 위 셀
                        // -------------------------------------------------
                        topLeft = shape.TopLeftCell;


                        // =================================================
                        // 4. 병합셀 확인
                        // =================================================
                        bool isMerged = false;

                        try
                        {
                            object mergeCells = topLeft.MergeCells;

                            if (mergeCells is bool b)
                                isMerged = b;
                        }
                        catch
                        {
                            isMerged = false;
                        }


                        // =================================================
                        // 5. 실제 기준 영역 결정
                        // =================================================
                        if (isMerged)
                        {
                            // 병합셀인 경우
                            // TopLeftCell의 MergeArea 전체를 기준으로 사용
                            area = topLeft.MergeArea;
                        }
                        else
                        {
                            // 일반 셀인 경우
                            bottomRight = shape.BottomRightCell;

                            area = targetWs.Range[
                                topLeft,
                                bottomRight];
                        }


                        // =================================================
                        // 6. 회전 확인
                        // =================================================
                        double rot = 0;

                        try
                        {
                            rot = shape.Rotation % 360;

                            if (rot < 0)
                                rot += 360;
                        }
                        catch
                        {
                            rot = 0;
                        }

                        bool rotated90 =
                            Math.Abs(rot - 90) < 1.0 ||
                            Math.Abs(rot - 270) < 1.0;


                        // =================================================
                        // 7. 사진 위치 및 크기 조정
                        // =================================================
                        if (rotated90)
                        {
                            // -------------------------------------------------
                            // 90 / 270도 회전된 이미지
                            // Width / Height 교체
                            // -------------------------------------------------
                            float targetW =
                                (float)area.Height;

                            float targetH =
                                (float)area.Width;


                            float newWidth =
                                targetW -
                                gapLeft -
                                gapRight;

                            float newHeight =
                                targetH -
                                gapTop -
                                gapBottom;


                            // 음수 방지
                            if (newWidth < 0)
                                newWidth = 0;

                            if (newHeight < 0)
                                newHeight = 0;


                            shape.Width = newWidth;
                            shape.Height = newHeight;


                            shape.Left =
                                (float)area.Left +
                                ((float)area.Width - targetW) / 2f +
                                gapLeft;

                            shape.Top =
                                (float)area.Top +
                                ((float)area.Height - targetH) / 2f +
                                gapTop;
                        }
                        else
                        {
                            // -------------------------------------------------
                            // 일반 이미지
                            // -------------------------------------------------
                            float newWidth =
                                (float)area.Width -
                                gapLeft -
                                gapRight;

                            float newHeight =
                                (float)area.Height -
                                gapTop -
                                gapBottom;


                            // 음수 방지
                            if (newWidth < 0)
                                newWidth = 0;

                            if (newHeight < 0)
                                newHeight = 0;


                            shape.Left =
                                (float)area.Left +
                                gapLeft;

                            shape.Top =
                                (float)area.Top +
                                gapTop;

                            shape.Width =
                                newWidth;

                            shape.Height =
                                newHeight;
                        }


                        // =================================================
                        // 8. 비율 고정 해제
                        // =================================================
                        shape.LockAspectRatio =
                            Microsoft.Office.Core.MsoTriState.msoFalse;


                        // =================================================
                        // 9. 셀 크기/위치 변경 시 이미지도 같이 이동
                        // =================================================
                        shape.Placement =
                            Excel.XlPlacement.xlMoveAndSize;
                    }
                    catch (Exception ex)
                    {
                        AddLog(
                            $"이미지 스냅 처리 중 오류 발생 ({i}): {ex.Message}");
                    }
                    finally
                    {
                        // =================================================
                        // COM 객체 해제
                        // =================================================

                        if (area != null)
                        {
                            Marshal.ReleaseComObject(area);
                            area = null;
                        }

                        if (bottomRight != null)
                        {
                            Marshal.ReleaseComObject(bottomRight);
                            bottomRight = null;
                        }

                        if (topLeft != null)
                        {
                            Marshal.ReleaseComObject(topLeft);
                            topLeft = null;
                        }

                        if (shape != null)
                        {
                            Marshal.ReleaseComObject(shape);
                            shape = null;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                AddLog(
                    $"이미지 스냅 처리 중 오류 발생: {ex.Message}");
            }
            finally
            {
                // =====================================================
                // Shapes 해제
                // =====================================================
                if (shapes != null)
                {
                    Marshal.ReleaseComObject(shapes);
                    shapes = null;
                }


                // =====================================================
                // 이 함수에서 직접 Excel을 열었다면 여기서 종료
                // =====================================================
                if (openedHere)
                {
                    // -------------------------------------------------
                    // Worksheet
                    // -------------------------------------------------
                    if (targetWs != null)
                    {
                        Marshal.ReleaseComObject(targetWs);
                        targetWs = null;
                    }


                    // -------------------------------------------------
                    // Workbook
                    // -------------------------------------------------
                    if (wb != null)
                    {
                        try
                        {
                            wb.Save();
                        }
                        catch
                        {
                        }

                        try
                        {
                            wb.Close(
                                SaveChanges: false);
                        }
                        catch
                        {
                        }

                        Marshal.ReleaseComObject(wb);
                        wb = null;
                    }


                    // -------------------------------------------------
                    // Excel Application
                    // -------------------------------------------------
                    if (xlApp != null)
                    {
                        try
                        {
                            xlApp.Quit();
                        }
                        catch
                        {
                        }

                        Marshal.ReleaseComObject(xlApp);
                        xlApp = null;
                    }


                    // -------------------------------------------------
                    // COM RCW 최종 정리
                    // -------------------------------------------------
                    GC.Collect();
                    GC.WaitForPendingFinalizers();

                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                }
            }
        }

        #region [엑셀 시트에서 그림 삭제]
        public void RemovePictures(Excel.Worksheet ws)
        {
            Excel.Shapes shapes = null;

            try
            {
                shapes = ws.Shapes;

                for (int i = shapes.Count; i >= 1; i--)
                {
                    Excel.Shape shape = null;

                    try
                    {
                        shape = shapes.Item(i);

                        if (shape.Type == Office.MsoShapeType.msoPicture ||
                            shape.Type == Office.MsoShapeType.msoLinkedPicture)
                        {
                            shape.Delete();
                        }
                    }
                    finally
                    {
                        if (shape != null)
                        {
                            Marshal.ReleaseComObject(shape);
                            shape = null;
                        }
                    }
                }
            }
            finally
            {
                if (shapes != null)
                {
                    Marshal.ReleaseComObject(shapes);
                    shapes = null;
                }
            }
        }

        public void RemovePictures(Excel.Worksheet ws, int? keepPictureIndex = null)
        {
            var pictureNames = new List<string>();
            Excel.Shapes shapes = ws.Shapes;

            if (keepPictureIndex == null) return;
            if (keepPictureIndex < 2) return;

            // 이미지 이름 수집
            for (int i = 1; i <= ws.Shapes.Count; i++)
            {
                Excel.Shape shape = null;

                try
                {
                    shape = shapes.Item(i);
                    if (shape.Type == Office.MsoShapeType.msoPicture ||
                        shape.Type == Office.MsoShapeType.msoLinkedPicture)
                    {
                        pictureNames.Add(shape.Name);
                    }
                }
                finally
                {
                    if (shape != null) Marshal.ReleaseComObject(shape);
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

        #region [열화상 이미지 분기 시트 삽입]
        private void EnsureFeverPicturePages(
            Excel.Worksheet ws,
            int imageCount)
        {
            if (ws == null || imageCount <= 0)
                return;

            const int PAGE_ROW_OFFSET = 56;

            // 한 페이지의 사진 위치
            const int IMAGE_START_ROW = 27;
            const int IMAGE_END_ROW = 43;

            // 보고서 마지막 열
            const string LAST_COLUMN = "AD";

            Excel.Range printAreaRange = null;

            try
            {
                //----------------------------------------------------------
                // 1. 사진이 들어가야 하는 마지막 행 계산
                //----------------------------------------------------------

                // 페이지당 사진 2장
                int requiredPages =
                    (int)Math.Ceiling(imageCount / 2.0);

                // 예:
                // 1 page -> 43
                // 2 page -> 99
                // 3 page -> 155
                // 4 page -> 211
                int requiredImageLastRow =
                    IMAGE_END_ROW +
                    ((requiredPages - 1) * PAGE_ROW_OFFSET);


                //----------------------------------------------------------
                // 2. 현재 PrintArea의 마지막 행 확인
                //----------------------------------------------------------

                string printArea = ws.PageSetup.PrintArea;

                if (string.IsNullOrWhiteSpace(printArea))
                {
                    throw new Exception(
                        $"[{ws.Name}] 인쇄영역(PrintArea)이 설정되어 있지 않습니다.");
                }

                printAreaRange = ws.Range[printArea];

                int printStartRow = printAreaRange.Row;

                int currentPrintLastRow =
                    printAreaRange.Row +
                    printAreaRange.Rows.Count - 1;


                Debug.WriteLine(
                    $"[{ws.Name}] " +
                    $"현재 PrintArea 마지막 행={currentPrintLastRow}, " +
                    $"사진 필요 마지막 행={requiredImageLastRow}");


                //----------------------------------------------------------
                // 3. 이미 충분하면 아무것도 하지 않음
                //----------------------------------------------------------

                if (requiredImageLastRow <= currentPrintLastRow)
                {
                    Debug.WriteLine(
                        $"[{ws.Name}] 페이지 추가 필요 없음.");

                    return;
                }


                //----------------------------------------------------------
                // 4. 필요한 만큼 페이지 추가
                //----------------------------------------------------------

                while (requiredImageLastRow > currentPrintLastRow)
                {
                    Excel.Range sourceRange = null;
                    Excel.Range destRange = null;
                    Excel.Range breakCell = null;

                    try
                    {
                        /*
                         * 현재 마지막 페이지 56행을 복사
                         *
                         * 예:
                         *
                         * PrintArea가 A1:AD168 이라면
                         *
                         * 마지막 페이지:
                         * 113 ~ 168
                         *
                         * 복사 위치:
                         * 169 ~ 224
                         */

                        int sourceStartRow =
                            currentPrintLastRow -
                            PAGE_ROW_OFFSET + 1;

                        int sourceEndRow =
                            currentPrintLastRow;

                        int destStartRow =
                            currentPrintLastRow + 1;

                        int destEndRow =
                            currentPrintLastRow +
                            PAGE_ROW_OFFSET;


                        Debug.WriteLine(
                            $"[{ws.Name}] 페이지 복사: " +
                            $"{sourceStartRow}~{sourceEndRow} " +
                            $"→ {destStartRow}~{destEndRow}");


                        //--------------------------------------------------
                        // 마지막 페이지 복사
                        //--------------------------------------------------

                        sourceRange = ws.Range[
                            $"A{sourceStartRow}",
                            $"{LAST_COLUMN}{sourceEndRow}"
                        ];

                        destRange = ws.Range[
                            $"A{destStartRow}",
                            $"{LAST_COLUMN}{destEndRow}"
                        ];

                        sourceRange.Copy(destRange);


                        //--------------------------------------------------
                        // 행 높이 복사
                        //
                        // Range.Copy만으로 행 높이가 정확히 복사되지
                        // 않는 경우가 있어서 별도로 맞춤
                        //--------------------------------------------------

                        for (int i = 0; i < PAGE_ROW_OFFSET; i++)
                        {
                            Excel.Range srcRow = null;
                            Excel.Range dstRow = null;

                            try
                            {
                                srcRow = ws.Rows[sourceStartRow + i];
                                dstRow = ws.Rows[destStartRow + i];

                                dstRow.RowHeight = srcRow.RowHeight;
                            }
                            finally
                            {
                                if (dstRow != null)
                                    Marshal.ReleaseComObject(dstRow);

                                if (srcRow != null)
                                    Marshal.ReleaseComObject(srcRow);
                            }
                        }


                        //--------------------------------------------------
                        // 새 페이지 시작 위치에 페이지 나누기
                        //--------------------------------------------------

                        breakCell = ws.Cells[destStartRow, 1];

                        bool breakExists = false;

                        int breakCount = ws.HPageBreaks.Count;

                        for (int i = 1; i <= breakCount; i++)
                        {
                            Excel.HPageBreak hp = null;
                            Excel.Range location = null;

                            try
                            {
                                hp = ws.HPageBreaks[i];

                                location = hp.Location;

                                if (location.Row == destStartRow)
                                {
                                    breakExists = true;
                                    break;
                                }
                            }
                            finally
                            {
                                if (location != null)
                                    Marshal.ReleaseComObject(location);

                                if (hp != null)
                                    Marshal.ReleaseComObject(hp);
                            }
                        }

                        if (!breakExists)
                        {
                            ws.HPageBreaks.Add(
                                Before: breakCell);
                        }


                        //--------------------------------------------------
                        // PrintArea 확장
                        //--------------------------------------------------

                        currentPrintLastRow = destEndRow;

                        ws.PageSetup.PrintArea =
                            $"$A${printStartRow}:${LAST_COLUMN}${currentPrintLastRow}";


                        Debug.WriteLine(
                            $"[{ws.Name}] PrintArea 확장 → " +
                            $"A{printStartRow}:{LAST_COLUMN}{currentPrintLastRow}");
                    }
                    finally
                    {
                        if (breakCell != null)
                            Marshal.ReleaseComObject(breakCell);

                        if (destRange != null)
                            Marshal.ReleaseComObject(destRange);

                        if (sourceRange != null)
                            Marshal.ReleaseComObject(sourceRange);
                    }
                }


                //----------------------------------------------------------
                // 5. 최종 상태
                //----------------------------------------------------------

                Debug.WriteLine(
                    $"[{ws.Name}] 페이지 확장 완료. " +
                    $"최종 마지막 행={currentPrintLastRow}");
            }
            finally
            {
                if (printAreaRange != null)
                    Marshal.ReleaseComObject(printAreaRange);
            }
        }

        public void ProcFeverPicture(
            string sheetName, 
            bool bCheckBoxOcr,
            Excel.Application xlApp,
            Excel.Workbook wb,
            string baseFolder,
            string pictureFolder)
        {
            Excel.Worksheet sourceWs = null;

            try
            {
                sourceWs = ExcelComHelper.GetWorksheetByLastName(wb, sheetName);

                if (sourceWs == null)
                {
                    throw new Exception("분기 시트를 찾을 수 없습니다.");
                }

                // =====================================================
                // 기존 그림 제거
                // =====================================================
                RemovePictures(sourceWs);

                // =====================================================
                // 이미지 파일
                // =====================================================
                string picturePath =
                    Path.Combine(baseFolder, pictureFolder);

                string[] files = Directory
                    .GetFiles(picturePath, "*.jpg")
                    .OrderBy(f =>
                    {
                        return int.TryParse(
                            Path.GetFileNameWithoutExtension(f),
                            out int n)
                            ? n
                            : int.MaxValue;
                    })
                    .ToArray();

                // =====================================================
                // 필요한 페이지 생성
                // =====================================================
                EnsureFeverPicturePages(sourceWs, files.Length);

                // =====================================================
                // 이미지 삽입
                // =====================================================
                int imageIndex = 0;

                for (int page = 0;
                     imageIndex < files.Length;
                     page++)
                {
                    int startRow = 27 + (page * 56);
                    int endRow = 43 + (page * 56);

                    string[] fromCols = { "A", "R" };
                    string[] toCols = { "Q", "AC" };

                    for (int i = 0;
                         i < fromCols.Length;
                         i++)
                    {
                        if (imageIndex >= files.Length)
                            break;

                        string fromCell =
                            $"{fromCols[i]}{startRow}";

                        string toCell =
                            $"{toCols[i]}{endRow}";

                        // =================================================
                        // 원본 통합문서
                        // =================================================
                        using (var inserter =
                            new ImageInserter(
                                sourceWs,
                                files[imageIndex]))
                        {
                            inserter.InsertFit(
                                fromCell,
                                toCell,
                                new ImageInsertOptions
                                {
                                    KeepAspectRatio = false
                                });
                        }

                        imageIndex++;
                    }
                }

                // =====================================================
                // OCR
                // =====================================================
                if (bCheckBoxOcr)
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
                    }
                }

                // =====================================================
                // 저장
                // =====================================================
                wb.Save();
            }
            catch (Exception ex)
            {
                AddLog(
                    $"분기 이미지 삽입 실패: {ex.Message}");
            }
            finally
            {
                AddLog(
                    "분기 이미지 삽입 완료", "Info");
                // -----------------------------------------------------
                // sourceWs
                // -----------------------------------------------------
                if (sourceWs != null)
                {
                    try
                    {
                        Marshal.FinalReleaseComObject(sourceWs);
                    }
                    catch { }

                    sourceWs = null;
                }
            }
        }
        #endregion


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

        private void ReplaceA13FromSample(
            Excel.Worksheet targetWs,
            string sampleFilePath)
        {
            Excel.Application app = null;
            Excel.Workbook sampleWb = null;
            Excel.Worksheet sampleWs = null;

            Excel.Range targetCell = null;
            Excel.Range sampleCell = null;

            Excel.Shapes targetShapes = null;
            Excel.Shapes sampleShapes = null;

            try
            {
                if (!File.Exists(sampleFilePath))
                    throw new FileNotFoundException(
                        "sample.xlsx 파일을 찾을 수 없습니다.",
                        sampleFilePath);

                // =====================================================
                // 대상 A13
                // =====================================================

                targetCell = targetWs.Range["A13"];

                // 병합셀이라면 MergeArea 전체를 대상으로 사용
                if (targetCell.MergeCells)
                {
                    Excel.Range mergeArea = targetCell.MergeArea;

                    ExcelComHelper.Release(targetCell);
                    targetCell = mergeArea;
                }


                // =====================================================
                // 현재 시트의 A13 영역에 걸쳐 있는 Shape 제거
                // =====================================================

                targetShapes = targetWs.Shapes;

                // 삭제하면서 Count가 변하므로 뒤에서부터 처리
                for (int i = targetShapes.Count; i >= 1; i--)
                {
                    Excel.Shape shape = null;

                    try
                    {
                        shape = targetShapes.Item(i);

                        double shapeLeft = shape.Left;
                        double shapeTop = shape.Top;
                        double shapeRight =
                            shape.Left + shape.Width;
                        double shapeBottom =
                            shape.Top + shape.Height;

                        double cellLeft = targetCell.Left;
                        double cellTop = targetCell.Top;
                        double cellRight =
                            targetCell.Left + targetCell.Width;
                        double cellBottom =
                            targetCell.Top + targetCell.Height;

                        // Shape와 A13 영역이 겹치는지 검사
                        bool overlaps =
                            shapeLeft < cellRight &&
                            shapeRight > cellLeft &&
                            shapeTop < cellBottom &&
                            shapeBottom > cellTop;

                        if (overlaps)
                        {
                            Debug.WriteLine(
                                $"A13 영역의 Shape 삭제: {shape.Name}");

                            shape.Delete();
                        }
                    }
                    finally
                    {
                        ExcelComHelper.Release(shape);
                    }
                }


                // =====================================================
                // sample.xlsx 열기
                // =====================================================

                app = new Excel.Application
                {
                    Visible = false,
                    DisplayAlerts = false
                };

                sampleWb = ExcelComHelper.OpenWorkbook(
                    app,
                    sampleFilePath,
                    true);       // ReadOnly

                sampleWs = ExcelComHelper.GetWorksheet(
                    sampleWb,
                    targetWs.Name);

                if (sampleWs == null)
                    throw new Exception(
                        $"sample.xlsx에서 '{targetWs.Name}' 시트를 찾을 수 없습니다.");


                // =====================================================
                // sample A13 셀 값 복사
                // =====================================================

                sampleCell = sampleWs.Range["A13"];

                targetCell.Value = sampleCell.Value;


                // =====================================================
                // sample A13에 걸쳐 있는 Shape 찾기
                // =====================================================

                sampleShapes = sampleWs.Shapes;

                for (int i = 1; i <= sampleShapes.Count; i++)
                {
                    Excel.Shape sampleShape = null;
                    Excel.Range compareCell = sampleCell;


                    try
                    {
                        sampleShape = sampleShapes.Item(i);

                        if ((bool)sampleCell.MergeCells)
                        {
                            compareCell = sampleCell.MergeArea;
                        }

                        double shapeLeft = sampleShape.Left;
                        double shapeTop = sampleShape.Top;
                        double shapeRight =
                            sampleShape.Left + sampleShape.Width;
                        double shapeBottom =
                            sampleShape.Top + sampleShape.Height;

                        double cellLeft = compareCell.Left;
                        double cellTop = compareCell.Top;
                        double cellRight = cellLeft + compareCell.Width;
                        double cellBottom = cellTop + compareCell.Height;

                        bool overlaps =
                            shapeLeft < cellRight &&
                            shapeRight > cellLeft &&
                            shapeTop < cellBottom &&
                            shapeBottom > cellTop;

                        if (!overlaps)
                            continue;


                        // =================================================
                        // 이미지 복사
                        // =================================================

                        float left = sampleShape.Left;
                        float top = sampleShape.Top;
                        float width = sampleShape.Width;
                        float height = sampleShape.Height;

                        sampleShape.Copy();

                        // 현재 시트에 붙여넣기
                        targetWs.Paste();

                        // 방금 붙여넣은 Shape
                        Excel.Shape newShape = null;

                        try
                        {
                            newShape =
                                targetWs.Shapes.Item(
                                    targetWs.Shapes.Count);

                            // sample과 동일한 위치/크기
                            newShape.Left = left;
                            newShape.Top = top;
                            newShape.Width = width;
                            newShape.Height = height;
                        }
                        finally
                        {
                            ExcelComHelper.Release(newShape);
                        }

                        // A13에 해당하는 이미지는 하나만 복사
                        break;
                    }
                    finally
                    {
                        ExcelComHelper.Release(sampleShape);
                    }
                }
            }
            finally
            {
                ExcelComHelper.Release(sampleShapes);
                ExcelComHelper.Release(targetShapes);

                ExcelComHelper.Release(sampleCell);
                ExcelComHelper.Release(targetCell);

                ExcelComHelper.Release(sampleWs);

                ExcelComHelper.CloseWorkbook(
                    ref sampleWb,
                    false);

                ExcelComHelper.QuitApplication(
                    ref app);

                ExcelComHelper.Cleanup();
            }
        }

        public ProcResult relocatePictures(string filePath, string sampleFilePath = null)
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
                // A13 내용을 sample.xlsx 기준으로 교체
                // =====================================================

                //if (!string.IsNullOrEmpty(sampleFilePath))
                //{
                //    ReplaceA13FromSample(
                //        ws,
                //        sampleFilePath);
                //}


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
                    // picture.Shadow.Visible = Microsoft.Office.Core.MsoTriState.msoFalse;

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
        private void AddLog(string msg, string part = "Error")
        {
            try
            {
                if (mainForm != null)
                    mainForm.AddLog(part, msg);
                else
                    MessageBox.Show(msg, part, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (Exception ex)
            {
                MessageBox.Show(msg, part, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        #endregion
    }
}
