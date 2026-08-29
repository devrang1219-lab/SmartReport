using System;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace WindowsFormsApp1.Comm
{
    public static class ExcelComHelper
    {
        // ============================================================
        // COM Release
        // ============================================================

        public static void Release(object obj)
        {
            if (obj == null)
                return;

            try
            {
                if (Marshal.IsComObject(obj))
                    Marshal.FinalReleaseComObject(obj);
            }
            catch
            {
                // 종료 과정에서 발생하는 COM 예외는 무시
            }
        }


        // ============================================================
        // Worksheet
        // ============================================================

        /// <summary>
        /// 이름에 특정 문자열이 포함된 Worksheet를 반환.
        /// 반환된 Worksheet는 호출한 쪽에서 Release 해야 함.
        /// </summary>
        public static Excel.Worksheet GetWorksheet(
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

                        string name = sheet.Name;

                        if (name.Trim().IndexOf(
                                sheetName,
                                StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            // 호출자에게 소유권 넘김
                            Excel.Worksheet result = sheet;
                            sheet = null;

                            return result;
                        }
                    }
                    finally
                    {
                        Release(sheet);
                    }
                }

                return null;
            }
            finally
            {
                Release(sheets);
            }
        }


        // ============================================================
        // Cell
        // ============================================================

        public static string GetCellText(
            Excel.Worksheet ws,
            int row,
            int col)
        {
            Excel.Range cells = null;
            Excel.Range cell = null;

            try
            {
                cells = ws.Cells;
                cell = (Excel.Range)cells[row, col];

                return Convert.ToString(cell.Text)?.Trim() ?? "";
            }
            finally
            {
                Release(cell);
                Release(cells);
            }
        }


        public static object GetCellValue(
            Excel.Worksheet ws,
            int row,
            int col)
        {
            Excel.Range cells = null;
            Excel.Range cell = null;

            try
            {
                cells = ws.Cells;
                cell = (Excel.Range)cells[row, col];

                return cell.Value2;
            }
            finally
            {
                Release(cell);
                Release(cells);
            }
        }


        public static void SetCellValue(
            Excel.Worksheet ws,
            int row,
            int col,
            object value)
        {
            Excel.Range cells = null;
            Excel.Range cell = null;

            try
            {
                cells = ws.Cells;
                cell = (Excel.Range)cells[row, col];

                cell.Value2 = value;
            }
            finally
            {
                Release(cell);
                Release(cells);
            }
        }


        // ============================================================
        // Range
        // ============================================================

        public static Excel.Range GetRange(
            Excel.Worksheet ws,
            string address)
        {
            Excel.Range range = null;

            try
            {
                range = ws.Range[address];

                Excel.Range result = range;
                range = null;

                return result;
            }
            finally
            {
                Release(range);
            }
        }


        public static Excel.Range GetRange(
            Excel.Worksheet ws,
            string from,
            string to)
        {
            Excel.Range range = null;

            try
            {
                range = ws.Range[from, to];

                Excel.Range result = range;
                range = null;

                return result;
            }
            finally
            {
                Release(range);
            }
        }


        public static string GetRangeText(
            Excel.Worksheet ws,
            string address)
        {
            Excel.Range range = null;

            try
            {
                range = ws.Range[address];

                return Convert.ToString(range.Text)?.Trim() ?? "";
            }
            finally
            {
                Release(range);
            }
        }


        public static object GetRangeValue(
            Excel.Worksheet ws,
            string address)
        {
            Excel.Range range = null;

            try
            {
                range = ws.Range[address];

                return range.Value2;
            }
            finally
            {
                Release(range);
            }
        }


        public static void SetRangeValue(
            Excel.Worksheet ws,
            string address,
            object value)
        {
            Excel.Range range = null;

            try
            {
                range = ws.Range[address];

                range.Value2 = value;
            }
            finally
            {
                Release(range);
            }
        }


        // ============================================================
        // Workbook Open
        // ============================================================

        /// <summary>
        /// Workbook을 연다.
        /// 반환된 Workbook은 호출자가 Close + Release 해야 함.
        /// </summary>
        public static Excel.Workbook OpenWorkbook(
            Excel.Application app,
            string filePath,
            bool readOnly = false)
        {
            Excel.Workbooks workbooks = null;
            Excel.Workbook wb = null;

            try
            {
                workbooks = app.Workbooks;

                wb = workbooks.Open(
                    filePath,
                    ReadOnly: readOnly);

                Excel.Workbook result = wb;
                wb = null;

                return result;
            }
            finally
            {
                Release(wb);
                Release(workbooks);
            }
        }


        // ============================================================
        // Workbook Close
        // ============================================================

        public static void CloseWorkbook(
            ref Excel.Workbook wb,
            bool saveChanges = false)
        {
            if (wb == null)
                return;

            try
            {
                wb.Close(saveChanges);
            }
            catch
            {
            }
            finally
            {
                Release(wb);
                wb = null;
            }
        }


        // ============================================================
        // Application Quit
        // ============================================================

        public static void QuitApplication(
            ref Excel.Application app)
        {
            if (app == null)
                return;

            try
            {
                app.Quit();
            }
            catch
            {
            }
            finally
            {
                Release(app);
                app = null;
            }
        }


        // ============================================================
        // GC
        // ============================================================

        public static void Cleanup()
        {
            GC.Collect();
            GC.WaitForPendingFinalizers();

            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
    }
}
