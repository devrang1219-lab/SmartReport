using System;
using Excel = Microsoft.Office.Interop.Excel;

namespace SmartReport
{
    public class OpinionMaker
    {
        public static void FillOpinionFromInsulation(Excel.Workbook wb)
        {
            Excel.Worksheet wsSrc = wb.Worksheets["절연"];   // 원본
            Excel.Worksheet wsDst = wb.Worksheets["의견"];   // 결과

            int lastRow = wsSrc.Cells[wsSrc.Rows.Count, "A"].End[Excel.XlDirection.xlUp].Row;

            // 의견 시트 16행부터 삽입
            int insertRow = 16;

            // 검사 대상 열: I(9), S(19)
            int[] targetCols = { 9, 19 };

            for (int r = 1; r <= lastRow; r++)
            {
                foreach (int c in targetCols)
                {
                    string cellValue = Convert.ToString((wsSrc.Cells[r, c] as Excel.Range)?.Value2)?.Trim();

                    if (cellValue == "점검요")
                    {
                        // 1) 블록 설비명 찾기
                        string equipName = FindBlockEquipmentName(wsSrc, r);

                        // 2) H열에 넣을 문구 만들기
                        string side = GetCellText(wsSrc, r, c - 8);   // 우
                        string no = GetCellText(wsSrc, r, c - 7);     // 2
                        string type = GetCellText(wsSrc, r, c - 6);   // MCCB
                        string value = GetCellText(wsSrc, r, c - 2);  // 0

                        string defectText = $"{side}측 {no}번패 {type} 절연저항 {value} MΩ";

                        // 3) 페이지 번호 구하기
                        int pageNo = GetPageNumber(wsSrc, r, c);

                        // 4) 의견 시트 16행에 행 삽입
                        Excel.Range rowRange = (Excel.Range)wsDst.Rows[insertRow];
                        rowRange.Insert(Excel.XlInsertShiftDirection.xlShiftDown);

                        // 필요 값 입력
                        wsDst.Cells[insertRow, "C"] = equipName;     // 부적합 설비
                        wsDst.Cells[insertRow, "H"] = defectText;    // 부적합 내용
                        wsDst.Cells[insertRow, "T"] = pageNo;        // 페이지 번호

                        // 순번(A열)도 넣고 싶으면 아래 추가
                        // wsDst.Cells[insertRow, "A"] = insertRow - 14;

                        insertRow++;
                    }
                }
            }
        }

        /// <summary>
        /// 점검요가 있는 행이 속한 블록의 설비명 찾기
        /// 예: 위로 올라가며 B열/또는 특정 열에 있는 "ME-104" 같은 제목 찾기
        /// </summary>
        private static string FindBlockEquipmentName(Excel.Worksheet ws, int currentRow)
        {
            // 현재 화면 기준으로는 블록 제목이 대체로 B~C 부근에 있음
            // 위로 올라가며 B열 값이 있으면 설비명으로 판단
            for (int r = currentRow; r >= 1; r--)
            {
                string b = GetCellText(ws, r, 2); // B열
                string c = GetCellText(ws, r, 3); // C열

                // 예: "ME-104" 같은 블록명
                if (!string.IsNullOrWhiteSpace(b) && IsEquipmentName(b))
                    return b;

                if (!string.IsNullOrWhiteSpace(c) && IsEquipmentName(c))
                    return c;
            }

            return "";
        }

        private static bool IsEquipmentName(string text)
        {
            // 필요하면 더 엄격하게 바꿔도 됨
            // 예: "ME-104", "PT 절연유" 같은 블록명 판별
            if (string.IsNullOrWhiteSpace(text)) return false;

            // 숫자/하이픈 포함된 설비명 우선
            if (text.Contains("-")) return true;

            // PT 절연유 같은 경우도 허용
            if (text.Contains("PT") || text.Contains("VCB") || text.Contains("TR") || text.Contains("판넬"))
                return true;

            return false;
        }

        private static string GetCellText(Excel.Worksheet ws, int row, int col)
        {
            if (row < 1 || col < 1) return "";

            object v = (ws.Cells[row, col] as Excel.Range)?.Value2;
            if (v == null) return "";

            // 숫자가 2.0처럼 들어오면 보기 좋게 정리
            if (v is double d)
            {
                if (Math.Abs(d - Math.Round(d)) < 0.0000001)
                    return ((int)Math.Round(d)).ToString();
                return d.ToString();
            }

            return v.ToString().Trim();
        }

        /// <summary>
        /// 특정 셀이 몇 페이지에 속하는지 계산
        /// HPageBreaks / VPageBreaks 기준
        /// </summary>
        private static int GetPageNumber(Excel.Worksheet ws, int row, int col)
        {
            int hPage = 1;
            int vPage = 1;

            // 가로 페이지 번호
            foreach (Excel.HPageBreak pb in ws.HPageBreaks)
            {
                int breakRow = pb.Location.Row;
                if (row >= breakRow) hPage++;
                else break;
            }

            // 세로 페이지 번호
            foreach (Excel.VPageBreak pb in ws.VPageBreaks)
            {
                int breakCol = pb.Location.Column;
                if (col >= breakCol) vPage++;
                else break;
            }

            // 보통 페이지는 좌->우, 위->아래 순으로 증가한다고 보고 계산
            // 세로 페이지 수 계산
            int totalVPages = ws.VPageBreaks.Count + 1;

            return (hPage - 1) * totalVPages + vPage;
        }
    }
}
