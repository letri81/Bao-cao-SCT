using System;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Web;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace BaoCao.SanLuongDien
{
    public static class ExcelReportBuilder
    {
        private static string GetConnectionString()
        {
            var cs = ConfigurationManager.ConnectionStrings["ReportTourismDB"];
            if (cs != null && !string.IsNullOrEmpty(cs.ConnectionString))
            {
                return cs.ConnectionString;
            }
            return "Server=10.57.30.10;Database=baocao.sct.cenit.vn;User Id=baocao.sct.cenit.vn;Password=gh@1s22@Fg;";
        }

        public static DataTable ExecuteProcedure(string procName, int reportYear, int periodMonth)
        {
            var dt = new DataTable();
            try
            {
                using (var conn = new SqlConnection(GetConnectionString()))
                {
                    using (var cmd = new SqlCommand(procName, conn))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.CommandTimeout = 120;
                        cmd.Parameters.AddWithValue("@ReportYear", reportYear);
                        cmd.Parameters.AddWithValue("@PeriodMonth", periodMonth);
                        using (var da = new SqlDataAdapter(cmd))
                        {
                            da.Fill(dt);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                // Fallback empty if database query fails
                System.Diagnostics.Debug.WriteLine("ExecuteProcedure error: " + ex.Message);
            }
            return dt;
        }

        public static void ExportFullWorkbook(HttpResponseBase response, int rptYear, int rptPeriod, DataTable dataChung, DataTable dataDetailCurrent)
        {
            if (dataChung == null || dataChung.Rows.Count == 0)
            {
                dataChung = ExecuteProcedure("usp_Report_SanLuongDien_Chung", rptYear, rptPeriod);
            }
            if (dataDetailCurrent == null || dataDetailCurrent.Rows.Count == 0)
            {
                dataDetailCurrent = ExecuteProcedure("usp_Report_SanLuongDien", rptYear, rptPeriod);
            }
            var dataDetailPrev = ExecuteProcedure("usp_Report_SanLuongDien", rptYear - 1, rptPeriod);

            using (var package = new ExcelPackage())
            {
                // Sheet 1: BC năm {rptYear - 1}
                BuildYearlyDetailSheet(package, $"BC năm {rptYear - 1}", rptYear - 1, rptPeriod, dataDetailPrev, isPreviousYear: true);

                // Sheet 2: BC năm {rptYear}
                BuildYearlyDetailSheet(package, $"BC năm {rptYear}", rptYear, rptPeriod, dataDetailCurrent, isPreviousYear: false);

                // Sheet 3: BC chung
                BuildGeneralSummarySheet(package, "BC chung", rptYear, rptPeriod, dataChung);

                try
                {
                    package.Workbook.Calculate();
                }
                catch
                {
                    // Ignore calculation error if any formula reference differs
                }

                var bytes = package.GetAsByteArray();
                var reportFilename = $"BaoCao_SanLuongDien_{rptYear}_{DateTime.Now:yyyyMMddHHmmss}.xlsx";
                response.Clear();
                response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                response.AddHeader("Content-Disposition", "attachment; filename=" + reportFilename);
                response.BinaryWrite(bytes);
                response.Flush();
                response.End();
            }
        }

        private static void BuildGeneralSummarySheet(ExcelPackage package, string sheetName, int rptYear, int rptPeriod, DataTable dataChung)
        {
            var ws = package.Workbook.Worksheets.Add(sheetName);
            ws.View.ShowGridLines = true;

            // Title
            ws.Cells["A1:I1"].Merge = true;
            ws.Cells["A1"].Value = "THEO DÕI SẢN LƯỢNG ĐIỆN 74 NHÀ MÁY ĐIỆN VÀ ĐIỆN MẶT TRỜI MÁI NHÀ";
            ws.Cells["A1"].Style.Font.Bold = true;
            ws.Cells["A1"].Style.Font.Size = 14;
            ws.Cells["A1"].Style.Font.Name = "Times New Roman";
            ws.Cells["A1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            ws.Cells["A1"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            ws.Row(1).Height = 28;

            // Table Headers (Row 3 & 4)
            int h1 = 3;
            int h2 = 4;

            ws.Cells[h1, 1, h2, 1].Merge = true;
            ws.Cells[h1, 1].Value = "STT";

            ws.Cells[h1, 2, h2, 2].Merge = true;
            ws.Cells[h1, 2].Value = "Tên nhà máy điện";

            ws.Cells[h1, 3, h1, 4].Merge = true;
            ws.Cells[h1, 3].Value = "Công suất";
            ws.Cells[h2, 3].Value = "MWp";
            ws.Cells[h2, 4].Value = "MW";

            ws.Cells[h1, 5, h2, 5].Merge = true;
            ws.Cells[h1, 5].Value = $"Thực tế năm {rptYear - 1} (kWh)";

            ws.Cells[h1, 6, h2, 6].Merge = true;
            ws.Cells[h1, 6].Value = $"Thực tế năm {rptYear} (kWh)";

            ws.Cells[h1, 7, h2, 7].Merge = true;
            ws.Cells[h1, 7].Value = "Sản lượng điện so với cùng kỳ (kWh)";

            ws.Cells[h1, 8, h2, 8].Merge = true;
            ws.Cells[h1, 8].Value = "Tỷ lệ";

            ws.Cells[h1, 9, h2, 9].Merge = true;
            ws.Cells[h1, 9].Value = "Tăng/giảm";

            using (var range = ws.Cells[h1, 1, h2, 9])
            {
                range.Style.Font.Bold = true;
                range.Style.Font.Size = 14;
                range.Style.Font.Name = "Times New Roman";
                range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                range.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                range.Style.WrapText = true;
            }

            // Set column widths
            ws.Column(1).Width = 7;
            ws.Column(2).Width = 38;
            ws.Column(3).Width = 10;
            ws.Column(4).Width = 13;
            ws.Column(5).Width = 26;
            ws.Column(6).Width = 25;
            ws.Column(7).Width = 24;
            ws.Column(8).Width = 12;
            ws.Column(9).Width = 14;

            // Group row definitions
            // Group 1: rows 5-11 (Gió)
            // Group 2: rows 12-18 (Mặt trời)
            // Group 3: rows 19-25 (Thủy điện)
            // Group 4: rows 26-32 (Nhiệt điện)
            // Group 5: rows 33-39 (Sinh khối)
            // Group 6: rows 40-46 (ĐMTMN)
            int[,] groupRowRanges = new int[,] {
                { 5, 11 },   // Group 1
                { 12, 18 },  // Group 2
                { 19, 25 },  // Group 3
                { 26, 32 },  // Group 4
                { 33, 39 },  // Group 5
                { 40, 46 }   // Group 6
            };

            // Index data rows by (GroupOrder, ItemOrder)
            DataRow[,] mappedRows = new DataRow[9, 7];
            if (dataChung != null)
            {
                foreach (DataRow r in dataChung.Rows)
                {
                    int go = Convert.ToInt32(r["GroupOrder"]);
                    int io = Convert.ToInt32(r["ItemOrder"]);
                    if (go >= 1 && go <= 8 && io >= 0 && io <= 6)
                    {
                        mappedRows[go, io] = r;
                    }
                }
            }

            for (int gIdx = 0; gIdx < 6; gIdx++)
            {
                int gOrder = gIdx + 1;
                int startR = groupRowRanges[gIdx, 0];
                int endR = groupRowRanges[gIdx, 1];

                // 1. Group Header Row (startR)
                var gRowData = mappedRows[gOrder, 0];
                ws.Cells[startR, 1].Value = gRowData != null ? gRowData["STT"] : GetRomanNumeral(gOrder);
                ws.Cells[startR, 2].Value = gRowData != null ? gRowData["Title"] : $"NHÓM {gOrder}";
                ws.Cells[startR, 1, startR, 9].Style.Font.Bold = true;
                ws.Cells[startR, 1, startR, 9].Style.Font.Size = 14;
                ws.Cells[startR, 1, startR, 9].Style.Font.Name = "Times New Roman";
                ws.Cells[startR, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ws.Cells[startR, 2].Style.WrapText = true;

                if (gRowData != null)
                {
                    if (gRowData["CapacityMWp"] != DBNull.Value && Convert.ToDecimal(gRowData["CapacityMWp"]) > 0)
                    {
                        ws.Cells[startR, 3].Value = Convert.ToDecimal(gRowData["CapacityMWp"]);
                        ws.Cells[startR, 3].Style.Numberformat.Format = "#,##0.000";
                        ws.Cells[startR, 3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    }
                    if (gRowData["CapacityMW"] != DBNull.Value && Convert.ToDecimal(gRowData["CapacityMW"]) > 0)
                    {
                        ws.Cells[startR, 4].Value = Convert.ToDecimal(gRowData["CapacityMW"]);
                        ws.Cells[startR, 4].Style.Numberformat.Format = "#,##0.000";
                        ws.Cells[startR, 4].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    }
                }

                // 2. Detail Rows (Quarters I - IV: startR + 1 .. startR + 4)
                string[] qNames = { "Quý I", "Quý II", "Quý III", $"Quý IV (số ước năm {rptYear})" };
                for (int q = 1; q <= 4; q++)
                {
                    int rIndex = startR + q;
                    var qData = mappedRows[gOrder, q];
                    ws.Cells[rIndex, 1].Value = q;
                    ws.Cells[rIndex, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    ws.Cells[rIndex, 2].Value = qNames[q - 1];

                    if (qData != null)
                    {
                        if (qData["OutputLastYear"] != DBNull.Value)
                            ws.Cells[rIndex, 5].Value = Convert.ToDecimal(qData["OutputLastYear"]);
                        if (qData["OutputCurrentYear"] != DBNull.Value)
                            ws.Cells[rIndex, 6].Value = Convert.ToDecimal(qData["OutputCurrentYear"]);
                    }

                    ws.Cells[rIndex, 7].Formula = $"F{rIndex}-E{rIndex}";
                    ws.Cells[rIndex, 8].Formula = $"F{rIndex}/E{rIndex}";
                    ws.Cells[rIndex, 9].Formula = $"H{rIndex}-1";

                    ApplyRowStyle(ws, rIndex, isBold: false, fillYellow: false, fillGold: false);
                }

                // 3. Cả năm (startR + 5)
                int rTotYear = startR + 5;
                ws.Cells[rTotYear, 1].Value = 5;
                ws.Cells[rTotYear, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ws.Cells[rTotYear, 2].Value = gOrder == 3 ? "Cả năm" : $"Cả năm (số ước năm {rptYear})";

                ws.Cells[rTotYear, 5].Formula = $"SUM(E{startR + 1}:E{startR + 4})";
                ws.Cells[rTotYear, 6].Formula = $"SUM(F{startR + 1}:F{startR + 4})";
                ws.Cells[rTotYear, 7].Formula = $"F{rTotYear}-E{rTotYear}";
                ws.Cells[rTotYear, 8].Formula = $"F{rTotYear}/E{rTotYear}";
                ws.Cells[rTotYear, 9].Formula = $"H{rTotYear}-1";

                ApplyRowStyle(ws, rTotYear, isBold: true, fillYellow: false, fillGold: false);

                // 4. Kỳ lũy kế (startR + 6, e.g. 09 tháng đầu năm)
                int rAcc = startR + 6;
                ws.Cells[rAcc, 1].Value = 6;
                ws.Cells[rAcc, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ws.Cells[rAcc, 2].Value = $"{rptPeriod:00} tháng đầu năm";

                // For 9 months, sum Q1 + Q2 + Q3 (startR + 1 .. startR + 3)
                int qCount = Math.Min(4, Math.Max(1, (int)Math.Ceiling(rptPeriod / 3.0)));
                string sumRangeE = qCount == 1 ? $"E{startR + 1}" : $"SUM(E{startR + 1}:E{startR + qCount})";
                string sumRangeF = qCount == 1 ? $"F{startR + 1}" : $"SUM(F{startR + 1}:F{startR + qCount})";

                ws.Cells[rAcc, 5].Formula = sumRangeE;
                ws.Cells[rAcc, 6].Formula = sumRangeF;
                ws.Cells[rAcc, 7].Formula = $"F{rAcc}-E{rAcc}";
                ws.Cells[rAcc, 8].Formula = $"F{rAcc}/E{rAcc}";
                ws.Cells[rAcc, 9].Formula = $"H{rAcc}-1";

                ApplyRowStyle(ws, rAcc, isBold: true, fillYellow: true, fillGold: false);
            }

            // Summary Row 47: Tổng cả năm các nhà máy điện
            int r47 = 47;
            ws.Cells[r47, 2].Value = "Tổng cả năm các nhà máy điện";
            ws.Cells[r47, 3].Formula = "SUM(C5:C45)";
            ws.Cells[r47, 4].Formula = "SUM(D5:D45)";
            ws.Cells[r47, 5].Formula = "E10+E17+E24+E31+E38+E45";
            ws.Cells[r47, 6].Formula = "F10+F17+F24+F31+F38+F45";
            ws.Cells[r47, 7].Formula = $"F{r47}-E{r47}";
            ws.Cells[r47, 8].Formula = $"F{r47}/E{r47}";
            ws.Cells[r47, 9].Formula = $"H{r47}-1";
            ApplyRowStyle(ws, r47, isBold: true, fillYellow: false, fillGold: true);

            // Summary Row 48: Tổng 09 tháng đầu năm các nhà máy điện
            int r48 = 48;
            ws.Cells[r48, 2].Value = $"Tổng {rptPeriod:00} tháng đầu năm các nhà máy điện";
            ws.Cells[r48, 5].Formula = "E11+E18+E25+E32+E39+E46";
            ws.Cells[r48, 6].Formula = "F11+F18+F25+F32+F39+F46";
            ws.Cells[r48, 7].Formula = $"F{r48}-E{r48}";
            ws.Cells[r48, 8].Formula = $"F{r48}/E{r48}";
            ws.Cells[r48, 9].Formula = $"H{r48}-1";
            ApplyRowStyle(ws, r48, isBold: true, fillYellow: false, fillGold: true);

            // Apply all borders for entire table A3:I48
            using (var allTable = ws.Cells[3, 1, 48, 9])
            {
                allTable.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                allTable.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                allTable.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                allTable.Style.Border.Right.Style = ExcelBorderStyle.Thin;
            }
        }

        private static void ApplyRowStyle(ExcelWorksheet ws, int r, bool isBold, bool fillYellow, bool fillGold)
        {
            ws.Cells[r, 1, r, 9].Style.Font.Name = "Times New Roman";
            ws.Cells[r, 1, r, 9].Style.Font.Size = 14;
            ws.Cells[r, 1, r, 9].Style.Font.Bold = isBold;
            ws.Cells[r, 1, r, 9].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

            // Number formats
            ws.Cells[r, 3, r, 4].Style.Numberformat.Format = "#,##0.000";
            ws.Cells[r, 3, r, 4].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            ws.Cells[r, 5, r, 7].Style.Numberformat.Format = "#,##0.0";
            ws.Cells[r, 5, r, 7].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

            ws.Cells[r, 8].Style.Numberformat.Format = "0.00%";
            ws.Cells[r, 8].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

            ws.Cells[r, 9].Style.Numberformat.Format = "0.00%";
            ws.Cells[r, 9].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

            if (fillYellow)
            {
                using (var rng = ws.Cells[r, 1, r, 9])
                {
                    rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(255, 255, 0));
                }
            }
            else if (fillGold)
            {
                using (var rng = ws.Cells[r, 1, r, 9])
                {
                    rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(255, 192, 0));
                }
            }
        }

        private static void BuildYearlyDetailSheet(ExcelPackage package, string sheetName, int year, int rptPeriod, DataTable data, bool isPreviousYear)
        {
            var ws = package.Workbook.Worksheets.Add(sheetName);
            ws.View.ShowGridLines = true;

            int maxCol = isPreviousYear ? 20 : 24;
            string lastColLetter = isPreviousYear ? "T" : "X";

            // Title
            ws.Cells[$"A1:{lastColLetter}1"].Merge = true;
            ws.Cells["A1"].Value = $"BÁO CÁO SẢN LƯỢNG ĐIỆN NĂM {year}";
            ws.Cells["A1"].Style.Font.Bold = true;
            ws.Cells["A1"].Style.Font.Size = 14;
            ws.Cells["A1"].Style.Font.Name = "Times New Roman";
            ws.Cells["A1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            ws.Cells[$"A2:{lastColLetter}2"].Merge = true;
            ws.Cells["A2"].Value = $"(Kèm theo Công văn số ………/SCT-QLNL ngày ……/ ……/ {year} của Sở Công Thương)";
            ws.Cells["A2"].Style.Font.Italic = true;
            ws.Cells["A2"].Style.Font.Size = 11;
            ws.Cells["A2"].Style.Font.Name = "Times New Roman";
            ws.Cells["A2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            // Headers
            int h1 = 4;
            int h2 = 5;

            string[] singleHeaders = { "STT", "Tên nhà máy điện", "Tên đơn vị", "Địa chỉ" };
            for (int c = 1; c <= 4; c++)
            {
                ws.Cells[h1, c, h2, c].Merge = true;
                ws.Cells[h1, c].Value = singleHeaders[c - 1];
            }

            ws.Cells[h1, 5, h1, 6].Merge = true;
            ws.Cells[h1, 5].Value = "Công suất";
            ws.Cells[h2, 5].Value = "MWp";
            ws.Cells[h2, 6].Value = "MW";

            int monthStartCol = 7;
            if (!isPreviousYear)
            {
                ws.Cells[h1, 7, h2, 7].Merge = true;
                ws.Cells[h1, 7].Value = $"Kế hoạch {year} (kWh)";
                ws.Cells[h1, 7].Style.WrapText = true;
                monthStartCol = 8;
            }

            int monthEndCol = monthStartCol + 11;
            ws.Cells[h1, monthStartCol, h1, monthEndCol].Merge = true;
            ws.Cells[h1, monthStartCol].Value = $"Sản lượng theo tháng năm {year} (kWh)";
            for (int m = 1; m <= 12; m++)
            {
                ws.Cells[h2, monthStartCol + m - 1].Value = $"Tháng {m}";
            }

            int totYearCol = monthEndCol + 1;
            ws.Cells[h1, totYearCol, h2, totYearCol].Merge = true;
            ws.Cells[h1, totYearCol].Value = isPreviousYear ? "Tổng cộng cả năm\n(kWh)" : "Tổng cộng\n(kWh)";
            ws.Cells[h1, totYearCol].Style.WrapText = true;

            if (isPreviousYear)
            {
                int accCol = totYearCol + 1;
                ws.Cells[h1, accCol, h2, accCol].Merge = true;
                ws.Cells[h1, accCol].Value = $"Tổng cộng\ntháng {rptPeriod}";
                ws.Cells[h1, accCol].Style.WrapText = true;
            }
            else
            {
                int c21 = totYearCol + 1;
                ws.Cells[h1, c21, h2, c21].Merge = true;
                ws.Cells[h1, c21].Value = $"Tổng {rptPeriod} tháng\nnăm {year - 1}";
                ws.Cells[h1, c21].Style.WrapText = true;

                int c22 = c21 + 1;
                ws.Cells[h1, c22, h2, c22].Merge = true;
                ws.Cells[h1, c22].Value = $"Tổng {rptPeriod} tháng\nnăm {year}";
                ws.Cells[h1, c22].Style.WrapText = true;

                int c23 = c22 + 1;
                ws.Cells[h1, c23, h2, c23].Merge = true;
                ws.Cells[h1, c23].Value = "Tỷ lệ đạt được\nso với cùng kỳ";
                ws.Cells[h1, c23].Style.WrapText = true;

                int c24 = c23 + 1;
                ws.Cells[h1, c24, h2, c24].Merge = true;
                ws.Cells[h1, c24].Value = "Tỷ lệ tăng/giảm\nso với cùng kỳ";
                ws.Cells[h1, c24].Style.WrapText = true;
            }

            using (var range = ws.Cells[h1, 1, h2, maxCol])
            {
                range.Style.Font.Bold = true;
                range.Style.Font.Name = "Times New Roman";
                range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(217, 226, 243));
                range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                range.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                range.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                range.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                range.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                range.Style.Border.Right.Style = ExcelBorderStyle.Thin;
            }

            int curRow = 6;
            string currentGroup = "";

            if (data != null && data.Rows.Count > 0)
            {
                foreach (DataRow row in data.Rows)
                {
                    var groupName = Convert.ToString(row["GroupName"]);
                    if (groupName != currentGroup)
                    {
                        currentGroup = groupName;
                        ws.Cells[curRow, 1, curRow, maxCol].Merge = true;
                        ws.Cells[curRow, 1].Value = currentGroup;
                        ws.Cells[curRow, 1].Style.Font.Bold = true;
                        ws.Cells[curRow, 1].Style.Font.Name = "Times New Roman";
                        ws.Cells[curRow, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        ws.Cells[curRow, 1].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(242, 242, 242));
                        curRow++;
                    }

                    ws.Cells[curRow, 1].Value = row["STT"];
                    ws.Cells[curRow, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    ws.Cells[curRow, 2].Value = row["PlantName"];
                    ws.Cells[curRow, 3].Value = row["CompanyName"];
                    ws.Cells[curRow, 4].Value = row["Address"];

                    if (row["CapacityMWp"] != DBNull.Value && Convert.ToDecimal(row["CapacityMWp"]) > 0)
                    {
                        ws.Cells[curRow, 5].Value = Convert.ToDecimal(row["CapacityMWp"]);
                        ws.Cells[curRow, 5].Style.Numberformat.Format = "#,##0.00";
                    }
                    if (row["CapacityMW"] != DBNull.Value && Convert.ToDecimal(row["CapacityMW"]) > 0)
                    {
                        ws.Cells[curRow, 6].Value = Convert.ToDecimal(row["CapacityMW"]);
                        ws.Cells[curRow, 6].Style.Numberformat.Format = "#,##0.00";
                    }

                    if (!isPreviousYear)
                    {
                        ws.Cells[curRow, 7].Value = Convert.ToDecimal(row["PlanYear"]);
                        ws.Cells[curRow, 7].Style.Numberformat.Format = "#,##0";
                    }

                    for (int m = 1; m <= 12; m++)
                    {
                        var val = Convert.ToDecimal(row[$"Month{m}"]);
                        if (val > 0)
                        {
                            ws.Cells[curRow, monthStartCol + m - 1].Value = val;
                            ws.Cells[curRow, monthStartCol + m - 1].Style.Numberformat.Format = "#,##0";
                        }
                    }

                    ws.Cells[curRow, totYearCol].Value = Convert.ToDecimal(row["TotalYear"]);
                    ws.Cells[curRow, totYearCol].Style.Numberformat.Format = "#,##0";

                    if (isPreviousYear)
                    {
                        ws.Cells[curRow, totYearCol + 1].Value = Convert.ToDecimal(row["AccumulatedCurrentYear"]);
                        ws.Cells[curRow, totYearCol + 1].Style.Numberformat.Format = "#,##0";
                    }
                    else
                    {
                        ws.Cells[curRow, totYearCol + 1].Value = Convert.ToDecimal(row["AccumulatedLastYear"]);
                        ws.Cells[curRow, totYearCol + 1].Style.Numberformat.Format = "#,##0";

                        ws.Cells[curRow, totYearCol + 2].Value = Convert.ToDecimal(row["AccumulatedCurrentYear"]);
                        ws.Cells[curRow, totYearCol + 2].Style.Numberformat.Format = "#,##0";

                        if (row["RatioAchieved"] != DBNull.Value)
                        {
                            ws.Cells[curRow, totYearCol + 3].Value = Convert.ToDecimal(row["RatioAchieved"]);
                            ws.Cells[curRow, totYearCol + 3].Style.Numberformat.Format = "0.00%";
                        }

                        if (row["RatioGrowth"] != DBNull.Value)
                        {
                            ws.Cells[curRow, totYearCol + 4].Value = Convert.ToDecimal(row["RatioGrowth"]);
                            ws.Cells[curRow, totYearCol + 4].Style.Numberformat.Format = "0.00%";
                        }
                    }

                    using (var r = ws.Cells[curRow, 1, curRow, maxCol])
                    {
                        r.Style.Font.Name = "Times New Roman";
                        r.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        r.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                        r.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                        r.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    }

                    curRow++;
                }
            }

            ws.Cells.AutoFitColumns(10, 50);
        }

        private static string GetRomanNumeral(int number)
        {
            switch (number)
            {
                case 1: return "I";
                case 2: return "II";
                case 3: return "III";
                case 4: return "IV";
                case 5: return "V";
                case 6: return "VI";
                case 7: return "VII";
                case 8: return "VIII";
                default: return number.ToString();
            }
        }
    }
}

