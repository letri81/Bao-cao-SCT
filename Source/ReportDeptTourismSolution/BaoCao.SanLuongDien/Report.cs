using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Data;
using System.Globalization;
using System.IO;
using System.Web;
using System.Web.UI.WebControls;
using CenIT.ReportTourism.Core.Attributes;
using CenIT.ReportTourism.Core.Interfaces;
using Microsoft.Reporting.WebForms;

namespace BaoCao.SanLuongDien
{
    [ReportPlugin("_03BaoCao_SanLuongDien", "03 - Báo cáo chi tiết sản lượng điện", "03 - Báo cáo chi tiết sản lượng điện các nhà máy điện và điện mặt trời mái nhà")]
    public class Report : IPlugableReport
    {
        public int IdentifyReport => 3;
        private string TemplateReport => "_03BaoCao_SanLuongDien.rdlc";
        private readonly string _dataSourceName = "Data_SanLuongDien";

        public string ReportKey => "_03BaoCao_SanLuongDien";
        public string ReportName => "03 - Báo cáo chi tiết sản lượng điện";
        public string Description => "03 - Báo cáo chi tiết sản lượng điện các nhà máy điện và điện mặt trời mái nhà";
        public string ViewName => "_03BaoCao_SanLuongDien";
        public string StoreName => "usp_Report_SanLuongDien";

        private object[] ReportParams { get; set; }

        public object[] CreateParams(NameValueCollection param)
        {
            int year = DateTime.Now.Year;
            if (!string.IsNullOrEmpty(param["ReportYear"]))
            {
                int.TryParse(param["ReportYear"], out year);
            }
            else if (!string.IsNullOrEmpty(param["OnMonth"]))
            {
                DateTime d;
                if (DateTime.TryParseExact(param["OnMonth"], "MM/yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out d))
                {
                    year = d.Year;
                }
            }

            int periodMonth = 9;
            if (!string.IsNullOrEmpty(param["PeriodMonth"]))
            {
                int.TryParse(param["PeriodMonth"], out periodMonth);
            }

            ReportParams = new object[] { year, periodMonth };
            return new object[] { year, periodMonth };
        }

        public void Export(HttpResponseBase response, DataTable data, string urlPathReport)
        {
            int rptYear = ReportParams != null && ReportParams.Length > 0 ? Convert.ToInt32(ReportParams[0]) : DateTime.Now.Year;
            int rptPeriod = ReportParams != null && ReportParams.Length > 1 ? Convert.ToInt32(ReportParams[1]) : 9;

            ExcelReportBuilder.ExportFullWorkbook(response, rptYear, rptPeriod, null, data);
        }

        public ReportViewer CreateReport(DataTable data, string urlPathReport)
        {
            var fullPathRdlc = Path.Combine(urlPathReport, TemplateReport);

            int rptYear = ReportParams != null && ReportParams.Length > 0 ? Convert.ToInt32(ReportParams[0]) : DateTime.Now.Year;
            int rptPeriod = ReportParams != null && ReportParams.Length > 1 ? Convert.ToInt32(ReportParams[1]) : 9;

            var listParams = new List<ReportParameter>
            {
                new ReportParameter("p_ReportYear", rptYear.ToString()),
                new ReportParameter("p_PeriodMonth", rptPeriod.ToString())
            };

            var reportViewer = new ReportViewer
            {
                ProcessingMode = ProcessingMode.Local,
                SizeToReportContent = true,
                ZoomMode = ZoomMode.PageWidth,
                Width = Unit.Percentage(99),
                Height = Unit.Pixel(1000),
                AsyncRendering = false,
                PageCountMode = PageCountMode.Estimate
            };

            reportViewer.LocalReport.ReportPath = fullPathRdlc;
            reportViewer.LocalReport.DataSources.Clear();
            reportViewer.LocalReport.DataSources.Add(new ReportDataSource(_dataSourceName, data));
            reportViewer.LocalReport.SetParameters(listParams);

            return reportViewer;
        }
    }
}
