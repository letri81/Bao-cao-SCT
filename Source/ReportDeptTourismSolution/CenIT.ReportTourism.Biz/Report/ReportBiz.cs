using System.Data;
using TSFramework.App.Processors;

namespace CenIT.ReportTourism.Biz.Report
{
    public class ReportBiz
    {
        private const string DATA_PROVIDER_NAME = "ReportTourismProvider";

        public DataTable GetDataReport(string procedureName, params object[] p)
        {
            var spName = ReportIndustryScope.Procedure(procedureName);
            var spParams = ReportIndustryScope.ParametersFor(procedureName, p);
            var dataReport = AppProcessor.ProcedureProvider.ExecuteProcedure(spName, DATA_PROVIDER_NAME, spParams);
            return dataReport ?? new DataTable();
        }
    }
}
