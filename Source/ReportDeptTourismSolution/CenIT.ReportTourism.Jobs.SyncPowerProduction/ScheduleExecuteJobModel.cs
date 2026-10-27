using System;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using Quartz;
using TSFramework.App.Processors;
using TSFramework.Core.Members.Job;

namespace CenIT.ReportTourism.Jobs.SyncPowerProduction
{
    public class ScheduleExecuteJobModel : JobModel
    {
        private const string SP_NAME = "sp_Job_SyncPowerProductionMonthly";
        private const string DATA_PROVIDER_NAME = "ReportTourismProvider";

        public override void Execute(IJobExecutionContext context)
        {
            try
            {
                RunSync();
            }
            catch (Exception ex)
            {
                AppProcessor.Logger.Error(ex);
            }
        }

        public override void ExecuteNow(object data)
        {
            try
            {
                if (data != null && data is Tuple<int, int> period)
                {
                    RunSync(period.Item1, period.Item2);
                }
                else
                {
                    RunSync();
                }
            }
            catch (Exception ex)
            {
                AppProcessor.Logger.Error(ex);
            }
        }

        public void RunSync(int? specifiedYear = null, int? specifiedMonth = null)
        {
            AppProcessor.Logger.Message("========================================================================");
            AppProcessor.Logger.Message("========== Bắt đầu Job đồng bộ sản lượng điện sản xuất hàng tháng ======");
            AppProcessor.Logger.Message("========================================================================");

            if (specifiedYear.HasValue && specifiedMonth.HasValue)
            {
                SyncMonth(specifiedYear.Value, specifiedMonth.Value);
            }
            else
            {
                var now = DateTime.Now;
                var prevMonthDate = now.AddMonths(-1);

                // 1. Đồng bộ tháng trước (kỳ báo cáo chính hoàn tất gần nhất)
                SyncMonth(prevMonthDate.Year, prevMonthDate.Month);

                // 2. Đồng bộ tháng hiện tại (nếu có doanh nghiệp cập nhật sớm)
                SyncMonth(now.Year, now.Month);
            }

            AppProcessor.Logger.Message("========================================================================");
            AppProcessor.Logger.Message("========== Hoàn tất Job đồng bộ sản lượng điện sản xuất hàng tháng =====");
            AppProcessor.Logger.Message("========================================================================");
        }

        public void SyncMonth(int targetYear, int targetMonth)
        {
            AppProcessor.Logger.Message($"[JobSyncPowerProduction] Bắt đầu đồng bộ kỳ báo cáo tháng {targetMonth:D2}/{targetYear}...");

            bool executed = false;

            // 1. Thử gọi qua ProcedureProvider
            try
            {
                if (AppProcessor.ProcedureProvider != null)
                {
                    AppProcessor.ProcedureProvider.Execute(SP_NAME, DATA_PROVIDER_NAME, targetYear, targetMonth);
                    executed = true;
                    AppProcessor.Logger.Message($"[JobSyncPowerProduction] Thực thi qua ProcedureProvider thành công kỳ tháng {targetMonth:D2}/{targetYear}.");
                }
            }
            catch (Exception ex)
            {
                AppProcessor.Logger.Message($"[JobSyncPowerProduction] ProcedureProvider báo lỗi: {ex.Message}. Chuyển sang kết nối DB trực tiếp.");
            }

            // 2. Dự phòng gọi trực tiếp nếu ProcedureProvider chưa chạy được
            if (!executed)
            {
                try
                {
                    var connStr = ConfigurationManager.ConnectionStrings["ReportTourismDB"]?.ConnectionString
                                  ?? ConfigurationManager.ConnectionStrings["BaseApp"]?.ConnectionString;
                    if (string.IsNullOrEmpty(connStr))
                    {
                        connStr = "Data Source=10.57.30.10;Initial Catalog=baocao.sct.cenit.vn;Persist Security Info=True;User Id=baocao.sct.cenit.vn;Password=gh@1s22@Fg;Connect Timeout=200; Pooling=true; Max Pool Size=200;";
                    }

                    if (!string.IsNullOrEmpty(connStr))
                    {
                        using (var conn = new SqlConnection(connStr))
                        {
                            conn.Open();
                            using (var cmd = conn.CreateCommand())
                            {
                                cmd.CommandType = CommandType.StoredProcedure;
                                cmd.CommandText = "dbo.sp_Job_SyncPowerProductionMonthly";
                                cmd.CommandTimeout = 300;
                                cmd.Parameters.Add(new SqlParameter("@ReportYear", SqlDbType.Int) { Value = targetYear });
                                cmd.Parameters.Add(new SqlParameter("@ReportMonth", SqlDbType.Int) { Value = targetMonth });
                                cmd.ExecuteNonQuery();
                            }
                        }
                        AppProcessor.Logger.Message($"[JobSyncPowerProduction] Thực thi trực tiếp qua SqlConnection thành công kỳ tháng {targetMonth:D2}/{targetYear}.");
                    }
                    else
                    {
                        AppProcessor.Logger.Message("[JobSyncPowerProduction] Lỗi: Không tìm thấy connection string trong cấu hình.");
                    }
                }
                catch (Exception sqlEx)
                {
                    AppProcessor.Logger.Error(sqlEx);
                }
            }
        }
    }
}
