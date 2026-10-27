using System;
using TSFramework.Core.Members.Job;
using TSFramework.Core.Providers;

namespace CenIT.ReportTourism.Jobs.SyncPowerProduction
{
    [JobPlugin("JobSyncPowerProduction", "Job tự động đồng bộ dữ liệu sản lượng điện hàng tháng")]
    public class JobSyncPowerProduction : IJobPlugable
    {
        public string PluginName
        {
            get { return "Job tự động đồng bộ dữ liệu sản lượng điện hàng tháng"; }
            set { }
        }

        public Type JobType { get; set; }

        public JobModel MainJob { get; set; }

        public void ExecuteJobNow(string sJobId, string sDescription, params object[] dataObjects)
        {
            var sPreTriggerId = JobProvider.GenerateTriggerId(5);
            var sPreGroupId = JobProvider.GenerateGroupId(6);
            var sTriggerId = $"{sPreTriggerId}-{sJobId}";
            var sGroupId = $"{sPreGroupId}-{sJobId}";

            var kpiScheduleJob = new ScheduleExecuteJobModel()
                .SetJobId(sJobId)
                .SetGroupId(sGroupId)
                .SetDescription(sDescription)
                .SetTriggerId(sTriggerId)
                .StartNow();

            var dataContent = dataObjects != null && dataObjects.Length > 0 ? dataObjects[0] : null;
            kpiScheduleJob.ExecuteNow(dataContent);
        }

        public IJobPlugable BuildJob(string sJobId, string sDescription, string sCronExpression,
            params object[] dataObjects)
        {
            var sPreTriggerId = JobProvider.GenerateTriggerId(5);
            var sPreGroupId = JobProvider.GenerateGroupId(6);
            var sTriggerId = $"{sPreTriggerId}-{sJobId}";
            var sGroupId = $"{sPreGroupId}-{sJobId}";

            var kpiScheduleJob = new ScheduleExecuteJobModel()
                .SetJobId(sJobId)
                .SetGroupId(sGroupId)
                .SetDescription(sDescription)
                .SetTriggerId(sTriggerId)
                .WithCronSchedule(sCronExpression);

            return new JobSyncPowerProduction { MainJob = kpiScheduleJob, JobType = typeof(ScheduleExecuteJobModel) };
        }
    }
}

