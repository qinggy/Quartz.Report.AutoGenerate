using Autofac;
using Topshelf;
using Topshelf.Autofac;

namespace Esd.Report.AutoGenerate
{
    class Program
    {
        static void Main(string[] args)
        {
            HostFactory.Run(config =>
            {
                config.SetServiceName(JobService.ServiceName);
                config.SetDisplayName(JobService.DisplayServiceName);
                config.SetDescription(JobService.ServiceDescription);
                config.UseLog4Net();
                config.UseAutofacContainer(JobService.Container);

                config.Service<JobService>(setting =>
                {
                    JobService.InitSchedule(setting);
                    setting.ConstructUsingAutofacContainer();
                    setting.WhenStarted(o => o.Start());
                    setting.WhenStopped(o => o.Stop());
                });
            });
        }
    }
}
