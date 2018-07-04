using Autofac;
using Esd.Report.AutoGenerate.Application;
using log4net;
using System;
using System.IO;
using System.Reflection;
using Topshelf;
using Topshelf.Autofac;
using Topshelf.ServiceConfigurators;

namespace Esd.Report.AutoGenerate
{
    class Program
    {
        static void Main(string[] args)
        {
            HostFactory.Run(config =>
            {
                config.UseLog4Net();
                config.RunAsLocalSystem();
                config.UseAutofacContainer(ServiceRunner.Container);
                config.EnablePauseAndContinue();
                config.SetServiceName(ServiceRunner.ServiceName);
                config.SetDisplayName(ServiceRunner.DisplayServiceName);
                config.SetDescription(ServiceRunner.ServiceDescription);

                config.Service<ServiceRunner>(s =>
                {
                    //ServiceRunner.InitSchedule(s);
                    s.ConstructUsingAutofacContainer();
                    s.WhenStarted((sr, hct) => sr.Start(hct));
                    s.WhenStopped((sr, hct) => sr.Stop(hct));
                    s.WhenContinued((sr, hct) => sr.Continue(hct));
                    s.WhenPaused((sr, hct) => sr.Pause(hct));
                });
            });

            Console.ReadLine();
        }
    }
}
