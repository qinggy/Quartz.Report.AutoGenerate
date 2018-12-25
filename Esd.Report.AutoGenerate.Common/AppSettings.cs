using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;

namespace Esd.Report.AutoGenerate.Service
{
    public class AppSettings
    {
        private static string serviceName = "ServiceName";
        private static string displayName = "DisplayServiceName";
        private static string description = "ServiceDescription";
        private static string ifServiceExport = "ifServiceExport";
        public static string poolType = "quartz.threadPool.type";
        public static string threadCount = "quartz.threadPool.threadCount";
        public static string threadPriority = "quartz.threadPool.threadPriority";
        public static string misfire = "quartz.jobStore.misfireThreshold";
        public static string storeType = "quartz.jobStore.type";
        //=============Export============
        public static string instanceName = "quartz.scheduler.instanceName";
        public static string exportType = "quartz.scheduler.exporter.type";
        public static string exportPort = "quartz.scheduler.exporter.port";
        public static string bindName = "quartz.scheduler.exporter.bindName";
        public static string channelType = "quartz.scheduler.exporter.channelType";
        //=============================
        private static string Configuration(string key)
        {
            return ConfigurationManager.AppSettings[key];
        }

        public static string ServiceName { get { return Configuration(serviceName); } }
        public static string DisplayServiceName { get { return Configuration(displayName); } }
        public static string ServiceDescription { get { return Configuration(description); } }
        public static string PoolType { get { return Configuration(poolType); } }
        public static string ThreadCount { get { return Configuration(threadCount); } }
        public static string ThreadPriority { get { return Configuration(threadPriority); } }
        public static string MissFire { get { return Configuration(misfire); } }
        public static string StoreType { get { return Configuration(storeType); } }
        public static string ExportType { get { return Configuration(exportType); } }
        public static string ExportPort { get { return Configuration(exportPort); } }
        public static string BindName { get { return Configuration(bindName); } }
        public static string ChannelType { get { return Configuration(channelType); } }
        public static bool IfServiceExport { get { return bool.Parse(Configuration(ifServiceExport)); } }
        public static string InstanceName { get { return Configuration(instanceName); } }
        public static string JobGroupName
        {
            get { return "报表生成作业处理"; }
        }
        public static string BaseJobGroupName
        {
            get { return "主服务定时刷新Job清单作业处理"; }
        }

        public static string GetValue(string key)
        {
            return ConfigurationManager.AppSettings[key];
        }
    }
}
