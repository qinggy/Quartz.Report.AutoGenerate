using Esd.Cloud.Common;
using EsdPec.Cloud.DbModel;
using EsdPec.Cloud.Infrastructure.Utilities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Esd.Report.AutoGenerate.Application
{
    public class DataHelper
    {
        public static List<MonRecord> GetMonRecordByWhere(Trans trans, List<Guid> bmfIdlist, Guid company)
        {
            DateTime stime = new DateTime(int.Parse(trans.stime), 1, 1);
            DateTime etime = new DateTime(int.Parse(trans.stime), 12, 31);
            var factory = ServiceFactory.CreateService<IMonRecordService>();
            var list = factory.GetMonRecordsByWhere(stime, etime, bmfIdlist, company);
            factory.CloseService();
            return list;
        }

        public static List<MonRecord> GetMonRecordFree(Trans trans, List<Guid> bmfIdlist, Guid company)
        {
            DateTime stime = DateTime.Parse(trans.stime);
            DateTime etime = DateTime.Parse(trans.etime);
            var factory = ServiceFactory.CreateService<IMonRecordService>();

            var list = factory.GetMonRecordsByWhere(stime, etime, bmfIdlist, company);
            factory.CloseService();
            return list;
        }

        public static List<DayRecord> GetDayRecordByWhere(Trans trans, List<Guid> bmfIdlist, Guid company)
        {
            DateTime stime = DateTime.Parse(trans.stime);
            DateTime etime = DateTime.Parse(trans.etime);
            var factory = ServiceFactory.CreateService<IDayRecordService>();
            var list = factory.GetDayRecordsByWhere(stime, etime, bmfIdlist, company);
            factory.CloseService();
            return list;
        }

        public static List<MinRecord> GetMinthRecordByWhere(Trans trans, List<Guid> bmfIdlist, Guid company)
        {
            DateTime stime = DateTime.Parse(trans.stime);
            DateTime etime = DateTime.Parse(trans.etime);
            var factory = ServiceFactory.CreateService<IMinRecordService>();
            var list = factory.GetMinRecordsByWhere(stime, etime, bmfIdlist, company);
            return list;
        }

        public static List<HourRecord> GetHourRecordByWhere(Trans trans, List<Guid> bmfIdlist, Guid company)
        {
            DateTime stime = DateTime.Parse(trans.stime + ":00");
            DateTime etime = DateTime.Parse(trans.etime + ":00");
            var factory = ServiceFactory.CreateService<IHourRecordService>();
            var list = factory.GetHourRecordsByWhere(stime, etime, bmfIdlist, company);
            factory.CloseService();
            return list;
        }

        public static List<YearRecord> GetYearRecordByWhere(Trans trans, List<Guid> bmfIdlist, Guid company)
        {
            DateTime stime = DateTime.Parse(trans.stime);
            DateTime etime = DateTime.Parse(trans.etime);
            var factory = ServiceFactory.CreateService<IYearRecordService>();
            var list = factory.GetYearRecordsByWhere(stime, etime, bmfIdlist, company);
            factory.CloseService();
            return list;
        }

        public static List<DayRecord> GetDayRecordByFreeTime(Trans trans, List<Guid> bmfIdlist, Guid company)
        {
            DateTime stime = DateTime.Parse(trans.stime);
            DateTime etime = DateTime.Parse(trans.etime);
            var factory = ServiceFactory.CreateService<IDayRecordService>();
            var list = factory.GetDayRecordsByWhere(stime, etime, bmfIdlist, company);
            factory.CloseService();
            return list;
        }

        public static DayRecord GetDayRecord(Guid mfId, Guid companyId, DateTime dtime)
        {
            var factory = ServiceFactory.CreateService<IDayRecordService>();
            var list = factory.GetLastDataByMfid(companyId, mfId, dtime);
            factory.CloseService();
            return list;
        }

        public static MonRecord GetMonthRecord(Guid mfId, Guid companyId, DateTime dtime)
        {
            var factory = ServiceFactory.CreateService<IMonRecordService>();
            var list = factory.GetLastDataByMfid(companyId, mfId, dtime);
            factory.CloseService();
            return list;
        }

    }

    public class Trans
    {
        public Guid rptId { get; set; }
        public string type { get; set; }
        public string stime { get; set; }
        public string etime { get; set; }
        public int duration { get; set; }
        public int CurentPage { get; set; }//分页
        public string tempetime { get; set; }//临时变量结束时间
    }

    public class RealTimeData
    {
        public double PreData { get; set; }
        public double NextData { get; set; }
        public double Sum { get; set; }
        public string mftId { get; set; }
        public string LastTime { get; set; }
        public string Htime { get; set; }
    }

    public class TranResult : BaseRecord
    {
        public string Time { get; set; }
    }

    /// <summary>
    /// 对比类
    /// </summary>
    public class ComparsionClass
    {
        public double YesterDayVal { get; set; }
        public double CurrentVal { get; set; }
        public double CurrentMonthVal { get; set; }
        public double LastMonthVal { get; set; }
        public Guid ReferenceId { get; set; }
    }

    public class ComparsionDto
    {
        public string MfId { get; set; }
        public DateTime HTime { get; set; }
        public double LastData { get; set; }
        public double CurrentData { get; set; }
        public double Rate { get; set; }
    }

    public class Fgp
    {
        public string Name { get; set; }
        public double Price { get; set; }
        public double TotalData { get; set; }
    }

    public class FgpEnergy
    {
        public string HTime { get; set; }
        public double SumVal { get; set; }
        public List<Fgp> FgpList { get; set; }
    }

    public class MfFgpEnergy
    {
        public Guid MfId { get; set; }
        public List<FgpEnergy> FgpList { get; set; }
    }
}