using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Esd.Report.AutoGenerate.Service.Model
{
    public class DatabaseExpression
    {
        public string QuerySqlJobWaitForExecute { get; set; } = "select * from t_Rptcenter_RptTaskRecord where ExecuteStatus = 0 and ActionStatus = 0";
        public string UpdateExecuteStatus { get; set; } = "update t_Rptcenter_RptTaskRecord set ExecuteStatus = {0} where id = {1}";
        public string UpdateActionStatus { get; set; } = "update t_Rptcenter_RptTaskRecord set ActionStatus = {0} where id = {1}";
        public string RemoveWaitForDeleteJob { get; set; } = "delete from t_Rptcenter_RptTaskRecord where id ={0}";
    }
}
