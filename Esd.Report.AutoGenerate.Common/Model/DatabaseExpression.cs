using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Esd.Report.AutoGenerate.Application.Model
{
    public class DatabaseExpression
    {
        public string QuerySqlJobWaitForExecute { get; set; } = "select * from t_Rptcenter_RptTaskRecord where ExecuteStatus = 0 and ActionStatus = 0 and CornLike is not NULL AND CornLike <> ''";
        public string UpdateExecuteStatus { get; set; } = "update t_Rptcenter_RptTaskRecord set ExecuteStatus = {0} where id = '{1}'";
        public string UpdateActionStatus { get; set; } = "update t_Rptcenter_RptTaskRecord set ActionStatus = {0} where id = '{1}'";
        public string RemoveWaitForDeleteJob { get; set; } = "delete from t_Rptcenter_RptTaskRecord where id ='{0}'";
        public string GetEmailSet { get; set; } = @"SELECT CASE WHEN (SELECT 1 FROM dbo.t_Rptcenter_RptSendEmailInfo WHERE ReportId = '{0}' AND IsDefault = 0) = 1 THEN (SELECT * FROM dbo.t_Rptcenter_RptSendEmailInfo WHERE ReportId = '{0}' AND IsDefault = 0) ELSE(SELECT * FROM dbo.t_Rptcenter_RptSendEmailInfo WHERE CompanyId = '{1}' AND IsDefault = 1) end";
    }
}
