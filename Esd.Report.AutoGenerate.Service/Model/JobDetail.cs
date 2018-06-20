using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Esd.Report.AutoGenerate.Common
{
    public class JobDetail
    {
        public Guid Id { get; set; }
        public string Name { get; set; }
        public string CornLike { get; set; }
        public Guid RptId { get; set; }
        public Guid CompanyId { get; set; }
        public int ActionStatus { get; set; }
        public int ExecuteStatus { get; set; }
    }
}
