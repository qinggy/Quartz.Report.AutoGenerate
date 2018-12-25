using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Esd.Report.AutoGenerate.Application.Model
{
    public class RptRendRecord
    {
        public Guid Id { get; set; }
        public DateTime SendTime { get; set; }
        public Guid CompanyId { get; set; }
        public Guid ReportId { get; set; }
        public string ReportFile { get; set; }
        public bool Status { get; set; }
    }
}
