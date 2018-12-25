using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Esd.Report.AutoGenerate.Application
{
    public class EmailDetail
    {
        public Guid Id { get; set; }
        public string Theme { get; set; }
        public string SendAddress { get; set; }
        public string SendServer { get; set; }
        public string Password { get; set; }
        public string Port { get; set; }
        public string Recipients { get; set; }
        public string Cc { get; set; }
        public string Body { get; set; }
        public Guid CompanyId { get; set; }
        public Guid ReportId { get; set; }
        public bool IsDefault { get; set; }
        public string Attachment { get; set; }
    }
}
