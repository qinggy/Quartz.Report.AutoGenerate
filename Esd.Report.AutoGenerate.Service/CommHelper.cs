using log4net;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;

namespace Esd.Report.AutoGenerate.Service
{
    public class CommHelper
    {
        public readonly static ILog AppLogger = LogManager.GetLogger(MethodBase.GetCurrentMethod().DeclaringType);
    }
}
