using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data.Entity;
using System.Linq;
using System.Text;

namespace Esd.Report.AutoGenerate.Service
{
    public class EnergyPecDbContext : DbContext
    {
        private static readonly string _connectionStr;

        static EnergyPecDbContext()
        {
            _connectionStr = ConfigurationManager.ConnectionStrings["ConnectionStr"].ConnectionString;
        }

        public EnergyPecDbContext() : base(_connectionStr) { }
    }
}
