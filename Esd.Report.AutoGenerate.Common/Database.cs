using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Esd.Report.AutoGenerate.Application
{
    public class Database
    {
        private readonly Lazy<EnergyPecDbContext> _dbContext;

        public Database()
        {
            _dbContext = new Lazy<EnergyPecDbContext>(() => new EnergyPecDbContext());
        }

        public EnergyPecDbContext DbContext { get { return _dbContext.Value; } }
    }
}
