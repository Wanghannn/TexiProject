using SendEmailMVC.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace SendEmailMVC.Service
{
    public interface IImportDataHelper
    {
        CheckResult CheckImportData(string fileName, List<Student> students);

        public void SaveImportData(IEnumerable<Student> students);
    }
}
