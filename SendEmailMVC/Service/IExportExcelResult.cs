using SendEmailMVC.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Http;
using SendEmailMVC.Data;

namespace SendEmailMVC.Service
{
    public interface IExportExcelResult
    {
        public void ExecuteResult(ExportExcelResult dt);

        public void ExportExcelEventHandler(ExportExcelResult dt);
    }
}