using SendEmailMVC.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.IO;
using LinqToExcel;
using System.Text;
using SendEmailMVC.Data;
using ClosedXML.Excel;
using System.Web;

namespace SendEmailMVC.Service
{
    public class ExportExcelResult : IExportExcelResult
    {
        public string SheetName { get; set; }
        public string FileName { get; set; }
        public List<Student> ExportData { get; set; }

        public ExportExcelResult()
        {
            //_context = context;
        }

        /// <summary>
        /// 檢查匯入的 Excel 資料.
        /// </summary>
        public void ExecuteResult(ExportExcelResult dt)
        {
            if (dt.ExportData == null)
            {
                throw new InvalidDataException("ExportData");
            }
            if (string.IsNullOrWhiteSpace(dt.SheetName))
            {
                this.SheetName = "outputfile";
            }
            if (string.IsNullOrWhiteSpace(dt.FileName))
            {
                this.FileName = string.Concat(
                    "ExportData_",
                    DateTime.Now.ToString("yyyyMMddHHmmss"),
                    ".xlsx");
            }

            this.ExportExcelEventHandler(dt);
        }

        /// <summary>
        /// Exports the excel event handler.
        /// </summary>
        public void ExportExcelEventHandler(ExportExcelResult dt)
        {
            try
            {
                var workbook = new XLWorkbook();

                IXLWorksheet worksheet = workbook.Worksheets.Add("outputfile");
                //worksheet.Cell(1, 1).Value = "Id";
                worksheet.Cell(1, 1).Value = "LastName";
                worksheet.Cell(1, 2).Value = "FirstMidName";
                worksheet.Cell(1, 3).Value = "EnrollmentDate";
                for (int index = 1; index <= dt.ExportData.Count; index++)
                {
                    //worksheet.Cell(index + 1, 1).Value = dt.ExportData[index - 1].ID;
                    worksheet.Cell(index + 1, 1).Value = dt.ExportData[index - 1].LastName;
                    worksheet.Cell(index + 1, 2).Value = dt.ExportData[index - 1].FirstMidName;
                    worksheet.Cell(index + 1, 3).Value = dt.ExportData[index - 1].EnrollmentDate;
                }

                //using (var stream = new MemoryStream())
                //{
                //    workbook.SaveAs(stream);
                //    var content = stream.ToArray();
                //    return File(content, contentType, fileName);
                //}

                //workbook.SaveAs(@"C:\Users\v-ivowan\Desktop\" + dt.FileName);
            }
            catch (Exception ex)
            {
                throw(ex);
            }
        }
    }
}
