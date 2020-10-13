using SendEmailMVC.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.IO;
using LinqToExcel;
using System.Text;
using SendEmailMVC.Data;

namespace SendEmailMVC.Service
{
    public class ImportDataHelper : IImportDataHelper
    {

        private readonly SchoolContext _context;

        public ImportDataHelper(SchoolContext context)
        {
            _context = context;
        }


        /// <summary>
        /// 檢查匯入的 Excel 資料.
        /// </summary>
        /// <param name="fileName">Name of the file.</param>
        /// <param name="importStudents">The import zip codes.</param>
        public CheckResult CheckImportData(string fileName, List<Student> importStudents)
        {
            var result = new CheckResult();
            var targetFile = new FileInfo(fileName);

            if (!targetFile.Exists)
            {
                //result.ID = Guid.NewGuid();
                result.Success = false;
                result.ErrorCount = 0;
                result.ErrorMessage = "匯入的資料檔案不存在";
                return result;
            }

            var excelFile = new ExcelQueryFactory(fileName);

            //欄位對映
            excelFile.AddMapping<Student>(x => x.ID, "ID");
            excelFile.AddMapping<Student>(x => x.LastName, "LastName");
            excelFile.AddMapping<Student>(x => x.FirstMidName, "FirstMidName");
            excelFile.AddMapping<Student>(x => x.EnrollmentDate, "EnrollmentDate");

            //SheetName
            var excelContent = excelFile.Worksheet<Student>("outputfile");

            int errorCount = 0;
            int rowIndex = 1;
            var importErrorMessages = new List<string>();

            //檢查資料
            foreach (var row in excelContent)
            {
                var errorMessage = new StringBuilder();
                var student = new Student();

                //student.ID = row.ID;
                student.LastName = row.LastName;
                student.FirstMidName = row.FirstMidName;
                student.EnrollmentDate =  DateTime.Now;

                //LastName
                if (string.IsNullOrWhiteSpace(row.LastName))
                {
                    errorMessage.Append("LastName - 不可空白. ");
                }
                student.LastName = row.LastName;

                //FirstMidName
                if (string.IsNullOrWhiteSpace(row.FirstMidName))
                {
                    errorMessage.Append("FirstMidName - 不可空白. ");
                }
                student.FirstMidName = row.FirstMidName;

                //=============================================================================
                if (errorMessage.Length > 0)
                {
                    errorCount += 1;
                    importErrorMessages.Add(string.Format(
                        "第 {0} 列資料發現錯誤：{1}{2}",
                        rowIndex,
                        errorMessage,
                        "<br/>"));
                }
                importStudents.Add(student);
                rowIndex += 1;
            }

            try
            {
                //result.ID = Guid.NewGuid();
                result.Success = errorCount.Equals(0);
                result.RowCount = importStudents.Count;
                result.ErrorCount = errorCount;

                string allErrorMessage = string.Empty;

                foreach (var message in importErrorMessages)
                {
                    allErrorMessage += message;
                }

                result.ErrorMessage = allErrorMessage;

                return result;
            }
            catch (Exception ex)
            {
                throw(ex);
            }
        }

        /// <summary>
        /// Saves the import data.
        /// </summary>
        /// <param name="students">The import zip codes.</param>
        /// <exception cref="System.NotImplementedException"></exception>
        public void SaveImportData(IEnumerable<Student> students)
        {
            try
            {
                //先砍掉全部資料
                foreach (var item in _context.Students.OrderBy(x => x.ID))
                {
                    _context.Students.Remove(item);
                }
                    _context.SaveChanges();

                //再把匯入的資料給存到資料庫
                foreach (var item in students)
                {
                    _context.Students.Add(item);
                }
                    _context.SaveChanges();
            }
            catch (Exception ex)
            {
                throw(ex);
            }
        }
    }
}
