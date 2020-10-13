using System;
using System.Configuration;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;
using Microsoft.EntityFrameworkCore;
using SendEmailMVC.Data;
using SendEmailMVC.Models;
using System.IO;
using DocumentFormat.OpenXml.Drawing.Diagrams;
using Microsoft.Extensions.FileProviders;
using Microsoft.AspNetCore.Http;
using Newtonsoft.Json.Linq;
using Newtonsoft.Json;
using System.Web;
using LinqToExcel;
using SendEmailMVC.Service;
using ClosedXML.Excel;

namespace SendEmailMVC.Controllers
{
    public class StudentsController : Controller
    {
        private readonly SchoolContext _context;
        private IImportDataHelper _importDataHelper;
        private IExportExcelResult _exportExcelResult;

        public StudentsController(
            SchoolContext context,
            IImportDataHelper iimportDataHelper,
            IExportExcelResult _iexportExcelResult)//ImportDataHelper importDataHelper
        {
            _context = context;
            _importDataHelper = iimportDataHelper;
            _exportExcelResult = _iexportExcelResult;
        }

        // GET: Students
        public async Task<IActionResult> Index(string sortOrder, string searchString, string importData)
        { 
            ViewData["NameSortParm"] = String.IsNullOrEmpty(sortOrder) ? "name_desc" : "";
            ViewData["DateSortParm"] = sortOrder == "Date" ? "date_desc" : "Date";
            ViewData["CurrentFilter"] = searchString;
            ViewData["Data"] = importData;
            var students = from s in _context.Students
                           select s;
            if (!String.IsNullOrEmpty(searchString))
            {
                students = students.Where(s => s.LastName.Contains(searchString)
                                       || s.FirstMidName.Contains(searchString));
            }
            switch (sortOrder)
            {
                case "name_desc":
                    students = students.OrderByDescending(s => s.LastName);
                    break;
                case "Date":
                    students = students.OrderBy(s => s.EnrollmentDate);
                    break;
                case "date_desc":
                    students = students.OrderByDescending(s => s.EnrollmentDate);
                    break;
                default:
                    students = students.OrderBy(s => s.LastName);
                    break;
            }
            return View(await students.AsNoTracking().ToListAsync());
        }

        private string fileSavedPath = ConfigurationManager.AppSettings["UploadPath"];
        [HttpPost]
        public ActionResult Upload(IFormFile file)
        {
            var jo = new JObject();
            string result = string.Empty;


            if (file == null)
            {
                jo.Add("Result", false);
                jo.Add("Msg", "請上傳檔案!");
                result = JsonConvert.SerializeObject(jo);
                return Content(result, "application/json");
            }
            if (file.Length <= 0)
            {
                jo.Add("Result", false);
                jo.Add("Msg", "請上傳正確的檔案.");
                result = JsonConvert.SerializeObject(jo);
                return Content(result, "application/json");
            }

            string fileExtName = Path.GetExtension(file.FileName).ToLower();

            if (!fileExtName.Equals(".xls", StringComparison.OrdinalIgnoreCase)
                &&
                !fileExtName.Equals(".xlsx", StringComparison.OrdinalIgnoreCase))
            {
                jo.Add("Result", false);
                jo.Add("Msg", "請上傳 .xls 或 .xlsx 格式的檔案");
                result = JsonConvert.SerializeObject(jo);
                return Content(result, "application/json");
            }

            try
            {
                var uploadResult = this.FileUploadHandler(file);

                var fileName = string.Concat(uploadResult);

                var students = new List<Student>();

                //var helper = ImportDataHelper();
                var checkResult = _importDataHelper.CheckImportData(fileName, students);

                jo.Add("Result", checkResult.Success);
                jo.Add("Msg", checkResult.Success ? string.Empty : checkResult.ErrorMessage);

                if (checkResult.Success)
                {
                    //儲存匯入的資料
                    _importDataHelper.SaveImportData(students);
                }
                result = JsonConvert.SerializeObject(jo);
            }
            catch (Exception ex)
            {
                throw (ex);
            }
            return RedirectToAction(nameof(Index));
        }
        private string FileUploadHandler(IFormFile file)
        {
            string result = string.Empty;

            if (file == null)
            {
                throw new ArgumentNullException("file", "上傳失敗：沒有檔案！");
            }
            if (file.Length <= 0)
            {
                throw new InvalidOperationException("上傳失敗：檔案沒有內容！");
            }

            try
            {
                //string virtualBaseFilePath = Url.Content(fileSavedPath);
                string filePath = Path.Combine(Environment.CurrentDirectory, "Upload");
                
                //建立資料夾
                if (!Directory.Exists(filePath))
                {
                    Directory.CreateDirectory(filePath);
                }

                //string newFileName = string.Concat(
                //    DateTime.Now.ToString("yyyyMMddHHmmssfff"),
                //    Path.GetExtension(file.FileName).ToLower());


                string fullFilePath = Path.Combine(filePath, file.FileName);

                using(var stream = System.IO.File.Create(fullFilePath))
                {
                   file.CopyTo(stream);
                }

                result = fullFilePath;//file.FileName;
            }
            catch (Exception ex)
            {
                throw(ex);
            }
            return result;
        }

        [HttpPost]
        public ActionResult Export()
        {
            JObject jo = new JObject();
            var hasData = !_context.Students.Count().Equals(0);
            if (!hasData)
            {
                jo.Add("Msg", hasData.ToString());
                return Content(JsonConvert.SerializeObject(jo), "application/json");
            }
            else
            {
                var exportFileName = string.Concat(
                     "Students_",
                     DateTime.Now.ToString("yyyyMMddHHmmss"),
                     ".xlsx");

                var exportData = new ExportExcelResult
                {
                   SheetName = "outputfile",
                   FileName = exportFileName,
                   ExportData = _context.Students.ToList()
                };

                //_exportExcelResult.ExecuteResult(exportData);
                try
                {
                    var workbook = new XLWorkbook();

                    IXLWorksheet worksheet = workbook.Worksheets.Add("outputfile");
                    //worksheet.Cell(1, 1).Value = "Id";
                    worksheet.Cell(1, 1).Value = "LastName";
                    worksheet.Cell(1, 2).Value = "FirstMidName";
                    worksheet.Cell(1, 3).Value = "EnrollmentDate";
                    for (int index = 1; index <= exportData.ExportData.Count; index++)
                    {
                        //worksheet.Cell(index + 1, 1).Value = exportData.ExportData[index - 1].ID;
                        worksheet.Cell(index + 1, 1).Value = exportData.ExportData[index - 1].LastName;
                        worksheet.Cell(index + 1, 2).Value = exportData.ExportData[index - 1].FirstMidName;
                        worksheet.Cell(index + 1, 3).Value = exportData.ExportData[index - 1].EnrollmentDate;
                    }

                    var contentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

                    using (var stream = new MemoryStream())
                    {
                        workbook.SaveAs(stream);
                        var content = stream.ToArray();
                        return File(content, contentType, exportData.FileName);
                    }

                    //workbook.SaveAs(@"C:\Users\v-ivowan\Desktop\" + dt.FileName);
                }
                catch (Exception ex)
                {
                    throw (ex);
                }

                //return RedirectToAction(nameof(Index));
            }
            //var exportSpource = this.GetExportData();
            //var dt = JsonConvert.DeserializeObject<Student>(exportSpource.ToString()); 
        }

        //private JArray GetExportData()
        //{
        //    var query = _context.Students.OrderBy(x => x.ID);

        //    JArray jObjects = new JArray();

        //    foreach (var item in query)
        //    {
        //        var jo = new JObject();
        //        jo.Add("ID", item.ID);
        //        jo.Add("LastName", item.LastName);
        //        jo.Add("FirstMidName", item.FirstMidName);
        //        jo.Add("EnrollmentDate", item.EnrollmentDate);
        //        jObjects.Add(jo);
        //    }
        //    return jObjects;
        //}

        // GET: Students/Details/5
        public async Task<IActionResult> Details(int? id)
        {
            if (id == null)
            {
                return NotFound();
            }

            var student = await _context.Students
                .Include(s => s.Enrollments)
                    .ThenInclude(e => e.Course)
                .AsNoTracking()
                .FirstOrDefaultAsync(m => m.ID == id);
            if (student == null)
            {
                return NotFound();
            }

            return View(student);
        }

        // GET: Students/Create
        public IActionResult Create()
        {
            return View();
        }

        // POST: Students/Create
        // To protect from overposting attacks, enable the specific properties you want to bind to, for 
        // more details, see http://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Create([Bind("ID,LastName,FirstMidName,EnrollmentDate")] Student student)
        {
            try
            {
                if (ModelState.IsValid)
                {
                    _context.Add(student);
                    await _context.SaveChangesAsync();
                    return RedirectToAction(nameof(Index));
                }
            }
            catch (DbUpdateException /* ex */)
            {
                //Log the error (uncomment ex variable name and write a log.
                ModelState.AddModelError("", "Unable to save changes. " +
                    "Try again, and if the problem persists " +
                    "see your system administrator.");
            }
            return View(student);
        }

        // GET: Students/Edit/5
        public async Task<IActionResult> Edit(int? id)
        {
            if (id == null)
            {
                return NotFound();
            }

            var student = await _context.Students.FindAsync(id);
            if (student == null)
            {
                return NotFound();
            }
            return View(student);
        }

        // POST: Students/Edit/5
        // To protect from overposting attacks, enable the specific properties you want to bind to, for 
        // more details, see http://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost, ActionName("Edit")]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> EditPost(int? id)
        {
            if (id == null)
            {
                return NotFound();
            }
            var studentToUpdate = await _context.Students.FirstOrDefaultAsync(s => s.ID == id);
            if (await TryUpdateModelAsync<Student>(
                studentToUpdate,
                "",
                s => s.FirstMidName, s => s.LastName, s => s.EnrollmentDate))
            {
                try
                {
                    await _context.SaveChangesAsync();
                    return RedirectToAction(nameof(Index));
                }
                catch (DbUpdateException /* ex */)
                {
                    //Log the error (uncomment ex variable name and write a log.)
                    ModelState.AddModelError("", "Unable to save changes. " +
                        "Try again, and if the problem persists, " +
                        "see your system administrator.");
                }
            }
            return View(studentToUpdate);
        }

        // GET: Students/Delete/5
        public async Task<IActionResult> Delete(int? id, bool? saveChangesError = false)
        {
            if (id == null)
            {
                return NotFound();
            }

            var student = await _context.Students
                .AsNoTracking()
                .FirstOrDefaultAsync(m => m.ID == id);
            if (student == null)
            {
                return NotFound();
            }

            if (saveChangesError.GetValueOrDefault())
            {
                ViewData["ErrorMessage"] =
                    "Delete failed. Try again, and if the problem persists " +
                    "see your system administrator.";
            }

            return View(student);
        }

        // POST: Students/Delete/5
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> DeleteConfirmed(int id)
        {
            var student = await _context.Students.FindAsync(id);
            if (student == null)
            {
                return RedirectToAction(nameof(Index));
            }

            try
            {
                _context.Students.Remove(student);
                await _context.SaveChangesAsync();
                return RedirectToAction(nameof(Index));
            }
            catch (DbUpdateException /* ex */)
            {
                //Log the error (uncomment ex variable name and write a log.)
                return RedirectToAction(nameof(Delete), new { id = id, saveChangesError = true });
            }
        }


    }
}
