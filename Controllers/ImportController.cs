using ExcelFileHandling_WebTech.Services;
using OfficeOpenXml;
using System;
using System.Data;
using System.Web;
using System.Web.Mvc;

namespace ExcelFileHandling_WebTech.Controllers
{
    public class ImportController : Controller
    {
        private readonly ImportServices _importService;

        public ImportController()
        {
            _importService = new ImportServices();
        }

        // GET: Import
        public ActionResult Index()
        {
            return View();
        }

        [HttpPost]
        public ActionResult Create(HttpPostedFileBase file, string tableName)
        {
            try
            {
                if (file == null || file.ContentLength == 0)
                {
                    ViewBag.ErrorMessage = "No file uploaded.";
                }
                else if (!file.ContentType.Equals("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", StringComparison.OrdinalIgnoreCase))
                {
                    ViewBag.ErrorMessage = "Invalid file format. Please upload an Excel file.";
                }
                else if (string.IsNullOrEmpty(tableName))
                {
                    ViewBag.ErrorMessage = "Table name cannot be empty.";
                }
                else
                {
                    // Extract data from Excel file
                    using (var package = new ExcelPackage(file.InputStream))
                    {
                        foreach (var worksheet in package.Workbook.Worksheets)
                        {
                            DataTable data = _importService.ExtractDataFromExcel(worksheet);

                            // Process Excel data and insert into the database
                            _importService.ProcessExcelData(tableName, data, worksheet.Name);
                        }
                    }

                    ViewBag.Message = $"File uploaded and data processed for tables with prefix '{tableName}'.";
                }

                return View("Index");
            }
            catch (Exception ex)
            {
                // Implement robust error handling
                ViewBag.ErrorMessage = $"An error occurred: {ex.Message}";
                return View("Index");
            }
        }
    }
}
