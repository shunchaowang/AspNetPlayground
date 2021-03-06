using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using AspNetCore.Controllers.Utils;
using AspNetPlayground.Models;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace AspNetPlayground.Controllers
{
    public class NpoiController : Controller
    {
        private readonly ILogger<NpoiController> logger;
        private readonly IWebHostEnvironment environment;

        public NpoiController(ILogger<NpoiController> logger, IWebHostEnvironment environment)
        {
            this.logger = logger;
            this.environment = environment;
        }

        public IActionResult Index()
        {
            logger.LogInformation("Npoi Index");
            return View();
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<HttpResult<List<User>>> Import(IFormFile file, CancellationToken cancellationToken)
        {
            logger.LogDebug($"Post file {file.FileName}");
            if (file == null || file.Length <= 0)
            {
                return HttpResult<List<User>>.GetResult(-1, "File is empty.");
            }
            string fileExtension = Path.GetExtension(file.FileName);
            if (!(fileExtension.Equals(".xlsx", StringComparison.OrdinalIgnoreCase)
                || fileExtension.Equals(".xlsx", StringComparison.OrdinalIgnoreCase)))
            {

                return HttpResult<List<User>>.GetResult(-1, "Not Support file extension");
            }

            var list = new List<User>();
            using (var stream = new MemoryStream())
            {
                await file.CopyToAsync(stream, cancellationToken);
                stream.Position = 0;

                ISheet sheet;

                if (fileExtension == ".xls") // excel 97-2000 
                {
                    sheet = new HSSFWorkbook(stream).GetSheetAt(0);
                }
                else // excel 2007 format
                {
                    sheet = new XSSFWorkbook(stream).GetSheetAt(0);
                }

                // test if file is empty
                if (sheet.LastRowNum < 1)
                {

                    return HttpResult<List<User>>.GetResult(-1, "File contains no data.");
                }

                IRow headerRow = sheet.GetRow(0); // Row 0 is the header
                if (headerRow.LastCellNum != 4) // there should be 4 columns
                {
                    return HttpResult<List<User>>.GetResult(-1, "File should have 4 columns");
                }

                // validate all fields are the right type
                for (int i = sheet.FirstRowNum + 1; i <= sheet.LastRowNum; ++i)
                {
                    IRow row = sheet.GetRow(i);
                    if (row == null) continue;

                    // get formatted cell string value and trim
                    string[] cellValues = new string[4];
                    for (int j = 0; j < 4; ++j)
                    {
                        cellValues[j] = row.GetCell(j) == null ? "" : row.GetCell(j).GetFormattedCellValue().Trim();
                    }

                    // parse values and add to list
                    int age;
                    DateTime dob;
                    if (!int.TryParse(cellValues[2], out age))
                    {
                        return HttpResult<List<User>>.GetResult(-1, "Age is not numeric.");
                    }
                    if (!DateTime.TryParse(cellValues[3], out dob))
                    {
                        return HttpResult<List<User>>.GetResult(-1, "Dob is not date.");
                    }

                    list.Add(new User
                    {
                        Name = cellValues[0],
                        Email = cellValues[1],
                        Age = age,
                        Dob = dob
                    });

                }

            }

            return HttpResult<List<User>>.GetResult(0, "OK", list);
        }

        [HttpGet]
        public async Task<HttpResult<string>> Export(CancellationToken cancellationToken)
        {
            string webRoot = environment.WebRootPath;
            string fileName = $"UserList-{DateTime.Now.ToString("yyyyMMddHHmmssfff")}.xlsx";
            string downloadUrl = string.Format("{0}://{1}{2}", Request.Scheme, Request.Host, fileName);

            FileInfo file = new FileInfo(Path.Combine(webRoot, fileName));

            if (file.Exists)
            {
                file.Delete();
                // file = new FileInfo(Path.Combine(webRoot, fileName));
            }

            using (var fs = new FileStream(Path.Combine(webRoot, fileName), FileMode.Create, FileAccess.Write))
            {
                IWorkbook workbook = new XSSFWorkbook();
                ISheet sheet = workbook.CreateSheet("user");
                IRow row = sheet.CreateRow(0);
                row.CreateCell(0).SetCellValue("Name");
                row.CreateCell(1).SetCellValue("Email");
                row.CreateCell(2).SetCellValue("Age");
                row.CreateCell(3).SetCellValue("Dob");

                // query data from database  
                await Task.Yield();

                // data, in reality this should be loaded from database
                //TODO load data from the database
                // here I create a random 10 rows
                for (int i = 1; i <= 10; ++i)
                {
                    row = sheet.CreateRow(i);
                    row.CreateCell(0).SetCellValue($"John {i}");
                    row.CreateCell(1).SetCellValue($"john.{i}@me.com");
                    row.CreateCell(2).SetCellValue(new Random().Next(20, 60));
                    row.CreateCell(3).SetCellValue(DateTime.Now.AddYears(-new Random().Next(30, 50)));
                }

                workbook.Write(fs);
            }

            // using(var stream = new FileStream(Path.Combine(webRoot, fileName), FileMode.Create)) {
            //     await stream.CopyToAsync(memory);
            // }

            return HttpResult<string>.GetResult(0, "OK", downloadUrl);
        }

        [HttpGet]
        public async Task<IActionResult> Download(CancellationToken cancellationToken)
        {
            string webRoot = environment.WebRootPath;
            string fileName = $"UserList-{DateTime.Now.ToString("yyyyMMddHHmmssfff")}.xlsx";
            string downloadUrl = string.Format("{0}://{1}{2}", Request.Scheme, Request.Host, fileName);

            FileInfo file = new FileInfo(Path.Combine(webRoot, fileName));

            if (file.Exists)
            {
                file.Delete();
            }

            using (var fs = new FileStream(Path.Combine(webRoot, fileName), FileMode.Create, FileAccess.Write))
            {
                IWorkbook workbook = new XSSFWorkbook();
                ISheet sheet = workbook.CreateSheet("user");
                IRow row = sheet.CreateRow(0);
                row.CreateCell(0).SetCellValue("Name");
                row.CreateCell(1).SetCellValue("Email");
                row.CreateCell(2).SetCellValue("Age");
                row.CreateCell(3).SetCellValue("Dob");

                // query data from database  
                await Task.Yield();

                // data, in reality this should be loaded from database
                //TODO load data from the database
                // here I create a random 10 rows
                for (int i = 1; i <= 10; ++i)
                {
                    row = sheet.CreateRow(i);
                    row.CreateCell(0).SetCellValue($"John {i}");
                    row.CreateCell(1).SetCellValue($"john.{i}@me.com");
                    row.CreateCell(2).SetCellValue(new Random().Next(20, 60));
                    row.CreateCell(3).SetCellValue(DateTime.Now.AddYears(-new Random().Next(30, 50)));
                }

                workbook.Write(fs);
            }

            byte[] fileBytes = System.IO.File.ReadAllBytes(downloadUrl);
            var stream = new MemoryStream(fileBytes);
            stream.Position = 0;

            return File(stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
        }
    }
}