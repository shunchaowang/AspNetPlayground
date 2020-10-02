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

        [HttpPost("import")]
        [AutoValidateAntiforgeryToken]
        public async Task<HttpResult<List<User>>> Import(IFormFile file, CancellationToken cancellationToken)
        {
            if (file == null || file.Length <= 0)
            {
                return HttpResult<List<User>>.GetResult(-1, "File is empty.");
            }

            return null;
        }

        [HttpPost("importsync")]
        [AutoValidateAntiforgeryToken]
        public HttpResult<List<User>> Import(IFormFile file)
        {
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
                file.CopyToAsync(stream);
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
                    list.Add(new User
                    {
                        Name = cellValues[0],
                        Email = cellValues[1],
                        Age = int.Parse(cellValues[3]),
                        Dob = DateTime.Parse(cellValues[3])
                    });

                }

            }

            return HttpResult<List<User>>.GetResult(0, "OK", list);
        }

        [HttpGet("export")]
        public async Task<HttpResult<string>> Export(CancellationToken cancellationToken)
        {
            return null;
        }

        [HttpGet("download")]
        public async Task<IActionResult> Download(CancellationToken cancellationToken)
        {
            return null;
        }
    }
}