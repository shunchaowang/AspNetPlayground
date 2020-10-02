using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
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

                IRow headerRow = sheet.GetRow(0); // Row 0 is the header
                if (headerRow.LastCellNum != 4) // there should be 4 columns
                {
                    return HttpResult<List<User>>.GetResult(-1, "File should have 4 columns");
                }
            }

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