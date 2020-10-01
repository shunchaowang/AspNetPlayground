using System.Collections.Generic;
using System.Diagnostics;
using System.Threading;
using System.Threading.Tasks;
using AspNetPlayground.Models;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;

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
            return null;
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