using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;

namespace AspNetPlayground.Controllers
{
    public class NpoiController : Controller
    {
        private readonly ILogger<NpoiController> logger;

        public NpoiController(ILogger<NpoiController> logger)
        {
            this.logger = logger;
        }
    }
}