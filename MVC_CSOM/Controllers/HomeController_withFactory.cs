// using System.Diagnostics;
// using System.Threading.Tasks;
// using Microsoft.AspNetCore.Mvc;
// using Microsoft.Extensions.Logging;
// using MVC_CSOM.Models;
// using MVC_CSOM.Services;

// namespace MVC_CSOM.Controllers
// {
//     public class HomeController : Controller
//     {
//         private readonly ILogger<HomeController> _logger;
//         private readonly ISharePointClientContextFactory _spClientContextFactory;

//         public HomeController(ILogger<HomeController> logger, ISharePointClientContextFactory spClientContextFactory)
//         {
//             _spClientContextFactory = spClientContextFactory;
//             _logger = logger;
//         }

//         public async Task<IActionResult> Index()
//         {
//             using (var ctx = _spClientContextFactory.GetClientContext())
//             {
//                 ctx.Load(ctx.Web);
//                 await ctx.ExecuteQueryAsync();
//                 ViewBag.WebTitle = ctx.Web.Title;
//                 return View();
//             }
//         }

//         public IActionResult Privacy()
//         {
//             return View();
//         }

//         [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
//         public IActionResult Error()
//         {
//             return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
//         }
//     }
// }
