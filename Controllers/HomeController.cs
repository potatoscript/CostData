using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using CostNag.Models;
using CostNag.Helper;
using System.Net.Http;
using Newtonsoft.Json;

namespace CostNag.Controllers
{
    public class HomeController : Controller
    {


        CostAPI _api = new CostAPI();
         
        public async Task<IActionResult> Index()
        {
            List<Cost> cost = new List<Cost>();
            HttpClient client = _api.Initial();
            HttpResponseMessage res = await client.GetAsync("api/data/get-all-costs");
            if (res.IsSuccessStatusCode)
            {
                var result = res.Content.ReadAsStringAsync().Result;
                cost = JsonConvert.DeserializeObject<List<Cost>>(result);
            }
            return View(cost);
        }

        public IActionResult Privacy()
        {
            return View();
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}
