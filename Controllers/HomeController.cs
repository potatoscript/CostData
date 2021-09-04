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
using System.Text;
using System.Dynamic;

namespace CostNag.Controllers
{
    public class HomeController : Controller
    {


        CostAPI _api = new CostAPI();
         
        public async Task<IActionResult> Index()
        {

            ViewBag.plant = "-";
            ViewBag.item_spec = "-";
            ViewBag.issue_date = DateTime.Now.AddDays(0).ToString("dd-MM-yyyy");
            ViewBag.section = "-";
            ViewBag.doc_no = "-";
            ViewBag.wr_no = "-";
            ViewBag.sales = "-";
            ViewBag.revision_no = "-";
            ViewBag.checked_date = "-";
            ViewBag.approved_by = "-";
            ViewBag.customer = "-";
            ViewBag.parts_code = "-";
            ViewBag.item = "-";
            ViewBag.product = "-";
            ViewBag.product_type = "-";
            ViewBag.size = "-";
            ViewBag.business_type = "-";
            ViewBag.qty_month = 0;
            ViewBag.exchange_rate = 1;
            ViewBag.target_price_bht = 0.0;
            ViewBag.target_price_export = 0.0;
            ViewBag.production_qty_day = 0;
            ViewBag.working_day = 21;


            ListModel list = new ListModel();
            ViewData["plant"] = list.plant;
            ViewData["item_spec"] = list.item_spec;

            List<Cost> cost = new List<Cost>();
            HttpClient client = _api.Initial();
            HttpResponseMessage res = await client.GetAsync("api/data/get-all-costs");
            if (res.IsSuccessStatusCode)
            {
                var result = res.Content.ReadAsStringAsync().Result;
                cost = JsonConvert.DeserializeObject<List<Cost>>(result);
                foreach (var o in cost)
                {
                    list.partscode.Add(new ListModel
                    {
                        code = o.parts_code,
                        id = o.CostId.ToString()

                    });
                }
                
            }
            List<ListModel> model = list.partscode.ToList();

            ViewData["partscode"] = model;

            dynamic mymodel = new ExpandoObject();
            mymodel.code = model;
            return View(mymodel);
        }


        public async Task<IActionResult> Read(int id)
        {
            ViewBag.CostId = id;

            ListModel list = new ListModel();
            ViewData["plant"] = list.plant;
            ViewData["item_spec"] = list.item_spec;
            List<Cost> cost = new List<Cost>();
            HttpClient client = _api.Initial();
            HttpResponseMessage res = await client.GetAsync("api/data/get-all-costs");
            if (res.IsSuccessStatusCode)
            {
                var result = res.Content.ReadAsStringAsync().Result;
                cost = JsonConvert.DeserializeObject<List<Cost>>(result);
                foreach (var o in cost)
                {
                    list.partscode.Add(new ListModel
                    {
                        code = o.parts_code,
                        id = o.CostId.ToString()

                    });
                }
            }
            List<ListModel> model = list.partscode.ToList();
            ViewData["partscode"] = model;

            Cost costdata = new Cost();
            
            //HttpClient clientdata = _api.Initial();
            //HttpResponseMessage resdata = await clientdata.GetAsync("api/data/get-cost-by-id/5" + id);


            HttpClient clientdata = _api.Initial();

            var action = "api/data/get-cost-by-id/5";
            HttpResponseMessage resdata = await clientdata.GetAsync(action).ConfigureAwait(false);
           // HttpResponseMessage resdata = await clientdata.GetAsync("api/data/get-all-costs");

            resdata.EnsureSuccessStatusCode();

            if (resdata.IsSuccessStatusCode)
            {
                var resultdata = resdata.Content.ReadAsStringAsync().Result;
                costdata = JsonConvert.DeserializeObject<Cost>(resultdata);

                ViewBag.customer = costdata.customer;
                ViewBag.parts_code = costdata.parts_code;

 

            }

            dynamic mymodel = new ExpandoObject();
           // mymodel.code = costdata;
            return View("Index",mymodel); 
        }

        public IActionResult Privacy()
        {
            return View();
        }



        public async Task<ActionResult<Cost>> Save(Cost model)
        {

            HttpClient client = _api.Initial();

            var content = new StringContent(JsonConvert.SerializeObject(model),Encoding.UTF8, "application/json");

            var action = "api/data/update-cost-by-id/"+ model.CostId;
            if (model.CostId == 0)
            {
               action = "api/data/add-cost";
            }

            HttpResponseMessage res = await client.PostAsync(action, content).ConfigureAwait(false);


            res.EnsureSuccessStatusCode();
            if (res.IsSuccessStatusCode)
            {

                var result = res.Content.ReadAsStringAsync().Result;

                return Json(new
                {
                    isValid = true
                });

            }

            return Json(new
            {
                isValid = false
            });
        }



        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}
