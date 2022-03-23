using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using CostNag.Helper;
using CostNag.Models;
using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json;

namespace CostNag.Controllers
{
    public class ToolingController : Controller
    {

        CostAPI _api = new CostAPI();

        public async Task<IActionResult> Index()
        {

            ViewBag.description = "-";
            ViewBag.source = "-";
            ViewBag.qty = 1;
            ViewBag.unit = "-";
            ViewBag.price = 0;
            ViewBag.ToolingId = 0;
            ViewBag.od = 0;
            ViewBag.od_max = 0;
            ViewBag.type = "-";


            Tooling list = new Tooling();
            List<Tooling> tooling_ = new List<Tooling>();

            HttpClient clientdata = _api.Initial();

            ViewBag.PageCount = 1;
            ViewBag.CurrentPageIndex = 1;

            ViewBag.first = 1;
            ViewBag.last = 10;

            var action = "api/tooling/get-tooling-by-page/" + 1;
            HttpResponseMessage resdata = await clientdata.GetAsync(action).ConfigureAwait(false);

            resdata.EnsureSuccessStatusCode();

            if (resdata.IsSuccessStatusCode)
            {
                var resultdata = resdata.Content.ReadAsStringAsync().Result;
                tooling_ = JsonConvert.DeserializeObject<List<Tooling>>(resultdata);

                ViewBag.PageCount = tooling_[0].PageCount;
                ViewBag.CurrentPageIndex = tooling_[0].CurrentPageIndex;

                foreach (var o in tooling_)
                {
                    list.data.Add(new Tooling
                    {
                        description = o.description,
                        source = o.source,
                        qty = o.qty,
                        unit = o.unit,
                        price = o.price,
                        od = o.od,
                        od_max = o.od_max,
                        type = o.type,
                        ToolingId = o.ToolingId

                    });
                }

            }
            List<Tooling> model = list.data.ToList();
            ViewData["data"] = model;

            return View();
        }

        public async Task<IActionResult> IndexPage(int CurrentPage, int first, int last)
        {

            ViewBag.description = "-";
            ViewBag.source = "-";
            ViewBag.qty = 1;
            ViewBag.unit = "-";
            ViewBag.price = 0;
            ViewBag.ToolingId = 0;
            ViewBag.od = 0;
            ViewBag.od_max = 0;
            ViewBag.type = "-";

            Tooling list = new Tooling();
            List<Tooling> tooling_ = new List<Tooling>();

            HttpClient clientdata = _api.Initial();

            var action = "api/tooling/get-tooling-by-page/" + CurrentPage;

            HttpResponseMessage resdata = await clientdata.GetAsync(action);

            resdata.EnsureSuccessStatusCode();

            ViewBag.PageCount = 1;
            ViewBag.CurrentPageIndex = 1;

            ViewBag.first = first;
            ViewBag.last = last;

            resdata.EnsureSuccessStatusCode();

            if (resdata.IsSuccessStatusCode)
            {
                var resultdata = resdata.Content.ReadAsStringAsync().Result;
                tooling_ = JsonConvert.DeserializeObject<List<Tooling>>(resultdata);

                ViewBag.PageCount = tooling_[0].PageCount;
                ViewBag.CurrentPageIndex = tooling_[0].CurrentPageIndex;

                foreach (var o in tooling_)
                {
                    list.data.Add(new Tooling
                    {
                        description = o.description,
                        source = o.source,
                        qty = o.qty,
                        unit = o.unit,
                        price = o.price,
                        od = o.od,
                        od_max = o.od_max,
                        type = o.type,
                        ToolingId = o.ToolingId

                    });
                }

            }

            List<Tooling> model = list.data.ToList();
            ViewData["data"] = model;

            return View("Index");
        }

        public async Task<ActionResult<Tooling>> Save(Tooling model)
        {

            HttpClient client = _api.Initial();

            var content = new StringContent(JsonConvert.SerializeObject(model), Encoding.UTF8, "application/json");

            var action = "api/tooling/add-tooling-data";
            if (model.ToolingId != 0)
            {
                action = "api/tooling/update-tooling-by-id/" + model.ToolingId;
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

        public async void Delete(
           Process model,
           bool confirm,
           int Id
       )
        {
            if (Id == null || Id == 0)  //this is used for the validation as well but in the server side
            {
                //return NotFound();
            }
            else
            {
                if (ModelState.IsValid && confirm == true)
                {
                    HttpClient client = _api.Initial();
                    var content = new StringContent(JsonConvert.SerializeObject(model), Encoding.UTF8, "application/json");
                    var action = "api/tooling/delete-tooling-by-id/" + Id;
                    HttpResponseMessage res = await client.PostAsync(action, content).ConfigureAwait(false);
                    res.EnsureSuccessStatusCode();
                    if (res.IsSuccessStatusCode)
                    {
                        var result = res.Content.ReadAsStringAsync().Result;
                    }
                }
            }

        }


    }
}