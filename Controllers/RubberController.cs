using System;
using System.Collections.Generic;
using System.Dynamic;
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
    public class RubberController : Controller
    {
        CostAPI _api = new CostAPI();

        public async Task<IActionResult> Index()
        {
            ViewBag.material_name = "-";
            ViewBag.price_kg = 0;
            ViewBag.mixing_process_cost = 0;
            ViewBag.weight_g = 0;
            ViewBag.yield_rate = 0;



            Rubber list = new Rubber();
            List<Rubber> rubber_ = new List<Rubber>();

            HttpClient clientdata = _api.Initial();

            var action = "api/rubber/get-all-rubbers";
            HttpResponseMessage resdata = await clientdata.GetAsync(action).ConfigureAwait(false);

            resdata.EnsureSuccessStatusCode();

            if (resdata.IsSuccessStatusCode)
            {
                var resultdata = resdata.Content.ReadAsStringAsync().Result;
                rubber_ = JsonConvert.DeserializeObject<List<Rubber>>(resultdata);
                foreach (var o in rubber_)
                {
                    list.data.Add(new Rubber
                    {
                        material_name = o.material_name,
                        price_kg = o.price_kg,
                        mixing_process_cost = o.mixing_process_cost,
                        weight_g = o.weight_g,
                        yield_rate = o.yield_rate,
                        RubberId = o.RubberId

                    });
                }

            }
            List<Rubber> model = list.data.ToList();
            ViewData["data"] = model;


            return View();
        }


        public async Task<ActionResult<Rubber>> Save(Rubber model)
        {

            HttpClient client = _api.Initial();

            var content = new StringContent(JsonConvert.SerializeObject(model), Encoding.UTF8, "application/json");

            var action = "api/rubber/update-rubber-by-id/" + model.RubberId;
            if (model.RubberId == 0)
            {
                action = "api/rubber/add-rubber";
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
                    var action = "api/rubber/delete-rubber-by-id/" + Id;
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