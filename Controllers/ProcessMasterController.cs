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
    public class ProcessMasterController : Controller
    {

        CostAPI _api = new CostAPI();

        public async Task<IActionResult> Index(string p_doc_no, int p_od)
        {
            CostAPI _api = new CostAPI();

            ViewBag.p_doc_no = p_doc_no;
            ViewBag.p_od = p_od;
            ViewBag.p_process_cost = 0;
            ViewBag.p_process_name = "-";
            ViewBag.p_process_type = "-";
            ViewBag.p_overhead_cost = 0;
            ViewBag.p_machine_cost = 0;
            ViewBag.p_labor_cost = 0;
            ViewBag.p_total_cost = 0;

            if(p_od>=5 && p_od <= 50)
            {
                ViewBag.p_od_min = 5;
                ViewBag.p_od_max = 50;
            }
            if (p_od >= 51 && p_od <= 100)
            {
                ViewBag.p_od_min = 51;
                ViewBag.p_od_max = 100;
            }
            if (p_od >= 101 && p_od <= 120)
            {
                ViewBag.p_od_min = 101;
                ViewBag.p_od_max = 120;
            }
            if (p_od >= 121 && p_od <= 150)
            {
                ViewBag.p_od_min = 121;
                ViewBag.p_od_max = 150;
            }
            if (p_od >= 151 )
            {
                ViewBag.p_od_min = 151;
                ViewBag.p_od_max = 10000000;
            }


            ProcessMaster list = new ProcessMaster();


            List<ProcessMaster> data_ = new List<ProcessMaster>();
            List<string> processname_ = new List<string>();

            HttpClient clientdata = _api.Initial();

            var action = "api/processmaster/get-processmaster-by-od/" + p_od;
            HttpResponseMessage resdata = await clientdata.GetAsync(action).ConfigureAwait(false);

            resdata.EnsureSuccessStatusCode();

            if (resdata.IsSuccessStatusCode)
            {
                var resultdata = resdata.Content.ReadAsStringAsync().Result;
                data_ = JsonConvert.DeserializeObject<List<ProcessMaster>>(resultdata);
                foreach (var o in data_)
                {
                    list.data.Add(new ProcessMaster
                    {
                        process_name = o.process_name,
                        process_type = o.process_type,
                        od_min = o.od_min,
                        od_max = o.od_max,
                        overhead_cost = o.overhead_cost,
                        machine_cost = o.machine_cost,
                        labor_cost = o.labor_cost,
                        total_cost = o.total_cost,
                        ProcessMasterId = o.ProcessMasterId
                    });
                }

            }
            List<ProcessMaster> model = list.data.ToList();
            ViewData["data"] = model;

            /*
            var action2 = "api/processmaster/get-processname-by-od/" + p_od;
            HttpResponseMessage resdata2 = await clientdata.GetAsync(action2).ConfigureAwait(false);

            resdata2.EnsureSuccessStatusCode();

            if (resdata2.IsSuccessStatusCode)
            {
                var resultdata2 = resdata2.Content.ReadAsStringAsync().Result;
                processname_ = JsonConvert.DeserializeObject<List<string>>(resultdata2);
                foreach (var o in processname_)
                {
                    list.process.Add(o);
                }

            }
            List<string> model2 = list.process.ToList();
            ViewData["processname"] = model2;
            */


            return View();
        }


        public async Task<ActionResult<ProcessMaster>> Save(ProcessMaster model)
        {

            HttpClient client = _api.Initial();

            var content = new StringContent(JsonConvert.SerializeObject(model), Encoding.UTF8, "application/json");

            var action = "api/processmaster/update-processmaster-by-id/" + model.ProcessMasterId;
            if (model.ProcessMasterId == 0)
            {
                action = "api/processmaster/add-processmaster";
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
                    var action = "api/processmaster/delete-processmaster-by-id/" + Id;
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