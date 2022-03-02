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
using Microsoft.AspNetCore.Mvc.Rendering;
using Newtonsoft.Json;

namespace CostNag.Controllers
{
    public class ProcessMasterController : Controller
    {

        CostAPI _api = new CostAPI();

        public async Task<IActionResult> Index(string p_type)
        {
            CostAPI _api = new CostAPI();

            ListModel m = new ListModel();

            ViewBag.Types = m.GetTypes().Select(x => new SelectListItem()
            {
                Text = x.processType,
                Value = x.processType
            }).ToList();


            ViewBag.p_od_min = 0;
            ViewBag.p_od_max = 0;
            ViewBag.p_process_cost = 0;
            ViewBag.p_process_name = "-";
            ViewBag.p_process_type = "-";
            ViewBag.p_overhead_cost = 0;
            ViewBag.p_machine_cost = 0;
            ViewBag.p_labor_cost = 0;
            ViewBag.p_total_cost = 0;
            ViewBag.ProcessId = 0;


            ViewBag.p_od_min = 5;
            ViewBag.p_od_max = 50;



            ProcessMaster list = new ProcessMaster();
            List<ProcessMaster> data_ = new List<ProcessMaster>();

            CostProcess list2 = new CostProcess();
            List<CostProcess> data2_ = new List<CostProcess>();

            HttpClient clientdata = _api.Initial();

            var action = "api/processmaster/get-processmaster-by-type/" + p_type;
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
                        overhead_cost = double.Parse(o.overhead_cost.ToString("0.0000")),
                        machine_cost = double.Parse(o.machine_cost.ToString("0.0000")),
                        labor_cost = double.Parse(o.labor_cost.ToString("0.0000")),
                        total_cost = double.Parse(o.total_cost.ToString("0.0000")),
                        ProcessMasterId = o.ProcessMasterId
                    });
                }

            }
            List<ProcessMaster> model = list.data.ToList();
            ViewData["data"] = model;

            return View();
        }


        public async Task<ActionResult<ProcessMaster>> Save(ProcessMaster model)
        {

            HttpClient client = _api.Initial();

            var content = new StringContent(JsonConvert.SerializeObject(model), Encoding.UTF8, "application/json");

            var action = "api/processmaster/add-processmaster"; 
            if (model.ProcessMasterId != 0)
            {
                action = "api/processmaster/update-processmaster-by-id/" + model.ProcessMasterId;
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