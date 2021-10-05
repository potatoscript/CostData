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
    public class ProcessController : Controller
    {

        CostAPI _api = new CostAPI();

        public async Task<IActionResult> Index(string p_doc_no)
        {
            ViewBag.p_doc_no = p_doc_no;
            ViewBag.p_process_cost = 0.0;

            ViewBag.p_process_name = "-";
            ViewBag.p_working_day = 0.0;
            ViewBag.p_working_time_day = 0.0;
            ViewBag.p_working_time_month = 0.0;
            ViewBag.p_shift = 0.0;
            ViewBag.p_worker = 0.0;
            ViewBag.p_direct_labour = 0.0;
            ViewBag.p_total_labour_cost = 0.0;
            ViewBag.p_machine_qty = 0.0;
            ViewBag.p_area = 0.0;
            ViewBag.p_special_material = 0.0;
            ViewBag.p_plant_maintenance = 0.0;
            ViewBag.p_plant_maintenance_unit = 0.0;
            ViewBag.p_total_machine_cost = 0.0;
            ViewBag.p_machine_usage_day = 96;
            ViewBag.p_machine_cost_month = 0.0;
            ViewBag.p_machine_cost_month_percentage = 0.0;
            ViewBag.p_machine_cost_month_percentage_unit = 10;
            ViewBag.p_consumption_kwh = 0.0;
            ViewBag.p_consumption_unit = 0.24;
            ViewBag.p_consumption_sgd = 0.0;
            ViewBag.p_consumption_rate = 0.0;
            ViewBag.p_utility_electric = 0.0;
            ViewBag.p_machine_utility_cost = 0.0;
            ViewBag.p_labour_electric_cost = 0.0;
            ViewBag.p_charge = 0.0;
            ViewBag.p_cycle_time = 0.0;
            ViewBag.p_cycle_time_unit = 0.0;
            ViewBag.p_time = 0.0;
            ViewBag.p_capacity = 0.0;
            ViewBag.p_time_g = 0.0;
            ViewBag.p_efficiency = 0.0;
            ViewBag.p_production_capacity = 0.0;
            ViewBag.p_production_cycle_time = 0.0;
            ViewBag.p_special_input = 0.0;
            ViewBag.p_direct_process_cost = 0.0;
            ViewBag.p_labour_cost_percentage = 0.0;
            ViewBag.p_machine_cost_percentage = 0.0;
            ViewBag.p_overhead_cost_percentage = 0.0;
            ViewBag.p_total_cost_percentage = 0.0;


            ListModel list = new ListModel();
            List<Process> process_ = new List<Process>();
            Process processdata = new Process();

            HttpClient clientdata = _api.Initial();

            var action = "api/process/get-process-by-id/" + p_doc_no;
            HttpResponseMessage resdata = await clientdata.GetAsync(action).ConfigureAwait(false);

            resdata.EnsureSuccessStatusCode();

            if (resdata.IsSuccessStatusCode)
            {
                var resultdata = resdata.Content.ReadAsStringAsync().Result;
                //processdata = JsonConvert.DeserializeObject<Process>(resultdata);

                //ViewBag.p_process_cost = processdata.process_cost;

                process_ = JsonConvert.DeserializeObject<List<Process>>(resultdata);
                foreach (var o in process_)
                {
                    list.process.Add(new ListModel
                    {
                        name = o.process_name,
                        cost = o.direct_process_cost.ToString()

                    });
                }

            }


            List<ListModel> model = list.process.ToList();
            ViewData["process"] = model;

            dynamic mymodel = new ExpandoObject();
            mymodel.process = model;


            return View(mymodel);
        }

        public async Task<ActionResult<Process>> Save(Process model)
        {

            HttpClient client = _api.Initial();

            var content = new StringContent(JsonConvert.SerializeObject(model), Encoding.UTF8, "application/json");

            var action = "api/process/update-process-by-id/" + model.ProcessId;
            if (model.ProcessId == 0)
            {
                action = "api/process/add-process";
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

    }
}