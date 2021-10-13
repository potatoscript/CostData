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
    public class ProcessController : Controller
    {

        CostAPI _api = new CostAPI();

        public async Task<IActionResult> Index(string p_doc_no,int p_id)
        {
            ViewBag.ProcessId = p_id;

            ViewBag.p_doc_no = p_doc_no;
            ViewBag.p_process_cost = 0.0;

            ViewBag.p_process_name = "-";
            ViewBag.p_working_day = 0.0;
            ViewBag.p_working_time_day = 0.0;
            ViewBag.p_working_time_month = 0.0;
            ViewBag.p_shift = 0.0;
            ViewBag.p_worker = 0.0;
            ViewBag.p_direct_labour_unit = 550;
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
            ViewBag.p_machine_cost_month2 = 0.0;
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
            ViewBag.p_cycle_time_unit = 5.43;
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

            /*
            List<string> master_ = new List<string>();

            HttpClient client = _api.Initial();
            HttpResponseMessage res = await client.GetAsync("api/process/get-process-master");
            if (res.IsSuccessStatusCode)
            {
                var result = res.Content.ReadAsStringAsync().Result;
                master_ = JsonConvert.DeserializeObject<List<string>>(result);
                
            }
            List<string> model_process = master_.ToList();
            ViewData["processMaster"] = model_process;
            */



            List<Process> process_ = new List<Process>();

            HttpClient clientdata = _api.Initial();

            var action = "api/process/get-process-by-docno/" + p_doc_no;
            HttpResponseMessage resdata = await clientdata.GetAsync(action).ConfigureAwait(false);

            resdata.EnsureSuccessStatusCode();

            if (resdata.IsSuccessStatusCode)
            {
                var resultdata = resdata.Content.ReadAsStringAsync().Result;
                process_ = JsonConvert.DeserializeObject<List<Process>>(resultdata);
                foreach (var o in process_)
                {
                    list.process.Add(new ListModel
                    {
                        name = o.process_name,
                        cost = o.direct_process_cost.ToString(),
                        id = o.ProcessId.ToString()

                    });
                }

            }
            List<ListModel> model = list.process.ToList();
            ViewData["process"] = model;


            try
            {
                //get process by id
                Process p = new Process();

                var action2 = "api/process/get-process-by-id/" + p_id;
                HttpResponseMessage resdata2 = await clientdata.GetAsync(action2).ConfigureAwait(false);

                resdata2.EnsureSuccessStatusCode();

                if (resdata2.IsSuccessStatusCode)
                {
                    var resultdata2 = resdata2.Content.ReadAsStringAsync().Result;
                    p = JsonConvert.DeserializeObject<Process>(resultdata2);

                    //ViewBag.p_doc_no = p.doc_no;
                    ViewBag.p_process_cost = p.direct_process_cost;

                    ViewBag.p_process_name = p.process_name;
                    ViewBag.p_working_day = p.working_day;
                    ViewBag.p_working_time_day = p.working_time_day;
                    ViewBag.p_working_time_month = p.working_time_month;
                    ViewBag.p_shift = p.shift;
                    ViewBag.p_worker = p.worker;
                    ViewBag.p_direct_labour_unit = p.direct_labour_unit;
                    ViewBag.p_direct_labour = p.direct_labour;
                    ViewBag.p_total_labour_cost = p.total_labour_cost;
                    ViewBag.p_machine_qty = p.machine_qty;
                    ViewBag.p_area = p.area;
                    ViewBag.p_special_material = p.special_material;
                    ViewBag.p_plant_maintenance = p.plant_maintenance;
                    ViewBag.p_plant_maintenance_unit = p.plant_maintenance_unit;
                    ViewBag.p_total_machine_cost = p.total_machine_cost;
                    ViewBag.p_machine_usage_day = p.machine_usage_day;
                    ViewBag.p_machine_cost_month = p.machine_cost_month;
                    ViewBag.p_machine_cost_month2 = p.machine_cost_month2;
                    ViewBag.p_machine_cost_month_percentage = p.machine_cost_month_percentage;
                    ViewBag.p_machine_cost_month_percentage_unit = p.machine_cost_month_percentage_unit;
                    ViewBag.p_consumption_kwh = p.consumption_kwh;
                    ViewBag.p_consumption_unit = p.consumption_unit;
                    ViewBag.p_consumption_sgd = p.consumption_sgd;
                    ViewBag.p_consumption_rate = p.consumption_rate;
                    ViewBag.p_utility_electric = p.utility_electric;
                    ViewBag.p_machine_utility_cost = p.machine_utility_cost;
                    ViewBag.p_labour_electric_cost = p.labour_electric_cost;
                    ViewBag.p_charge = p.charge;
                    ViewBag.p_cycle_time = p.cycle_time;
                    ViewBag.p_cycle_time_unit = p.cycle_time_unit;
                    ViewBag.p_time = p.time;
                    ViewBag.p_capacity = p.capacity;
                    ViewBag.p_time_g = p.time_g;
                    ViewBag.p_efficiency = p.efficiency;
                    ViewBag.p_production_capacity = p.production_capacity;
                    ViewBag.p_production_cycle_time = p.production_cycle_time;
                    ViewBag.p_special_input = p.special_input;
                    ViewBag.p_direct_process_cost = p.direct_process_cost;
                    ViewBag.p_labour_cost_percentage = p.labour_cost_percentage;
                    ViewBag.p_machine_cost_percentage = p.machine_cost_percentage;
                    ViewBag.p_overhead_cost_percentage = p.overhead_cost_percentage;
                    ViewBag.p_total_cost_percentage = p.total_cost_percentage;

                }
            }
            catch(Exception e) { }
            
            

            //the following dynamic model was keep for review in future use
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
                    isValid = true,
                    process_cost = model.process_cost
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

                    var action = "api/process/delete-process-by-id/" + Id;

                    HttpResponseMessage res = await client.PostAsync(action, content).ConfigureAwait(false);


                    res.EnsureSuccessStatusCode();
                    if (res.IsSuccessStatusCode)
                    {

                        var result = res.Content.ReadAsStringAsync().Result;

                        //string returnUrl = Url.Content("~/");

                        //return LocalRedirect(returnUrl);
                        

                    }


                }
            }

            //return View();
        }

    }
}