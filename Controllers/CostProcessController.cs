using System;
using System.Collections.Generic;
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
    public class CostProcessController : Controller
    {
        CostAPI _api = new CostAPI();

        

        public async Task<IActionResult> Index(string p_doc_no, int p_od, string p_process_type, float p_rubber_weight)
        {
            CostAPI _api = new CostAPI();

            ListModel m = new ListModel();

            ViewBag.Types = m.GetTypes().Select(x => new SelectListItem()
            {
                Text = x.processType,
                Value = x.processType
            }).ToList();


            ViewBag.p_doc_no = p_doc_no;
            ViewBag.p_od = p_od;
            ViewBag.p_process_cost = 0;
            ViewBag.p_process_name = "-";
            ViewBag.p_process_type = p_process_type;
            ViewBag.p_overhead_cost = 0;
            ViewBag.p_machine_cost = 0;
            ViewBag.p_total_labor_cost = 0;
            ViewBag.p_total_overhead_cost = 0;
            ViewBag.p_total_machine_cost = 0;
            ViewBag.p_labor_cost = 0;
            ViewBag.p_total_cost = 0;
            ViewBag.p_rubber_weight = p_rubber_weight;

            if (p_od >= 5 && p_od <= 50)
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
            if (p_od >= 151)
            {
                ViewBag.p_od_min = 151;
                ViewBag.p_od_max = 10000000;
            }


            ProcessMaster list = new ProcessMaster();
            List<ProcessMaster> data_ = new List<ProcessMaster>();

            CostProcess list2 = new CostProcess();
            List<CostProcess> data2_ = new List<CostProcess>();

            HttpClient clientdata = _api.Initial();


            var action = "api/processmaster/get-processmaster-by-odtype/" + p_od + "/" + p_process_type;
            if (p_od < 5)
                action = "api/processmaster/get-processmaster-by-type/" + p_process_type;

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


            var action2 = "api/costprocess/get-costprocess-by-docno/" + p_doc_no;
            HttpResponseMessage resdata2 = await clientdata.GetAsync(action2).ConfigureAwait(false);

            resdata2.EnsureSuccessStatusCode();

            if (resdata2.IsSuccessStatusCode)
            {
                var resultdata2 = resdata2.Content.ReadAsStringAsync().Result;
                data2_ = JsonConvert.DeserializeObject<List<CostProcess>>(resultdata2);
                foreach (var o in data2_)
                {
                    list2.data.Add(new CostProcess
                    {
                        process_name = o.process_name,
                        process_type = o.process_type,
                        item_od = o.item_od,
                        overhead_cost = double.Parse(o.overhead_cost.ToString("0.0000")),
                        machine_cost = double.Parse(o.machine_cost.ToString("0.0000")),
                        labor_cost = double.Parse(o.labor_cost.ToString("0.0000")),
                        total_cost = double.Parse(o.total_cost.ToString("0.0000")),
                        CostProcessId = o.CostProcessId
                    });
                }

            }
            List<CostProcess> model2 = list2.data.ToList();
            ViewData["processdata"] = model2;

            return View();
        }

        public async Task<IActionResult> IndexGetType(string p_doc_no, int p_od, string p_process_type, float p_rubber_weight)
        {
            CostAPI _api = new CostAPI();

            ListModel m = new ListModel();

            ViewBag.Types = m.GetTypes().Select(x => new SelectListItem()
            {
                Text = x.processType,
                Value = x.processType
            }).ToList();


            ViewBag.type = p_process_type;


            ViewBag.p_doc_no = p_doc_no;
            ViewBag.p_od = p_od;
            ViewBag.p_process_cost = 0;
            ViewBag.p_process_name = "-";
            ViewBag.p_process_type = "-";
            ViewBag.p_overhead_cost = 0;
            ViewBag.p_machine_cost = 0;
            ViewBag.p_labor_cost = 0;
            ViewBag.p_total_cost = 0;
            ViewBag.p_rubber_weight = p_rubber_weight;

            if (p_od >= 5 && p_od <= 50)
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
            if (p_od >= 151)
            {
                ViewBag.p_od_min = 151;
                ViewBag.p_od_max = 10000000;
            }


            ProcessMaster list = new ProcessMaster();
            List<ProcessMaster> data_ = new List<ProcessMaster>();

            CostProcess list2 = new CostProcess();
            List<CostProcess> data2_ = new List<CostProcess>();

            HttpClient clientdata = _api.Initial();

            //var action = "api/processmaster/get-processmaster-by-od-type/" + p_od;
            var action = "api/processmaster/get-processmaster-by-odtype/" + p_od+"/"+ p_process_type;
            if (p_od < 5)
                action = "api/processmaster/get-processmaster-by-type/" + p_process_type;


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


            var action2 = "api/costprocess/get-costprocess-by-docno/" + p_doc_no;
            HttpResponseMessage resdata2 = await clientdata.GetAsync(action2).ConfigureAwait(false);

            resdata2.EnsureSuccessStatusCode();

            if (resdata2.IsSuccessStatusCode)
            {
                var resultdata2 = resdata2.Content.ReadAsStringAsync().Result;
                data2_ = JsonConvert.DeserializeObject<List<CostProcess>>(resultdata2);
                foreach (var o in data2_)
                {
                    list2.data.Add(new CostProcess
                    {
                        process_name = o.process_name,
                        process_type = o.process_type,
                        item_od = o.item_od,
                        overhead_cost = double.Parse(o.overhead_cost.ToString("0.0000")),
                        machine_cost = double.Parse(o.machine_cost.ToString("0.0000")),
                        labor_cost = double.Parse(o.labor_cost.ToString("0.0000")),
                        total_cost = double.Parse(o.total_cost.ToString("0.0000")),
                        CostProcessId = o.CostProcessId
                    });
                }

            }
            List<CostProcess> model2 = list2.data.ToList();
            ViewData["processdata"] = model2;



            return View("Index");
        }


        public async Task<ActionResult<CostProcess>> Save(CostProcess model)
        {

            HttpClient client = _api.Initial();

            var content = new StringContent(JsonConvert.SerializeObject(model), Encoding.UTF8, "application/json");

            var action = "api/costprocess/add-costprocess";

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

        public async void Delete(CostProcess model,bool confirm,int Id)
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
                    var action = "api/costprocess/delete-costprocess-by-id/" + Id;
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