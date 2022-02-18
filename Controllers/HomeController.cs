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
using Process = CostNag.Models.Process;
using Microsoft.AspNetCore.Mvc.Rendering;
using ClosedXML.Excel;
using System.IO;

namespace CostNag.Controllers
{
    public class HomeController : Controller
    {


        CostAPI _api = new CostAPI();
         
        public async Task<IActionResult> Index()
        {

            Rubber listRubber = new Rubber();
            List<Rubber> rubber = new List<Rubber>();


            HttpClient r = _api.Initial();
            HttpResponseMessage resRubber = await r.GetAsync("api/rubber/get-all-rubbers");
            if (resRubber.IsSuccessStatusCode)
            {
                var result = resRubber.Content.ReadAsStringAsync().Result;
                rubber = JsonConvert.DeserializeObject<List<Rubber>>(result);
                //listRubber.data.Clear();

                //listRubber.data.Add(new Rubber
                //{
                //    material_name ="-",
                //    price_kg = 0,
                //    mixing_process_cost = 0,
                //    weight_g = 0,
                //    yield_rate = 0
                //});

                //foreach (var o in rubber)
                //{
                //    listRubber.data.Add(new Rubber
                //    {
                //        material_name = o.material_name,
                //        price_kg = o.price_kg,
                //        mixing_process_cost = o.mixing_process_cost,
                //        weight_g = o.weight_g,
                //        yield_rate = o.yield_rate
                //    });
                //}
            }

            //List<Rubber> rubberModel = listRubber.data.ToList();

            ViewBag.Rubber = rubber.Select(x => new SelectListItem()
            {
                Text = x.material_name,
                Value = x.material_name + "," + x.price_kg + "," + x.mixing_process_cost + "," + x.weight_g + "," + x.yield_rate

            }).ToList();



            Tooling listTooling = new Tooling();
            List<Tooling> tooling = new List<Tooling>();

            HttpClient t = _api.Initial();
            HttpResponseMessage resTooling = await t.GetAsync("api/tooling/get-all-toolings");
            if (resTooling.IsSuccessStatusCode)
            {
                var result = resTooling.Content.ReadAsStringAsync().Result;
                tooling = JsonConvert.DeserializeObject<List<Tooling>>(result);
            }
            ViewBag.Tooling = tooling.Select(x => new SelectListItem()
            {
                Text = x.description,
                Value = x.description + "," + x.source + "," + x.qty + "," + x.unit + "," + x.price

            }).ToList();


            ViewBag.process_type = "Process Type";

            ViewBag.plant = "-";
            ViewBag.item_spec = "-";
            ViewBag.issue_date = DateTime.Now.AddDays(0).ToString("dd-MM-yyyy");
            ViewBag.section = "-";
            ViewBag.doc_no = "-";
            ViewBag.wr_no = "-";
            ViewBag.sales = "-";
            ViewBag.revision_no = "-";
            ViewBag.checked_date = DateTime.Now.AddDays(0).ToString("dd-MM-yyyy");
            ViewBag.approved_by = DateTime.Now.AddDays(0).ToString("dd-MM-yyyy");
            ViewBag.expired_by = DateTime.Now.AddDays(180).ToString("dd-MM-yyyy");
            ViewBag.customer = "-";
            ViewBag.parts_code = "-";
            ViewBag.item = "-";
            ViewBag.product = "-";
            ViewBag.product_type = "-";
            ViewBag.size = "-";
            ViewBag.item_id = 0;
            ViewBag.item_od = 0;
            ViewBag.item_w = 0;
            ViewBag.item_w2 = 0;
            ViewBag.business_type = "-";
            ViewBag.qty_month = 0;

            ViewBag.exchange_rate_jpy = 84.13;
            ViewBag.exchange_rate_usd = 0.74;
            ViewBag.exchange_rate_eud = 0.65;

            ViewBag.target_price_sgd = 0.0;
            ViewBag.target_price_usd = 0.0;
            ViewBag.target_price_eud = 0.0;
            ViewBag.target_price_wr_sgd = 0.0;
            ViewBag.target_price_wr_usd = 0.0;
            ViewBag.target_price_wr_eud = 0.0;

            ViewBag.production_qty_day = 0;
            ViewBag.working_day = 21;

            ViewBag.rubber_weight_g_total = 0;

            ViewBag.rubber_material_name = "-";
            ViewBag.rubber_database_price_current = 0;
            ViewBag.rubber_database_price_new = 0;
            ViewBag.rubber_price_kg = 0;
            ViewBag.rubber_mixing_process_cost = 0;
            ViewBag.rubber_weight_g = 0;
            ViewBag.rubber_weight_kg = 0;
            ViewBag.rubber_yield_rate = 0;
            ViewBag.rubber_weight_kg_yieldrate = 0;
            ViewBag.rubber_cost_sgd = 0;
            ViewBag.rubber_percentage_target_price = 0;

            ViewBag.rubber_material_name2 = "-";
            ViewBag.rubber_database_price_current2 = 0;
            ViewBag.rubber_database_price_new2 = 0;
            ViewBag.rubber_price_kg2 = 0;
            ViewBag.rubber_mixing_process_cost2 = 0;
            ViewBag.rubber_weight_g2 = 0;
            ViewBag.rubber_weight_kg2 = 0;
            ViewBag.rubber_yield_rate2 = 0;
            ViewBag.rubber_weight_kg_yieldrate2 = 0;
            ViewBag.rubber_cost_sgd2 = 0;
            ViewBag.rubber_percentage_target_price2 = 0;

            ViewBag.material_inhouse_name_1 = "-";
            ViewBag.material_inhouse_info_1 = "-";
            ViewBag.material_inhouse_value_1 = 0;
            ViewBag.material_inhouse_info_1b = "gram/pcs       $/kg";
            ViewBag.material_inhouse_value_1b = 0;
            ViewBag.material_inhouse_cost_sgd_1 = 0;
            ViewBag.material_inhouse_percentage_target_price_1 = 0;

            ViewBag.material_inhouse_name_2 = "-";
            ViewBag.material_inhouse_info_2 = "-";
            ViewBag.material_inhouse_value_2 = 0;
            ViewBag.material_inhouse_info_2b = "gram/pcs       $/kg";
            ViewBag.material_inhouse_value_2b = 0;
            ViewBag.material_inhouse_cost_sgd_2 = 0;
            ViewBag.material_inhouse_percentage_target_price_2 = 0;

            ViewBag.material_inhouse_name_3 = "-";
            ViewBag.material_inhouse_info_3 = "-";
            ViewBag.material_inhouse_value_3 = 0;
            ViewBag.material_inhouse_info_3b = "gram/pcs       $/kg";
            ViewBag.material_inhouse_value_3b = 0;
            ViewBag.material_inhouse_cost_sgd_3 = 0;
            ViewBag.material_inhouse_percentage_target_price_3 = 0;

            ViewBag.material_outside_name_1 = "-";
            ViewBag.material_outside_info_1 = "-";
            ViewBag.material_outside_value_1 = 0;
            ViewBag.material_outside_info_1b = "gram/pcs       $/kg";
            ViewBag.material_outside_value_1b = 0;
            ViewBag.material_outside_cost_sgd_1 = 0;
            ViewBag.material_outside_percentage_target_price_1 = 0;

            ViewBag.material_outside_name_2 = "-";
            ViewBag.material_outside_info_2 = "-";
            ViewBag.material_outside_value_2 = 0;
            ViewBag.material_outside_info_2b = "gram/pcs       $/kg";
            ViewBag.material_outside_value_2b = 0;
            ViewBag.material_outside_cost_sgd_2 = 0;
            ViewBag.material_outside_percentage_target_price_2 = 0;

            ViewBag.material_outside_name_3 = "-";
            ViewBag.material_outside_info_3 = "-";
            ViewBag.material_outside_value_3 = 0;
            ViewBag.material_outside_info_3b = "gram/pcs       $/kg";
            ViewBag.material_outside_value_3b = 0;
            ViewBag.material_outside_cost_sgd_3 = 0;
            ViewBag.material_outside_percentage_target_price_3 = 0;

            ViewBag.direct_material_cost = 0;
            ViewBag.direct_material_cost_percentage = 0;
            ViewBag.sub_material_percentage = 10;
            ViewBag.sub_material_cost = 0;
            ViewBag.sub_material_cost_percentage = 0;
            ViewBag.direct_process_cost = 0;
            ViewBag.direct_process_cost_percentage = 0;
            ViewBag.total_direct_cost = 0;
            ViewBag.total_direct_cost_percentage = 0;
            ViewBag.defective_percentage = 3;
            ViewBag.defective_cost = 0;
            ViewBag.defective_cost_percentage = 0;
            ViewBag.indirect_percentage = 15;
            ViewBag.indirect_cost = 0;
            ViewBag.indirect_cost_percentage = 0;
            ViewBag.packing_material_percentage = 5;
            ViewBag.special_package_cost = 0;
            ViewBag.packing_material_cost = 0;
            ViewBag.packing_material_cost_percentage = 0;
            ViewBag.administration_percentage = 10;
            ViewBag.administration_cost = 0;
            ViewBag.administration_cost_percentage = 0;
            ViewBag.plant_retail_percentage = 5;
            ViewBag.plant_retail_cost = 0;
            ViewBag.plant_retail_cost_percentage = 0;
            ViewBag.moldjig_percentage = 0;
            ViewBag.moldjig_cost = 0;
            ViewBag.moldjig_cost_percentage = 0;
            ViewBag.die_percentage = 0;
            ViewBag.die_cost = 0;
            ViewBag.die_cost_percentage = 0;
            ViewBag.note = "-";
            ViewBag.net_included_tooling_cost = 0;
            ViewBag.net_included_tooling_cost_percentage = 0;
            ViewBag.net_exclude_tooling_cost = 0;
            ViewBag.net_exclude_tooling_cost_percentage = 0;

            ViewBag.tooling_list_description_1 = "-";
            ViewBag.tooling_list_type_1 = "-";
            ViewBag.tooling_list_source_1 = "-";
            ViewBag.tooling_list_qty_1 = 0;
            ViewBag.tooling_list_unit_1 = "-";
            ViewBag.tooling_list_price_1 = 0;
            ViewBag.tooling_list_amount_jpy_1 = 0;
            ViewBag.tooling_list_amount_sgd_1 = 0;

            ViewBag.tooling_list_description_2 = "-";
            ViewBag.tooling_list_type_2 = "-";
            ViewBag.tooling_list_source_2 = "-";
            ViewBag.tooling_list_qty_2 = 0;
            ViewBag.tooling_list_unit_2 = "-";
            ViewBag.tooling_list_price_2 = 0;
            ViewBag.tooling_list_amount_jpy_2 = 0;
            ViewBag.tooling_list_amount_sgd_2 = 0;

            ViewBag.tooling_list_description_3 = "-";
            ViewBag.tooling_list_type_3 = "-";
            ViewBag.tooling_list_source_3 = "-";
            ViewBag.tooling_list_qty_3 = 0;
            ViewBag.tooling_list_unit_3 = "-";
            ViewBag.tooling_list_price_3 = 0;
            ViewBag.tooling_list_amount_jpy_3 = 0;
            ViewBag.tooling_list_amount_sgd_3 = 0;

            ViewBag.tooling_list_description_4 = "-";
            ViewBag.tooling_list_type_4 = "-";
            ViewBag.tooling_list_source_4 = "-";
            ViewBag.tooling_list_qty_4 = 0;
            ViewBag.tooling_list_unit_4 = "-";
            ViewBag.tooling_list_price_4 = 0;
            ViewBag.tooling_list_amount_jpy_4 = 0;
            ViewBag.tooling_list_amount_sgd_4 = 0;

            ViewBag.tooling_list_description_5 = "-";
            ViewBag.tooling_list_type_5 = "-";
            ViewBag.tooling_list_source_5 = "-";
            ViewBag.tooling_list_qty_5 = 0;
            ViewBag.tooling_list_unit_5 = "-";
            ViewBag.tooling_list_price_5 = 0;
            ViewBag.tooling_list_amount_jpy_5 = 0;
            ViewBag.tooling_list_amount_sgd_5 = 0;

            ViewBag.tooling_list_description_6 = "-";
            ViewBag.tooling_list_type_6 = "-";
            ViewBag.tooling_list_source_6 = "-";
            ViewBag.tooling_list_qty_6 = 0;
            ViewBag.tooling_list_unit_6 = "-";
            ViewBag.tooling_list_price_6 = 0;
            ViewBag.tooling_list_amount_jpy_6 = 0;
            ViewBag.tooling_list_amount_sgd_6 = 0;

            ViewBag.tooling_list_description_7 = "-";
            ViewBag.tooling_list_type_7 = "-";
            ViewBag.tooling_list_source_7 = "-";
            ViewBag.tooling_list_qty_7 = 0;
            ViewBag.tooling_list_unit_7 = "-";
            ViewBag.tooling_list_price_7 = 0;
            ViewBag.tooling_list_amount_jpy_7 = 0;
            ViewBag.tooling_list_amount_sgd_7 = 0;

            ViewBag.tooling_list_description_8 = "-";
            ViewBag.tooling_list_type_8 = "-";
            ViewBag.tooling_list_source_8 = "-";
            ViewBag.tooling_list_qty_8 = 0;
            ViewBag.tooling_list_unit_8 = "-";
            ViewBag.tooling_list_price_8 = 0;
            ViewBag.tooling_list_amount_jpy_8 = 0;
            ViewBag.tooling_list_amount_sgd_8 = 0;

            ViewBag.tooling_list_description_9 = "-";
            ViewBag.tooling_list_type_9 = "-";
            ViewBag.tooling_list_source_9 = "-";
            ViewBag.tooling_list_qty_9 = 0;
            ViewBag.tooling_list_unit_9 = "-";
            ViewBag.tooling_list_price_9 = 0;
            ViewBag.tooling_list_amount_jpy_9 = 0;
            ViewBag.tooling_list_amount_sgd_9 = 0;

            ViewBag.tooling_list_description_10 = "-";
            ViewBag.tooling_list_type_10 = "-";
            ViewBag.tooling_list_source_10 = "-";
            ViewBag.tooling_list_qty_10 = 0;
            ViewBag.tooling_list_unit_10 = "-";
            ViewBag.tooling_list_price_10 = 0;
            ViewBag.tooling_list_amount_jpy_10 = 0;
            ViewBag.tooling_list_amount_sgd_10 = 0;

            ViewBag.tooling_list_description_11 = "-";
            ViewBag.tooling_list_type_11 = "-";
            ViewBag.tooling_list_source_11 = "-";
            ViewBag.tooling_list_qty_11 = 0;
            ViewBag.tooling_list_unit_11 = "-";
            ViewBag.tooling_list_price_11 = 0;
            ViewBag.tooling_list_amount_jpy_11 = 0;
            ViewBag.tooling_list_amount_sgd_11 = 0;

            ViewBag.tooling_list_description_12 = "-";
            ViewBag.tooling_list_type_12 = "-";
            ViewBag.tooling_list_source_12 = "-";
            ViewBag.tooling_list_qty_12 = 0;
            ViewBag.tooling_list_unit_12 = "-";
            ViewBag.tooling_list_price_12 = 0;
            ViewBag.tooling_list_amount_jpy_12 = 0;
            ViewBag.tooling_list_amount_sgd_12 = 0;

            ViewBag.tooling_list_description_13 = "-";
            ViewBag.tooling_list_type_13 = "-";
            ViewBag.tooling_list_source_13 = "-";
            ViewBag.tooling_list_qty_13 = 0;
            ViewBag.tooling_list_unit_13 = "-";
            ViewBag.tooling_list_price_13 = 0;
            ViewBag.tooling_list_amount_jpy_13 = 0;
            ViewBag.tooling_list_amount_sgd_13 = 0;

            ViewBag.tooling_list_description_14 = "-";
            ViewBag.tooling_list_type_14 = "-";
            ViewBag.tooling_list_source_14 = "-";
            ViewBag.tooling_list_qty_14 = 0;
            ViewBag.tooling_list_unit_14 = "-";
            ViewBag.tooling_list_price_14 = 0;
            ViewBag.tooling_list_amount_jpy_14 = 0;
            ViewBag.tooling_list_amount_sgd_14 = 0;

            ViewBag.tooling_list_description_15 = "-";
            ViewBag.tooling_list_type_15 = "-";
            ViewBag.tooling_list_source_15 = "-";
            ViewBag.tooling_list_qty_15 = 0;
            ViewBag.tooling_list_unit_15 = "-";
            ViewBag.tooling_list_price_15 = 0;
            ViewBag.tooling_list_amount_jpy_15 = 0;
            ViewBag.tooling_list_amount_sgd_15 = 0;

            ViewBag.tooling_list_total_amount_sgd = 0;

            ViewBag.direct_raw_material= 0;
            ViewBag.direct_raw_material_p= 0;
            ViewBag.sub_material= 0;
            ViewBag.sub_material_p= 0;
            ViewBag.raw_material_cost_sub_total= 0;
            ViewBag.raw_material_cost_sub_total_p= 0;
            ViewBag.labor_cost= 0;
            ViewBag.labor_cost_p= 0;
            ViewBag.machine_cost= 0;
            ViewBag.machine_cost_p= 0;
            ViewBag.overhead_cost= 0;
            ViewBag.overhead_cost_p= 0;
            ViewBag.process_cost_sub_total= 0;
            ViewBag.process_cost_sub_total_p= 0;
            ViewBag.defectives= 0;
            ViewBag.defectives_p= 0;
            ViewBag.admin_engin_qc= 0;
            ViewBag.admin_engin_qc_p= 0;
            ViewBag.tooling_cost= 0;
            ViewBag.tooling_cost_p= 0;
            ViewBag.process_margin_adjust= 0;
            ViewBag.process_margin_adjust_p= 0;
            ViewBag.other_fixed_cost_sub_total= 0;
            ViewBag.other_fixed_cost_sub_total_p= 0;
            ViewBag.grand_total_cost= 0;
            ViewBag.grand_total_cost_p= 0;

            ViewBag.production_capacity = 0;
            ViewBag.actual_working_time = 0;
            ViewBag.cycle_time = 0;
            ViewBag.efficiency = 0;
            ViewBag.daily_qty_days = 0;
            ViewBag.daily_qty_days_p = 0;
            ViewBag.daily_amount = 0;


            ListModel list = new ListModel();
            ViewData["plant"] = list.plant;
            ViewData["item_spec"] = list.item_spec;

            List<Cost> cost = new List<Cost>();
            HttpClient client = _api.Initial();
            try
            {
                HttpResponseMessage res = await client.GetAsync("api/data/get-all-costs");
                if (res.IsSuccessStatusCode)
                {
                    var result = res.Content.ReadAsStringAsync().Result;
                    cost = JsonConvert.DeserializeObject<List<Cost>>(result);
                    foreach (var o in cost)
                    {
                        list.partscode.Add(new ListModel
                        {
                            doc_no = o.doc_no,
                            code = o.parts_code,
                            id = o.CostId.ToString()

                        });
                    }

                }
                
            }
            catch(Exception e) { }

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
                        doc_no = o.doc_no,
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

            var action = "api/data/get-cost-by-id/"+id;
            HttpResponseMessage resdata = await clientdata.GetAsync(action).ConfigureAwait(false);
           // HttpResponseMessage resdata = await clientdata.GetAsync("api/data/get-all-costs");

            resdata.EnsureSuccessStatusCode();

            if (resdata.IsSuccessStatusCode)
            {
                var resultdata = resdata.Content.ReadAsStringAsync().Result;
                costdata = JsonConvert.DeserializeObject<Cost>(resultdata);


                Rubber listRubber = new Rubber();
                List<Rubber> rubber = new List<Rubber>();

                HttpClient r = _api.Initial();
                HttpResponseMessage resRubber = await r.GetAsync("api/rubber/get-all-rubbers");
                if (resRubber.IsSuccessStatusCode)
                {
                    var result = resRubber.Content.ReadAsStringAsync().Result;
                    rubber = JsonConvert.DeserializeObject<List<Rubber>>(result);
                }

                //List<Rubber> rubberModel = listRubber.data.ToList();

                ViewBag.Rubber = rubber.Select(x => new SelectListItem()
                {
                    Text = x.material_name,
                    Value = x.material_name+","+x.price_kg + "," +x.mixing_process_cost + "," +x.weight_g + "," +x.yield_rate

                }).ToList();



                Tooling listTooling = new Tooling();
                List<Tooling> tooling = new List<Tooling>();

                HttpClient t = _api.Initial();
                HttpResponseMessage resTooling = await t.GetAsync("api/tooling/get-all-toolings");
                if (resTooling.IsSuccessStatusCode)
                {
                    var result = resTooling.Content.ReadAsStringAsync().Result;
                    tooling = JsonConvert.DeserializeObject<List<Tooling>>(result);
                }
                ViewBag.Tooling = tooling.Select(x => new SelectListItem()
                {
                    Text = x.description,
                    Value = x.description + "," + x.source + "," + x.qty + "," + x.unit + "," + x.price

                }).ToList();


                ViewBag.plant = costdata.plant;
                ViewBag.item_spec = costdata.item_spec;
                ViewBag.issue_date = costdata.issue_date;
                ViewBag.section = costdata.section;
                ViewBag.doc_no = costdata.doc_no;
                ViewBag.wr_no = costdata.wr_no;
                ViewBag.sales = costdata.sales;
                ViewBag.revision_no = costdata.revision_no;
                ViewBag.checked_date = costdata.checked_date;
                ViewBag.approved_by = costdata.approved_by;
                ViewBag.expired_by = costdata.expired_by;
                ViewBag.customer = costdata.customer;
                ViewBag.parts_code = costdata.parts_code;
                ViewBag.item = costdata.item;
                ViewBag.product =costdata.product;
                ViewBag.product_type = costdata.product_type;
                ViewBag.size = costdata.size;
                ViewBag.item_id = costdata.item_id;
                ViewBag.item_od = costdata.item_od;
                ViewBag.item_w = costdata.item_w;
                ViewBag.item_w2 = costdata.item_w2;
                ViewBag.business_type = costdata.business_type;
                ViewBag.qty_month = costdata.qty_month;

                ViewBag.exchange_rate_jpy = costdata.exchange_rate_jpy;
                ViewBag.exchange_rate_usd = costdata.exchange_rate_usd;
                ViewBag.exchange_rate_eud = costdata.exchange_rate_eud;


                ViewBag.target_price_sgd = costdata.target_price_sgd;
                ViewBag.target_price_usd = costdata.target_price_usd;
                ViewBag.target_price_eud = costdata.target_price_eud;

                ViewBag.target_price_wr_sgd = costdata.target_price_wr_sgd;
                ViewBag.target_price_wr_usd = costdata.target_price_wr_usd;
                ViewBag.target_price_wr_eud = costdata.target_price_wr_eud;

                ViewBag.production_qty_day = costdata.production_qty_day;
                ViewBag.working_day = costdata.working_day;

                var total_rubber_weight = float.Parse(costdata.rubber_weight_g.ToString()) + float.Parse(costdata.rubber_weight_g2.ToString());
                ViewBag.rubber_weight_g_total = total_rubber_weight;

                ViewBag.rubber_material_name = costdata.rubber_material_name;
                ViewBag.rubber_database_price_current = costdata.rubber_database_price_current;
                ViewBag.rubber_database_price_new = costdata.rubber_database_price_new;
                ViewBag.rubber_price_kg = costdata.rubber_price_kg;
                ViewBag.rubber_mixing_process_cost = costdata.rubber_mixing_process_cost;
                ViewBag.rubber_weight_g = costdata.rubber_weight_g;
                ViewBag.rubber_weight_kg = costdata.rubber_weight_kg;
                ViewBag.rubber_yield_rate = costdata.rubber_yield_rate;
                ViewBag.rubber_weight_kg_yieldrate = costdata.rubber_weight_kg_yieldrate;
                ViewBag.rubber_cost_sgd = costdata.rubber_cost_sgd;
                ViewBag.rubber_percentage_target_price = costdata.rubber_percentage_target_price;

                ViewBag.rubber_material_name2 = costdata.rubber_material_name2;
                ViewBag.rubber_database_price_current2 = costdata.rubber_database_price_current2;
                ViewBag.rubber_database_price_new2 = costdata.rubber_database_price_new2;
                ViewBag.rubber_price_kg2 = costdata.rubber_price_kg2;
                ViewBag.rubber_mixing_process_cost2 = costdata.rubber_mixing_process_cost2;
                ViewBag.rubber_weight_g2 = costdata.rubber_weight_g2;
                ViewBag.rubber_weight_kg2 = costdata.rubber_weight_kg2;
                ViewBag.rubber_yield_rate2 = costdata.rubber_yield_rate2;
                ViewBag.rubber_weight_kg_yieldrate2 = costdata.rubber_weight_kg_yieldrate2;
                ViewBag.rubber_cost_sgd2 = costdata.rubber_cost_sgd2;
                ViewBag.rubber_percentage_target_price2 = costdata.rubber_percentage_target_price2;

                ViewBag.material_inhouse_name_1 = costdata.material_inhouse_name_1;
                ViewBag.material_inhouse_info_1 = costdata.material_inhouse_info_1;
                ViewBag.material_inhouse_value_1 = costdata.material_inhouse_value_1;
                ViewBag.material_inhouse_info_1b = costdata.material_inhouse_info_1b;
                ViewBag.material_inhouse_value_1b = costdata.material_inhouse_value_1b;
                ViewBag.material_inhouse_cost_sgd_1 = costdata.material_inhouse_cost_sgd_1;
                ViewBag.material_inhouse_percentage_target_price_1 = costdata.material_inhouse_percentage_target_price_1;

                ViewBag.material_inhouse_name_2 = costdata.material_inhouse_name_2;
                ViewBag.material_inhouse_info_2 = costdata.material_inhouse_info_2;
                ViewBag.material_inhouse_value_2 = costdata.material_inhouse_value_2;
                ViewBag.material_inhouse_info_2b = costdata.material_inhouse_info_2b;
                ViewBag.material_inhouse_value_2b = costdata.material_inhouse_value_2b;
                ViewBag.material_inhouse_cost_sgd_2 = costdata.material_inhouse_cost_sgd_2;
                ViewBag.material_inhouse_percentage_target_price_2 = costdata.material_inhouse_percentage_target_price_2;

                ViewBag.material_inhouse_name_3 = costdata.material_inhouse_name_3;
                ViewBag.material_inhouse_info_3 = costdata.material_inhouse_info_3;
                ViewBag.material_inhouse_value_3 = costdata.material_inhouse_value_3;
                ViewBag.material_inhouse_info_3b = costdata.material_inhouse_info_3b;
                ViewBag.material_inhouse_value_3b = costdata.material_inhouse_value_3b;
                ViewBag.material_inhouse_cost_sgd_3 = costdata.material_inhouse_cost_sgd_3;
                ViewBag.material_inhouse_percentage_target_price_3 = costdata.material_inhouse_percentage_target_price_3;

                ViewBag.material_outside_name_1 = costdata.material_outside_name_1;
                ViewBag.material_outside_info_1 = costdata.material_outside_info_1;
                ViewBag.material_outside_value_1 = costdata.material_outside_value_1;
                ViewBag.material_outside_info_1b = costdata.material_outside_info_1b;
                ViewBag.material_outside_value_1b = costdata.material_outside_value_1b;
                ViewBag.material_outside_cost_sgd_1 = costdata.material_outside_cost_sgd_1;
                ViewBag.material_outside_percentage_target_price_1 = costdata.material_outside_percentage_target_price_1;

                ViewBag.material_outside_name_2 = costdata.material_outside_name_2;
                ViewBag.material_outside_info_2 = costdata.material_outside_info_2;
                ViewBag.material_outside_value_2 = costdata.material_outside_value_2;
                ViewBag.material_outside_info_2b = costdata.material_outside_info_2b;
                ViewBag.material_outside_value_2b = costdata.material_outside_value_2b;
                ViewBag.material_outside_cost_sgd_2 = costdata.material_outside_cost_sgd_2;
                ViewBag.material_outside_percentage_target_price_2 = costdata.material_outside_percentage_target_price_2;

                ViewBag.material_outside_name_3 = costdata.material_outside_name_3;
                ViewBag.material_outside_info_3 = costdata.material_outside_info_3;
                ViewBag.material_outside_value_3 = costdata.material_outside_value_3;
                ViewBag.material_outside_info_3b = costdata.material_outside_info_3b;
                ViewBag.material_outside_value_3b = costdata.material_outside_value_3b;
                ViewBag.material_outside_cost_sgd_3 = costdata.material_outside_cost_sgd_3;
                ViewBag.material_outside_percentage_target_price_3 = costdata.material_outside_percentage_target_price_3;

                ViewBag.direct_material_cost = costdata.direct_material_cost;
                ViewBag.direct_material_cost_percentage = costdata.direct_material_cost_percentage;
                ViewBag.sub_material_percentage = costdata.sub_material_percentage;
                ViewBag.sub_material_cost = costdata.sub_material_cost;
                ViewBag.sub_material_cost_percentage = costdata.sub_material_cost_percentage;
                ViewBag.direct_process_cost = costdata.direct_process_cost;
                ViewBag.direct_process_cost_percentage = costdata.direct_process_cost_percentage;
                ViewBag.total_direct_cost = costdata.total_direct_cost;
                ViewBag.total_direct_cost_percentage = costdata.total_direct_cost_percentage;
                ViewBag.defective_percentage = costdata.defective_percentage;
                ViewBag.defective_cost = costdata.defective_cost;
                ViewBag.defective_cost_percentage = costdata.defective_cost_percentage;
                ViewBag.indirect_percentage = costdata.indirect_percentage;
                ViewBag.indirect_cost = costdata.indirect_cost;
                ViewBag.indirect_cost_percentage = costdata.indirect_cost_percentage;
                ViewBag.packing_material_percentage = costdata.packing_material_percentage;
                ViewBag.special_package_cost = costdata.special_package_cost;
                ViewBag.packing_material_cost = costdata.packing_material_cost;
                ViewBag.packing_material_cost_percentage = costdata.packing_material_cost_percentage;
                ViewBag.administration_percentage = costdata.administration_percentage;
                ViewBag.administration_cost = costdata.administration_cost;
                ViewBag.administration_cost_percentage = costdata.administration_cost_percentage;
                ViewBag.plant_retail_percentage = costdata.plant_retail_percentage;
                ViewBag.plant_retail_cost = costdata.plant_retail_cost;
                ViewBag.plant_retail_cost_percentage = costdata.plant_retail_cost_percentage;
                ViewBag.moldjig_percentage = costdata.moldjig_percentage;
                ViewBag.moldjig_cost = costdata.moldjig_cost;
                ViewBag.moldjig_cost_percentage = costdata.moldjig_cost_percentage;
                ViewBag.die_percentage = costdata.die_percentage;
                ViewBag.die_cost = costdata.die_cost;
                ViewBag.die_cost_percentage = costdata.die_cost_percentage;
                ViewBag.note = costdata.note;
                ViewBag.net_included_tooling_cost = costdata.net_included_tooling_cost;
                ViewBag.net_included_tooling_cost_percentage = costdata.net_included_tooling_cost_percentage;
                ViewBag.net_exclude_tooling_cost = costdata.net_exclude_tooling_cost;
                ViewBag.net_exclude_tooling_cost_percentage = costdata.net_exclude_tooling_cost_percentage;

                ViewBag.tooling_list_description_1 = costdata.tooling_list_description_1;
                ViewBag.tooling_list_type_1 = costdata.tooling_list_type_1;
                ViewBag.tooling_list_source_1 = costdata.tooling_list_source_1;
                ViewBag.tooling_list_qty_1 = costdata.tooling_list_qty_1;
                ViewBag.tooling_list_unit_1 = costdata.tooling_list_unit_1;
                ViewBag.tooling_list_price_1 = costdata.tooling_list_price_1;
                ViewBag.tooling_list_amount_jpy_1 = costdata.tooling_list_amount_jpy_1.ToString("#,##0");
                ViewBag.tooling_list_amount_usd_1 = costdata.tooling_list_amount_usd_1.ToString("#,##0"); 

                ViewBag.tooling_list_description_2 = costdata.tooling_list_description_2;
                ViewBag.tooling_list_type_2 = costdata.tooling_list_type_2;
                ViewBag.tooling_list_source_2 = costdata.tooling_list_source_2;
                ViewBag.tooling_list_qty_2 = costdata.tooling_list_qty_2;
                ViewBag.tooling_list_unit_2 = costdata.tooling_list_unit_2;
                ViewBag.tooling_list_price_2 = costdata.tooling_list_price_2;
                ViewBag.tooling_list_amount_jpy_2 = costdata.tooling_list_amount_jpy_2.ToString("#,##0");
                ViewBag.tooling_list_amount_usd_2 = costdata.tooling_list_amount_usd_2.ToString("#,##0");

                ViewBag.tooling_list_description_3 = costdata.tooling_list_description_3;
                ViewBag.tooling_list_type_3 = costdata.tooling_list_type_3;
                ViewBag.tooling_list_source_3 = costdata.tooling_list_source_3;
                ViewBag.tooling_list_qty_3 = costdata.tooling_list_qty_3;
                ViewBag.tooling_list_unit_3 = costdata.tooling_list_unit_3;
                ViewBag.tooling_list_price_3 = costdata.tooling_list_price_3;
                ViewBag.tooling_list_amount_jpy_3 = costdata.tooling_list_amount_jpy_3.ToString("#,##0");
                ViewBag.tooling_list_amount_usd_3 = costdata.tooling_list_amount_usd_3.ToString("#,##0");

                ViewBag.tooling_list_description_4 = costdata.tooling_list_description_4;
                ViewBag.tooling_list_type_4 = costdata.tooling_list_type_4;
                ViewBag.tooling_list_source_4 = costdata.tooling_list_source_4;
                ViewBag.tooling_list_qty_4 = costdata.tooling_list_qty_4;
                ViewBag.tooling_list_unit_4 = costdata.tooling_list_unit_4;
                ViewBag.tooling_list_price_4 = costdata.tooling_list_price_4;
                ViewBag.tooling_list_amount_jpy_4 = costdata.tooling_list_amount_jpy_4.ToString("#,##0");
                ViewBag.tooling_list_amount_usd_4 = costdata.tooling_list_amount_usd_4.ToString("#,##0");

                ViewBag.tooling_list_description_5 = costdata.tooling_list_description_5;
                ViewBag.tooling_list_type_5 = costdata.tooling_list_type_5;
                ViewBag.tooling_list_source_5 = costdata.tooling_list_source_5;
                ViewBag.tooling_list_qty_5 = costdata.tooling_list_qty_5;
                ViewBag.tooling_list_unit_5 = costdata.tooling_list_unit_5;
                ViewBag.tooling_list_price_5 = costdata.tooling_list_price_5;
                ViewBag.tooling_list_amount_jpy_5 = costdata.tooling_list_amount_jpy_5.ToString("#,##0");
                ViewBag.tooling_list_amount_usd_5 = costdata.tooling_list_amount_usd_5.ToString("#,##0");

                ViewBag.tooling_list_description_6 = costdata.tooling_list_description_6;
                ViewBag.tooling_list_type_6 = costdata.tooling_list_type_6;
                ViewBag.tooling_list_source_6 = costdata.tooling_list_source_6;
                ViewBag.tooling_list_qty_6 = costdata.tooling_list_qty_6;
                ViewBag.tooling_list_unit_6 = costdata.tooling_list_unit_6;
                ViewBag.tooling_list_price_6 = costdata.tooling_list_price_6;
                ViewBag.tooling_list_amount_jpy_6 = costdata.tooling_list_amount_jpy_6.ToString("#,##0");
                ViewBag.tooling_list_amount_usd_6 = costdata.tooling_list_amount_usd_6.ToString("#,##0");

                ViewBag.tooling_list_description_7 = costdata.tooling_list_description_7;
                ViewBag.tooling_list_type_7 = costdata.tooling_list_type_7;
                ViewBag.tooling_list_source_7 = costdata.tooling_list_source_7;
                ViewBag.tooling_list_qty_7 = costdata.tooling_list_qty_7;
                ViewBag.tooling_list_unit_7 = costdata.tooling_list_unit_7;
                ViewBag.tooling_list_price_7 = costdata.tooling_list_price_7;
                ViewBag.tooling_list_amount_jpy_7 = costdata.tooling_list_amount_jpy_7.ToString("#,##0");
                ViewBag.tooling_list_amount_usd_7 = costdata.tooling_list_amount_usd_7.ToString("#,##0");

                ViewBag.tooling_list_description_8 = costdata.tooling_list_description_8;
                ViewBag.tooling_list_type_8 = costdata.tooling_list_type_8;
                ViewBag.tooling_list_source_8 = costdata.tooling_list_source_8;
                ViewBag.tooling_list_qty_8 = costdata.tooling_list_qty_8;
                ViewBag.tooling_list_unit_8 = costdata.tooling_list_unit_8;
                ViewBag.tooling_list_price_8 = costdata.tooling_list_price_8;
                ViewBag.tooling_list_amount_jpy_8 = costdata.tooling_list_amount_jpy_8.ToString("#,##0");
                ViewBag.tooling_list_amount_usd_8 = costdata.tooling_list_amount_usd_8.ToString("#,##0");

                ViewBag.tooling_list_description_9 = costdata.tooling_list_description_9;
                ViewBag.tooling_list_type_9 = costdata.tooling_list_type_9;
                ViewBag.tooling_list_source_9 = costdata.tooling_list_source_9;
                ViewBag.tooling_list_qty_9 = costdata.tooling_list_qty_9;
                ViewBag.tooling_list_unit_9 = costdata.tooling_list_unit_9;
                ViewBag.tooling_list_price_9 = costdata.tooling_list_price_9;
                ViewBag.tooling_list_amount_jpy_9 = costdata.tooling_list_amount_jpy_9.ToString("#,##0");
                ViewBag.tooling_list_amount_usd_9 = costdata.tooling_list_amount_usd_9.ToString("#,##0");

                ViewBag.tooling_list_description_10 = costdata.tooling_list_description_10;
                ViewBag.tooling_list_type_10 = costdata.tooling_list_type_10;
                ViewBag.tooling_list_source_10 = costdata.tooling_list_source_10;
                ViewBag.tooling_list_qty_10 = costdata.tooling_list_qty_10;
                ViewBag.tooling_list_unit_10 = costdata.tooling_list_unit_10;
                ViewBag.tooling_list_price_10 = costdata.tooling_list_price_10;
                ViewBag.tooling_list_amount_jpy_10 = costdata.tooling_list_amount_jpy_10.ToString("#,##0");
                ViewBag.tooling_list_amount_usd_10 = costdata.tooling_list_amount_usd_10.ToString("#,##0");

                ViewBag.tooling_list_description_11 = costdata.tooling_list_description_11;
                ViewBag.tooling_list_type_11 = costdata.tooling_list_type_11;
                ViewBag.tooling_list_source_11 = costdata.tooling_list_source_11;
                ViewBag.tooling_list_qty_11 = costdata.tooling_list_qty_11;
                ViewBag.tooling_list_unit_11 = costdata.tooling_list_unit_11;
                ViewBag.tooling_list_price_11 = costdata.tooling_list_price_11;
                ViewBag.tooling_list_amount_jpy_11 = costdata.tooling_list_amount_jpy_11.ToString("#,##0");
                ViewBag.tooling_list_amount_usd_11 = costdata.tooling_list_amount_usd_11.ToString("#,##0");

                ViewBag.tooling_list_description_12 = costdata.tooling_list_description_12;
                ViewBag.tooling_list_type_12 = costdata.tooling_list_type_12;
                ViewBag.tooling_list_source_12 = costdata.tooling_list_source_12;
                ViewBag.tooling_list_qty_12 = costdata.tooling_list_qty_12;
                ViewBag.tooling_list_unit_12 = costdata.tooling_list_unit_12;
                ViewBag.tooling_list_price_12 = costdata.tooling_list_price_12;
                ViewBag.tooling_list_amount_jpy_12 = costdata.tooling_list_amount_jpy_12.ToString("#,##0");
                ViewBag.tooling_list_amount_usd_12 = costdata.tooling_list_amount_usd_12.ToString("#,##0");

                ViewBag.tooling_list_description_13 = costdata.tooling_list_description_13;
                ViewBag.tooling_list_type_13 = costdata.tooling_list_type_13;
                ViewBag.tooling_list_source_13 = costdata.tooling_list_source_13;
                ViewBag.tooling_list_qty_13 = costdata.tooling_list_qty_13;
                ViewBag.tooling_list_unit_13 = costdata.tooling_list_unit_13;
                ViewBag.tooling_list_price_13 = costdata.tooling_list_price_13;
                ViewBag.tooling_list_amount_jpy_13 = costdata.tooling_list_amount_jpy_13.ToString("#,##0");
                ViewBag.tooling_list_amount_usd_13 = costdata.tooling_list_amount_usd_13.ToString("#,##0");

                ViewBag.tooling_list_description_14 = costdata.tooling_list_description_14;
                ViewBag.tooling_list_type_14 = costdata.tooling_list_type_14;
                ViewBag.tooling_list_source_14 = costdata.tooling_list_source_14;
                ViewBag.tooling_list_qty_14 = costdata.tooling_list_qty_14;
                ViewBag.tooling_list_unit_14 = costdata.tooling_list_unit_14;
                ViewBag.tooling_list_price_14 = costdata.tooling_list_price_14;
                ViewBag.tooling_list_amount_jpy_14 = costdata.tooling_list_amount_jpy_14.ToString("#,##0");
                ViewBag.tooling_list_amount_usd_14 = costdata.tooling_list_amount_usd_14.ToString("#,##0");

                ViewBag.tooling_list_description_15 = costdata.tooling_list_description_15;
                ViewBag.tooling_list_type_15 = costdata.tooling_list_type_15;
                ViewBag.tooling_list_source_15 = costdata.tooling_list_source_15;
                ViewBag.tooling_list_qty_15 = costdata.tooling_list_qty_15;
                ViewBag.tooling_list_unit_15 = costdata.tooling_list_unit_15;
                ViewBag.tooling_list_price_15 = costdata.tooling_list_price_15;
                ViewBag.tooling_list_amount_jpy_15 = costdata.tooling_list_amount_jpy_15.ToString("#,##0");
                ViewBag.tooling_list_amount_usd_15 = costdata.tooling_list_amount_usd_15.ToString("#,##0");

                ViewBag.tooling_list_total_amount_usd = costdata.tooling_list_total_amount_usd.ToString("#,##0");


                ViewBag.direct_raw_material = costdata.direct_raw_material;
                ViewBag.direct_raw_material_p = costdata.direct_raw_material_p;
                ViewBag.sub_material = costdata.sub_material;
                ViewBag.sub_material_p = costdata.sub_material_p;
                ViewBag.raw_material_cost_sub_total = costdata.raw_material_cost_sub_total;
                ViewBag.raw_material_cost_sub_total_p = costdata.raw_material_cost_sub_total_p;
                ViewBag.labor_cost = costdata.labor_cost;
                ViewBag.labor_cost_p = costdata.labor_cost_p;
                ViewBag.machine_cost = costdata.machine_cost;
                ViewBag.machine_cost_p = costdata.machine_cost_p;
                ViewBag.overhead_cost = costdata.overhead_cost;
                ViewBag.overhead_cost_p = costdata.overhead_cost_p;
                ViewBag.process_cost_sub_total = costdata.process_cost_sub_total;
                ViewBag.process_cost_sub_total_p = costdata.process_cost_sub_total_p;
                ViewBag.defectives = costdata.defectives;
                ViewBag.defectives_p = costdata.defectives_p;
                ViewBag.admin_engin_qc = costdata.admin_engin_qc;
                ViewBag.admin_engin_qc_p = costdata.admin_engin_qc_p;
                ViewBag.tooling_cost = costdata.tooling_cost;
                ViewBag.tooling_cost_p = costdata.tooling_cost_p;
                ViewBag.process_margin_adjust = costdata.process_margin_adjust;
                ViewBag.process_margin_adjust_p = costdata.process_margin_adjust_p;
                ViewBag.other_fixed_cost_sub_total = costdata.other_fixed_cost_sub_total;
                ViewBag.other_fixed_cost_sub_total_p = costdata.other_fixed_cost_sub_total_p;
                ViewBag.grand_total_cost = costdata.grand_total_cost;
                ViewBag.grand_total_cost_p = costdata.grand_total_cost_p;

                ViewBag.production_capacity = costdata.production_capacity;
                ViewBag.actual_working_time = costdata.actual_working_time;
                ViewBag.cycle_time = costdata.cycle_time;
                ViewBag.efficiency = costdata.efficiency;
                ViewBag.daily_qty_days = costdata.daily_qty_days;
                ViewBag.daily_qty_days_p = costdata.daily_qty_days_p;
                ViewBag.daily_amount = costdata.daily_amount;

            }

            dynamic mymodel = new ExpandoObject();
           // mymodel.code = costdata;
            return View("Index",mymodel); 
        }


        public IActionResult ExcelData(
            string customer,
            string parts_code,
            string item,
            string product,
            string product_type,
            string item_id,
            string item_od,
            string item_w,
            string item_w2,
            string business_type,
            string qty_month,
            string target_price_sgd,
            string target_price_usd,
            string target_price_eud,
            string exchange_rate_jpy,
            string exchange_rate_usd,
            string exchange_rate_eud,
            string production_qty_day,
            string working_day,
            string issue_date,
            string section,
            string doc_no,
            string wr_no,
            string sales,
            string revision_no,
            string checked_date,
            string approved_by,
            string expired_by,
            string rubber_material_name,
            string rubber_price_kg,
            string rubber_mixing_process_cost,
            string rubber_weight_g,
            string rubber_weight_kg,
            string rubber_yield_rate,
            string rubber_weight_kg_yieldrate,
            string rubber_cost_sgd,
            string rubber_percentage_target_price,
            string rubber_material_name2,
            string rubber_price_kg2,
            string rubber_mixing_process_cost2,
            string rubber_weight_g2,
            string rubber_weight_kg2,
            string rubber_yield_rate2,
            string rubber_weight_kg_yieldrate2,
            string rubber_cost_sgd2,
            string rubber_percentage_target_price2,
            string material_inhouse_name_1,
            string material_inhouse_info_1,
            string material_inhouse_value_1,
            string material_inhouse_info_1b,
            string material_inhouse_value_1b,
            string material_inhouse_cost_sgd_1,
            string material_inhouse_percentage_target_price_1,
            string material_inhouse_name_2,
            string material_inhouse_info_2,
            string material_inhouse_value_2,
            string material_inhouse_info_2b,
            string material_inhouse_value_2b,
            string material_inhouse_cost_sgd_2,
            string material_inhouse_percentage_target_price_2,
            string material_inhouse_name_3,
            string material_inhouse_info_3,
            string material_inhouse_value_3,
            string material_inhouse_info_3b,
            string material_inhouse_value_3b,
            string material_inhouse_cost_sgd_3,
            string material_inhouse_percentage_target_price_3,
            string material_outside_name_1,
            string material_outside_info_1,
            string material_outside_value_1,
            string material_outside_info_1b,
            string material_outside_value_1b,
            string material_outside_cost_sgd_1,
            string material_outside_percentage_target_price_1,
            string material_outside_name_2,
            string material_outside_info_2,
            string material_outside_value_2,
            string material_outside_info_2b,
            string material_outside_value_2b,
            string material_outside_cost_sgd_2,
            string material_outside_percentage_target_price_2,
            string material_outside_name_3,
            string material_outside_info_3,
            string material_outside_value_3,
            string material_outside_info_3b,
            string material_outside_value_3b,
            string material_outside_cost_sgd_3,
            string material_outside_percentage_target_price_3,
            string direct_material_cost,
            string direct_material_cost_percentage,
            string sub_material_percentage,
            string sub_material_cost,
            string sub_material_cost_percentage, 
            string direct_process_cost,
            string direct_process_cost_percentage,
            string total_direct_cost,
            string total_direct_cost_percentage,
            string defective_percentage,
            string defective_cost,
            string defective_cost_percentage,
            string indirect_percentage,
            string indirect_cost,
            string indirect_cost_percentage,
            string packing_material_percentage,
            string special_package_cost,
            string packing_material_cost,
            string packing_material_cost_percentage,
            string administration_percentage,
            string administration_cost,
            string administration_cost_percentage,
            string plant_retail_percentage,
            string plant_retail_cost,
            string plant_retail_cost_percentage,
            string moldjig_percentage,
            string moldjig_cost,
            string moldjig_cost_percentage,
            string die_percentage,
            string die_cost,
            string die_cost_percentage,
            string note,
            string net_included_tooling_cost,
            string net_included_tooling_cost_percentage,
            string net_exclude_tooling_cost,
            string net_exclude_tooling_cost_percentage,
            string tooling_list_description_1,
            string tooling_list_description_2,
            string tooling_list_description_3,
            string tooling_list_description_4,
            string tooling_list_description_5,
            string tooling_list_description_6,
            string tooling_list_description_7,
            string tooling_list_description_8,
            string tooling_list_description_9,
            string tooling_list_description_10,
            string tooling_list_description_11,
            string tooling_list_description_12,
            string tooling_list_description_13,
            string tooling_list_description_14,
            string tooling_list_description_15,
            string tooling_list_type_1,
            string tooling_list_type_2,
            string tooling_list_type_3,
            string tooling_list_type_4,
            string tooling_list_type_5,
            string tooling_list_type_6,
            string tooling_list_type_7,
            string tooling_list_type_8,
            string tooling_list_type_9,
            string tooling_list_type_10,
            string tooling_list_type_11,
            string tooling_list_type_12,
            string tooling_list_type_13,
            string tooling_list_type_14,
            string tooling_list_type_15,
            string tooling_list_source_1,
            string tooling_list_source_2,
            string tooling_list_source_3,
            string tooling_list_source_4,
            string tooling_list_source_5,
            string tooling_list_source_6,
            string tooling_list_source_7,
            string tooling_list_source_8,
            string tooling_list_source_9,
            string tooling_list_source_10,
            string tooling_list_source_11,
            string tooling_list_source_12,
            string tooling_list_source_13,
            string tooling_list_source_14,
            string tooling_list_source_15,
            string tooling_list_qty_1,
            string tooling_list_qty_2,
            string tooling_list_qty_3,
            string tooling_list_qty_4,
            string tooling_list_qty_5,
            string tooling_list_qty_6,
            string tooling_list_qty_7,
            string tooling_list_qty_8,
            string tooling_list_qty_9,
            string tooling_list_qty_10,
            string tooling_list_qty_11,
            string tooling_list_qty_12,
            string tooling_list_qty_13,
            string tooling_list_qty_14,
            string tooling_list_qty_15,
            string tooling_list_unit_1,
            string tooling_list_unit_2,
            string tooling_list_unit_3,
            string tooling_list_unit_4,
            string tooling_list_unit_5,
            string tooling_list_unit_6,
            string tooling_list_unit_7,
            string tooling_list_unit_8,
            string tooling_list_unit_9,
            string tooling_list_unit_10,
            string tooling_list_unit_11,
            string tooling_list_unit_12,
            string tooling_list_unit_13,
            string tooling_list_unit_14,
            string tooling_list_unit_15,
            string tooling_list_price_1,
            string tooling_list_price_2,
            string tooling_list_price_3,
            string tooling_list_price_4,
            string tooling_list_price_5,
            string tooling_list_price_6,
            string tooling_list_price_7,
            string tooling_list_price_8,
            string tooling_list_price_9,
            string tooling_list_price_10,
            string tooling_list_price_11,
            string tooling_list_price_12,
            string tooling_list_price_13,
            string tooling_list_price_14,
            string tooling_list_price_15,
            string tooling_list_amount_jpy_1,
            string tooling_list_amount_jpy_2,
            string tooling_list_amount_jpy_3,
            string tooling_list_amount_jpy_4,
            string tooling_list_amount_jpy_5,
            string tooling_list_amount_jpy_6,
            string tooling_list_amount_jpy_7,
            string tooling_list_amount_jpy_8,
            string tooling_list_amount_jpy_9,
            string tooling_list_amount_jpy_10,
            string tooling_list_amount_jpy_11,
            string tooling_list_amount_jpy_12,
            string tooling_list_amount_jpy_13,
            string tooling_list_amount_jpy_14,
            string tooling_list_amount_jpy_15,
            string tooling_list_amount_usd_1,
            string tooling_list_amount_usd_2,
            string tooling_list_amount_usd_3,
            string tooling_list_amount_usd_4,
            string tooling_list_amount_usd_5,
            string tooling_list_amount_usd_6,
            string tooling_list_amount_usd_7,
            string tooling_list_amount_usd_8,
            string tooling_list_amount_usd_9,
            string tooling_list_amount_usd_10,
            string tooling_list_amount_usd_11,
            string tooling_list_amount_usd_12,
            string tooling_list_amount_usd_13,
            string tooling_list_amount_usd_14,
            string tooling_list_amount_usd_15,
            string tooling_list_total_amount_usd,
            string direct_raw_material,
            string direct_raw_material_p,
            string sub_material,
            string sub_material_p,
            string raw_material_cost_sub_total,
            string raw_material_cost_sub_total_p,
            string labor_cost,
            string labor_cost_p,
            string machine_cost,
            string machine_cost_p,
            string overhead_cost,
            string overhead_cost_p,
            string process_cost_sub_total,
            string process_cost_sub_total_p,
            string defectives,
            string defectives_p,
            string admin_engin_qc,
            string admin_engin_qc_p,
            string tooling_cost,
            string tooling_cost_p,
            string process_margin_adjust,
            string process_margin_adjust_p,
            string other_fixed_cost_sub_total,
            string other_fixed_cost_sub_total_p,
            string grand_total_cost,
            string grand_total_cost_p
        )
        {
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Cost");

                
                worksheet.Columns().Style.Font.FontSize = 13;
                worksheet.Columns().Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
                worksheet.Columns().Style.Alignment.SetVertical(XLAlignmentVerticalValues.Center);
                //worksheet.Columns().AdjustToContents();
                worksheet.ColumnWidth = 15;
                worksheet.RowHeight = 25;
                

                worksheet.Columns("A:A").Width = 3;
                worksheet.Columns("B:B").Width = 2;
                worksheet.Columns("C:C").Width = 2;

                worksheet.Cell(2, 2).Value = "1. Object";
                worksheet.Cell(2, 2).Style.Font.FontSize = 18;
                worksheet.Cell(2, 2).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Left);
                worksheet.Range("B2:D2").Merge(true);

                worksheet.Cell(3, 2).Value = "Customer";
                worksheet.Range("B3:D3").Merge(true);
                worksheet.Range("B3:D3").Style.Fill.BackgroundColor = XLColor.LightGray;
                worksheet.Range("B3:D3").Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Range("B3:D3").Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Range("B3:D3").Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Range("B3:D3").Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(3, 5).Value = customer;
                worksheet.Range("E3:J3").Merge(true);
                worksheet.Range("E3:J3").Style.Fill.BackgroundColor = XLColor.LightYellow;
                worksheet.Range("E3:J3").Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Range("E3:J3").Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Range("E3:J3").Style.Border.RightBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(4, 2).Value = "Parts code";
                worksheet.Range("B4:D4").Merge(true);
                worksheet.Range("B4:D4").Style.Fill.BackgroundColor = XLColor.LightGray;
                worksheet.Range("B4:D4").Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Range("B4:D4").Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Range("B4:D4").Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Range("B4:D4").Style.Border.LeftBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(4, 5).Value = parts_code;
                worksheet.Range("E4:G4").Merge(true);
                worksheet.Range("E4:G4").Style.Fill.BackgroundColor = XLColor.LightYellow;
                worksheet.Range("E4:G4").Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Range("E4:G4").Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Range("E4:G4").Style.Border.RightBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(4, 8).Value = "Item: (If have)";
                worksheet.Cell(4, 8).Style.Fill.BackgroundColor = XLColor.LightGray;
                worksheet.Cell(4, 8).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(4, 8).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(4, 8).Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(4, 8).Style.Border.LeftBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(4, 9).Value = item;
                worksheet.Cell(4, 9).Style.Fill.BackgroundColor = XLColor.LightYellow;
                worksheet.Cell(4, 9).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(4, 9).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(4, 9).Style.Border.RightBorder = XLBorderStyleValues.Thin;

                worksheet.Range("J4:J5").Merge(true);
                worksheet.Range("J4:J5").Style.Fill.BackgroundColor = XLColor.LightYellow;
                worksheet.Range("J4:J5").Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Range("J4:J5").Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Range("J4:J5").Style.Border.RightBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(5, 2).Value = "Product";
                worksheet.Range("B5:D5").Merge(true);
                worksheet.Range("B5:D5").Style.Fill.BackgroundColor = XLColor.LightGray;
                worksheet.Range("B5:D5").Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Range("B5:D5").Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Range("B5:D5").Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Range("B5:D5").Style.Border.LeftBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(5, 5).Value = product;
                worksheet.Range("E5:G5").Merge(true);
                worksheet.Range("E5:G5").Style.Fill.BackgroundColor = XLColor.LightYellow;
                worksheet.Range("E5:G5").Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Range("E5:G5").Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Range("E5:G5").Style.Border.RightBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(5, 8).Value = "Product Type:";
                worksheet.Cell(5, 8).Style.Fill.BackgroundColor = XLColor.LightGray;
                worksheet.Cell(5, 8).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(5, 8).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(5, 8).Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(5, 8).Style.Border.LeftBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(5, 9).Value = product_type;
                worksheet.Cell(5, 9).Style.Fill.BackgroundColor = XLColor.LightYellow;
                worksheet.Cell(5, 9).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(5, 9).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(5, 9).Style.Border.RightBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(6, 2).Value = "Size(IDxODxW1xW2)";
                worksheet.Range("B6:D6").Merge(true);
                worksheet.Range("B6:D6").Style.Fill.BackgroundColor = XLColor.LightGray;
                worksheet.Range("B6:D6").Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Range("B6:D6").Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Range("B6:D6").Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Range("B6:D6").Style.Border.LeftBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(6, 5).Value = item_id;
                worksheet.Range("E6:G6").Merge(true);
                worksheet.Range("E6:G6").Style.Fill.BackgroundColor = XLColor.LightYellow;
                worksheet.Range("E6:G6").Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Range("E6:G6").Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Range("E6:G6").Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(6, 8).Value = item_od;
                worksheet.Cell(6, 8).Style.Fill.BackgroundColor = XLColor.LightYellow;
                worksheet.Cell(6, 8).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(6, 8).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(6, 8).Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(6, 9).Value = item_w;
                worksheet.Cell(6, 9).Style.Fill.BackgroundColor = XLColor.LightYellow;
                worksheet.Cell(6, 9).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(6, 9).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(6, 9).Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(6, 10).Value = item_w2;
                worksheet.Cell(6, 10).Style.Fill.BackgroundColor = XLColor.LightYellow;
                worksheet.Cell(6, 10).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(6, 10).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(6, 10).Style.Border.RightBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(8, 2).Value = "Business Type:";
                worksheet.Range("B8:D8").Merge(true);
                worksheet.Range("B8:D8").Style.Fill.BackgroundColor = XLColor.LightGray;
                worksheet.Range("B8:D8").Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Range("B8:D8").Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Range("B8:D8").Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Range("B8:D8").Style.Border.LeftBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(8, 5).Value = business_type;
                worksheet.Range("E8:L8").Merge(true);
                worksheet.Range("E8:L8").Style.Fill.BackgroundColor = XLColor.LightYellow;
                worksheet.Range("E8:L8").Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Range("E8:L8").Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Range("E8:L8").Style.Border.RightBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(9, 2).Value = "Quantity/month:";
                worksheet.Range("B9:D9").Merge(true);
                worksheet.Range("B9:D9").Style.Fill.BackgroundColor = XLColor.LightGray;
                worksheet.Range("B9:D9").Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Range("B9:D9").Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Range("B9:D9").Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Range("B9:D9").Style.Border.LeftBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(9, 5).Value = qty_month;
                worksheet.Range("E9:L9").Merge(true);
                worksheet.Range("E9:L9").Style.Fill.BackgroundColor = XLColor.LightYellow;
                worksheet.Range("E9:L9").Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Range("E9:L9").Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Range("E9:L9").Style.Border.RightBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(10, 2).Value = "Target Price(SGD):";
                worksheet.Range("B10:D10").Merge(true);
                worksheet.Range("B10:D10").Style.Fill.BackgroundColor = XLColor.LightGray;
                worksheet.Range("B10:D10").Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Range("B10:D10").Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(11, 2).Value = "Included Tooling";
                worksheet.Range("B11:D11").Merge(true);
                worksheet.Range("B11:D11").Style.Fill.BackgroundColor = XLColor.LightGray;
                worksheet.Range("B11:D11").Style.Fill.BackgroundColor = XLColor.LightGray;
                worksheet.Range("B11:D11").Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Range("B11:D11").Style.Border.RightBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(10, 5).Value = "SGD";
                worksheet.Cell(10, 5).Style.Fill.BackgroundColor = XLColor.LightGray;
                worksheet.Cell(10, 5).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(10, 5).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(10, 5).Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(10, 5).Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(10, 6).Value = "USD";
                worksheet.Cell(10, 6).Style.Fill.BackgroundColor = XLColor.LightGray;
                worksheet.Cell(10, 6).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(10, 6).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(10, 6).Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(10, 6).Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(10, 7).Value = "EUR";
                worksheet.Cell(10, 7).Style.Fill.BackgroundColor = XLColor.LightGray;
                worksheet.Cell(10, 7).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(10, 7).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(10, 7).Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(10, 7).Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(10, 8).Value = "Exchange Rate:";
                worksheet.Range("H10:I11").Merge(true);
                worksheet.Range("H10:I11").Style.Alignment.SetVertical(XLAlignmentVerticalValues.Top);
                worksheet.Range("H10:I10").Style.Font.FontColor = XLColor.Red;
                worksheet.Range("H10:I10").Style.Fill.BackgroundColor = XLColor.LightGray;
                worksheet.Range("H10:I10").Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Range("H10:I10").Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Range("H10:I10").Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Range("H10:I10").Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(10, 10).Value = "JPY";
                worksheet.Cell(10, 10).Style.Fill.BackgroundColor = XLColor.LightGray;
                worksheet.Cell(10, 10).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(10, 10).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(10, 10).Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(10, 10).Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(10, 11).Value = "USD";
                worksheet.Cell(10, 11).Style.Fill.BackgroundColor = XLColor.LightGray;
                worksheet.Cell(10, 11).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(10, 11).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(10, 11).Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(10, 11).Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(10, 12).Value = "EUR";
                worksheet.Cell(10, 12).Style.Fill.BackgroundColor = XLColor.LightGray;
                worksheet.Cell(10, 12).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(10, 12).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(10, 12).Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(10, 12).Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(11, 5).Value = target_price_sgd;
                worksheet.Cell(11, 5).Style.Fill.BackgroundColor = XLColor.LightGreen;
                worksheet.Cell(11, 5).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(11, 5).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(11, 5).Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(11, 5).Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(11, 6).Value = target_price_usd;
                worksheet.Cell(11, 6).Style.Fill.BackgroundColor = XLColor.LightGreen;
                worksheet.Cell(11, 6).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(11, 6).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(11, 6).Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(11, 6).Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(11, 7).Value = target_price_eud;
                worksheet.Cell(11, 7).Style.Fill.BackgroundColor = XLColor.LightGreen;
                worksheet.Cell(11, 7).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(11, 7).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(11, 7).Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(11, 7).Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(11, 10).Value = exchange_rate_jpy;
                worksheet.Cell(11, 10).Style.Fill.BackgroundColor = XLColor.LightYellow;
                worksheet.Cell(11, 10).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(11, 10).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(11, 10).Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(11, 10).Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(11, 11).Value = exchange_rate_usd;
                worksheet.Cell(11, 11).Style.Fill.BackgroundColor = XLColor.LightYellow;
                worksheet.Cell(11, 11).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(11, 11).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(11, 11).Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(11, 11).Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(11, 12).Value = exchange_rate_eud;
                worksheet.Cell(11, 12).Style.Fill.BackgroundColor = XLColor.LightYellow;
                worksheet.Cell(11, 12).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(11, 12).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(11, 12).Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(11, 12).Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(12, 2).Value = "Production Quantity/day";
                worksheet.Range("B12:G12").Merge(true);
                worksheet.Range("B12:G12").Style.Fill.BackgroundColor = XLColor.LightGray;
                worksheet.Range("B12:G12").Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Range("B12:G12").Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Range("B12:G12").Style.Border.RightBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(12, 8).Value = production_qty_day + "  pc/day, 3 Shifts";
                worksheet.Range("H12:L12").Merge(true);
                worksheet.Range("H12:L12").Style.Fill.BackgroundColor = XLColor.LightGreen;
                worksheet.Range("H12:L12").Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Range("H12:L12").Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Range("H12:L12").Style.Border.RightBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(13, 2).Value = "Working Day";
                worksheet.Range("B13:G13").Merge(true);
                worksheet.Range("B13:G13").Style.Fill.BackgroundColor = XLColor.LightGray;
                worksheet.Range("B13:G13").Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Range("B13:G13").Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Range("B13:G13").Style.Border.RightBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(13, 8).Value = working_day + "  day/month";
                worksheet.Range("H13:L13").Merge(true);
                worksheet.Range("H13:L13").Style.Fill.BackgroundColor = XLColor.LightYellow;
                worksheet.Range("H13:L13").Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Range("H13:L13").Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Range("H13:L13").Style.Border.RightBorder = XLBorderStyleValues.Thin;

                //    worksheet.Cell(currentRow, n).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
                //    worksheet.Cell(currentRow, n).Style.Alignment.SetVertical(XLAlignmentVerticalValues.Center);

                worksheet.Columns("M:M").Width = 3;
                worksheet.Columns("N:N").Width = 20;
                worksheet.Columns("O:O").Width = 25;

                worksheet.Cell(3, 14).Value = "Issue Date: ";
                worksheet.Cell(3, 14).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Left);
                worksheet.Cell(3, 14).Style.Fill.BackgroundColor = XLColor.LightGray;
                worksheet.Cell(3, 14).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(3, 14).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(3, 14).Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(3, 14).Style.Border.LeftBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(3, 15).Value = issue_date;
                worksheet.Cell(3, 15).Style.Fill.BackgroundColor = XLColor.LightYellow;
                worksheet.Cell(3, 15).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(3, 15).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(3, 15).Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(3, 15).Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(4, 14).Value = "Section: ";
                worksheet.Cell(4, 14).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Left);
                worksheet.Cell(4, 14).Style.Fill.BackgroundColor = XLColor.LightGray;
                worksheet.Cell(4, 14).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(4, 14).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(4, 14).Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(4, 14).Style.Border.LeftBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(4, 15).Value = section;
                worksheet.Cell(4, 15).Style.Fill.BackgroundColor = XLColor.LightYellow;
                worksheet.Cell(4, 15).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(4, 15).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(4, 15).Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(4, 15).Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(5, 14).Value = "Doc No: ";
                worksheet.Cell(5, 14).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Left);
                worksheet.Cell(5, 14).Style.Fill.BackgroundColor = XLColor.LightGray;
                worksheet.Cell(5, 14).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(5, 14).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(5, 14).Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(5, 14).Style.Border.LeftBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(5, 15).Value = doc_no;
                worksheet.Cell(5, 15).Style.Fill.BackgroundColor = XLColor.LightYellow;
                worksheet.Cell(5, 15).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(5, 15).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(5, 15).Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(5, 15).Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(6, 14).Value = "WR No: ";
                worksheet.Cell(6, 14).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Left);
                worksheet.Cell(6, 14).Style.Fill.BackgroundColor = XLColor.LightGray;
                worksheet.Cell(6, 14).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(6, 14).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(6, 14).Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(6, 14).Style.Border.LeftBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(6, 15).Value = wr_no;
                worksheet.Cell(6, 15).Style.Fill.BackgroundColor = XLColor.LightYellow;
                worksheet.Cell(6, 15).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(6, 15).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(6, 15).Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(6, 15).Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(7, 14).Value = "Sales: ";
                worksheet.Cell(7, 14).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Left);
                worksheet.Cell(7, 14).Style.Fill.BackgroundColor = XLColor.LightGray;
                worksheet.Cell(7, 14).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(7, 14).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(7, 14).Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(7, 14).Style.Border.LeftBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(7, 15).Value = sales;
                worksheet.Cell(7, 15).Style.Fill.BackgroundColor = XLColor.LightYellow;
                worksheet.Cell(7, 15).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(7, 15).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(7, 15).Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(7, 15).Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(8, 14).Value = "Revision no: ";
                worksheet.Cell(8, 14).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Left);
                worksheet.Cell(8, 14).Style.Fill.BackgroundColor = XLColor.LightGray;
                worksheet.Cell(8, 14).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(8, 14).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(8, 14).Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(8, 14).Style.Border.LeftBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(8, 15).Value = revision_no;
                worksheet.Cell(8, 15).Style.Fill.BackgroundColor = XLColor.LightYellow;
                worksheet.Cell(8, 15).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(8, 15).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(8, 15).Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(8, 15).Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(9, 14).Value = "Checked by/date: ";
                worksheet.Cell(9, 14).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Left);
                worksheet.Cell(9, 14).Style.Fill.BackgroundColor = XLColor.LightGray;
                worksheet.Cell(9, 14).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(9, 14).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(9, 14).Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(9, 14).Style.Border.LeftBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(9, 15).Value = checked_date;
                worksheet.Cell(9, 15).Style.Fill.BackgroundColor = XLColor.LightYellow;
                worksheet.Cell(9, 15).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(9, 15).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(9, 15).Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(9, 15).Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(10, 14).Value = "Approved by/date: ";
                worksheet.Cell(10, 14).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Left);
                worksheet.Cell(10, 14).Style.Fill.BackgroundColor = XLColor.LightGray;
                worksheet.Cell(10, 14).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(10, 14).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(10, 14).Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(10, 14).Style.Border.LeftBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(10, 15).Value = approved_by;
                worksheet.Cell(10, 15).Style.Fill.BackgroundColor = XLColor.LightYellow;
                worksheet.Cell(10, 15).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(10, 15).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(10, 15).Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(10, 15).Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(11, 14).Value = "Expired by/date: ";
                worksheet.Cell(11, 14).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Left);
                worksheet.Cell(11, 14).Style.Fill.BackgroundColor = XLColor.LightGray;
                worksheet.Cell(11, 14).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(11, 14).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(11, 14).Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(11, 14).Style.Border.LeftBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(11, 15).Value = expired_by;
                worksheet.Cell(11, 15).Style.Fill.BackgroundColor = XLColor.LightYellow;
                worksheet.Cell(11, 15).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(11, 15).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(11, 15).Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(11, 15).Style.Border.LeftBorder = XLBorderStyleValues.Thin;


                worksheet.Cell(15, 2).Value = "2. Calculation";
                worksheet.Cell(15, 2).Style.Font.FontSize = 18;
                worksheet.Cell(15, 2).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Left);
                worksheet.Range("B15:D15").Merge(true);

                worksheet.Cell(16, 2).Value = "M a t e r  i  a l";
                worksheet.Range("B16:B26").Merge(true);
                worksheet.Range("B16:B26").Style.Alignment.WrapText = true;
                worksheet.Range("B16:B26").Style.Fill.BackgroundColor = XLColor.LightGray;
                worksheet.Range("B16:B16").Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Range("B26:B26").Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Range("B16:B26").Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Range("B16:B26").Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(16, 3).Value = "R u b b e r";
                worksheet.Range("C16:C20").Merge(true);
                worksheet.Range("C16:C20").Style.Alignment.WrapText = true;
                worksheet.Range("C16:C20").Style.Fill.BackgroundColor = XLColor.LightGray;
                worksheet.Range("C16:C20").Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Range("C19:C20").Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Range("C16:C20").Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Range("C16:C20").Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(16, 4).Value = "Material name";
                worksheet.Range("D16:D17").Merge(true);
                worksheet.Range("D16:D17").Style.Fill.BackgroundColor = XLColor.LightGray;
                worksheet.Range("D16:D16").Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Range("D17:D17").Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Range("D16:D17").Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Range("D16:D17").Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(18, 4).Value = rubber_material_name;
                worksheet.Cell(18, 4).Style.Fill.BackgroundColor = XLColor.LightYellow;
                worksheet.Cell(18, 4).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(18, 4).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(18, 4).Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(18, 4).Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(19, 4).Value = rubber_material_name2;
                worksheet.Cell(19, 4).Style.Fill.BackgroundColor = XLColor.LightYellow;
                worksheet.Cell(19, 4).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(19, 4).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(19, 4).Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(19, 4).Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(16, 5).Value = "price(kg)";
                worksheet.Range("E16:E17").Merge(true);
                worksheet.Range("E16:E17").Style.Fill.BackgroundColor = XLColor.LightGray;
                worksheet.Range("E16:E16").Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Range("E17:E17").Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Range("E16:E17").Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Range("E16:E17").Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(18, 5).Value = rubber_price_kg;
                worksheet.Cell(18, 5).Style.Fill.BackgroundColor = XLColor.LightYellow;
                worksheet.Cell(18, 5).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(18, 5).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(18, 5).Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(18, 5).Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(19, 5).Value = rubber_price_kg2;
                worksheet.Cell(19, 5).Style.Fill.BackgroundColor = XLColor.LightYellow;
                worksheet.Cell(19, 5).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(19, 5).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(19, 5).Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(19, 5).Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(16, 6).Value = "Mixing Process Cost";
                worksheet.Range("F16:F17").Merge(true);
                worksheet.Range("F16:F17").Style.Fill.BackgroundColor = XLColor.LightGray;
                worksheet.Range("F16:F16").Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Range("F17:F17").Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Range("F16:F17").Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Range("F16:F17").Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(18, 6).Value = rubber_mixing_process_cost;
                worksheet.Cell(18, 6).Style.Fill.BackgroundColor = XLColor.LightYellow;
                worksheet.Cell(18, 6).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(18, 6).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(18, 6).Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(18, 6).Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(19, 6).Value = rubber_mixing_process_cost2;
                worksheet.Cell(19, 6).Style.Fill.BackgroundColor = XLColor.LightYellow;
                worksheet.Cell(19, 6).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(19, 6).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(19, 6).Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(19, 6).Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(16, 7).Value = "Weight";
                worksheet.Range("G16:H16").Merge(true);
                worksheet.Range("G16:H16").Style.Fill.BackgroundColor = XLColor.LightGray;
                worksheet.Range("G16:H16").Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Range("G16:H16").Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Range("H16:H16").Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Range("G16:G16").Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(17, 7).Value = "g/pc";
                worksheet.Range("G17:G17").Style.Fill.BackgroundColor = XLColor.LightGray;
                worksheet.Range("G17:G17").Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Range("G17:G17").Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Range("G17:G17").Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Range("G17:G17").Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(18, 7).Value = rubber_weight_g;
                worksheet.Cell(18, 7).Style.Fill.BackgroundColor = XLColor.LightYellow;
                worksheet.Cell(18, 7).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(18, 7).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(18, 7).Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(18, 7).Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(19, 7).Value = rubber_weight_g2;
                worksheet.Cell(19, 7).Style.Fill.BackgroundColor = XLColor.LightYellow;
                worksheet.Cell(19, 7).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(19, 7).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(19, 7).Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(19, 7).Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(17, 8).Value = "kg/pc";
                worksheet.Range("H17:H17").Style.Fill.BackgroundColor = XLColor.LightGray;
                worksheet.Range("H17:H17").Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Range("H17:H17").Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Range("H17:H17").Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Range("H17:H17").Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(18, 8).Value = rubber_weight_kg;
                worksheet.Cell(18, 8).Style.Fill.BackgroundColor = XLColor.LightGreen;
                worksheet.Cell(18, 8).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(18, 8).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(18, 8).Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(18, 8).Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(19, 8).Value = rubber_weight_kg2;
                worksheet.Cell(19, 8).Style.Fill.BackgroundColor = XLColor.LightGreen;
                worksheet.Cell(19, 8).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(19, 8).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(19, 8).Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(19, 8).Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(16, 9).Value = "Yield Rate(%)";
                worksheet.Range("I16:I17").Merge(true);
                worksheet.Range("I16:I17").Style.Fill.BackgroundColor = XLColor.LightGray;
                worksheet.Range("I16:I16").Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Range("I17:I17").Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Range("I16:I17").Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Range("I16:I17").Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(18, 9).Value = rubber_yield_rate;
                worksheet.Cell(18, 9).Style.Fill.BackgroundColor = XLColor.LightYellow;
                worksheet.Cell(18, 9).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(18, 9).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(18, 9).Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(18, 9).Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(19, 9).Value = rubber_yield_rate2;
                worksheet.Cell(19, 9).Style.Fill.BackgroundColor = XLColor.LightYellow;
                worksheet.Cell(19, 9).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(19, 9).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(19, 9).Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(19, 9).Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(16, 10).Value = "Weight";
                worksheet.Cell(17, 10).Value = "(kg/pcs)";
                worksheet.Range("J16:J17").Style.Fill.BackgroundColor = XLColor.LightGray;
                worksheet.Range("J16:J16").Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Range("J17:J17").Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Range("J16:J17").Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Range("J16:J17").Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(18, 10).Value = rubber_weight_kg_yieldrate;
                worksheet.Cell(18, 10).Style.Fill.BackgroundColor = XLColor.LightGreen;
                worksheet.Cell(18, 10).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(18, 10).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(18, 10).Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(18, 10).Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(19, 10).Value = rubber_weight_kg_yieldrate2;
                worksheet.Cell(19, 10).Style.Fill.BackgroundColor = XLColor.LightGreen;
                worksheet.Cell(19, 10).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(19, 10).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(19, 10).Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(19, 10).Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(16, 11).Value = "SGD";
                worksheet.Range("K16:K17").Merge(true);
                worksheet.Range("K16:K17").Style.Fill.BackgroundColor = XLColor.LightGray;
                worksheet.Range("K16:K16").Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Range("K17:K17").Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Range("K16:K17").Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Range("K16:K17").Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(18, 11).Value = rubber_cost_sgd;
                worksheet.Cell(18, 11).Style.Fill.BackgroundColor = XLColor.LightGreen;
                worksheet.Cell(18, 11).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(18, 11).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(18, 11).Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(18, 11).Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(19, 11).Value = rubber_cost_sgd2;
                worksheet.Cell(19, 11).Style.Fill.BackgroundColor = XLColor.LightGreen;
                worksheet.Cell(19, 11).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(19, 11).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(19, 11).Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(19, 11).Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(16, 12).Value = "% of Target";
                worksheet.Cell(17, 12).Value = "price";
                worksheet.Range("L16:L17").Style.Fill.BackgroundColor = XLColor.LightGray;
                worksheet.Range("L16:L16").Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Range("L17:L17").Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Range("L16:L17").Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Range("L16:L17").Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(18, 12).Value = rubber_percentage_target_price;
                worksheet.Cell(18, 12).Style.Fill.BackgroundColor = XLColor.LightGreen;
                worksheet.Cell(18, 12).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(18, 12).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(18, 12).Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(18, 12).Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(19, 12).Value = rubber_percentage_target_price2;
                worksheet.Cell(19, 12).Style.Fill.BackgroundColor = XLColor.LightGreen;
                worksheet.Cell(19, 12).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(19, 12).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(19, 12).Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(19, 12).Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(20, 4).Value = "Material Information";
                worksheet.Range("D20:L20").Merge(true);
                worksheet.Range("D20:L20").Style.Fill.BackgroundColor = XLColor.LightGray;
                worksheet.Range("D20:L20").Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Range("D20:L20").Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Range("D20:L20").Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Range("D20:L20").Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(21, 3).Value = "I n H o u s e";
                worksheet.Range("C21:C23").Merge(true);
                worksheet.Range("C21:C23").Style.Alignment.WrapText = true;
                worksheet.Range("C21:C23").Style.Fill.BackgroundColor = XLColor.LightGray;
                worksheet.Range("C21:C23").Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Range("C21:C23").Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Range("C21:C23").Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Range("C21:C23").Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(21, 4).Value = material_inhouse_name_1;
                worksheet.Cell(21, 4).Style.Fill.BackgroundColor = XLColor.LightYellow;
                worksheet.Cell(21, 4).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(21, 4).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(21, 4).Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(21, 4).Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(21, 5).Value = material_inhouse_info_1;
                worksheet.Range("E21:F21").Merge(true);
                worksheet.Range("E21:F21").Style.Fill.BackgroundColor = XLColor.LightYellow;
                worksheet.Range("E21:F21").Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Range("E21:F21").Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Range("F21:F21").Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Range("E21:E21").Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(21, 7).Value = material_inhouse_value_1;
                worksheet.Cell(21, 7).Style.Fill.BackgroundColor = XLColor.LightYellow;
                worksheet.Cell(21, 7).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(21, 7).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(21, 7).Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(21, 7).Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(21, 8).Value = material_inhouse_info_1b;
                worksheet.Range("H21:I21").Merge(true);
                worksheet.Range("H21:I21").Style.Fill.BackgroundColor = XLColor.LightYellow;
                worksheet.Range("H21:I21").Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Range("H21:I21").Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Range("I21:I21").Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Range("H21:H21").Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(21, 10).Value = material_inhouse_value_1b;
                worksheet.Cell(21, 10).Style.Fill.BackgroundColor = XLColor.LightYellow;
                worksheet.Cell(21, 10).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(21, 10).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(21, 10).Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(21, 10).Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(21, 11).Value = material_inhouse_cost_sgd_1;
                worksheet.Cell(21, 11).Style.Fill.BackgroundColor = XLColor.LightGreen;
                worksheet.Cell(21, 11).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(21, 11).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(21, 11).Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(21, 11).Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(21, 12).Value = material_inhouse_percentage_target_price_1;
                worksheet.Cell(21, 12).Style.Fill.BackgroundColor = XLColor.LightGreen;
                worksheet.Cell(21, 12).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(21, 12).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(21, 12).Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(21, 12).Style.Border.LeftBorder = XLBorderStyleValues.Thin;


                worksheet.Cell(22, 4).Value = material_inhouse_name_2;
                worksheet.Cell(22, 4).Style.Fill.BackgroundColor = XLColor.LightYellow;
                worksheet.Cell(22, 4).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(22, 4).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(22, 4).Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(22, 4).Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(22, 5).Value = material_inhouse_info_2;
                worksheet.Range("E22:F22").Merge(true);
                worksheet.Range("E22:F22").Style.Fill.BackgroundColor = XLColor.LightYellow;
                worksheet.Range("E22:F22").Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Range("E22:F22").Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Range("F22:F22").Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Range("E22:E22").Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(22, 7).Value = material_inhouse_value_2;
                worksheet.Cell(22, 7).Style.Fill.BackgroundColor = XLColor.LightYellow;
                worksheet.Cell(22, 7).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(22, 7).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(22, 7).Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(22, 7).Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(22, 8).Value = material_inhouse_info_2b;
                worksheet.Range("H22:I22").Merge(true);
                worksheet.Range("H22:I22").Style.Fill.BackgroundColor = XLColor.LightYellow;
                worksheet.Range("H22:I22").Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Range("H22:I22").Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Range("I22:I22").Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Range("H22:H22").Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(22, 10).Value = material_inhouse_value_2b;
                worksheet.Cell(22, 10).Style.Fill.BackgroundColor = XLColor.LightYellow;
                worksheet.Cell(22, 10).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(22, 10).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(22, 10).Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(22, 10).Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(22, 11).Value = material_inhouse_cost_sgd_2;
                worksheet.Cell(22, 11).Style.Fill.BackgroundColor = XLColor.LightGreen;
                worksheet.Cell(22, 11).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(22, 11).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(22, 11).Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(22, 11).Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(22, 12).Value = material_inhouse_percentage_target_price_2;
                worksheet.Cell(22, 12).Style.Fill.BackgroundColor = XLColor.LightGreen;
                worksheet.Cell(22, 12).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(22, 12).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(22, 12).Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(22, 12).Style.Border.LeftBorder = XLBorderStyleValues.Thin;



                worksheet.Cell(23, 4).Value = material_inhouse_name_3;
                worksheet.Cell(23, 4).Style.Fill.BackgroundColor = XLColor.LightYellow;
                worksheet.Cell(23, 4).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(23, 4).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(23, 4).Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(23, 4).Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(23, 5).Value = material_inhouse_info_3;
                worksheet.Range("E23:F23").Merge(true);
                worksheet.Range("E23:F23").Style.Fill.BackgroundColor = XLColor.LightYellow;
                worksheet.Range("E23:F23").Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Range("E23:F23").Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Range("F23:F23").Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Range("E23:E23").Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(23, 7).Value = material_inhouse_value_3;
                worksheet.Cell(23, 7).Style.Fill.BackgroundColor = XLColor.LightYellow;
                worksheet.Cell(23, 7).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(23, 7).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(23, 7).Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(23, 7).Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(23, 8).Value = material_inhouse_info_3b;
                worksheet.Range("H23:I23").Merge(true);
                worksheet.Range("H23:I23").Style.Fill.BackgroundColor = XLColor.LightYellow;
                worksheet.Range("H23:I23").Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Range("H23:I23").Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Range("I23:I23").Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Range("H23:H23").Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(23, 10).Value = material_inhouse_value_3b;
                worksheet.Cell(23, 10).Style.Fill.BackgroundColor = XLColor.LightYellow;
                worksheet.Cell(23, 10).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(23, 10).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(23, 10).Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(23, 10).Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(23, 11).Value = material_inhouse_cost_sgd_3;
                worksheet.Cell(23, 11).Style.Fill.BackgroundColor = XLColor.LightGreen;
                worksheet.Cell(23, 11).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(23, 11).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(23, 11).Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(23, 11).Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(23, 12).Value = material_inhouse_percentage_target_price_3;
                worksheet.Cell(23, 12).Style.Fill.BackgroundColor = XLColor.LightGreen;
                worksheet.Cell(23, 12).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(23, 12).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(23, 12).Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(23, 12).Style.Border.LeftBorder = XLBorderStyleValues.Thin;


                worksheet.Cell(24, 3).Value = "O u t s i d e";
                worksheet.Range("C24:C26").Merge(true);
                worksheet.Range("C24:C26").Style.Alignment.WrapText = true;
                worksheet.Range("C24:C26").Style.Fill.BackgroundColor = XLColor.LightGray;
                worksheet.Range("C24:C26").Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Range("C24:C26").Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Range("C24:C26").Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Range("C24:C26").Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(24, 4).Value = material_outside_name_1;
                worksheet.Cell(24, 4).Style.Fill.BackgroundColor = XLColor.LightYellow;
                worksheet.Cell(24, 4).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(24, 4).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(24, 4).Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(24, 4).Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(24, 5).Value = material_outside_info_1;
                worksheet.Range("E24:F24").Merge(true);
                worksheet.Range("E24:F24").Style.Fill.BackgroundColor = XLColor.LightYellow;
                worksheet.Range("E24:F24").Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Range("E24:F24").Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Range("F24:F24").Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Range("E24:E24").Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(24, 7).Value = material_outside_value_1;
                worksheet.Cell(24, 7).Style.Fill.BackgroundColor = XLColor.LightYellow;
                worksheet.Cell(24, 7).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(24, 7).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(24, 7).Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(24, 7).Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(24, 8).Value = material_outside_info_1b;
                worksheet.Range("H24:I24").Merge(true);
                worksheet.Range("H24:I24").Style.Fill.BackgroundColor = XLColor.LightYellow;
                worksheet.Range("H24:I24").Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Range("H24:I24").Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Range("I24:I24").Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Range("H24:H24").Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(24, 10).Value = material_outside_value_1b;
                worksheet.Cell(24, 10).Style.Fill.BackgroundColor = XLColor.LightYellow;
                worksheet.Cell(24, 10).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(24, 10).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(24, 10).Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(24, 10).Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(24, 11).Value = material_outside_cost_sgd_1;
                worksheet.Cell(24, 11).Style.Fill.BackgroundColor = XLColor.LightGreen;
                worksheet.Cell(24, 11).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(24, 11).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(24, 11).Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(24, 11).Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(24, 12).Value = material_outside_percentage_target_price_1;
                worksheet.Cell(24, 12).Style.Fill.BackgroundColor = XLColor.LightGreen;
                worksheet.Cell(24, 12).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(24, 12).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(24, 12).Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(24, 12).Style.Border.LeftBorder = XLBorderStyleValues.Thin;


                worksheet.Cell(25, 4).Value = material_outside_name_2;
                worksheet.Cell(25, 4).Style.Fill.BackgroundColor = XLColor.LightYellow;
                worksheet.Cell(25, 4).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(25, 4).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(25, 4).Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(25, 4).Style.Border.LeftBorder = XLBorderStyleValues.Thin;
                
                worksheet.Cell(25, 5).Value = material_outside_info_2;
                worksheet.Range("E25:F25").Merge(true);
                worksheet.Range("E25:F25").Style.Fill.BackgroundColor = XLColor.LightYellow;
                worksheet.Range("E25:F25").Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Range("E25:F25").Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Range("F25:F25").Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Range("E25:E25").Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(25, 7).Value = material_outside_value_2;
                worksheet.Cell(25, 7).Style.Fill.BackgroundColor = XLColor.LightYellow;
                worksheet.Cell(25, 7).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(25, 7).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(25, 7).Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(25, 7).Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(25, 8).Value = material_outside_info_2b;
                worksheet.Range("H25:I25").Merge(true);
                worksheet.Range("H25:I25").Style.Fill.BackgroundColor = XLColor.LightYellow;
                worksheet.Range("H25:I25").Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Range("H25:I25").Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Range("I25:I25").Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Range("H25:H25").Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(25, 10).Value = material_outside_value_2b;
                worksheet.Cell(25, 10).Style.Fill.BackgroundColor = XLColor.LightYellow;
                worksheet.Cell(25, 10).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(25, 10).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(25, 10).Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(25, 10).Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(25, 11).Value = material_outside_cost_sgd_2;
                worksheet.Cell(25, 11).Style.Fill.BackgroundColor = XLColor.LightGreen;
                worksheet.Cell(25, 11).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(25, 11).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(25, 11).Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(25, 11).Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(25, 12).Value = material_outside_percentage_target_price_2;
                worksheet.Cell(25, 12).Style.Fill.BackgroundColor = XLColor.LightGreen;
                worksheet.Cell(25, 12).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(25, 12).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(25, 12).Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(25, 12).Style.Border.LeftBorder = XLBorderStyleValues.Thin;



                worksheet.Cell(26, 4).Value = material_outside_name_3;
                worksheet.Cell(26, 4).Style.Fill.BackgroundColor = XLColor.LightYellow;
                worksheet.Cell(26, 4).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(26, 4).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(26, 4).Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(26, 4).Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(26, 5).Value = material_outside_info_3;
                worksheet.Range("E26:F26").Merge(true);
                worksheet.Range("E26:F26").Style.Fill.BackgroundColor = XLColor.LightYellow;
                worksheet.Range("E26:F26").Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Range("E26:F26").Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Range("F26:F26").Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Range("E26:E26").Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(26, 7).Value = material_outside_value_3;
                worksheet.Cell(26, 7).Style.Fill.BackgroundColor = XLColor.LightYellow;
                worksheet.Cell(26, 7).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(26, 7).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(26, 7).Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(26, 7).Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(26, 8).Value = material_outside_info_3b;
                worksheet.Range("H26:I26").Merge(true);
                worksheet.Range("H26:I26").Style.Fill.BackgroundColor = XLColor.LightYellow;
                worksheet.Range("H26:I26").Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Range("H26:I26").Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Range("I26:I26").Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Range("H26:H26").Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(26, 10).Value = material_outside_value_3b;
                worksheet.Cell(26, 10).Style.Fill.BackgroundColor = XLColor.LightYellow;
                worksheet.Cell(26, 10).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(26, 10).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(26, 10).Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(26, 10).Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(26, 11).Value = material_outside_cost_sgd_3;
                worksheet.Cell(26, 11).Style.Fill.BackgroundColor = XLColor.LightGreen;
                worksheet.Cell(26, 11).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(26, 11).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(26, 11).Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(26, 11).Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(26, 12).Value = material_outside_percentage_target_price_3;
                worksheet.Cell(26, 12).Style.Fill.BackgroundColor = XLColor.LightGreen;
                worksheet.Cell(26, 12).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(26, 12).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(26, 12).Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(26, 12).Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(27, 2).Value = "　Direct material cost";
                worksheet.Cell(27, 2).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Left);
                worksheet.Range("B27:I27").Merge(true);
                worksheet.Range("B27:I27").Style.Fill.BackgroundColor = XLColor.White;
                worksheet.Range("B27:I27").Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Range("B27:I27").Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Range("I27:I27").Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Range("B27:B27").Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(27, 10).Value = "A";
                worksheet.Cell(27, 10).Style.Fill.BackgroundColor = XLColor.LightGray;
                worksheet.Cell(27, 10).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(27, 10).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(27, 10).Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(27, 10).Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(27, 11).Value = direct_material_cost;
                worksheet.Cell(27, 11).Style.Fill.BackgroundColor = XLColor.LightGreen;
                worksheet.Cell(27, 11).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(27, 11).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(27, 11).Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(27, 11).Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(27, 12).Value = direct_material_cost_percentage;
                worksheet.Cell(27, 12).Style.Fill.BackgroundColor = XLColor.LightGreen;
                worksheet.Cell(27, 12).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(27, 12).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(27, 12).Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(27, 12).Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(28, 2).Value = "　Sub Material";
                worksheet.Cell(28, 2).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Left);
                worksheet.Range("B28:D28").Merge(true);
                worksheet.Range("B28:D28").Style.Fill.BackgroundColor = XLColor.White;
                worksheet.Range("B28:D28").Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Range("B28:D28").Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Range("D28:D28").Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Range("B28:B28").Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(28, 5).Value = sub_material_percentage;
                worksheet.Cell(28, 5).Style.Fill.BackgroundColor = XLColor.LightYellow;
                worksheet.Cell(28, 5).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(28, 5).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(28, 5).Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(28, 5).Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(28, 6).Value = "　×　A";
                worksheet.Cell(28, 6).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Left);
                worksheet.Cell(28, 6).Style.Font.FontColor = XLColor.Red;
                worksheet.Range("F28:J28").Merge(true);
                worksheet.Range("F28:J28").Style.Fill.BackgroundColor = XLColor.White;
                worksheet.Range("F28:J28").Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Range("F28:J28").Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Range("J28:J28").Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Range("F28:F28").Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(28, 11).Value = sub_material_cost;
                worksheet.Cell(28, 11).Style.Fill.BackgroundColor = XLColor.LightGreen;
                worksheet.Cell(28, 11).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(28, 11).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(28, 11).Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(28, 11).Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(28, 12).Value = sub_material_cost_percentage;
                worksheet.Cell(28, 12).Style.Fill.BackgroundColor = XLColor.LightGreen;
                worksheet.Cell(28, 12).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(28, 12).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(28, 12).Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(28, 12).Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(29, 2).Value = "　Direct process cost";
                worksheet.Cell(29, 2).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Left);
                worksheet.Range("B29:E29").Merge(true);
                worksheet.Range("B29:E29").Style.Fill.BackgroundColor = XLColor.LightBlue;
                worksheet.Range("B29:E29").Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Range("B29:E29").Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Range("E29:E29").Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Range("B29:B29").Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(29, 6).Value = "　(Detail refer to Direct process cost sheets";
                worksheet.Cell(29, 6).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Left);
                worksheet.Range("F29:I29").Merge(true);
                worksheet.Range("F29:I29").Style.Fill.BackgroundColor = XLColor.LightBlue;
                worksheet.Range("F29:I29").Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Range("F29:I29").Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Range("I29:I29").Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Range("F29:I29").Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(29, 10).Value = "B";
                worksheet.Cell(29, 10).Style.Fill.BackgroundColor = XLColor.LightBlue;
                worksheet.Cell(29, 10).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(29, 10).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(29, 10).Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(29, 10).Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(29, 11).Value = direct_process_cost;
                worksheet.Cell(29, 11).Style.Fill.BackgroundColor = XLColor.LightGreen;
                worksheet.Cell(29, 11).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(29, 11).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(29, 11).Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(29, 11).Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(29, 12).Value = direct_process_cost_percentage;
                worksheet.Cell(29, 12).Style.Fill.BackgroundColor = XLColor.LightGreen;
                worksheet.Cell(29, 12).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(29, 12).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(29, 12).Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(29, 12).Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(30, 2).Value = "　Total direct cost　　　　　　( A + B )";
                worksheet.Cell(30, 2).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Left);
                worksheet.Range("B30:I30").Merge(true);
                worksheet.Range("B30:I30").Style.Fill.BackgroundColor = XLColor.White;
                worksheet.Range("B30:I30").Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Range("B30:I30").Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Range("I30:I30").Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Range("B30:B30").Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(30, 10).Value = "C";
                worksheet.Cell(30, 10).Style.Fill.BackgroundColor = XLColor.LightGray;
                worksheet.Cell(30, 10).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(30, 10).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(30, 10).Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(30, 10).Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(30, 11).Value = total_direct_cost;
                worksheet.Cell(30, 11).Style.Fill.BackgroundColor = XLColor.LightGreen;
                worksheet.Cell(30, 11).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(30, 11).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(30, 11).Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(30, 11).Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(30, 12).Value = total_direct_cost_percentage;
                worksheet.Cell(30, 12).Style.Fill.BackgroundColor = XLColor.LightGreen;
                worksheet.Cell(30, 12).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(30, 12).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(30, 12).Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(30, 12).Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(31, 2).Value = "　Defective";
                worksheet.Cell(31, 2).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Left);
                worksheet.Range("B31:D31").Merge(true);
                worksheet.Range("B31:D31").Style.Fill.BackgroundColor = XLColor.White;
                worksheet.Range("B31:D31").Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Range("B31:D31").Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Range("D31:D31").Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Range("B31:B31").Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(31, 5).Value = defective_percentage;
                worksheet.Cell(31, 5).Style.Fill.BackgroundColor = XLColor.LightYellow;
                worksheet.Cell(31, 5).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(31, 5).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(31, 5).Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(31, 5).Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(31, 6).Value = "　×　C　　　　　　　　　　　　　　　　　　　　=　　";
                worksheet.Cell(31, 6).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Left);
                worksheet.Range("F31:I31").Merge(true);
                worksheet.Range("F31:I31").Style.Fill.BackgroundColor = XLColor.White;
                worksheet.Range("F31:I31").Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Range("F31:I31").Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Range("I31:I31").Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Range("F31:F31").Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(31, 10).Value = "D";
                worksheet.Cell(31, 10).Style.Fill.BackgroundColor = XLColor.LightGray;
                worksheet.Cell(31, 10).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(31, 10).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(31, 10).Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(31, 10).Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(31, 11).Value = defective_cost;
                worksheet.Cell(31, 11).Style.Fill.BackgroundColor = XLColor.LightGreen;
                worksheet.Cell(31, 11).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(31, 11).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(31, 11).Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(31, 11).Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(31, 12).Value = defective_cost_percentage;
                worksheet.Cell(31, 12).Style.Fill.BackgroundColor = XLColor.LightGreen;
                worksheet.Cell(31, 12).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(31, 12).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(31, 12).Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(31, 12).Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(32, 2).Value = "　Indirect Cost";
                worksheet.Cell(32, 2).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Left);
                worksheet.Range("B32:D32").Merge(true);
                worksheet.Range("B32:D32").Style.Fill.BackgroundColor = XLColor.White;
                worksheet.Range("B32:D32").Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Range("B32:D32").Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Range("D32:D32").Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Range("B32:B32").Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(32, 5).Value = indirect_percentage;
                worksheet.Cell(32, 5).Style.Fill.BackgroundColor = XLColor.LightYellow;
                worksheet.Cell(32, 5).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(32, 5).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(32, 5).Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(32, 5).Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(32, 6).Value = "　×　（C　＋　D）　 　　　　　　　　　　　　　=　　";
                worksheet.Cell(32, 6).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Left);
                worksheet.Range("F32:I32").Merge(true);
                worksheet.Range("F32:I32").Style.Fill.BackgroundColor = XLColor.White;
                worksheet.Range("F32:I32").Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Range("F32:I32").Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Range("I32:I32").Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Range("F32:F32").Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(32, 10).Value = "E";
                worksheet.Cell(32, 10).Style.Fill.BackgroundColor = XLColor.LightGray;
                worksheet.Cell(32, 10).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(32, 10).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(32, 10).Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(32, 10).Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(32, 11).Value = indirect_cost;
                worksheet.Cell(32, 11).Style.Fill.BackgroundColor = XLColor.LightGreen;
                worksheet.Cell(32, 11).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(32, 11).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(32, 11).Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(32, 11).Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(32, 12).Value = indirect_cost_percentage;
                worksheet.Cell(32, 12).Style.Fill.BackgroundColor = XLColor.LightGreen;
                worksheet.Cell(32, 12).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(32, 12).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(32, 12).Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(32, 12).Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(33, 2).Value = "　Packing Materials";
                worksheet.Cell(33, 2).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Left);
                worksheet.Range("B33:D33").Merge(true);
                worksheet.Range("B33:D33").Style.Fill.BackgroundColor = XLColor.White;
                worksheet.Range("B33:D33").Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Range("B33:D33").Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Range("D33:D33").Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Range("B33:B33").Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(33, 5).Value = packing_material_percentage;
                worksheet.Cell(33, 5).Style.Fill.BackgroundColor = XLColor.LightYellow;
                worksheet.Cell(33, 5).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(33, 5).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(33, 5).Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(33, 5).Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(33, 6).Value = "　×　A";
                worksheet.Cell(33, 6).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Left);
                worksheet.Cell(33, 6).Style.Font.FontColor = XLColor.Red;
                worksheet.Range("F33:F33").Merge(true);
                worksheet.Range("F33:F33").Style.Fill.BackgroundColor = XLColor.White;
                worksheet.Range("F33:F33").Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Range("F33:F33").Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Range("F33:F33").Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Range("F33:F33").Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(33, 7).Value = "Special package costs/pc (if any)";
                worksheet.Range("G33:I33").Merge(true);
                worksheet.Range("G33:I33").Style.Fill.BackgroundColor = XLColor.White;
                worksheet.Range("G33:I33").Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Range("G33:I33").Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Range("I33:I33").Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Range("G33:G33").Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(33, 10).Value = special_package_cost;
                worksheet.Cell(33, 10).Style.Fill.BackgroundColor = XLColor.LightYellow;
                worksheet.Cell(33, 10).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(33, 10).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(33, 10).Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(33, 10).Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(33, 11).Value = packing_material_cost;
                worksheet.Cell(33, 11).Style.Fill.BackgroundColor = XLColor.LightGreen;
                worksheet.Cell(33, 11).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(33, 11).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(33, 11).Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(33, 11).Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(33, 12).Value = packing_material_cost_percentage;
                worksheet.Cell(33, 12).Style.Fill.BackgroundColor = XLColor.LightGreen;
                worksheet.Cell(33, 12).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(33, 12).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(33, 12).Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(33, 12).Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(34, 2).Value = "　Administration";
                worksheet.Cell(34, 2).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Left);
                worksheet.Range("B34:D34").Merge(true);
                worksheet.Range("B34:D34").Style.Fill.BackgroundColor = XLColor.White;
                worksheet.Range("B34:D34").Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Range("B34:D34").Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Range("D34:D34").Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Range("B34:B34").Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(34, 5).Value = administration_percentage;
                worksheet.Cell(34, 5).Style.Fill.BackgroundColor = XLColor.LightYellow;
                worksheet.Cell(34, 5).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(34, 5).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(34, 5).Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(34, 5).Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(34, 6).Value = "　×　（C　＋　D　＋　E）　　　　　　　　　　　=　　";
                worksheet.Cell(34, 6).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Left);
                worksheet.Range("F34:I34").Merge(true);
                worksheet.Range("F34:I34").Style.Fill.BackgroundColor = XLColor.White;
                worksheet.Range("F34:I34").Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Range("F34:I34").Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Range("I34:I34").Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Range("F34:F34").Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(34, 10).Value = "F";
                worksheet.Cell(34, 10).Style.Fill.BackgroundColor = XLColor.LightGray;
                worksheet.Cell(34, 10).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(34, 10).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(34, 10).Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(34, 10).Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(34, 11).Value = administration_cost;
                worksheet.Cell(34, 11).Style.Fill.BackgroundColor = XLColor.LightGreen;
                worksheet.Cell(34, 11).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(34, 11).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(34, 11).Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(34, 11).Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(34, 12).Value = administration_cost_percentage;
                worksheet.Cell(34, 12).Style.Fill.BackgroundColor = XLColor.LightGreen;
                worksheet.Cell(34, 12).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(34, 12).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(34, 12).Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(34, 12).Style.Border.LeftBorder = XLBorderStyleValues.Thin;


                worksheet.Cell(35, 2).Value = "　Tooling Maintenance";
                worksheet.Cell(35, 2).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Left);
                worksheet.Cell(35, 2).Style.Font.FontColor = XLColor.Red;
                worksheet.Range("B35:D35").Style.Fill.BackgroundColor = XLColor.White;
                worksheet.Range("B35:D35").Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Range("B35:D35").Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Range("D35:D35").Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Range("B35:B35").Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(35, 5).Value = plant_retail_percentage;
                worksheet.Cell(35, 5).Style.Fill.BackgroundColor = XLColor.LightYellow;
                worksheet.Cell(35, 5).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(35, 5).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(35, 5).Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(35, 5).Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(35, 6).Value = "　×　（C　＋　D　＋　E　＋　Packing　＋　F）　";
                worksheet.Cell(35, 6).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Left);
                worksheet.Cell(35, 6).Style.Font.FontColor = XLColor.Red;
                worksheet.Range("F35:J35").Merge(true);
                worksheet.Range("F35:J35").Style.Fill.BackgroundColor = XLColor.White;
                worksheet.Range("F35:J35").Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Range("F35:J35").Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Range("J35:J35").Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Range("F35:F35").Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(35, 11).Value = plant_retail_cost;
                worksheet.Cell(35, 11).Style.Fill.BackgroundColor = XLColor.LightGreen;
                worksheet.Cell(35, 11).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(35, 11).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(35, 11).Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(35, 11).Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(35, 12).Value = plant_retail_cost_percentage;
                worksheet.Cell(35, 12).Style.Fill.BackgroundColor = XLColor.LightGreen;
                worksheet.Cell(35, 12).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(35, 12).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(35, 12).Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(35, 12).Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(36, 2).Value = "T o o l";
                worksheet.Range("B36:B39").Merge(true);
                worksheet.Range("B36:B39").Style.Alignment.WrapText = true;
                worksheet.Range("B36:B39").Style.Fill.BackgroundColor = XLColor.LightGray;
                worksheet.Range("B36:B39").Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Range("B36:B39").Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Range("B36:B39").Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Range("B36:B39").Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(36, 3).Value = "　Mold &  Jig";
                worksheet.Cell(36, 3).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Left);
                worksheet.Range("C36:D37").Merge(true);
                worksheet.Range("C36:D37").Style.Fill.BackgroundColor = XLColor.White;
                worksheet.Range("C36:D37").Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Range("C36:D37").Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Range("D36:D37").Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Range("C36:C37").Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(36, 5).Value = moldjig_percentage;
                worksheet.Range("E36:E37").Merge(true);
                worksheet.Range("E36:E37").Style.Fill.BackgroundColor = XLColor.LightYellow;
                worksheet.Range("E36:E37").Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Range("E36:E37").Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Range("E36:E37").Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Range("E36:E37").Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(36, 6).Value = "　÷　　　24　months　　　　　　　　　　=　　";
                worksheet.Range("F36:I37").Merge(true);
                worksheet.Range("F36:I37").Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Left);
                worksheet.Range("F36:I37").Style.Font.FontColor = XLColor.Red;
                worksheet.Range("F36:I37").Merge(true);
                worksheet.Range("F36:I37").Style.Fill.BackgroundColor = XLColor.White;
                worksheet.Range("F36:I37").Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Range("F36:I37").Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Range("I36:I37").Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Range("F36:F37").Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(36, 10).Value = "G";
                worksheet.Range("J36:J37").Merge(true);
                worksheet.Range("J36:J37").Style.Fill.BackgroundColor = XLColor.LightGray;
                worksheet.Range("J36:J37").Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Range("J36:J37").Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Range("J36:J37").Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Range("J36:J37").Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(36, 11).Value = moldjig_cost;
                worksheet.Range("K36:K37").Merge(true);
                worksheet.Range("K36:K37").Style.Fill.BackgroundColor = XLColor.LightGreen;
                worksheet.Range("K36:K37").Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Range("K36:K37").Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Range("K36:K37").Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Range("K36:K37").Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(36, 12).Value = moldjig_cost_percentage;
                worksheet.Range("L36:L37").Merge(true);
                worksheet.Range("L36:L37").Style.Fill.BackgroundColor = XLColor.LightGreen;
                worksheet.Range("L36:L37").Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Range("L36:L37").Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Range("L36:L37").Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Range("L36:L37").Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(38, 3).Value = "　Die";
                worksheet.Cell(38, 3).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Left);
                worksheet.Range("C38:D39").Merge(true);
                worksheet.Range("C38:D39").Style.Fill.BackgroundColor = XLColor.White;
                worksheet.Range("C38:D39").Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Range("C38:D39").Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Range("D38:D39").Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Range("C38:C39").Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(38, 5).Value = die_percentage;
                worksheet.Range("E38:E39").Merge(true);
                worksheet.Range("E38:E39").Style.Fill.BackgroundColor = XLColor.LightYellow;
                worksheet.Range("E38:E39").Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Range("E38:E39").Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Range("E38:E39").Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Range("E38:E39").Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(38, 6).Value = "　÷　　　24　months　　　　　　　　　　=　　";
                worksheet.Range("F38:I39").Merge(true);
                worksheet.Range("F38:I39").Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Left);
                worksheet.Range("F38:I39").Style.Font.FontColor = XLColor.Red;
                worksheet.Range("F38:I39").Merge(true);
                worksheet.Range("F38:I39").Style.Fill.BackgroundColor = XLColor.White;
                worksheet.Range("F38:I39").Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Range("F38:I39").Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Range("I38:I39").Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Range("F38:F39").Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(38, 10).Value = "G";
                worksheet.Range("J38:J39").Merge(true);
                worksheet.Range("J38:J39").Style.Fill.BackgroundColor = XLColor.LightGray;
                worksheet.Range("J38:J39").Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Range("J38:J39").Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Range("J38:J39").Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Range("J38:J39").Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(38, 11).Value = die_cost;
                worksheet.Range("K38:K39").Merge(true);
                worksheet.Range("K38:K39").Style.Fill.BackgroundColor = XLColor.LightGreen;
                worksheet.Range("K38:K39").Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Range("K38:K39").Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Range("K38:K39").Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Range("K38:K39").Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(38, 12).Value = die_cost_percentage;
                worksheet.Range("L38:L39").Merge(true);
                worksheet.Range("L38:L39").Style.Fill.BackgroundColor = XLColor.LightGreen;
                worksheet.Range("L38:L39").Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Range("L38:L39").Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Range("L38:L39").Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Range("L38:L39").Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(40, 2).Value = " Note :";
                worksheet.Cell(40, 2).Style.Font.FontSize = 18;
                worksheet.Cell(40, 2).Style.Font.SetUnderline();
                worksheet.Cell(40, 2).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Left);
                worksheet.Range("B40:F40").Merge(true);
                worksheet.Range("B40:F40").Style.Fill.BackgroundColor = XLColor.LightGray;
                worksheet.Range("B40:F40").Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Range("B40:F40").Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Range("B40:F40").Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(40, 7).Value = " Net Cost";
                worksheet.Cell(40, 7).Style.Font.FontSize = 16;
                worksheet.Cell(40, 7).Style.Font.SetUnderline();
                worksheet.Cell(40, 7).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Left);
                worksheet.Range("G40:H40").Merge(true);
                worksheet.Range("G40:H40").Style.Fill.BackgroundColor = XLColor.White;
                worksheet.Range("G40:H40").Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Range("G40:H40").Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(40, 9).Value = "Included tooling";
                worksheet.Cell(40, 9).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Right);
                worksheet.Range("I40:J40").Merge(true);
                worksheet.Range("I40:J40").Style.Fill.BackgroundColor = XLColor.White;
                worksheet.Range("I40:J40").Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Range("I40:J40").Style.Border.RightBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(40, 11).Value = net_included_tooling_cost;
                worksheet.Range("K40:K40").Merge(true);
                worksheet.Range("K40:K40").Style.Fill.BackgroundColor = XLColor.LightGreen;
                worksheet.Range("K40:K40").Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Range("K40:K40").Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Range("K40:K40").Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Range("K40:K40").Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(40, 12).Value = net_included_tooling_cost_percentage;
                worksheet.Range("L40:L40").Merge(true);
                worksheet.Range("L40:L40").Style.Fill.BackgroundColor = XLColor.LightGreen;
                worksheet.Range("L40:L40").Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Range("L40:L40").Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Range("L40:L40").Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Range("L40:L40").Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(41, 2).Value = note;
                worksheet.Cell(41, 2).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Left);
                worksheet.Range("B41:F41").Merge(true);
                worksheet.Range("B41:F41").Style.Fill.BackgroundColor = XLColor.LightYellow;
                worksheet.Range("B41:F41").Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Range("B41:F41").Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Range("B41:F41").Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(41, 7).Value = " ";
                worksheet.Cell(41, 7).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Left);
                worksheet.Range("G41:H41").Merge(true);
                worksheet.Range("G41:H41").Style.Fill.BackgroundColor = XLColor.White;
                worksheet.Range("G41:H41").Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Range("G41:H41").Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(41, 9).Value = "Exluded tooling";
                worksheet.Cell(41, 9).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Right);
                worksheet.Range("I41:J41").Merge(true);
                worksheet.Range("I41:J41").Style.Fill.BackgroundColor = XLColor.White;
                worksheet.Range("I41:J41").Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Range("I41:J41").Style.Border.RightBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(41, 11).Value = net_exclude_tooling_cost;
                worksheet.Range("K41:K41").Merge(true);
                worksheet.Range("K41:K41").Style.Fill.BackgroundColor = XLColor.LightGreen;
                worksheet.Range("K41:K41").Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Range("K41:K41").Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Range("K41:K41").Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Range("K41:K41").Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(41, 12).Value = net_exclude_tooling_cost_percentage;
                worksheet.Range("L41:L41").Merge(true);
                worksheet.Range("L41:L41").Style.Fill.BackgroundColor = XLColor.LightGreen;
                worksheet.Range("L41:L41").Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Range("L41:L41").Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Range("L41:L41").Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Range("L41:L41").Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(43, 2).Value = "3. Machine & Tooling List";
                worksheet.Cell(43, 2).Style.Font.FontSize = 18;
                worksheet.Cell(43, 2).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Left);
                worksheet.Range("B43:G43").Merge(true);

                worksheet.Cell(44, 2).Value = "No.";
                worksheet.Range("B44:C44").Merge(true);
                worksheet.Range("B44:C44").Style.Fill.BackgroundColor = XLColor.LightGray;
                worksheet.Range("B44:C44").Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Range("B44:C44").Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Range("B44:C44").Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(44, 4).Value = "Description";
                worksheet.Range("D44:F44").Merge(true);
                worksheet.Range("D44:F44").Style.Fill.BackgroundColor = XLColor.LightGray;
                worksheet.Range("D44:F44").Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Range("D44:F44").Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Range("D44:F44").Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(44, 7).Value = "Type of tool";
                worksheet.Cell(44, 7).Style.Fill.BackgroundColor = XLColor.LightGray;
                worksheet.Cell(44, 7).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(44, 7).Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(44, 7).Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(44, 8).Value = "Source";
                worksheet.Range("H44:H44").Style.Fill.BackgroundColor = XLColor.LightGray;
                worksheet.Range("H44:H44").Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Range("H44:H44").Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Range("H44:H44").Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(44, 9).Value = "Qty";
                worksheet.Range("I44:I44").Style.Fill.BackgroundColor = XLColor.LightGray;
                worksheet.Range("I44:I44").Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Range("I44:I44").Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Range("I44:I44").Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(44, 10).Value = "Unit";
                worksheet.Cell(44, 10).Style.Fill.BackgroundColor = XLColor.LightGray;
                worksheet.Cell(44, 10).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(44, 10).Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(44, 10).Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(44, 11).Value = "Price/unit";
                worksheet.Cell(44, 11).Style.Fill.BackgroundColor = XLColor.LightGray;
                worksheet.Cell(44, 11).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(44, 11).Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(44, 11).Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(44, 12).Value = "Amount(SGD)";
                worksheet.Cell(44, 12).Style.Fill.BackgroundColor = XLColor.LightGray;
                worksheet.Cell(44, 12).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(44, 12).Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(44, 12).Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                string[] description = { 
                    tooling_list_description_1, 
                    tooling_list_description_2, 
                    tooling_list_description_3, 
                    tooling_list_description_4, 
                    tooling_list_description_5, 
                    tooling_list_description_6, 
                    tooling_list_description_7, 
                    tooling_list_description_8, 
                    tooling_list_description_9, 
                    tooling_list_description_10, 
                    tooling_list_description_11, 
                    tooling_list_description_12, 
                    tooling_list_description_13, 
                    tooling_list_description_14, 
                    tooling_list_description_15 
                };
                string[] type = {
                    tooling_list_type_1,
                    tooling_list_type_2,
                    tooling_list_type_3,
                    tooling_list_type_4,
                    tooling_list_type_5,
                    tooling_list_type_6,
                    tooling_list_type_7,
                    tooling_list_type_8,
                    tooling_list_type_9,
                    tooling_list_type_10,
                    tooling_list_type_11,
                    tooling_list_type_12,
                    tooling_list_type_13,
                    tooling_list_type_14,
                    tooling_list_type_15
                };
                string[] source = {
                    tooling_list_source_1,
                    tooling_list_source_2,
                    tooling_list_source_3,
                    tooling_list_source_4,
                    tooling_list_source_5,
                    tooling_list_source_6,
                    tooling_list_source_7,
                    tooling_list_source_8,
                    tooling_list_source_9,
                    tooling_list_source_10,
                    tooling_list_source_11,
                    tooling_list_source_12,
                    tooling_list_source_13,
                    tooling_list_source_14,
                    tooling_list_source_15
                };
                string[] qty = {
                    tooling_list_qty_1,
                    tooling_list_qty_2,
                    tooling_list_qty_3,
                    tooling_list_qty_4,
                    tooling_list_qty_5,
                    tooling_list_qty_6,
                    tooling_list_qty_7,
                    tooling_list_qty_8,
                    tooling_list_qty_9,
                    tooling_list_qty_10,
                    tooling_list_qty_11,
                    tooling_list_qty_12,
                    tooling_list_qty_13,
                    tooling_list_qty_14,
                    tooling_list_qty_15
                };
                string[] unit = {
                    tooling_list_unit_1,
                    tooling_list_unit_2,
                    tooling_list_unit_3,
                    tooling_list_unit_4,
                    tooling_list_unit_5,
                    tooling_list_unit_6,
                    tooling_list_unit_7,
                    tooling_list_unit_8,
                    tooling_list_unit_9,
                    tooling_list_unit_10,
                    tooling_list_unit_11,
                    tooling_list_unit_12,
                    tooling_list_unit_13,
                    tooling_list_unit_14,
                    tooling_list_unit_15
                };
                string[] price = {
                    tooling_list_price_1,
                    tooling_list_price_2,
                    tooling_list_price_3,
                    tooling_list_price_4,
                    tooling_list_price_5,
                    tooling_list_price_6,
                    tooling_list_price_7,
                    tooling_list_price_8,
                    tooling_list_price_9,
                    tooling_list_price_10,
                    tooling_list_price_11,
                    tooling_list_price_12,
                    tooling_list_price_13,
                    tooling_list_price_14,
                    tooling_list_price_15
                };
                string[] amountjpy = {
                    tooling_list_amount_jpy_1,
                    tooling_list_amount_jpy_2,
                    tooling_list_amount_jpy_3,
                    tooling_list_amount_jpy_4,
                    tooling_list_amount_jpy_5,
                    tooling_list_amount_jpy_6,
                    tooling_list_amount_jpy_7,
                    tooling_list_amount_jpy_8,
                    tooling_list_amount_jpy_9,
                    tooling_list_amount_jpy_10,
                    tooling_list_amount_jpy_11,
                    tooling_list_amount_jpy_12,
                    tooling_list_amount_jpy_13,
                    tooling_list_amount_jpy_14,
                    tooling_list_amount_jpy_15
                };
                string[] amountusd = {
                    tooling_list_amount_usd_1,
                    tooling_list_amount_usd_2,
                    tooling_list_amount_usd_3,
                    tooling_list_amount_usd_4,
                    tooling_list_amount_usd_5,
                    tooling_list_amount_usd_6,
                    tooling_list_amount_usd_7,
                    tooling_list_amount_usd_8,
                    tooling_list_amount_usd_9,
                    tooling_list_amount_usd_10,
                    tooling_list_amount_usd_11,
                    tooling_list_amount_usd_12,
                    tooling_list_amount_usd_13,
                    tooling_list_amount_usd_14,
                    tooling_list_amount_usd_15
                };

                int c = 45;
                for (int i=0; i<15; i++)
                {
                    worksheet.Cell(c, 2).Value = i+1;
                    worksheet.Range("B" +c + ":C" +c).Merge(true);
                    worksheet.Range("B" +c + ":C" +c).Style.Fill.BackgroundColor = XLColor.White;
                    worksheet.Range("B" +c + ":C" +c).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                    worksheet.Range("B" +c + ":C" +c).Style.Border.RightBorder = XLBorderStyleValues.Thin;
                    worksheet.Range("B" +c + ":C" +c).Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                    worksheet.Cell(c, 4).Value = description[i];
                    worksheet.Range("D" +c + ":F" +c).Merge(true);
                    worksheet.Range("D" +c + ":F" +c).Style.Fill.BackgroundColor = XLColor.LightYellow;
                    worksheet.Range("D" +c + ":F" +c).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                    worksheet.Range("D" +c + ":F" +c).Style.Border.RightBorder = XLBorderStyleValues.Thin;
                    worksheet.Range("D" +c + ":F" +c).Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                    c++;

                    worksheet.Cell(45 + i, 7).Value = type[i];
                    worksheet.Cell(45 + i, 7).Style.Fill.BackgroundColor = XLColor.LightYellow;
                    worksheet.Cell(45 + i, 7).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                    worksheet.Cell(45 + i, 7).Style.Border.RightBorder = XLBorderStyleValues.Thin;
                    worksheet.Cell(45 + i, 7).Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                    worksheet.Cell(45 + i, 8).Value = source[i];
                    worksheet.Cell(45 + i, 8).Style.Fill.BackgroundColor = XLColor.LightYellow;
                    worksheet.Cell(45 + i, 8).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                    worksheet.Cell(45 + i, 8).Style.Border.RightBorder = XLBorderStyleValues.Thin;
                    worksheet.Cell(45 + i, 8).Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                    worksheet.Cell(45 + i, 9).Value = qty[i];
                    worksheet.Cell(45 + i, 9).Style.Fill.BackgroundColor = XLColor.LightYellow;
                    worksheet.Cell(45 + i, 9).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                    worksheet.Cell(45 + i, 9).Style.Border.RightBorder = XLBorderStyleValues.Thin;
                    worksheet.Cell(45 + i, 9).Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                    worksheet.Cell(45 + i, 10).Value = unit[i];
                    worksheet.Cell(45 + i, 10).Style.Fill.BackgroundColor = XLColor.LightYellow;
                    worksheet.Cell(45 + i, 10).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                    worksheet.Cell(45 + i, 10).Style.Border.RightBorder = XLBorderStyleValues.Thin;
                    worksheet.Cell(45 + i, 10).Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                    worksheet.Cell(45 + i, 11).Value = price[i];
                    worksheet.Cell(45 + i, 11).Style.Fill.BackgroundColor = XLColor.LightYellow;
                    worksheet.Cell(45 + i, 11).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                    worksheet.Cell(45 + i, 11).Style.Border.RightBorder = XLBorderStyleValues.Thin;
                    worksheet.Cell(45 + i, 11).Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                    worksheet.Cell(45 + i, 12).Value = amountusd[i];
                    worksheet.Cell(45 + i, 12).Style.Fill.BackgroundColor = XLColor.LightGreen;
                    worksheet.Cell(45 + i, 12).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                    worksheet.Cell(45 + i, 12).Style.Border.RightBorder = XLBorderStyleValues.Thin;
                    worksheet.Cell(45 + i, 12).Style.Border.LeftBorder = XLBorderStyleValues.Thin;
                }

                worksheet.Cell(60, 11).Value = "Total";
                worksheet.Cell(60, 11).Style.Fill.BackgroundColor = XLColor.LightGray;
                worksheet.Cell(60, 11).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(60, 11).Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(60, 11).Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(60, 12).Value = tooling_list_total_amount_usd;
                worksheet.Cell(60, 12).Style.Fill.BackgroundColor = XLColor.LightGreen;
                worksheet.Cell(60, 12).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(60, 12).Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Cell(60, 12).Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(62, 2).Value = "　Propose Price to NFS";
                worksheet.Cell(62, 2).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Left);
                worksheet.Range("B62:F62").Merge(true);
                worksheet.Range("B62:F62").Style.Fill.BackgroundColor = XLColor.White;
                worksheet.Range("B62:F62").Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Range("B62:F62").Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Range("B62:F62").Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Range("B62:F62").Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(62, 7).Value = "　";
                worksheet.Range("G62:H62").Merge(true);
                worksheet.Range("G62:H62").Style.Fill.BackgroundColor = XLColor.White;
                worksheet.Range("G62:H62").Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Range("G62:H62").Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Range("G62:H62").Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Range("G62:H62").Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(63, 2).Value = "　Tooling Price By NFS";
                worksheet.Cell(63, 2).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Left);
                worksheet.Range("B63:F63").Merge(true);
                worksheet.Range("B63:F63").Style.Fill.BackgroundColor = XLColor.White;
                worksheet.Range("B63:F63").Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Range("B63:F63").Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Range("B63:F63").Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Range("B63:F63").Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(63, 7).Value = tooling_list_total_amount_usd;
                worksheet.Range("G63:H63").Merge(true);
                worksheet.Range("G63:H63").Style.Fill.BackgroundColor = XLColor.White;
                worksheet.Range("G63:H63").Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Range("G63:H63").Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Range("G63:H63").Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Range("G63:H63").Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(64, 2).Value = "　Remark :";
                worksheet.Cell(64, 2).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Left);
                worksheet.Cell(64, 2).Style.Alignment.SetVertical(XLAlignmentVerticalValues.Top);
                worksheet.Range("B64:H67").Merge(true);
                worksheet.Range("B64:H67").Style.Fill.BackgroundColor = XLColor.White;
                worksheet.Range("B64:H67").Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Range("B64:H67").Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Range("B64:H67").Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Range("B64:H67").Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(62, 9).Value = "　";
                worksheet.Cell(62, 9).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Left);
                worksheet.Range("I62:L62").Merge(true);
                worksheet.Range("I62:L62").Style.Fill.BackgroundColor = XLColor.White;
                worksheet.Range("I62:L62").Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Range("I62:L62").Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Range("I62:L62").Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(63, 9).Value = "　Final Unit Price　　　　：_________________________________";
                worksheet.Cell(63, 9).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Left);
                worksheet.Range("I63:L63").Merge(true);
                worksheet.Range("I63:L63").Style.Fill.BackgroundColor = XLColor.White;
                worksheet.Range("I63:L63").Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Range("I63:L63").Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(64, 9).Value = "　Tooling Cost Paid by NFS ：_________________________________";
                worksheet.Cell(64, 9).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Left);
                worksheet.Range("I64:L64").Merge(true);
                worksheet.Range("I64:L64").Style.Fill.BackgroundColor = XLColor.White;
                worksheet.Range("I64:L64").Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Range("I64:L64").Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(65, 9).Value = "　Sign off of acceptance 　：_________________________________";
                worksheet.Cell(65, 9).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Left);
                worksheet.Range("I65:L65").Merge(true);
                worksheet.Range("I65:L65").Style.Fill.BackgroundColor = XLColor.White;
                worksheet.Range("I65:L65").Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Range("I65:L65").Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(66, 9).Value = "　By & Date　　　　　　：_________________________________";
                worksheet.Cell(66, 9).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Left);
                worksheet.Range("I66:L66").Merge(true);
                worksheet.Range("I66:L66").Style.Fill.BackgroundColor = XLColor.White;
                worksheet.Range("I66:L66").Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Range("I66:L66").Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                worksheet.Cell(67, 9).Value = "　";
                worksheet.Cell(67, 9).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Left);
                worksheet.Range("I67:L67").Merge(true);
                worksheet.Range("I67:L67").Style.Fill.BackgroundColor = XLColor.White;
                worksheet.Range("I67:L67").Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                worksheet.Range("I67:L67").Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Range("I67:L67").Style.Border.LeftBorder = XLBorderStyleValues.Thin;


                worksheet.Cell(16, 14).Value = "Raw Material Cost";
                worksheet.Cell(16, 14).Style.Font.FontSize = 13;
                worksheet.Cell(16, 14).Style.Font.SetUnderline();
                worksheet.Cell(16, 15).Value = "Unit Cost";
                worksheet.Cell(16, 15).Style.Font.FontSize = 13;
                worksheet.Cell(16, 15).Style.Font.SetUnderline();
                worksheet.Cell(16, 16).Value = "%";
                worksheet.Cell(16, 16).Style.Font.FontSize = 13;
                worksheet.Cell(16, 16).Style.Font.SetUnderline();


                worksheet.Cell(17, 14).Value = "Direct Raw Material";
                worksheet.Cell(17, 15).Value = direct_raw_material;
                worksheet.Cell(17, 16).Value = direct_raw_material_p + " %";

                worksheet.Cell(18, 14).Value = "Sub-Material";
                worksheet.Cell(18, 15).Value = sub_material;
                worksheet.Cell(18, 16).Value = sub_material_p + " %";

                worksheet.Cell(19, 14).Value = "Sub Total";
                worksheet.Cell(19, 14).Style.Fill.BackgroundColor = XLColor.LightGray;
                worksheet.Cell(19, 15).Value = raw_material_cost_sub_total;
                worksheet.Cell(19, 15).Style.Fill.BackgroundColor = XLColor.LightGray;
                worksheet.Cell(19, 16).Value = raw_material_cost_sub_total_p + " %";
                worksheet.Cell(19, 16).Style.Fill.BackgroundColor = XLColor.LightGray;


                worksheet.Cell(21, 14).Value = "Breakdown of Process Cost";
                worksheet.Cell(21, 14).Style.Font.FontSize = 13;
                worksheet.Cell(21, 14).Style.Font.SetUnderline();
                worksheet.Cell(21, 15).Value = "Unit Cost";
                worksheet.Cell(21, 15).Style.Font.FontSize = 13;
                worksheet.Cell(21, 15).Style.Font.SetUnderline();
                worksheet.Cell(21, 16).Value = "%";
                worksheet.Cell(21, 16).Style.Font.FontSize = 13;
                worksheet.Cell(21, 16).Style.Font.SetUnderline();

                worksheet.Cell(22, 14).Value = "Labor Cost";
                worksheet.Cell(22, 15).Value = labor_cost;
                worksheet.Cell(22, 16).Value = labor_cost_p + " %";

                worksheet.Cell(23, 14).Value = "Machine Cost";
                worksheet.Cell(23, 15).Value = machine_cost;
                worksheet.Cell(23, 16).Value = machine_cost_p + " %";

                worksheet.Cell(24, 14).Value = "Overhead Cost";
                worksheet.Cell(24, 15).Value = overhead_cost;
                worksheet.Cell(24, 16).Value = overhead_cost_p + " %";

                worksheet.Cell(25, 14).Value = "Sub Total";
                worksheet.Cell(25, 14).Style.Fill.BackgroundColor = XLColor.LightGray;
                worksheet.Cell(25, 15).Value = process_cost_sub_total;
                worksheet.Cell(25, 15).Style.Fill.BackgroundColor = XLColor.LightGray;
                worksheet.Cell(25, 16).Value = process_cost_sub_total_p + " %";
                worksheet.Cell(25, 16).Style.Fill.BackgroundColor = XLColor.LightGray;


                worksheet.Cell(27, 14).Value = "Other Fixed Cost";
                worksheet.Cell(27, 14).Style.Font.FontSize = 13;
                worksheet.Cell(27, 14).Style.Font.SetUnderline();
                worksheet.Cell(27, 15).Value = "Unit Cost";
                worksheet.Cell(27, 15).Style.Font.FontSize = 13;
                worksheet.Cell(27, 15).Style.Font.SetUnderline();
                worksheet.Cell(27, 16).Value = "%";
                worksheet.Cell(27, 16).Style.Font.FontSize = 13;
                worksheet.Cell(27, 16).Style.Font.SetUnderline();

                worksheet.Cell(28, 14).Value = "Defectives";
                worksheet.Cell(28, 15).Value = defectives;
                worksheet.Cell(28, 16).Value = defectives_p + " %";

                worksheet.Cell(29, 14).Value = "Admin A Engin, QC";
                worksheet.Cell(29, 15).Value = admin_engin_qc;
                worksheet.Cell(29, 16).Value = admin_engin_qc_p + " %";

                worksheet.Cell(30, 14).Value = "Tooling Cost";
                worksheet.Cell(30, 15).Value = tooling_cost;
                worksheet.Cell(30, 16).Value = tooling_cost_p + " %";

                worksheet.Cell(31, 14).Value = "Process Margin Adjust";
                worksheet.Cell(31, 15).Value = process_margin_adjust;
                worksheet.Cell(31, 16).Value = process_margin_adjust_p + " %";

                worksheet.Cell(32, 14).Value = "Sub Total";
                worksheet.Cell(32, 14).Style.Fill.BackgroundColor = XLColor.LightGray;
                worksheet.Cell(32, 15).Value = other_fixed_cost_sub_total;
                worksheet.Cell(32, 15).Style.Fill.BackgroundColor = XLColor.LightGray;
                worksheet.Cell(32, 16).Value = other_fixed_cost_sub_total_p + " %";
                worksheet.Cell(32, 16).Style.Fill.BackgroundColor = XLColor.LightGray;

                worksheet.Cell(34, 14).Value = "Grand Total Cost";
                worksheet.Cell(34, 14).Style.Fill.BackgroundColor = XLColor.LightGray;
                worksheet.Cell(34, 15).Value = grand_total_cost;
                worksheet.Cell(34, 15).Style.Fill.BackgroundColor = XLColor.LightGray;
                worksheet.Cell(34, 16).Value = grand_total_cost_p + " %";
                worksheet.Cell(34, 16).Style.Fill.BackgroundColor = XLColor.LightGray;

                using (var stream = new MemoryStream())
                {
                    workbook.SaveAs(stream);
                    var content = stream.ToArray();
                    return File(
                        content,
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        "Cost.xlsx");

                }
            }
        }

        public async void Delete(
            Cost model,
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

                    var action = "api/data/delete-cost-by-id/" + Id;

                    HttpResponseMessage res = await client.PostAsync(action, content).ConfigureAwait(false);

                    res.EnsureSuccessStatusCode();
                    if (res.IsSuccessStatusCode)
                    {

                        var result = res.Content.ReadAsStringAsync().Result;
                    }
                }
            }
        }

        public async Task<ActionResult<Cost>> Save(Cost model)
        {

            HttpClient client = _api.Initial();

            var content = new StringContent(JsonConvert.SerializeObject(model),Encoding.UTF8, "application/json");

            var action = "api/data/add-cost"; 
            if (model.CostId > 0)
            {
                action = "api/data/update-cost-by-id/" + model.CostId;
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
