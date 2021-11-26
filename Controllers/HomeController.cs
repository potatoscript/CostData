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
            ViewBag.sub_material_percentage = 0;
            ViewBag.sub_material_cost = 0;
            ViewBag.sub_material_cost_percentage = 0;
            ViewBag.direct_process_cost = 0;
            ViewBag.direct_process_cost_percentage = 0;
            ViewBag.total_direct_cost = 0;
            ViewBag.total_direct_cost_percentage = 0;
            ViewBag.defective_percentage = 0;
            ViewBag.defective_cost = 0;
            ViewBag.defective_cost_percentage = 0;
            ViewBag.indirect_percentage = 0;
            ViewBag.indirect_cost = 0;
            ViewBag.indirect_cost_percentage = 0;
            ViewBag.packing_material_percentage = 0;
            ViewBag.special_package_cost = 0;
            ViewBag.packing_material_cost = 0;
            ViewBag.packing_material_cost_percentage = 0;
            ViewBag.administration_percentage = 0;
            ViewBag.administration_cost = 0;
            ViewBag.administration_cost_percentage = 0;
            ViewBag.plant_retail_percentage = 0;
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
                    //listRubber.data.Clear();

                    //listRubber.data.Add(new Rubber
                    //{
                    //    material_name = "-",
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
                    Value = x.material_name+","+x.price_kg + "," +x.mixing_process_cost + "," +x.weight_g + "," +x.yield_rate

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
            string expired_by
        ){
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Cost");

                
                worksheet.Columns().Style.Font.FontSize = 13;
                worksheet.Columns().Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
                //worksheet.Columns().AdjustToContents();
                worksheet.ColumnWidth = 13;

                worksheet.Columns("A:A").Width = 3;
                worksheet.Columns("B:B").Width = 3;
                worksheet.Columns("C:C").Width = 3;

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

                worksheet.Cell(10, 7).Value = "EUD";
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

                worksheet.Cell(10, 12).Value = "EUD";
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
