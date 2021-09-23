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
            ViewBag.tooling_list_amount_usd_1 = 0;

            ViewBag.tooling_list_description_2 = "-";
            ViewBag.tooling_list_type_2 = "-";
            ViewBag.tooling_list_source_2 = "-";
            ViewBag.tooling_list_qty_2 = 0;
            ViewBag.tooling_list_unit_2 = "-";
            ViewBag.tooling_list_price_2 = 0;
            ViewBag.tooling_list_amount_jpy_2 = 0;
            ViewBag.tooling_list_amount_usd_2 = 0;

            ViewBag.tooling_list_description_3 = "-";
            ViewBag.tooling_list_type_3 = "-";
            ViewBag.tooling_list_source_3 = "-";
            ViewBag.tooling_list_qty_3 = 0;
            ViewBag.tooling_list_unit_3 = "-";
            ViewBag.tooling_list_price_3 = 0;
            ViewBag.tooling_list_amount_jpy_3 = 0;
            ViewBag.tooling_list_amount_usd_3 = 0;

            ViewBag.tooling_list_description_4 = "-";
            ViewBag.tooling_list_type_4 = "-";
            ViewBag.tooling_list_source_4 = "-";
            ViewBag.tooling_list_qty_4 = 0;
            ViewBag.tooling_list_unit_4 = "-";
            ViewBag.tooling_list_price_4 = 0;
            ViewBag.tooling_list_amount_jpy_4 = 0;
            ViewBag.tooling_list_amount_usd_4 = 0;

            ViewBag.tooling_list_description_5 = "-";
            ViewBag.tooling_list_type_5 = "-";
            ViewBag.tooling_list_source_5 = "-";
            ViewBag.tooling_list_qty_5 = 0;
            ViewBag.tooling_list_unit_5 = "-";
            ViewBag.tooling_list_price_5 = 0;
            ViewBag.tooling_list_amount_jpy_5 = 0;
            ViewBag.tooling_list_amount_usd_5 = 0;

            ViewBag.tooling_list_description_6 = "-";
            ViewBag.tooling_list_type_6 = "-";
            ViewBag.tooling_list_source_6 = "-";
            ViewBag.tooling_list_qty_6 = 0;
            ViewBag.tooling_list_unit_6 = "-";
            ViewBag.tooling_list_price_6 = 0;
            ViewBag.tooling_list_amount_jpy_6 = 0;
            ViewBag.tooling_list_amount_usd_6 = 0;

            ViewBag.tooling_list_description_7 = "-";
            ViewBag.tooling_list_type_7 = "-";
            ViewBag.tooling_list_source_7 = "-";
            ViewBag.tooling_list_qty_7 = 0;
            ViewBag.tooling_list_unit_7 = "-";
            ViewBag.tooling_list_price_7 = 0;
            ViewBag.tooling_list_amount_jpy_7 = 0;
            ViewBag.tooling_list_amount_usd_7 = 0;

            ViewBag.tooling_list_description_8 = "-";
            ViewBag.tooling_list_type_8 = "-";
            ViewBag.tooling_list_source_8 = "-";
            ViewBag.tooling_list_qty_8 = 0;
            ViewBag.tooling_list_unit_8 = "-";
            ViewBag.tooling_list_price_8 = 0;
            ViewBag.tooling_list_amount_jpy_8 = 0;
            ViewBag.tooling_list_amount_usd_8 = 0;

            ViewBag.tooling_list_description_9 = "-";
            ViewBag.tooling_list_type_9 = "-";
            ViewBag.tooling_list_source_9 = "-";
            ViewBag.tooling_list_qty_9 = 0;
            ViewBag.tooling_list_unit_9 = "-";
            ViewBag.tooling_list_price_9 = 0;
            ViewBag.tooling_list_amount_jpy_9 = 0;
            ViewBag.tooling_list_amount_usd_9 = 0;

            ViewBag.tooling_list_description_10 = "-";
            ViewBag.tooling_list_type_10 = "-";
            ViewBag.tooling_list_source_10 = "-";
            ViewBag.tooling_list_qty_10 = 0;
            ViewBag.tooling_list_unit_10 = "-";
            ViewBag.tooling_list_price_10 = 0;
            ViewBag.tooling_list_amount_jpy_10 = 0;
            ViewBag.tooling_list_amount_usd_10 = 0;

            ViewBag.tooling_list_description_11 = "-";
            ViewBag.tooling_list_type_11 = "-";
            ViewBag.tooling_list_source_11 = "-";
            ViewBag.tooling_list_qty_11 = 0;
            ViewBag.tooling_list_unit_11 = "-";
            ViewBag.tooling_list_price_11 = 0;
            ViewBag.tooling_list_amount_jpy_11 = 0;
            ViewBag.tooling_list_amount_usd_11 = 0;

            ViewBag.tooling_list_description_12 = "-";
            ViewBag.tooling_list_type_12 = "-";
            ViewBag.tooling_list_source_12 = "-";
            ViewBag.tooling_list_qty_12 = 0;
            ViewBag.tooling_list_unit_12 = "-";
            ViewBag.tooling_list_price_12 = 0;
            ViewBag.tooling_list_amount_jpy_12 = 0;
            ViewBag.tooling_list_amount_usd_12 = 0;

            ViewBag.tooling_list_description_13 = "-";
            ViewBag.tooling_list_type_13 = "-";
            ViewBag.tooling_list_source_13 = "-";
            ViewBag.tooling_list_qty_13 = 0;
            ViewBag.tooling_list_unit_13 = "-";
            ViewBag.tooling_list_price_13 = 0;
            ViewBag.tooling_list_amount_jpy_13 = 0;
            ViewBag.tooling_list_amount_usd_13 = 0;

            ViewBag.tooling_list_description_14 = "-";
            ViewBag.tooling_list_type_14 = "-";
            ViewBag.tooling_list_source_14 = "-";
            ViewBag.tooling_list_qty_14 = 0;
            ViewBag.tooling_list_unit_14 = "-";
            ViewBag.tooling_list_price_14 = 0;
            ViewBag.tooling_list_amount_jpy_14 = 0;
            ViewBag.tooling_list_amount_usd_14 = 0;

            ViewBag.tooling_list_description_15 = "-";
            ViewBag.tooling_list_type_15 = "-";
            ViewBag.tooling_list_source_15 = "-";
            ViewBag.tooling_list_qty_15 = 0;
            ViewBag.tooling_list_unit_15 = "-";
            ViewBag.tooling_list_price_15 = 0;
            ViewBag.tooling_list_amount_jpy_15 = 0;
            ViewBag.tooling_list_amount_usd_15 = 0;

            ViewBag.tooling_list_total_amount_usd = 0;


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
                ViewBag.customer = costdata.customer;
                ViewBag.parts_code = costdata.parts_code;
                ViewBag.item = costdata.item;
                ViewBag.product =costdata.product;
                ViewBag.product_type = costdata.product_type;
                ViewBag.size = costdata.size;
                ViewBag.business_type = costdata.business_type;
                ViewBag.qty_month = costdata.qty_month;
                ViewBag.exchange_rate = costdata.exchange_rate;
                ViewBag.target_price_bht = costdata.target_price_bht;
                ViewBag.target_price_export = costdata.target_price_export;
                ViewBag.production_qty_day = costdata.production_qty_day;
                ViewBag.working_day = costdata.working_day;

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
                ViewBag.tooling_list_amount_jpy_1 = costdata.tooling_list_amount_jpy_1;
                ViewBag.tooling_list_amount_usd_1 = costdata.tooling_list_amount_usd_1;

                ViewBag.tooling_list_description_2 = costdata.tooling_list_description_2;
                ViewBag.tooling_list_type_2 = costdata.tooling_list_type_2;
                ViewBag.tooling_list_source_2 = costdata.tooling_list_source_2;
                ViewBag.tooling_list_qty_2 = costdata.tooling_list_qty_2;
                ViewBag.tooling_list_unit_2 = costdata.tooling_list_unit_2;
                ViewBag.tooling_list_price_2 = costdata.tooling_list_price_2;
                ViewBag.tooling_list_amount_jpy_2 = costdata.tooling_list_amount_jpy_2;
                ViewBag.tooling_list_amount_usd_2 = costdata.tooling_list_amount_usd_2;

                ViewBag.tooling_list_description_3 = costdata.tooling_list_description_3;
                ViewBag.tooling_list_type_3 = costdata.tooling_list_type_3;
                ViewBag.tooling_list_source_3 = costdata.tooling_list_source_3;
                ViewBag.tooling_list_qty_3 = costdata.tooling_list_qty_3;
                ViewBag.tooling_list_unit_3 = costdata.tooling_list_unit_3;
                ViewBag.tooling_list_price_3 = costdata.tooling_list_price_3;
                ViewBag.tooling_list_amount_jpy_3 = costdata.tooling_list_amount_jpy_3;
                ViewBag.tooling_list_amount_usd_3 = costdata.tooling_list_amount_usd_3;

                ViewBag.tooling_list_description_4 = costdata.tooling_list_description_4;
                ViewBag.tooling_list_type_4 = costdata.tooling_list_type_4;
                ViewBag.tooling_list_source_4 = costdata.tooling_list_source_4;
                ViewBag.tooling_list_qty_4 = costdata.tooling_list_qty_4;
                ViewBag.tooling_list_unit_4 = costdata.tooling_list_unit_4;
                ViewBag.tooling_list_price_4 = costdata.tooling_list_price_4;
                ViewBag.tooling_list_amount_jpy_4 = costdata.tooling_list_amount_jpy_4;
                ViewBag.tooling_list_amount_usd_4 = costdata.tooling_list_amount_usd_4;

                ViewBag.tooling_list_description_5 = costdata.tooling_list_description_5;
                ViewBag.tooling_list_type_5 = costdata.tooling_list_type_5;
                ViewBag.tooling_list_source_5 = costdata.tooling_list_source_5;
                ViewBag.tooling_list_qty_5 = costdata.tooling_list_qty_5;
                ViewBag.tooling_list_unit_5 = costdata.tooling_list_unit_5;
                ViewBag.tooling_list_price_5 = costdata.tooling_list_price_5;
                ViewBag.tooling_list_amount_jpy_5 = costdata.tooling_list_amount_jpy_5;
                ViewBag.tooling_list_amount_usd_5 = costdata.tooling_list_amount_usd_5;

                ViewBag.tooling_list_description_6 = costdata.tooling_list_description_6;
                ViewBag.tooling_list_type_6 = costdata.tooling_list_type_6;
                ViewBag.tooling_list_source_6 = costdata.tooling_list_source_6;
                ViewBag.tooling_list_qty_6 = costdata.tooling_list_qty_6;
                ViewBag.tooling_list_unit_6 = costdata.tooling_list_unit_6;
                ViewBag.tooling_list_price_6 = costdata.tooling_list_price_6;
                ViewBag.tooling_list_amount_jpy_6 = costdata.tooling_list_amount_jpy_6;
                ViewBag.tooling_list_amount_usd_6 = costdata.tooling_list_amount_usd_6;

                ViewBag.tooling_list_description_7 = costdata.tooling_list_description_7;
                ViewBag.tooling_list_type_7 = costdata.tooling_list_type_7;
                ViewBag.tooling_list_source_7 = costdata.tooling_list_source_7;
                ViewBag.tooling_list_qty_7 = costdata.tooling_list_qty_7;
                ViewBag.tooling_list_unit_7 = costdata.tooling_list_unit_7;
                ViewBag.tooling_list_price_7 = costdata.tooling_list_price_7;
                ViewBag.tooling_list_amount_jpy_7 = costdata.tooling_list_amount_jpy_7;
                ViewBag.tooling_list_amount_usd_7 = costdata.tooling_list_amount_usd_7;

                ViewBag.tooling_list_description_8 = costdata.tooling_list_description_8;
                ViewBag.tooling_list_type_8 = costdata.tooling_list_type_8;
                ViewBag.tooling_list_source_8 = costdata.tooling_list_source_8;
                ViewBag.tooling_list_qty_8 = costdata.tooling_list_qty_8;
                ViewBag.tooling_list_unit_8 = costdata.tooling_list_unit_8;
                ViewBag.tooling_list_price_8 = costdata.tooling_list_price_8;
                ViewBag.tooling_list_amount_jpy_8 = costdata.tooling_list_amount_jpy_8;
                ViewBag.tooling_list_amount_usd_8 = costdata.tooling_list_amount_usd_8;

                ViewBag.tooling_list_description_9 = costdata.tooling_list_description_9;
                ViewBag.tooling_list_type_9 = costdata.tooling_list_type_9;
                ViewBag.tooling_list_source_9 = costdata.tooling_list_source_9;
                ViewBag.tooling_list_qty_9 = costdata.tooling_list_qty_9;
                ViewBag.tooling_list_unit_9 = costdata.tooling_list_unit_9;
                ViewBag.tooling_list_price_9 = costdata.tooling_list_price_9;
                ViewBag.tooling_list_amount_jpy_9 = costdata.tooling_list_amount_jpy_9;
                ViewBag.tooling_list_amount_usd_9 = costdata.tooling_list_amount_usd_9;

                ViewBag.tooling_list_description_10 = costdata.tooling_list_description_10;
                ViewBag.tooling_list_type_10 = costdata.tooling_list_type_10;
                ViewBag.tooling_list_source_10 = costdata.tooling_list_source_10;
                ViewBag.tooling_list_qty_10 = costdata.tooling_list_qty_10;
                ViewBag.tooling_list_unit_10 = costdata.tooling_list_unit_10;
                ViewBag.tooling_list_price_10 = costdata.tooling_list_price_10;
                ViewBag.tooling_list_amount_jpy_10 = costdata.tooling_list_amount_jpy_10;
                ViewBag.tooling_list_amount_usd_10 = costdata.tooling_list_amount_usd_10;

                ViewBag.tooling_list_description_11 = costdata.tooling_list_description_11;
                ViewBag.tooling_list_type_11 = costdata.tooling_list_type_11;
                ViewBag.tooling_list_source_11 = costdata.tooling_list_source_11;
                ViewBag.tooling_list_qty_11 = costdata.tooling_list_qty_11;
                ViewBag.tooling_list_unit_11 = costdata.tooling_list_unit_11;
                ViewBag.tooling_list_price_11 = costdata.tooling_list_price_11;
                ViewBag.tooling_list_amount_jpy_11 = costdata.tooling_list_amount_jpy_11;
                ViewBag.tooling_list_amount_usd_11 = costdata.tooling_list_amount_usd_11;

                ViewBag.tooling_list_description_12 = costdata.tooling_list_description_12;
                ViewBag.tooling_list_type_12 = costdata.tooling_list_type_12;
                ViewBag.tooling_list_source_12 = costdata.tooling_list_source_12;
                ViewBag.tooling_list_qty_12 = costdata.tooling_list_qty_12;
                ViewBag.tooling_list_unit_12 = costdata.tooling_list_unit_12;
                ViewBag.tooling_list_price_12 = costdata.tooling_list_price_12;
                ViewBag.tooling_list_amount_jpy_12 = costdata.tooling_list_amount_jpy_12;
                ViewBag.tooling_list_amount_usd_12 = costdata.tooling_list_amount_usd_12;

                ViewBag.tooling_list_description_13 = costdata.tooling_list_description_13;
                ViewBag.tooling_list_type_13 = costdata.tooling_list_type_13;
                ViewBag.tooling_list_source_13 = costdata.tooling_list_source_13;
                ViewBag.tooling_list_qty_13 = costdata.tooling_list_qty_13;
                ViewBag.tooling_list_unit_13 = costdata.tooling_list_unit_13;
                ViewBag.tooling_list_price_13 = costdata.tooling_list_price_13;
                ViewBag.tooling_list_amount_jpy_13 = costdata.tooling_list_amount_jpy_13;
                ViewBag.tooling_list_amount_usd_13 = costdata.tooling_list_amount_usd_13;

                ViewBag.tooling_list_description_14 = costdata.tooling_list_description_14;
                ViewBag.tooling_list_type_14 = costdata.tooling_list_type_14;
                ViewBag.tooling_list_source_14 = costdata.tooling_list_source_14;
                ViewBag.tooling_list_qty_14 = costdata.tooling_list_qty_14;
                ViewBag.tooling_list_unit_14 = costdata.tooling_list_unit_14;
                ViewBag.tooling_list_price_14 = costdata.tooling_list_price_14;
                ViewBag.tooling_list_amount_jpy_14 = costdata.tooling_list_amount_jpy_14;
                ViewBag.tooling_list_amount_usd_14 = costdata.tooling_list_amount_usd_14;

                ViewBag.tooling_list_description_15 = costdata.tooling_list_description_15;
                ViewBag.tooling_list_type_15 = costdata.tooling_list_type_15;
                ViewBag.tooling_list_source_15 = costdata.tooling_list_source_15;
                ViewBag.tooling_list_qty_15 = costdata.tooling_list_qty_15;
                ViewBag.tooling_list_unit_15 = costdata.tooling_list_unit_15;
                ViewBag.tooling_list_price_15 = costdata.tooling_list_price_15;
                ViewBag.tooling_list_amount_jpy_15 = costdata.tooling_list_amount_jpy_15;
                ViewBag.tooling_list_amount_usd_15 = costdata.tooling_list_amount_usd_15;

                ViewBag.tooling_list_total_amount_usd = 0;

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
