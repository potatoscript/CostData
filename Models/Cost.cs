using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Threading.Tasks;

namespace CostNag.Models
{
    public class Cost
    {


        public int CostId { get; set; } = 0;

        public string plant { get; set; } = "-";

        public string item_spec { get; set; } = "-";

        public string issue_date { get; set; } = DateTime.Now.AddDays(0).ToString("dd-MM-yyyy");

        public string section { get; set; } = "-";

        public string doc_no { get; set; } = "-";

        public string wr_no { get; set; } = "-";

        public string sales { get; set; } = "-";

        public string revision_no { get; set; } = "-";

        public string checked_date { get; set; } = DateTime.Now.AddDays(0).ToString("dd-MM-yyyy");

        public string approved_by { get; set; } = DateTime.Now.AddDays(0).ToString("dd-MM-yyyy");

        public string expired_by { get; set; } = DateTime.Now.AddDays(180).ToString("dd-MM-yyyy");

        public string customer { get; set; } = "-";

        public string parts_code { get; set; } = "-";

        public string item { get; set; } = "-";

        public string product { get; set; } = "-";

        public string product_type { get; set; } = "-";

        public string size { get; set; } = "-";

        
        public double item_id { get; set; } = 0;

        
        public double item_od { get; set; } = 0;

        
        public double item_w { get; set; } = 0;

        
        public double item_w2 { get; set; } = 0;

        public string business_type { get; set; } = "-";

        public int qty_month { get; set; } = 0;

        public double exchange_rate_jpy { get; set; } = 84.13;

        
        public double exchange_rate_usd { get; set; } = 0.74;

        
        public double exchange_rate_eud { get; set; } = 0.65;


        
        public double target_price_sgd { get; set; } = 0;

        
        public double target_price_wr_sgd { get; set; } = 0;

        
        public double target_price_usd { get; set; } = 0;

        
        public double target_price_wr_usd { get; set; } = 0;

        
        public double target_price_eud { get; set; } = 0;

        
        public double target_price_wr_eud { get; set; } = 0;

        public double production_qty_day { get; set; } = 0;

        public double working_day { get; set; } = 21;

        public string rubber_material_name { get; set; } = "-";


        public double rubber_database_price_current { get; set; } = 0;

        
        public double rubber_database_price_new { get; set; } = 0;

        
        public double rubber_price_kg { get; set; } = 0;

        
        public double rubber_mixing_process_cost { get; set; } = 0;

        
        public double rubber_weight_g { get; set; } = 0;

        
        public double rubber_weight_kg { get; set; } = 0;

        
        public double rubber_yield_rate { get; set; } = 0;

        
        public double rubber_weight_kg_yieldrate { get; set; } = 0;

        
        public double rubber_cost_sgd { get; set; } = 0;

        
        public double rubber_percentage_target_price { get; set; } = 0;


        public string rubber_material_name2 { get; set; } = "-";

        
        public double rubber_database_price_current2 { get; set; } = 0;

        
        public double rubber_database_price_new2 { get; set; } = 0;

        
        public double rubber_price_kg2 { get; set; } = 0;

        
        public double rubber_mixing_process_cost2 { get; set; } = 0;

        
        public double rubber_weight_g2 { get; set; } = 0;

        
        public double rubber_weight_kg2 { get; set; } = 0;

        
        public double rubber_yield_rate2 { get; set; } = 0;

        
        public double rubber_weight_kg_yieldrate2 { get; set; } = 0;

        
        public double rubber_cost_sgd2 { get; set; } = 0;

        
        public double rubber_percentage_target_price2 { get; set; } = 0;

        /// <summary>
        /// /////////////Material Inhouse
        /// </summary>
        [Column(TypeName = "character varying(30)")]
        public string material_inhouse_name_1 { get; set; } = "-";

        [Column(TypeName = "character varying(50)")]
        public string material_inhouse_info_1 { get; set; } = "-";

        
        public double material_inhouse_value_1 { get; set; } = 0;

        [Column(TypeName = "character varying(50)")]
        public string material_inhouse_info_1b { get; set; } = "gram/pcs       $/kg";

        
        public double material_inhouse_value_1b { get; set; } = 0;

        
        public double material_inhouse_cost_sgd_1 { get; set; } = 0;

        
        public double material_inhouse_percentage_target_price_1 { get; set; } = 0;

        [Column(TypeName = "character varying(30)")]
        public string material_inhouse_name_2 { get; set; } = "-";

        [Column(TypeName = "character varying(50)")]
        public string material_inhouse_info_2 { get; set; } = "-";

        
        public double material_inhouse_value_2 { get; set; } = 0;

        [Column(TypeName = "character varying(50)")]
        public string material_inhouse_info_2b { get; set; } = "gram/pcs       $/kg";

        
        public double material_inhouse_value_2b { get; set; } = 0;

        
        public double material_inhouse_cost_sgd_2 { get; set; } = 0;

        
        public double material_inhouse_percentage_target_price_2 { get; set; } = 0;

        [Column(TypeName = "character varying(30)")]
        public string material_inhouse_name_3 { get; set; } = "-";

        [Column(TypeName = "character varying(50)")]
        public string material_inhouse_info_3 { get; set; } = "-";

        
        public double material_inhouse_value_3 { get; set; } = 0;

        [Column(TypeName = "character varying(50)")]
        public string material_inhouse_info_3b { get; set; } = "gram/pcs       $/kg";

        
        public double material_inhouse_value_3b { get; set; } = 0;

        
        public double material_inhouse_cost_sgd_3 { get; set; } = 0;

        
        public double material_inhouse_percentage_target_price_3 { get; set; } = 0;


        ////////////Material Outside
        [Column(TypeName = "character varying(30)")]
        public string material_outside_name_1 { get; set; } = "-";

        [Column(TypeName = "character varying(50)")]
        public string material_outside_info_1 { get; set; } = "-";

        
        public double material_outside_value_1 { get; set; } = 0;

        [Column(TypeName = "character varying(50)")]
        public string material_outside_info_1b { get; set; } = "gram/pcs       $/kg";

        
        public double material_outside_value_1b { get; set; } = 0;

        
        public double material_outside_cost_sgd_1 { get; set; } = 0;

        
        public double material_outside_percentage_target_price_1 { get; set; } = 0;

        [Column(TypeName = "character varying(30)")]
        public string material_outside_name_2 { get; set; } = "-";

        [Column(TypeName = "character varying(50)")]
        public string material_outside_info_2 { get; set; } = "-";

        
        public double material_outside_value_2 { get; set; } = 0;

        [Column(TypeName = "character varying(50)")]
        public string material_outside_info_2b { get; set; } = "gram/pcs       $/kg";

        
        public double material_outside_value_2b { get; set; } = 0;

        
        public double material_outside_cost_sgd_2 { get; set; } = 0;

        
        public double material_outside_percentage_target_price_2 { get; set; } = 0;

        [Column(TypeName = "character varying(30)")]
        public string material_outside_name_3 { get; set; } = "-";

        [Column(TypeName = "character varying(50)")]
        public string material_outside_info_3 { get; set; } = "-";

        
        public double material_outside_value_3 { get; set; } = 0;

        [Column(TypeName = "character varying(50)")]
        public string material_outside_info_3b { get; set; } = "gram/pcs       $/kg";

        
        public double material_outside_value_3b { get; set; } = 0;

        
        public double material_outside_cost_sgd_3 { get; set; } = 0;

        
        public double material_outside_percentage_target_price_3 { get; set; } = 0;

        /// <summary>
        /// Direct Material 
        /// </summary>

        
        public double direct_material_cost { get; set; } = 0;

        
        public double direct_material_cost_percentage { get; set; } = 0;

        
        public double sub_material_percentage { get; set; } = 10;

        
        public double sub_material_cost { get; set; } = 0;

        
        public double sub_material_cost_percentage { get; set; } = 0;

        
        public double direct_process_cost { get; set; } = 0;

        
        public double direct_process_cost_percentage { get; set; } = 0;

        
        public double total_direct_cost { get; set; } = 0;

        
        public double total_direct_cost_percentage { get; set; } = 0;

        
        public double defective_percentage { get; set; } = 3;

        
        public double defective_cost { get; set; } = 0;

        
        public double defective_cost_percentage { get; set; } = 0;

        
        public double indirect_percentage { get; set; } = 15;

        
        public double indirect_cost { get; set; } = 0;

        
        public double indirect_cost_percentage { get; set; } = 0;

        
        public double packing_material_percentage { get; set; } = 5;

        
        public double special_package_cost { get; set; } = 0;

        
        public double packing_material_cost { get; set; } = 0;

        
        public double packing_material_cost_percentage { get; set; } = 0;

        
        public double administration_percentage { get; set; } = 10;

        
        public double administration_cost { get; set; } = 0;

        
        public double administration_cost_percentage { get; set; } = 0;

        
        public double plant_retail_percentage { get; set; } = 5;//Tooling Maintenace

        
        public double plant_retail_cost { get; set; } = 0;//Tooling Maintenace

        
        public double plant_retail_cost_percentage { get; set; } = 0;//Tooling Maintenace

        
        public double moldjig_percentage { get; set; } = 0;

        
        public double moldjig_cost { get; set; } = 0;

        public double moldjig_cost_percentage { get; set; } = 0;

        public double die_percentage { get; set; } = 0;

        public double die_cost { get; set; } = 0;
        
        public double die_cost_percentage { get; set; } = 0;

        public string note { get; set; } = "-";

        
        public double net_included_tooling_cost { get; set; } = 0;

        public double net_included_tooling_cost_percentage { get; set; } = 0;

        
        public double net_exclude_tooling_cost { get; set; } = 0;

        
        public double net_exclude_tooling_cost_percentage { get; set; } = 0;


        ////Machine & Tooling List
        public string tooling_list_description_1 { get; set; } = "-";
        public string tooling_list_type_1 { get; set; } = "-";
        public string tooling_list_source_1 { get; set; } = "-";
        public int tooling_list_qty_1 { get; set; } = 0;
        public string tooling_list_unit_1 { get; set; } = "-";
        public double tooling_list_price_1 { get; set; } = 0;
        public double tooling_list_amount_jpy_1 { get; set; } = 0;
        public double tooling_list_amount_usd_1 { get; set; } = 0;
        public double tooling_list_amount_sgd_1 { get; set; } = 0;

        public string tooling_list_description_2 { get; set; } = "-";
        public string tooling_list_type_2 { get; set; } = "-";
        public string tooling_list_source_2 { get; set; } = "-";
        public int tooling_list_qty_2 { get; set; } = 0;
        public string tooling_list_unit_2 { get; set; } = "-";
        public double tooling_list_price_2 { get; set; } = 0;
        public double tooling_list_amount_jpy_2 { get; set; } = 0;
        public double tooling_list_amount_usd_2 { get; set; } = 0;
        public double tooling_list_amount_sgd_2 { get; set; } = 0;

        public string tooling_list_description_3 { get; set; } = "-";
        public string tooling_list_type_3 { get; set; } = "-";
        public string tooling_list_source_3 { get; set; } = "-";
        public int tooling_list_qty_3 { get; set; } = 0;
        public string tooling_list_unit_3 { get; set; } = "-";
        public double tooling_list_price_3 { get; set; } = 0;
        public double tooling_list_amount_jpy_3 { get; set; } = 0;
        public double tooling_list_amount_usd_3 { get; set; } = 0;
        public double tooling_list_amount_sgd_3 { get; set; } = 0;

        public string tooling_list_description_4 { get; set; } = "-";
        public string tooling_list_type_4 { get; set; } = "-";
        public string tooling_list_source_4 { get; set; } = "-";
        public int tooling_list_qty_4 { get; set; } = 0;
        public string tooling_list_unit_4 { get; set; } = "-";
        public double tooling_list_price_4 { get; set; } = 0;
        public double tooling_list_amount_jpy_4 { get; set; } = 0;
        public double tooling_list_amount_usd_4 { get; set; } = 0;
        public double tooling_list_amount_sgd_4 { get; set; } = 0;

        public string tooling_list_description_5 { get; set; } = "-";
        public string tooling_list_type_5 { get; set; } = "-";
        public string tooling_list_source_5 { get; set; } = "-";
        public int tooling_list_qty_5 { get; set; } = 0;
        public string tooling_list_unit_5 { get; set; } = "-";
        public double tooling_list_price_5 { get; set; } = 0;
        public double tooling_list_amount_jpy_5 { get; set; } = 0;
        public double tooling_list_amount_usd_5 { get; set; } = 0;
        public double tooling_list_amount_sgd_5 { get; set; } = 0;

        public string tooling_list_description_6 { get; set; } = "-";
        public string tooling_list_type_6 { get; set; } = "-";
        public string tooling_list_source_6 { get; set; } = "-";
        public int tooling_list_qty_6 { get; set; } = 0;
        public string tooling_list_unit_6 { get; set; } = "-";
        public double tooling_list_price_6 { get; set; } = 0;
        public double tooling_list_amount_jpy_6 { get; set; } = 0;
        public double tooling_list_amount_usd_6 { get; set; } = 0;
        public double tooling_list_amount_sgd_6 { get; set; } = 0;

        public string tooling_list_description_7 { get; set; } = "-";
        public string tooling_list_type_7 { get; set; } = "-";
        public string tooling_list_source_7 { get; set; } = "-";
        public int tooling_list_qty_7 { get; set; } = 0;
        public string tooling_list_unit_7 { get; set; } = "-";
        public double tooling_list_price_7 { get; set; } = 0;
        public double tooling_list_amount_jpy_7 { get; set; } = 0;
        public double tooling_list_amount_usd_7 { get; set; } = 0;
        public double tooling_list_amount_sgd_7 { get; set; } = 0;

        public string tooling_list_description_8 { get; set; } = "-";
        public string tooling_list_type_8 { get; set; } = "-";
        public string tooling_list_source_8 { get; set; } = "-";
        public int tooling_list_qty_8 { get; set; } = 0;
        public string tooling_list_unit_8 { get; set; } = "-";
        public double tooling_list_price_8 { get; set; } = 0;
        public double tooling_list_amount_jpy_8 { get; set; } = 0;
        public double tooling_list_amount_usd_8 { get; set; } = 0;
        public double tooling_list_amount_sgd_8 { get; set; } = 0;

        public string tooling_list_description_9 { get; set; } = "-";
        public string tooling_list_type_9 { get; set; } = "-";
        public string tooling_list_source_9 { get; set; } = "-";
        public int tooling_list_qty_9 { get; set; } = 0;
        public string tooling_list_unit_9 { get; set; } = "-";
        public double tooling_list_price_9 { get; set; } = 0;
        public double tooling_list_amount_jpy_9 { get; set; } = 0;
        public double tooling_list_amount_usd_9 { get; set; } = 0;
        public double tooling_list_amount_sgd_9 { get; set; } = 0;

        public string tooling_list_description_10 { get; set; } = "-";
        public string tooling_list_type_10 { get; set; } = "-";
        public string tooling_list_source_10 { get; set; } = "-";
        public int tooling_list_qty_10 { get; set; } = 0;
        public string tooling_list_unit_10 { get; set; } = "-";
        public double tooling_list_price_10 { get; set; } = 0;
        public double tooling_list_amount_jpy_10 { get; set; } = 0;
        public double tooling_list_amount_usd_10 { get; set; } = 0;
        public double tooling_list_amount_sgd_10 { get; set; } = 0;

        public string tooling_list_description_11 { get; set; } = "-";
        public string tooling_list_type_11 { get; set; } = "-";
        public string tooling_list_source_11 { get; set; } = "-";
        public int tooling_list_qty_11 { get; set; } = 0;
        public string tooling_list_unit_11 { get; set; } = "-";
        public double tooling_list_price_11 { get; set; } = 0;
        public double tooling_list_amount_jpy_11 { get; set; } = 0;
        public double tooling_list_amount_usd_11 { get; set; } = 0;
        public double tooling_list_amount_sgd_11 { get; set; } = 0;

        public string tooling_list_description_12 { get; set; } = "-";
        public string tooling_list_type_12 { get; set; } = "-";
        public string tooling_list_source_12 { get; set; } = "-";
        public int tooling_list_qty_12 { get; set; } = 0;
        public string tooling_list_unit_12 { get; set; } = "-";
        public double tooling_list_price_12 { get; set; } = 0;
        public double tooling_list_amount_jpy_12 { get; set; } = 0;
        public double tooling_list_amount_usd_12 { get; set; } = 0;
        public double tooling_list_amount_sgd_12 { get; set; } = 0;

        public string tooling_list_description_13 { get; set; } = "-";
        public string tooling_list_type_13 { get; set; } = "-";
        public string tooling_list_source_13 { get; set; } = "-";
        public int tooling_list_qty_13 { get; set; } = 0;
        public string tooling_list_unit_13 { get; set; } = "-";
        public double tooling_list_price_13 { get; set; } = 0;
        public double tooling_list_amount_jpy_13 { get; set; } = 0;
        public double tooling_list_amount_usd_13 { get; set; } = 0;
        public double tooling_list_amount_sgd_13 { get; set; } = 0;

        public string tooling_list_description_14 { get; set; } = "-";
        public string tooling_list_type_14 { get; set; } = "-";
        public string tooling_list_source_14 { get; set; } = "-";
        public int tooling_list_qty_14 { get; set; } = 0;
        public string tooling_list_unit_14 { get; set; } = "-";
        public double tooling_list_price_14 { get; set; } = 0;
        public double tooling_list_amount_jpy_14 { get; set; } = 0;
        public double tooling_list_amount_usd_14 { get; set; } = 0;
        public double tooling_list_amount_sgd_14 { get; set; } = 0;

        public string tooling_list_description_15 { get; set; } = "-";
        public string tooling_list_type_15 { get; set; } = "-";
        public string tooling_list_source_15 { get; set; } = "-";
        public int tooling_list_qty_15 { get; set; } = 0;
        public string tooling_list_unit_15 { get; set; } = "-";
        public double tooling_list_price_15 { get; set; } = 0;
        public double tooling_list_amount_jpy_15 { get; set; } = 0;
        public double tooling_list_amount_usd_15 { get; set; } = 0;
        public double tooling_list_amount_sgd_15 { get; set; } = 0;

        //total machine and tooling list
        public double tooling_list_total_amount_usd { get; set; } = 0;

        public double tooling_list_total_amount_sgd { get; set; } = 0;

        //Raw Material Cost
        public double direct_raw_material { get; set; } = 0;

        public double direct_raw_material_p { get; set; } = 0;

        public double sub_material { get; set; } = 0;

        public double sub_material_p { get; set; } = 0;

        public double raw_material_cost_sub_total { get; set; } = 0;

        public double raw_material_cost_sub_total_p { get; set; } = 0;


        //Breakdown of Process Cost
        
        public double labor_cost { get; set; } = 0;

        
        public double labor_cost_p { get; set; } = 0;

        
        public double machine_cost { get; set; } = 0;

        
        public double machine_cost_p { get; set; } = 0;

        
        public double overhead_cost { get; set; } = 0;

        
        public double overhead_cost_p { get; set; } = 0;

        
        public double process_cost_sub_total { get; set; } = 0;

        
        public double process_cost_sub_total_p { get; set; } = 0;

        public double defectives { get; set; } = 0;

        public double defectives_p { get; set; } = 0;

        public double admin_engin_qc { get; set; } = 0;

        public double admin_engin_qc_p { get; set; } = 0;

        public double tooling_cost { get; set; } = 0;

        public double tooling_cost_p { get; set; } = 0;

        public double process_margin_adjust { get; set; } = 0;

        public double process_margin_adjust_p { get; set; } = 0;

        public double other_fixed_cost_sub_total { get; set; } = 0;

        public double other_fixed_cost_sub_total_p { get; set; } = 0;

        public double grand_total_cost { get; set; } = 0;

        public double grand_total_cost_p { get; set; } = 0;


        public double production_capacity { get; set; } = 0;
        public double actual_working_time { get; set; } = 0;

        public double cycle_time { get; set; } = 0;

        public double efficiency { get; set; } = 0;
        public double daily_qty_days { get; set; } = 0;

        public double daily_qty_days_p { get; set; } = 0;

        public double daily_amount { get; set; } = 0;


        //Navigation Properties
        public List<CostProcess> Cost_Processes { get; set; }


    }
}





