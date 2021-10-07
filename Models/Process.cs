using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Threading.Tasks;

namespace CostNag.Models
{
    public class Process
    {
        public double process_cost { get; set; }

        public int ProcessId { get; set; }

        public string doc_no { get; set; }

        public string process_name { get; set; }
        public double working_day { get; set; }
        public double working_time_day { get; set; }

        public double working_time_month { get; set; }
        public double shift { get; set; }

        [Column(TypeName = "double precision")]
        public double worker { get; set; }

        [Column(TypeName = "double precision")]
        public double direct_labour_unit { get; set; }

        [Column(TypeName = "double precision")]
        public double direct_labour { get; set; }

        [Column(TypeName = "double precision")]
        public double total_labour_cost { get; set; }

        [Column(TypeName = "double precision")]
        public double machine_qty { get; set; }

        [Column(TypeName = "double precision")]
        public double area { get; set; }

        [Column(TypeName = "double precision")]
        public double special_material { get; set; }

        [Column(TypeName = "double precision")]
        public double plant_maintenance { get; set; }

        [Column(TypeName = "double precision")]
        public double plant_maintenance_unit { get; set; }

        [Column(TypeName = "double precision")]
        public double total_machine_cost { get; set; }

        [Column(TypeName = "double precision")]
        public double machine_usage_day { get; set; }

        [Column(TypeName = "double precision")]
        public double machine_cost_month { get; set; }

        [Column(TypeName = "double precision")]
        public double machine_cost_month_percentage { get; set; }

        [Column(TypeName = "double precision")]
        public double machine_cost_month_percentage_unit { get; set; }

        [Column(TypeName = "double precision")]
        public double consumption_kwh { get; set; }

        [Column(TypeName = "double precision")]
        public double consumption_unit { get; set; }

        [Column(TypeName = "double precision")]
        public double consumption_sgd { get; set; }

        [Column(TypeName = "double precision")]
        public double consumption_rate { get; set; }

        [Column(TypeName = "double precision")]
        public double utility_electric { get; set; }

        [Column(TypeName = "double precision")]
        public double machine_utility_cost { get; set; }

        [Column(TypeName = "double precision")]
        public double labour_electric_cost { get; set; }

        [Column(TypeName = "double precision")]
        public double charge { get; set; }

        [Column(TypeName = "double precision")]
        public double cycle_time { get; set; }

        [Column(TypeName = "double precision")]
        public double cycle_time_unit { get; set; }

        [Column(TypeName = "double precision")]
        public double time { get; set; }

        [Column(TypeName = "double precision")]
        public double capacity { get; set; }

        [Column(TypeName = "double precision")]
        public double time_g { get; set; }

        [Column(TypeName = "double precision")]
        public double efficiency { get; set; }

        [Column(TypeName = "double precision")]
        public double production_capacity { get; set; }

        [Column(TypeName = "double precision")]
        public double production_cycle_time { get; set; }

        [Column(TypeName = "double precision")]
        public double special_input { get; set; }

        [Column(TypeName = "double precision")]
        public double direct_process_cost { get; set; }

        [Column(TypeName = "double precision")]
        public double labour_cost_percentage { get; set; }

        [Column(TypeName = "double precision")]
        public double machine_cost_percentage { get; set; }

        [Column(TypeName = "double precision")]
        public double overhead_cost_percentage { get; set; }

        [Column(TypeName = "double precision")]
        public double total_cost_percentage { get; set; }


    }
}
