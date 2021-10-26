using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Threading.Tasks;

namespace CostNag.Models
{
    public class Rubber
    {

        public int RubberId { get; set; }

        [Column(TypeName = "character varying(10)")]
        public string material_name { get; set; }

        [Column(TypeName = "double precision")]
        public double price_kg { get; set; }

        [Column(TypeName = "double precision")]
        public double mixing_process_cost { get; set; }

        [Column(TypeName = "double precision")]
        public double weight_g { get; set; }

        [Column(TypeName = "double precision")]
        public double weight_kg { get; set; }

        [Column(TypeName = "double precision")]
        public double yield_rate { get; set; }

        [Column(TypeName = "double precision")]
        public double weight_kg_pcs { get; set; }

        [Column(TypeName = "double precision")]
        public double rubber_sgd { get; set; }

        [Column(TypeName = "double precision")]
        public double rubber_target_price_percentage { get; set; }

        public List<Rubber> data = new List<Rubber>();

    }
}
