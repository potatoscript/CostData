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

        public string material_name { get; set; }

        public double price_kg { get; set; }

        public double mixing_process_cost { get; set; }

        public double weight_g { get; set; }

        public double yield_rate { get; set; }

        public List<Rubber> data = new List<Rubber>();

    }
}
