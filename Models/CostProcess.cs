using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Threading.Tasks;

namespace CostNag.Models
{
    public class CostProcess
    {
        public int CostProcessId { get; set; }


        [Column(TypeName = "character varying(20)")]
        public string doc_no { get; set; }

        [Column(TypeName = "character varying(100)")]
        public string process_name { get; set; }

        [Column(TypeName = "character varying(20)")]
        public string process_type { get; set; }

        public double item_od { get; set; }

        [Column(TypeName = "double precision")]
        public double overhead_cost { get; set; }

        [Column(TypeName = "double precision")]
        public double machine_cost { get; set; }

        [Column(TypeName = "double precision")]
        public double labor_cost { get; set; }

        [Column(TypeName = "double precision")]
        public double total_cost { get; set; }

        public List<CostProcess> data = new List<CostProcess>();
    }
}
