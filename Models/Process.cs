using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Threading.Tasks;

namespace CostNag.Models
{
    public class Process
    {
        public int ProcessId { get; set; }

        [Column(TypeName = "character varying(20)")]
        public string doc_no { get; set; }

        [Column(TypeName = "double precision")]
        public double process_cost { get; set; }

        [Column(TypeName = "double precision")]
        public double daily_amount { get; set; }

        [Column(TypeName = "integer")]
        public int daily_qty { get; set; }

        [Column(TypeName = "double precision")]
        public double daily_percentage { get; set; }

        [Column(TypeName = "integer")]
        public double working_day { get; set; }


    }
}
