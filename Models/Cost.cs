using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Threading.Tasks;

namespace CostNag.Models
{
    public class Cost
    {

        public int CostId { get; set; }

        [Column(TypeName = "character varying(5)")]
        public string plant { get; set; }

        [Column(TypeName = "character varying(20)")]
        public string item_spec { get; set; }

        [Column(TypeName = "character varying(10)")]
        public string issue_date { get; set; }

        [Column(TypeName = "character varying(10)")]
        public string section { get; set; }

        [Column(TypeName = "character varying(20)")]
        public string doc_no { get; set; }

        [Column(TypeName = "character varying(20)")]
        public string wr_no { get; set; }

        [Column(TypeName = "character varying(10)")]
        public string sales { get; set; }

        [Column(TypeName = "character varying(10)")]
        public string revision_no { get; set; }

        [Column(TypeName = "character varying(10)")]
        public string checked_date { get; set; }

        [Column(TypeName = "character varying(10)")]
        public string approved_by { get; set; }

        [Column(TypeName = "character varying(20)")]
        public string customer { get; set; }

        [Column(TypeName = "character varying(15)")]
        public string parts_code { get; set; }

        [Column(TypeName = "character varying(10)")]
        public string item { get; set; }

        [Column(TypeName = "character varying(20)")]
        public string product { get; set; }

        [Column(TypeName = "character varying(10)")]
        public string product_type { get; set; }

        [Column(TypeName = "character varying(10)")]
        public string size { get; set; }

        [Column(TypeName = "character varying(20)")]
        public string business_type { get; set; }

        [Column(TypeName = "integer")]
        public int qty_month { get; set; }

        [Column(TypeName = "double precision")]
        public double exchange_rate { get; set; }

        [Column(TypeName = "double precision")]
        public double target_price_bht { get; set; }

        [Column(TypeName = "double precision")]
        public double target_price_export { get; set; }

        [Column(TypeName = "integer")]
        public double production_qty_day { get; set; }

        [Column(TypeName = "integer")]
        public double working_day { get; set; }


    }
}





