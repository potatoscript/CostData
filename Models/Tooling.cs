using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace CostNag.Models
{
    public class Tooling
    {
        public int ToolingId { get; set; }

        public string description { get; set; }

        public string source { get; set; }

        public double qty { get; set; }

        public string unit { get; set; }

        public double price { get; set; }

        public List<Tooling> data = new List<Tooling>();
    }
}
