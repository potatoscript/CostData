using CostNag.Helper;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Threading.Tasks;

namespace CostNag.Models
{

    public class Filter
    {
        public string Name { get; set; }
    }
    public class ListModel
    {
        public Filter[] Filters { get; set; }

        public List<ListModel> partscode = new List<ListModel>();
        public string code { get; set; }
        public string doc_no { get; set; }
        public string id { get; set; }
        public string master_process { get; set; }
        public string processName { get; set; }
        public string processType { get; set; }

        public List<string> processMaster = new List<string>();

        public List<ListModel> process = new List<ListModel>();
        public string name { get; set; }
        public string cost { get; set; }

        public List<SelectListItem> plant = new List<SelectListItem>
        {
            new SelectListItem{Value = "BPK1", Text="BPK1"},
            new SelectListItem{Value = "BPK2", Text="BPK2"},
            new SelectListItem{Value = "BPK3", Text="BPK3"},
            new SelectListItem{Value = "PTN", Text="PTN"}
        };
        public List<SelectListItem> item_spec = new List<SelectListItem>
        {
            new SelectListItem{Value = "New RFQ", Text="New RFQ"},
            new SelectListItem{Value = "Current Item", Text="Current Item"},
            new SelectListItem{Value = "SM", Text="SM"}
        };

        public List<ListModel> GetTypes()
        {
            return new List<ListModel>()
            {
                new ListModel(){processType = "Process Type"},
                new ListModel(){processType = "Bonding"},
                new ListModel(){processType = "Combination Forming"},
                new ListModel(){processType = "Progressive"},
                new ListModel(){processType = "Rubber"},
                new ListModel(){processType = "Spring"},
                new ListModel(){processType = "Step Forming"},
                new ListModel(){processType = "Trimmer"}
            };
        }



    }
}
