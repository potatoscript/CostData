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
    public class ListModel
    {

        public List<ListModel> partscode = new List<ListModel>();
        public string code { get; set; }
        public string id { get; set; }

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
    }
}
