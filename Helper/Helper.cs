using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Threading.Tasks;

namespace CostNag.Helper
{
    public class CostAPI
    {
        public HttpClient Initial()
        {
            var Client = new HttpClient();
            //Client.BaseAddress = new Uri("http://localhost:58810");//for my pc
            Client.BaseAddress = new Uri("http://localhost/costnagapi/");//for nab server
            return Client;
        }
    }
}
