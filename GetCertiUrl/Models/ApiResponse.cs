using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace GetCertiUrl.Models
{

    public class ApiResponse
    {
        public string  code { get; set; }
        public ApiResponse_Inner data { get; set; }
    }
    public class ApiResponse_Inner
    {
        public string cert_url { get; set; }
        public string stone_id { get; set; }
    }
}