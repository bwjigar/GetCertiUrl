using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace GetCertiUrl.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            //string url = "https://pck.blob.core.windows.net/int-multimedia/Stone-multimedia/stone-multimedia.htm?stoneIds=A3C335AA-3E0D-4601-B0CA-06B9B911A344&showMediaType=PDF&mediaKey=QX99JES0BU";
            //System.Net.WebRequest req = System.Net.WebRequest.Create(url);
            //System.Net.WebResponse resp = req.GetResponse();
            //System.IO.StreamReader sr = new System.IO.StreamReader(resp.GetResponseStream());
            //string response = sr.ReadToEnd().Trim();
            return View();
        }

        public ActionResult About()
        {
            ViewBag.Message = "Your application description page.";

            return View();
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }
    }
}