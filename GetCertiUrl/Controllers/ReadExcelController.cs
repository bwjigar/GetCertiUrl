using GetCertiUrl.Models;
using Newtonsoft.Json;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.Common;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Web;
using System.Web.Mvc;
using System.Web.Script.Serialization;

namespace GetCertiUrl.Controllers
{
    public class ReadExcelController : Controller
    {
        // GET: ReadExcel
        public ActionResult Index()
        {
            return View();
        }

        [HttpPost]
        public ActionResult Index(ReadExcel readExcel)
        {
            if (ModelState.IsValid)
            {
                string ConStr = "";

                DateTime _now = DateTime.Now;
                string _date = _now.Day + "" + _now.Month + "" + _now.Year + "" + _now.Hour + "" + _now.Minute + "" + _now.Second;

                
                 string ProjectUrl = ConfigurationManager.AppSettings["ProjectURL"];
                string path = Server.MapPath("~/Content/Upload/" +_date + readExcel.file.FileName  );
                readExcel.file.SaveAs(path);


                if (Path.GetExtension(path).ToLower().Trim() == ".xls")
                {
                    ConStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + path + ";Extended Properties=\"Excel 8.0;HDR=Yes;IMEX=2\"";
                }
                else
                {
                    ConStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";Extended Properties=\"Excel 12.0;HDR=Yes;IMEX=2\"";
                    ViewBag.Message = "Only Xls File Allowed.";
                    return View();
                }

                string query = "SELECT * FROM [Sheet1$]";

                OleDbConnection conn = new OleDbConnection(ConStr);

                if (conn.State == ConnectionState.Closed)
                {
                    conn.Open();
                }

                OleDbCommand cmd = new OleDbCommand(query, conn);

                OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                DataTable dt = new DataTable();
                da.Fill(dt);

                DataTable FinalDt = new DataTable();
                FinalDt.Columns.Add("StoneID", typeof(string));
                FinalDt.Columns.Add("CertNo", typeof(string));
                FinalDt.Columns.Add("PdfLink", typeof(string));

                if (dt != null)
                {

                    if (dt.Rows.Count <= 500)
                    {


                        string API = "";
                        for (int i = 0; i < dt.Rows.Count; i++)
                        {
                            string Url = dt.Rows[i]["Lab Link"].ToString();
                            if (Url != "")
                            {
                                Uri myUri = new Uri(Url);
                                string stoneIds = HttpUtility.ParseQueryString(myUri.Query).Get("stoneIds");

                                string MediaKey = HttpUtility.ParseQueryString(myUri.Query).Get("mediaKey");
                                //                '7XGK2MJUGJ':'http://104.211.91.117:3025',
                                //'0CWF0LXIKN':'http://qa.srk.best',
                                //'8RFWY4ZN3H':'http://stage-int.srk.best',
                                //'QX99JES0BU':'https://int.srk.best'

                                if (MediaKey == "QX99JES0BU") { API = "https://int.srk.best"; }
                                else if (MediaKey == "8RFWY4ZN3H") { API = "http://stage-int.srk.best"; }
                                else if (MediaKey == "7XGK2MJUGJ") { API = "http://104.211.91.117:3025"; }
                                else if (MediaKey == "0CWF0LXIKN") { API = "http://qa.srk.best"; }

                                var customUrl = API + "/exposed/url/all/stone/v1/" + stoneIds;
                                string ResultReturn = "";

                                WebClient client = new WebClient();
                                client.Headers["Content-type"] = "application/x-www-form-urlencoded";
                                client.Encoding = Encoding.UTF8;
                                ServicePointManager.Expect100Continue = false;
                                ServicePointManager.SecurityProtocol |= SecurityProtocolType.Tls11 | SecurityProtocolType.Tls12;
                                string json;
                                DataRow dr = FinalDt.NewRow();
                                try
                                {
                                    json = client.DownloadString(customUrl);
                                    // List<ApiResponse> _data = new List<ApiResponse>();

                                    //_data = JsonConvert.DeserializeObject<List<ApiResponse>>(json.ToString());

                                    ApiResponse _data = new ApiResponse();
                                    _data = (new JavaScriptSerializer()).Deserialize<ApiResponse>(json);
                                    dr["StoneID"] = _data.data.stone_id.ToString();
                                    dr["PdfLink"] = _data.data.cert_url.ToString();
                                    string str = _data.data.cert_url.ToString();
                                    var m = Regex.Match(str, @"\S+(?=\.pdf)");
                                    string certno = m.Groups[0].Value.Split('/').Last();
                                    dr["CertNo"] = certno;
                                    FinalDt.Rows.Add(dr);

                                }
                                catch (WebException ex)
                                {
                                    //if (ex.Status)
                                    var webException = ex as WebException;


                                }
                                catch (Exception ex)
                                {
                                    json = ex.Message;

                                }
                            }
                        }

                        if (FinalDt != null)
                        {
                            //Exporting to Excel
                            //string folderPath = "C:\\Excel\\";
                            DateTime now = DateTime.Now;
                            string date = now.Day + "" + now.Month + "" + now.Year + "" + now.Hour + "" + now.Minute + "" + now.Second;
                            string folderPath = Server.MapPath("~/Content/FinalExcel/");
                           // string folderPath = ProjectUrl + "/Content/FinalExcel/";
                        
                            string filename = "CertUrl_" + date + ".xlsx";

                            if (!Directory.Exists(folderPath))
                            {
                                Directory.CreateDirectory(folderPath);
                            }
                            //using (XLWorkbook wb = new XLWorkbook())
                            //{
                            //    wb.Worksheets.Add(dt, "Report");
                            //    wb.SaveAs(folderPath + "" + filename);
                            //}

                            ExcelPackage.LicenseContext = LicenseContext.Commercial;
                            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                            using (var pck = new ExcelPackage(new FileInfo(folderPath + "" + filename)))
                            {
                                ExcelWorksheet ws = pck.Workbook.Worksheets.Add("CertUrl");
                                pck.Workbook.Properties.Title = "Certificate";
                                ws.Cells["A1"].LoadFromDataTable(FinalDt, true);
                                pck.Save();
                            }
                           // string folderPath = ProjectUrl + "/Content/FinalExcel/";
                            ViewBag.Message = ProjectUrl + "/Content/FinalExcel/"  + filename;
                        }


                    }
                    else
                    {
                        ViewBag.Message = "Max limit is 500 stone in excel.";
                    }
                }



            }
            return View("Index");
        }





    }
}