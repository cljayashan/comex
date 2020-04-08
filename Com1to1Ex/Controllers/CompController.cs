using Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace Com1to1Ex.Controllers
{
    public class CompController : Controller
    {
        // GET: Comp
        [HttpGet]
        public ActionResult Index()
        {
            return View();
        }

        //[HttpPost]
        public ActionResult Index(HttpPostedFileBase file1, HttpPostedFileBase file2)
        {
            if ((file1 != null && file1.ContentLength > 0) && (file2 != null && file2.ContentLength > 0))
            {
                Dictionary<string, string> result = new Dictionary<string, string>();

                Stream f1Stream = file1.InputStream;
                Stream f2Stream = file2.InputStream;

                IExcelDataReader f1Reader = null;
                IExcelDataReader f2Reader = null;

                f1Reader = ExcelReaderFactory.CreateBinaryReader(f1Stream);
                f2Reader = ExcelReaderFactory.CreateBinaryReader(f2Stream);

                f1Reader.IsFirstRowAsColumnNames = false;
                f2Reader.IsFirstRowAsColumnNames = false;

                DataSet f1Result = f1Reader.AsDataSet();
                DataSet f2Result = f2Reader.AsDataSet();

                DataTable f1Dt = f1Result.Tables[0];
                DataTable f2Dt = f2Result.Tables[0];

                for (int i = 0; i < f1Dt.Rows.Count; i++)
                {
                    for (int j = 0; j < f1Dt.Columns.Count; j++)
                    {
                        string f1Val = f1Dt.Rows[i][j].ToString().Trim().ToUpper();
                        string f2Val = f2Dt.Rows[i][j].ToString().Trim().ToUpper();

                        if (f1Val != f2Val)
                        {
                            string key = "Row : '" + (i + 1) + "' | Column : '" + (j + 1) + "'";
                            string val = "'" + f1Val + "'   !=   '" + f2Val + "'";
                            result.Add(key, val);
                        }
                    }
                }
                ViewBag.Result = result;
            }
            return View();
        }
    }
}