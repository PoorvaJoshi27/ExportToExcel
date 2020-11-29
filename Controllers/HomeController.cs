using ExportExcell.Models;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace ExportExcell.Controllers
{
    public class HomeController : Controller
    {
        BillingEntities db = new BillingEntities();

        public ActionResult Index()
        {
            List<BillModel> lst = db.Bills.Select(x => new BillModel
            {
                B_no = x.B_no,
                Product = x.Product,
                Price = (int)x.Price,
                Date = (DateTime)x.Date
            }).ToList();

            return View(lst);
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

        public void ExportToExcel()
        {
            List<BillModel> lst = db.Bills.Select(x => new BillModel
            {
                B_no = x.B_no,
                Product = x.Product,
                Price = (int)x.Price,
                Date = (DateTime)x.Date
            }).ToList();

            ExcelPackage pck = new ExcelPackage();
            ExcelWorksheet ws = pck.Workbook.Worksheets.Add("Report");

            ws.Cells["A1"].Value = "B_no";
            ws.Cells["B1"].Value = "Product";
            ws.Cells["C1"].Value = "Price";
            ws.Cells["D1"].Value = "Date";

            int rowStart = 7;
            foreach (var item in lst)
            {
                ws.Cells[string.Format("A{0}", rowStart)].Value = item.B_no;
                ws.Cells[string.Format("B{0}", rowStart)].Value = item.Product;
                ws.Cells[string.Format("C{0}", rowStart)].Value = item.Price;
                ws.Cells[string.Format("D{0}", rowStart)].Value = item.Date;
                rowStart++;
            }

            ws.Cells["A:AZ"].AutoFitColumns();
            Response.Clear();
            Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            Response.AddHeader("content-disposition", "attachment: filename=" + "ExcelReport.xlsx");
            Response.BinaryWrite(pck.GetAsByteArray());
            Response.End();
        }

        public ActionResult Display()
        {
            BillingEntities db = new BillingEntities();
            DataModel dt = new DataModel();
            var get = db.Bills.ToList();
            List<BillModel> lst = new List<BillModel>();

            foreach (var item in get)
            {
                lst.Add(new BillModel
                {
                    B_no = item.B_no,
                    Product = item.Product,
                    Price = (int)item.Price,
                    Date = (DateTime)item.Date
                });
            }
            dt.list = lst;
            return View(dt);
        }

        public ActionResult Update(int ID)
        {
            BillingEntities db = new BillingEntities();
            BillModel dt = new BillModel();

            var getdata = db.Bills.FirstOrDefault(m => m.B_no == ID);

            dt.Product = getdata.Product;
            dt.Price = getdata.Price;
            dt.Date = (DateTime)getdata.Date;

            return View(dt);

        }
        [HttpPost]
        public ActionResult Update(int ID, BillModel dt)
        {
            BillingEntities db = new BillingEntities();

            var getdata = db.Bills.FirstOrDefault(m => m.B_no == ID);
            getdata.Product = (string)dt.Product;
            getdata.Price = (Decimal)dt.Price;
            getdata.Date = (DateTime)dt.Date;
            db.SaveChanges();

            ViewBag.Text = "Details Updated";


            return RedirectToAction("Display", "Home");
        }

        public ActionResult Delete(int id)
        {
            BillingEntities db = new BillingEntities();
            var getdata = db.Bills.FirstOrDefault(m => m.B_no == id);
            db.Bills.Remove(getdata);
            db.SaveChanges();
            return RedirectToAction("Display", "Home");
        }

        public ActionResult Insert()
        {
            return View();
        }

        [HttpPost]
        public ActionResult Insert(BillModel d)
        {
            BillingEntities db = new BillingEntities();
            Bill B = new Bill();
            B.Product = d.Product;
            B.Price = d.Price;
            B.Date = d.Date;
            db.Bills.Add(B);
            db.SaveChanges();
            ViewBag.Text = "Details Inserted";

            return RedirectToAction("Display", "Home");
        }
    }
}