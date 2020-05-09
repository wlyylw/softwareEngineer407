using Microsoft.Ajax.Utilities;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Security.Cryptography;
using System.Web;
using System.Web.Mvc;

namespace coursedesign.Controllers
{
    public class HomeController : Controller
    {
        static DataEntity db = new DataEntity();
        List<engineering> list = (from c in db.engineering select c).ToList();

        public ActionResult Index()
        {
            return View();
        }
        [HttpGet]
        public ActionResult QueryData()
        {
            var listItem = new List<SelectListItem>
            {
                new SelectListItem{Text="根据姓名查询",Value="asname"},
                new SelectListItem{Text="根据编号查询",Value="asid"}
            };
            ViewBag.list = new SelectList(listItem, "Value", "Text", "");
            return View();
        }
     
        public ActionResult ShowQueryData()
        {
            var option = Request.Form["queryOption"];
            if (option.Equals("asname"))
            {
                string value = Request.Form["queryType"];
                List<engineering> list = (from c in db.engineering
                                          where c.name == value
                                          select c).ToList();
                ViewData["QueryList"] = list;
            }
            else if(option.Equals("asid"))
            {
                int value = Convert.ToInt32( Request.Form["queryType"]);
                List<engineering> list = (from c in db.engineering
                                          where c.id == value
                                          select c).ToList();
                ViewData["QueryList"] = list;
            }
            else
            {
                return Content("请输入信息");
            }
            return View();
        }
        [HttpGet]
        public ActionResult AddData()
        {
            return View();
        }
        [HttpPost]
        public ActionResult AddData(engineering model)
        {
            try
            {
                engineering e = new engineering();
                e.id = model.id;
                e.name = model.name.Trim();
                e.place = model.place.Trim();
                e.salary = model.salary;
                e.sex = model.sex;
                e.telephone = model.telephone.Trim();
                e.workage = model.workage.Trim();
                e.address = model.address.Trim();
                e.education = model.education;
                e.birth = model.birth;
                db.engineering.Add(e);
                db.SaveChanges();
                db.Configuration.ValidateOnSaveEnabled = true;
                return RedirectToAction("AddData", "Home");
            }
            catch (Exception)
            {
                //指定对应跳转的视图Test下的Test.cshtml文件
     /*           return RedirectToAction("AddData", "AddData");*/
                return Content("添加失败" );
            }
        }

        public ActionResult EditData()
        {
            ViewData["DataList"] = list;
            return View();
        }
        public ActionResult CalculateData()
        {
            int id = Convert.ToInt32(Request.QueryString["Cid"]);
            engineering engineering = (from e in db.engineering where id == e.id select e).SingleOrDefault();
            return View(engineering);
        }
        [HttpPost]
        public ActionResult CalculateData(engineering engineering)
        {
            int basicsalary = Convert.ToInt32(engineering.salary);
            int workage = Convert.ToInt32(engineering.workage);
            var effectiveDay = Convert.ToInt32(Request.Form["effectiveDay"]);
            var effectiveMonth = Convert.ToInt32(Request.Form["effectiveMonth"]);
            var insurace = Convert.ToInt32(Request.Form["insurance"]);
            if (effectiveDay == 0 || effectiveMonth == 0 || insurace == 0 || workage == 0 || basicsalary == 0)
            {
                ViewBag.salary = null;
            }
            else
                ViewBag.salary = (basicsalary + 10 * effectiveDay + effectiveMonth * workage / 100) * 0.9 - insurace;
            return View();
        }



        public ActionResult SaveData()
        {
            return View();
        }
        public ActionResult ClearData()
        {
            return View();
        }
        public ActionResult PrintData()
        {

            return View();
        }
        public ActionResult ReGetData()
        {
            return View();
        }
        public ActionResult ExitSys()
        {
            return View();
        }

        [HttpGet]
        public ActionResult Modify()
        {
            int id = Convert.ToInt32(Request.QueryString["Cid"]);
            engineering engineering = (from e in db.engineering where id == e.id select e).SingleOrDefault();
            return View(engineering);
        }
        [HttpPost]
        public ActionResult Modify(engineering model)
        {
            engineering e = db.engineering.Where(e1 => e1.id == model.id).ToList().FirstOrDefault();
            e.name = model.name.Trim();
            e.place = model.place.Trim(); 
            e.salary = model.salary;
            e.sex = model.sex;
            e.telephone = model.telephone.Trim();
            e.workage = model.workage.Trim();
            e.address = model.address.Trim();
            e.education = model.education;
            e.birth = model.birth;
            db.SaveChanges();
            db.Configuration.ValidateOnSaveEnabled = true;
            return RedirectToAction("EditData", "Home");
        }
        public ActionResult Remove()
        {
            try
            {
                int id = Convert.ToInt32(Request.QueryString["Cid"]);
                engineering engineering = (from e in db.engineering where e.id == id select e).SingleOrDefault();
                db.engineering.Remove(engineering);
                db.SaveChanges();
                return RedirectToAction("EditData", "Home");
            }
            catch (Exception)
            {
                return RedirectToAction("EditData", "EditData");
            }
        }
    }
}