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
        DataEntity db = new DataEntity();
        public ActionResult Index()
        {
            return View();
        }

        public ActionResult QueryData()
        {
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

            List<engineering> list = (from c in db.engineering select c).ToList();
            ViewData["DataList"] = list;
            return View();
        }
        public ActionResult CalculateData()
        {
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
                //需要一个实体对象参数
                //db.Customers.Remove(new Customer() {CustomerNo = id });
                //1,创建要删除的对象
                int id = Convert.ToInt32(Request.QueryString["Cid"]);
                engineering engineering = (from e in db.engineering where e.id == id select e).SingleOrDefault();
                db.engineering.Remove(engineering);
                db.SaveChanges();
                return RedirectToAction("EditData", "Home");
            }
            catch (Exception)
            {
                //指定对应跳转的视图Test下的Test.cshtml文件
                return RedirectToAction("EditData", "EditData");
                //return Content("删除失败" + ex.Message);
            }
        }
    }
}