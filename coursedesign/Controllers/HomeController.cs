using coursedesign.Models;
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
        public ActionResult Index()
        {
            return View();
        }

        public ActionResult QueryData()
        {
            return View();
        }
        public ActionResult AddData()
        {
            return View();
        }
        public ActionResult DeleteData()
        {
            return View();
        }
        public ActionResult AlertData()
        {
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
            DataEntity db = new DataEntity();
            List<user> list = (from c in db.user select c).ToList();
            ViewData["DataList"] = list;
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
    }
}