using ExcelDataReader;
using Microsoft.Ajax.Utilities;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Web;
using System.Web.Mvc;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace coursedesign.Controllers
{
    public class HomeController : Controller
    {
        DataEntity db = new DataEntity();
       

        public ActionResult Index() 
        {
            string username = Request.Form["username"];
            string password = Request.Form["password"];

            List<user> list = (from c in db.user
                                      where c.username.Trim() == username
                               select c).ToList();

            foreach(user e in list)
            {
                if(e.password.Trim().Equals(password))
                {
                    //TODO:登录成功添加session 
                    Session["username"] = username;
                    return Content("<script>alert('登录成功');window.location.href='../Home/Index';</script>");
                }
                else
                {
                    return Content("<script>alert('用户名或密码错误');history.go(-1);</script>");
                }
            }
            return View();

        }
        [HttpGet]
        public ActionResult QueryData()
        {
            if(Session["username"] == null)
            {
                return Content("<script>alert('请先登录');window.location.href='../Home/Index';</script>");
            }

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
            if (Session["username"] == null)
            {
                return Content("<script>alert('请先登录');window.location.href='../Home/Index';</script>");
            }
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
            if (Session["username"] == null)
            {
                return Content("<script>alert('请先登录');window.location.href='../Home/Index';</script>");
            }
            return View();
        }
        [HttpPost]
        public ActionResult AddData(engineering model)
        {
            if (Session["username"] == null)
            {
                return Content("<script>alert('请先登录');window.location.href='../Home/Index';</script>");
            }
            try
            {
                engineering e = new engineering();
                e.SetValue(e, model);
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
            if (Session["username"] == null)
            {
                return Content("<script>alert('请先登录');window.location.href='../Home/Index';</script>");
            }
            List<engineering> list = (from c in db.engineering select c).ToList();
            ViewData["DataList"] = list;
            var listItem = new List<SelectListItem>
            {
                new SelectListItem{Text="根据编号降序",Value="1"},
                new SelectListItem{Text="根据编号升序",Value="-1"},
                new SelectListItem{Text="根据姓名降序",Value="2"},
                new SelectListItem{Text="根据姓名升序",Value="-2"},
                new SelectListItem{Text="根据工龄降序",Value="3"},
                new SelectListItem{Text="根据工龄升序",Value="-3"}
            };
            ViewBag.orderlist = new SelectList(listItem, "Value", "Text", "");
            return View();
        }
        public ActionResult ShowOrderData()
        {
            if (Session["username"] == null)
            {
                return Content("<script>alert('请先登录');window.location.href='../Home/Index';</script>");
            }
            //TODO:升序降序操作
            List<engineering> list = (from c in db.engineering select c).ToList();
            var option = Request.Form["orderOption"];
            switch(option)
            {
                case "1":
                    list = list.OrderByDescending(t => t.id).ToList();
                    break;
                case "-1":
                    list = list.OrderBy(t => t.id).ToList();
                    break;
                case "2":
                    list = list.OrderByDescending(t => t.name).ToList();
                    break;
                case "-2":
                    list = list.OrderBy(t => t.name).ToList();
                    break;
                case "3":
                    list = list.OrderByDescending(t => t.workage).ToList();
                    break;
                case "-3":
                    list = list.OrderBy(t => t.workage).ToList();
                    break;
                default:
                    break;

            }
            ViewData["DataList"] = list;
            return View();
        }


        public ActionResult CalculateData()
        {
            if (Session["username"] == null)
            {
                return Content("<script>alert('请先登录');window.location.href='../Home/Index';</script>");
            }
            int id = Convert.ToInt32(Request.QueryString["Cid"]);
            engineering engineering = (from e in db.engineering where id == e.id select e).SingleOrDefault();
            return View(engineering);
        }
        [HttpPost]
        public ActionResult CalculateData(engineering engineering)
        {
            if (Session["username"] == null)
            {
                return Content("<script>alert('请先登录');window.location.href='../Home/Index';</script>");
            }
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
            if (Session["username"] == null)
            {
                return Content("<script>alert('请先登录');window.location.href='../Home/Index';</script>");
            }
            var query = (from u in db.engineering
                         select new
                         {
                             u.id,
                             u.name,
                             u.sex,
                             u.education,
                             u.place,
                             u.address,
                             u.telephone,
                             u.workage,
                             u.salary,
                             u.birth
                         }).ToList();


            string p = System.AppDomain.CurrentDomain.BaseDirectory;
            string sWebRootFolder = p+@"\excels\";
            string sFileName = $@"{DateTime.Now.ToString("yyyyMMddHHmmss")}.xlsx";
            var path = Path.Combine(sWebRootFolder, sFileName);
            FileInfo file = new FileInfo(path);
            if(file.Exists)
            {
                file.Delete();
                file = new FileInfo(path);
            }
            ExcelPackage.LicenseContext = LicenseContext.Commercial;
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (ExcelPackage package = new ExcelPackage(file))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("UserInfo");
                worksheet.Cells.LoadFromCollection(query, true);
                package.Save();
            }
            return Content("<script>alert('保存成功');history.go(-1);</script>");
        }
        public ActionResult ClearData()
        {
            if (Session["username"] == null)
            {
                return Content("<script>alert('请先登录');window.location.href='../Home/Index';</script>");
            }
            List<engineering> list = (from c in db.engineering                                 
                                      select c).ToList();
            foreach(engineering e in list)
            {
                db.engineering.Remove(e);
            }
            db.SaveChanges();
            return Content("<script>alert('清除成功');window.location.href='../Home/Index';</script>");
        }

        public ActionResult UpLoadData()
        {
            if (Session["username"] == null)
            {
                return Content("<script>alert('请先登录');window.location.href='../Home/Index';</script>");
            }
            return View();
        }
     
        [HttpPost]
        public ActionResult UploadData(HttpPostedFileBase file)
        {
            if (Session["username"] == null)
            {
                return Content("<script>alert('请先登录');window.location.href='../Home/Index';</script>");
            }
            var fileName = file.FileName;
            var filePath = Server.MapPath(string.Format("~/{0}", "UpLoadExcel"));
            string path = Path.Combine(filePath, fileName);
            FileInfo f = new FileInfo(path);
            if (f.Exists)
            {
                f.Delete();
            }
            file.SaveAs(path);
            ExcelPackage.LicenseContext = LicenseContext.Commercial;
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            List<engineering> es = new List<engineering>();

            using (ExcelPackage package = new ExcelPackage(new FileStream(path, FileMode.Open)))
            {
                //TODO:做一个列表的判断
                ExcelWorksheet sheet = package.Workbook.Worksheets[0];
                //index从1开始
                for(int row = 2;row<= sheet.Dimension.Rows;row++)
                {
                    engineering e = new engineering();
                    for (int colunmn = 1;colunmn<=sheet.Dimension.Columns;colunmn++)
                    {
                        string s = sheet.Cells[1, colunmn].Value.ToString();
                        switch (s)
                        {
                            case "id":
                                e.id = Convert.ToInt32(sheet.Cells[row, colunmn].Value.ToString());
                                break;
                            case "name":
                                e.name = sheet.Cells[row, colunmn].Value.ToString();
                                break;
                            case "sex":
                                e.sex = Convert.ToInt32(sheet.Cells[row, colunmn].Value.ToString());
                                break;
                            case "education":
                                e.education = Convert.ToInt32(sheet.Cells[row, colunmn].Value.ToString());
                                break;
                            case "place":
                                e.place = sheet.Cells[row, colunmn].Value.ToString();
                                break;
                            case "address":
                                e.address = sheet.Cells[row, colunmn].Value.ToString();
                                break;
                            case "telephone":
                                e.telephone = sheet.Cells[row, colunmn].Value.ToString();
                                break;
                            case "workage":
                                e.workage = sheet.Cells[row, colunmn].Value.ToString();
                                break;
                            case "salary":
                                e.salary = Convert.ToInt32(sheet.Cells[row, colunmn].Value.ToString());
                                break;
                            case "birth":
                                e.birth = Convert.ToDateTime(sheet.Cells[row, colunmn].Value.ToString()); 
                                break;
                        }   
                    }
                    es.Add(e);
                }
                this.TempData["list"] = es;
            }
           
            return RedirectToAction("GetData");
        }

        public ActionResult InsertExcelInfo()
        {
            if (Session["username"] == null)
            {
                return Content("<script>alert('请先登录');window.location.href='../Home/Index';</script>");
            }
            List<engineering> es = (List < engineering >) this.TempData["list"];
            try
            {
                foreach (engineering e in es)
                {
                    db.engineering.Add(e);
                    db.Configuration.ValidateOnSaveEnabled = false;
                }
                db.SaveChanges();

                db.Configuration.ValidateOnSaveEnabled = true;
                return Content("<script>alert('添加成功');history.go(-1);</script>");
            }
            catch (Exception ex)
            {
                return Content(ex.Message);
            }
        }



        public ActionResult GetData()
        {
            if (Session["username"] == null)
            {
                return Content("<script>alert('请先登录');window.location.href='../Home/Index';</script>");
            }
            ViewData["UpList"] = this.TempData["list"];
            this.TempData["list"] = this.TempData["list"];
            return View();
        }
        public ActionResult ExitSys()
        {
           
            if (Session["username"] == null)
            {
                return Content("<script>alert('尚未登陆');window.location.href='../Home/Index';</script>");
            }
            else
            {
                Session["username"] = null;
                return Content("<script>alert('退出成功');window.location.href='../Home/Index';</script>");
            }
        }

        [HttpGet]
        public ActionResult Modify()
        {
            if (Session["username"] == null)
            {
                return Content("<script>alert('请先登录');window.location.href='../Home/Index';</script>");
              }
            int id = Convert.ToInt32(Request.QueryString["Cid"]);
            engineering engineering = (from e in db.engineering where id == e.id select e).SingleOrDefault();
            return View(engineering);
        }
        [HttpPost]
        public ActionResult Modify(engineering model)
        {
            if (Session["username"] == null)
            {
                return Content("<script>alert('请先登录');window.location.href='../Home/Index';</script>");
            }
            engineering e = db.engineering.Where(e1 => e1.id == model.id).ToList().FirstOrDefault();
            e.SetValue(e, model);


            db.SaveChanges();
            db.Configuration.ValidateOnSaveEnabled = true;
            return RedirectToAction("EditData", "Home");
        }
        public ActionResult Remove()
        {
            if (Session["username"] == null)
            {
                return Content("<script>alert('请先登录');window.location.href='../Home/Index';</script>");
              }
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