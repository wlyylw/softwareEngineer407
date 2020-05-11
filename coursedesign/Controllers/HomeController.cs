using CourseDesign.Models;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Web;
using System.Web.Mvc;
using System.Web.UI.WebControls;



namespace CourseDesign.Controllers
{
    public class HomeController : Controller
    {
       
        DataModel db = new DataModel();

        [ActionFilter(IsLogin = true)]
        public ActionResult Index()
        {
            string username = Request.Form["username"];
            string password = Request.Form["password"];

            List<User> list = (from c in db.User
                               where c.Username.Trim() == username
                               select c).ToList();
            int flag = 0;
            foreach (User e in list)
            {
                if (e.Password.Trim().Equals(password))
                {
                    //TODO:登录成功添加session 
                    Session["username"] = username;
                    flag = 1;
                    return Content("<script>alert('登录成功');window.location.href='../Home/Index';</script>");
                }

            }
            if (flag == 0 && username != null)
            {
                return Content("<script>alert('用户名或密码错误');history.go(-1);</script>");
            }
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
                List<Engineering> list = (from c in db.Engineering
                                          where c.Name == value
                                          select c).ToList();
                ViewData["QueryList"] = list;
            }
            else if (option.Equals("asid"))
            {
                int value = Convert.ToInt32(Request.Form["queryType"]);
                List<Engineering> list = (from c in db.Engineering
                                          where c.Id == value
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
        public ActionResult AddData(Engineering model)
        {
            try
            {
                if (ModelState.IsValid)
                {
                    Engineering e = new Engineering();
                    e.Id = model.Id;
                    e.Name = model.Name.Trim();
                    e.Place = model.Place.Trim();
                    e.Salary = model.Salary;
                    e.Sex = model.Sex;
                    e.Telephone = model.Telephone.Trim();
                    e.Workage = model.Workage;
                    e.Address = model.Address.Trim();
                    e.Education = model.Education;
                    e.Birth = model.Birth;

                    db.Engineering.Add(e);
                    db.SaveChanges();
                    db.Configuration.ValidateOnSaveEnabled = true;
                    return RedirectToAction("AddData", "Home");

                }
                else
                {
                    return View();
                }
            }
            catch (Exception)
            {
                return Content("出现问题，数据添加失败");
            }
        }
        //TODO:

        public ActionResult EditData()
        {

            List<Engineering> list = (from c in db.Engineering select c).ToList();
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

            //TODO:升序降序操作
            List<Engineering> list = (from c in db.Engineering select c).ToList();
            var option = Request.Form["orderOption"];
            switch (option)
            {
                case "1":
                    list = list.OrderByDescending(t => t.Id).ToList();
                    break;
                case "-1":
                    list = list.OrderBy(t => t.Id).ToList();
                    break;
                case "2":
                    list = list.OrderByDescending(t => t.Name).ToList();
                    break;
                case "-2":
                    list = list.OrderBy(t => t.Name).ToList();
                    break;
                case "3":
                    list = list.OrderByDescending(t => t.Workage).ToList();
                    break;
                case "-3":
                    list = list.OrderBy(t => t.Workage).ToList();
                    break;
                default:
                    break;

            }
            ViewData["DataList"] = list;
            return View();
        }

        public ActionResult CalculateData()
        {

            int id = Convert.ToInt32(Request.QueryString["Cid"]);
            Engineering engineering = (from e in db.Engineering where id == e.Id select e).SingleOrDefault();
            return View(engineering);
        }

        [HttpPost]
        public ActionResult CalculateData(Engineering engineering)
        {

            int basicsalary = Convert.ToInt32(engineering.Salary);
            int workage = Convert.ToInt32(engineering.Workage);
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
            var query = (from u in db.Engineering
                         select new
                         {
                             u.Id,
                             u.Name,
                             u.Sex,
                             u.Education,
                             u.Place,
                             u.Address,
                             u.Telephone,
                             u.Workage,
                             u.Salary,
                             u.Birth
                         }).ToList();

            List<Engineering> list = (from c in db.Engineering select c).ToList();

            string sWebRootFolder = System.AppDomain.CurrentDomain.BaseDirectory + @"\excels\";
            string sFileName = $@"{DateTime.Now.ToString("yyyyMMddHHmmss")}.xlsx";
            //把项目名加到指定存放的路径
            FileInfo file = new FileInfo(Path.Combine(sWebRootFolder, sFileName));
            if (file.Exists)
            {
                file.Delete();
                file = new FileInfo(Path.Combine(sWebRootFolder, sFileName));
            }
            ExcelPackage.LicenseContext = LicenseContext.Commercial;

            // If you use EPPlus in a noncommercial context
            // according to the Polyform Noncommercial license:
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (ExcelPackage package = new ExcelPackage(file))
            {
                //添加worksheet的名字
                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("工程师管理系统");
                //添加表头名字
                worksheet.Cells[1, 1].Value = "编号";
                worksheet.Cells[1, 2].Value = "姓名";
                worksheet.Cells[1, 3].Value = "性别";
                worksheet.Cells[1, 4].Value = "学历";
                worksheet.Cells[1, 5].Value = "籍贯";
                worksheet.Cells[1, 6].Value = "地址";
                worksheet.Cells[1, 7].Value = "电话";
                worksheet.Cells[1, 8].Value = "工龄";
                worksheet.Cells[1, 9].Value = "工资";
                worksheet.Cells[1, 10].Value = "出生日期";
                //添加值
                var rowNum = 2; // rowNum 1 is head
                foreach (var message in list as List<Engineering>)
                {
                    worksheet.Cells["A" + rowNum].Value = message.Id;
                    worksheet.Cells["B" + rowNum].Value = message.Name;
                    switch (message.Sex)
                    {
                        case 0:
                            worksheet.Cells["C" + rowNum].Value = "女";
                            break;
                        case 1:
                            worksheet.Cells["C" + rowNum].Value = "男";
                            break;
                        default:
                            worksheet.Cells["C" + rowNum].Value = "  ";
                            break;
                    }
                    switch (message.Education)
                    {
                        case 0:
                            worksheet.Cells["D" + rowNum].Value = "高中";
                            break;
                        case 1:
                            worksheet.Cells["D" + rowNum].Value = "学士";
                            break;
                        case 2:
                            worksheet.Cells["D" + rowNum].Value = "硕士";
                            break;
                        case 3:
                            worksheet.Cells["D" + rowNum].Value = "博士";
                            break;
                        case 4:
                            worksheet.Cells["D" + rowNum].Value = "其他";
                            break;
                        default:
                            worksheet.Cells["D" + rowNum].Value = "无学历";
                            break;
                    }
                    worksheet.Cells["E" + rowNum].Value = message.Place;
                    worksheet.Cells["F" + rowNum].Value = message.Address;
                    worksheet.Cells["G" + rowNum].Value = message.Telephone;
                    worksheet.Cells["H" + rowNum].Value = message.Workage;
                    worksheet.Cells["I" + rowNum].Value = message.Salary;
                    worksheet.Cells["J" + rowNum].Value = message.Birth.ToString();
                    rowNum++;
                }
                package.Save();
            }
            return Content("<script>alert('保存成功');history.go(-1);</script>");
        }

        public ActionResult ClearData()
        {

            List<Engineering> list = (from c in db.Engineering
                                      select c).ToList();
            foreach (Engineering e in list)
            {
                db.Engineering.Remove(e);
            }
            db.SaveChanges();
            return Content("<script>alert('清除成功');window.location.href='../Home/Index';</script>");
        }
        public ActionResult UpLoadData()
        {
            return View();
        }

        [HttpPost]
        public ActionResult UploadData(HttpPostedFileBase file)
        {
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
            List<Engineering> es = new List<Engineering>();

            using (ExcelPackage package = new ExcelPackage(new FileStream(path, FileMode.Open)))
            {
                //TODO:做一个列表的判断
                ExcelWorksheet sheet = package.Workbook.Worksheets[0];
                //index从1开始
                for (int row = 2; row <= sheet.Dimension.Rows; row++)
                {
                    Engineering e = new Engineering();
                    for (int colunmn = 1; colunmn <= sheet.Dimension.Columns; colunmn++)
                    {
                        string s = sheet.Cells[1, colunmn].Value.ToString();
                        switch (s)
                        {
                            case "编号":
                                e.Id = Convert.ToInt32(sheet.Cells[row, colunmn].Value.ToString());
                                break;
                            case "姓名":
                                e.Name = sheet.Cells[row, colunmn].Value.ToString();
                                break;
                            case "性别":
                                if (sheet.Cells[row, colunmn].Value.ToString().Equals("男"))
                                    e.Sex = 1;
                                else
                                    e.Sex = 0;
                                break;
                            case "学历":
                                if (sheet.Cells[row, colunmn].Value.ToString().Equals("高中"))
                                    e.Education = 0;
                                else if (sheet.Cells[row, colunmn].Value.ToString().Equals("学士"))
                                    e.Education = 1;
                                else if (sheet.Cells[row, colunmn].Value.ToString().Equals("硕士"))
                                    e.Education = 2;
                                else if (sheet.Cells[row, colunmn].Value.ToString().Equals("博士"))
                                    e.Education = 3;
                                else if (sheet.Cells[row, colunmn].Value.ToString().Equals("其他"))
                                    e.Education = 4;
                                break;
                            case "籍贯":
                                e.Place = sheet.Cells[row, colunmn].Value.ToString();
                                break;
                            case "地址":
                                e.Address = sheet.Cells[row, colunmn].Value.ToString();
                                break;
                            case "电话":
                                e.Telephone = sheet.Cells[row, colunmn].Value.ToString();
                                break;
                            case "工龄":
                                e.Workage = Convert.ToInt32(sheet.Cells[row, colunmn].Value.ToString());
                                break;
                            case "工资":
                                e.Salary = Convert.ToInt32(sheet.Cells[row, colunmn].Value.ToString());
                                break;
                            case "出生日期":
                                e.Birth = Convert.ToDateTime(sheet.Cells[row, colunmn].Value.ToString());
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
            List<Engineering> es = (List<Engineering>)this.TempData["list"];
            try
            {
                foreach (Engineering e in es)
                {
                    db.Engineering.Add(e);
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
            ViewData["UpList"] = this.TempData["list"];
            this.TempData["list"] = this.TempData["list"];
            return View();
        }
        public ActionResult ExitSys()
        {
            Session["username"] = null;
            return Content("<script>alert('退出成功');window.location.href='../Home/Index';</script>");
        }
        [HttpGet]
        public ActionResult Modify()
        {
            int id = Convert.ToInt32(Request.QueryString["Cid"]);
            Engineering engineering = (from e in db.Engineering where id == e.Id select e).SingleOrDefault();
            return View(engineering);
        }
        [HttpPost]
        public ActionResult Modify(Engineering model)
        {
            Engineering e = db.Engineering.Where(e1 => e1.Id == model.Id).ToList().FirstOrDefault();
            e.SetValue(e, model);
            db.Configuration.ValidateOnSaveEnabled = false;
            db.SaveChanges();
            db.Configuration.ValidateOnSaveEnabled = true;
            return RedirectToAction("EditData", "Home");
        }
        public ActionResult Remove()
        {
            try
            {
                int id = Convert.ToInt32(Request.QueryString["Cid"]);
                Engineering engineering = (from e in db.Engineering where e.Id == id select e).SingleOrDefault();
                db.Engineering.Remove(engineering);
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