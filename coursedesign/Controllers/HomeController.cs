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
                if (ModelState.IsValid)
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
                else
                {
                    return View();
                }
            }
            catch (Exception)
            {
                //指定对应跳转的视图Test下的Test.cshtml文件
                /*           return RedirectToAction("AddData", "AddData");*/
                return Content("出现问题，数据添加失败");
            }
        }

        public ActionResult EditData()
        {
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
            //var query = (from u in db.engineering
            //             select new
            //             {
            //                 u.id,
            //                 u.name,
            //                 u.sex,
            //                 u.education,
            //                 u.place,
            //                 u.address,
            //                 u.telephone,
            //                 u.workage,
            //                 u.salary,
            //                 u.birth
            //             }).ToList();


            //string p = System.AppDomain.CurrentDomain.BaseDirectory;
            //string sWebRootFolder = p+@"\excels\";
            //string sFileName = $@"{DateTime.Now.ToString("yyyyMMddHHmmss")}.xlsx";
            //var path = Path.Combine(sWebRootFolder, sFileName);
            //FileInfo file = new FileInfo(path);
            //if(file.Exists)
            //{
            //    file.Delete();
            //    file = new FileInfo(path);
            //}
            //ExcelPackage.LicenseContext = LicenseContext.Commercial;
            //ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            //using (ExcelPackage package = new ExcelPackage(file))
            //{
            //    ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("UserInfo");
            //    worksheet.Cells.LoadFromCollection(query, true);
            //    package.Save();
            //}

            List<engineering> list = (from c in db.engineering select c).ToList();
            //ViewData["DataList"] = list;
            ////指定项目存放的路径
            //string sWebRootFolder = @"E:/";
            ////指定项目名字
            //string sFileName = "工程师管理系统" + $"{Guid.NewGuid()}.xlsx";
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
                foreach (var message in list as List<engineering>)
                {
                    worksheet.Cells["A" + rowNum].Value = message.id;
                    worksheet.Cells["B" + rowNum].Value = message.name;
                    switch (message.sex)
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
                    switch (message.education)
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
                    worksheet.Cells["E" + rowNum].Value = message.place;
                    worksheet.Cells["F" + rowNum].Value = message.address;
                    worksheet.Cells["G" + rowNum].Value = message.telephone;
                    worksheet.Cells["H" + rowNum].Value = message.workage;
                    worksheet.Cells["I" + rowNum].Value = message.salary;
                    worksheet.Cells["J" + rowNum].Value = message.birth.ToString();
                    rowNum++;
                }
                package.Save();
            }
            return Content("<script>alert('保存成功');history.go(-1);</script>");
        }
        public ActionResult ClearData()
        {
            return View();
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
                            case "编号":
                                e.id = Convert.ToInt32(sheet.Cells[row, colunmn].Value.ToString());
                                break;
                            case "姓名":
                                e.name = sheet.Cells[row, colunmn].Value.ToString();
                                break;
                            case "性别":
                                if (sheet.Cells[row, colunmn].Value.ToString().Equals("男"))
                                    e.sex = 1;
                                else
                                    e.sex = 0;
                                break;
                            case "学历":
                                if (sheet.Cells[row, colunmn].Value.ToString().Equals("高中"))
                                    e.education = 0;
                                else if(sheet.Cells[row, colunmn].Value.ToString().Equals("学士"))
                                    e.education = 1;
                                else if (sheet.Cells[row, colunmn].Value.ToString().Equals("硕士"))
                                    e.education = 2;
                                else if (sheet.Cells[row, colunmn].Value.ToString().Equals("博士"))
                                    e.education = 3;
                                else if (sheet.Cells[row, colunmn].Value.ToString().Equals("其他"))
                                    e.education = 4;
                                break;
                            case "籍贯":
                                e.place = sheet.Cells[row, colunmn].Value.ToString();
                                break;
                            case "地址":
                                e.address = sheet.Cells[row, colunmn].Value.ToString();
                                break;
                            case "电话":
                                e.telephone = sheet.Cells[row, colunmn].Value.ToString();
                                break;
                            case "工龄":
                                e.workage = sheet.Cells[row, colunmn].Value.ToString();
                                break;
                            case "工资":
                                e.salary = Convert.ToInt32(sheet.Cells[row, colunmn].Value.ToString());
                                break;
                            case "出生日期":
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
            ViewData["UpList"] = this.TempData["list"];
            this.TempData["list"] = this.TempData["list"];
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
            e.SetValue(e, model);


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