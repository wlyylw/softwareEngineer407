using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace CourseDesign.Models
{
    public class ActionFilter : ActionFilterAttribute
    {
        public bool IsLogin { get; set; }


        public override void OnActionExecuting(ActionExecutingContext filterContext)

        {

            var varget = filterContext.HttpContext.Session["username"];

            if (IsLogin == false)

            {

                if (varget == null)

                {
                    HttpContext.Current.Response.Write("<script>alert('请先登录');window.location.href='../Home/Index';</script>");
                }

                return;

            }

        }

    }
}