using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace project.Models
{
    public class Engineer
    {

        public DateTime Birth { get; set; }
        public int Sex { get; set; }//0 女/1 男
        public string Name { get; set; }
        public int Id { get; set; }
        public string Place { get; set; }
        public int Education { get; set; }//高中0、学士1、硕士2、博士3、其它为4
        public string Address { get; set; }
        public int Telephone { get; set; }
        public string Workage { get; set; }//(0,50]。
        public int Salary { get; set; }//不能为0
    }
}