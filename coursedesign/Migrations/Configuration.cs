namespace CourseDesign.Migrations
{
    using CourseDesign.Models;
    using System;
    using System.Data.Entity;
    using System.Data.Entity.Migrations;
    using System.Linq;

    internal sealed class Configuration : DbMigrationsConfiguration<CourseDesign.DataModel>
    {
        public Configuration()
        {
            AutomaticMigrationsEnabled = true;
        }

        protected override void Seed(CourseDesign.DataModel context)
        {
            //可以初始化数据
            context.User.AddOrUpdate(new User() { Password = "123456", Username = "admin" });
            context.Engineering.AddOrUpdate(new Engineering()
            {
                Id = 1,
                Sex = 0,
                Name = "初始化1",
                Place = "浙江",
                Education = 1,
                Address = "杭州",
                Telephone = "17326081111",
                Workage = 10,
                Birth = new DateTime(1999, 9, 9),
                Salary = 1000
            }) ;
            context.SaveChanges();
        }
    }
}

