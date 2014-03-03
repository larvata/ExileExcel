using System.Collections.Generic;

namespace HaruP.PlayGround
{
    internal class Program
    {
        private static void Main(string[] args)
        {

            #region anonymous list
            var list = new List<dynamic>
            {
                new
                {
                    Id = 1,
                    Company = "单位1",
                    Project = "项目",
                    Town = "镇",
                    ProjectType = "项目类别",
                    Amount = 100,
                    Region = 900,
                    Progress = "呵呵呵",
                    Target = "哈哈哈",
                    Month1 = "一月内容",
                }
                ,
                new
                {
                    Id = 2,
                    Company = "单位2",
                    Project = "项目2",
                    Town = "镇2",
                    ProjectType = "项目类别2",
                    Amount = 1002,
                    Region = 9002,
                    Progress = "呵呵呵2",
                    Target = "哈哈哈2",
                    Month1 = "一月内容2",
                }
            };

            var list2 = new List<dynamic>
            {
                new
                {
                    Id = 1,
                    Company = "单位a",
                    Project = "项目a",
                    Town = "镇a",
                    ProjectType = "项目类别a",
                    Amount = 1009,
                    Region = 9009,
                    Progress = "呵呵呵a",
                    Target = "哈哈哈a",
                    Month1 = "一月内容a",
                    Month2 = "二月内容a",
                    Month3 = "三月内容2b"
                }
                ,
                new
                {
                    Id = 2,
                    Company = "单位2b",
                    Project = "项目2b",
                    Town = "镇2b",
                    ProjectType = "项目类别2b",
                    Amount = 10029,
                    Region = 90029,
                    Progress = "呵呵呵2b",
                    Target = "哈哈哈2b",
                    Month1 = "一月内容2b",
                    Month2 = "二月内容2b",
                    Month3 = "三月内容2b"
                }
            };

            var stat = new
            {
                Count =  list.Count+" 个 呵呵"
            };
            var sstat = new
            {
                Count = list.Count+" 个 呵呵"
            };

            var extractor = new ExcelExtractor(@"template\t2.xls");
            extractor.PutData(list);
            extractor.PutData(list2,"S");
            extractor.PutData(stat);
            extractor.PutData(sstat,"S");
            extractor.Write("o2.xls");
//
            #endregion

            #region strong typed

//
//            var t = new Mytarget()
//            {
//                year=2014,
//                month = 12,
//                days = 13,
//                Id = "Id",
//                AnnualTarget = "AnnualTarget",
//                PName = "PName",
//                ConstructionUnit = "ConstructionUnit",
//                Town = "Town",
//                Type = "Type",
//                TotalInvestment = "TotalInvestment",
//                GFA = "GFA",
//                Progress = "Progress",
//                January = "January",
//                February = "February",
//                March = "March",
//                April = "April",
//                May = "May",
//                June = "June",
//                July = "July",
//                August = "August",
//                September = "September",
//                October = "October",
//                November = "November",
//                December = "December"
//            };
//            var ext = new ExcelExtractor(@"template\t1.xls");
//            ext.PutData(t);
//            ext.Write("03.xls");

            #endregion
        }
    }
}
