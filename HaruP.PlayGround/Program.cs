namespace HaruP.PlayGround
{
    internal class Program
    {
        private static void Main(string[] args)
        {

            #region anonymous list
//            var list = new List<dynamic>
//            {
//                new
//                {
//                    Id = 1,
//                    Company = "单位1",
//                    Project = "项目",
//                    Town = "镇",
//                    ProjectType = "项目类别",
//                    Amount = 100,
//                    Region = 900,
//                    Progress = "呵呵呵",
//                    Target = "哈哈哈",
//                    Month1 = "一月内容",
//                }
//                ,
//                new
//                {
//                    Id = 2,
//                    Company = "单位2",
//                    Project = "项目2",
//                    Town = "镇2",
//                    ProjectType = "项目类别2",
//                    Amount = 1002,
//                    Region = 9002,
//                    Progress = "呵呵呵2",
//                    Target = "哈哈哈2",
//                    Month1 = "一月内容2",
//                }
//            };
//
//            var list2 = new List<dynamic>
//            {
//                new
//                {
//                    Id2 = 1,
//                    Company2 = "单位a",
//                    Project2 = "项目a",
//                    Town2 = "镇a",
//                    ProjectType2 = "项目类别a",
//                    Amount2 = 1009,
//                    Region2 = 9009,
//                    Progress2 = "呵呵呵a",
//                    Target2 = "哈哈哈a",
//                    Month1_2 = "一月内容a",
//                    Month2_2 = "二月内容a",
//                    Month3_2 = "三月内容2b"
//                }
//                ,
//                new
//                {
//                    Id2 = 2,
//                    Company2 = "单位2b",
//                    Project2 = "项目2b",
//                    Town2 = "镇2b",
//                    ProjectType2 = "项目类别2b",
//                    Amount2 = 10029,
//                    Region2 = 90029,
//                    Progress2 = "呵呵呵2b",
//                    Target2 = "哈哈哈2b",
//                    Month1_2 = "一月内容2b",
//                    Month2_2 = "二月内容2b",
//                    Month3_2 = "三月内容2b"
//                }
//            };
//
//            var stat = new
//            {
//                PCount = string.Format("区配合项目（{0}个）", list.Count),
//                SCount = string.Format("区实施项目（{0}个）", list2.Count),
//            };
//
//
//            var extractor = new ExcelExtractor(@"template\t2.xls");
//            extractor.PutData(list);
//            extractor.PutData(list2);
//            extractor.PutData(stat);
//
//            extractor.Write("o2.xls");
//
            #endregion

            #region strong typed

            var t = new Mytarget()
            {
                Id = "Id",
                AnnualTarget = "AnnualTarget",
                PName = "PName",
                ConstructionUnit = "ConstructionUnit",
                Town = "Town",
                Type = "Type",
                TotalInvestment = "TotalInvestment",
                GFA = "GFA",
                Progress = "Progress",
                January = "January",
                February = "February",
                March = "March",
                April = "April",
                May = "May",
                June = "June",
                July = "July",
                August = "August",
                September = "September",
                October = "October",
                November = "November",
                December = "December"
            };
            var ext = new ExcelExtractor(@"template\t1.xls");
            ext.PutData(t);
            ext.Write("03.xls");

            #endregion
        }
    }
}
