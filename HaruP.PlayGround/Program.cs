using System;
using System.Collections.Generic;
using HaruP.Common;

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
                    formula="(1+2)"
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
                    formula="(3+4)"
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
                    ProjectType = "项目类别a678678678678567467tdy5du6ur65udr65yhdr56udr6u6ud6d6udc6u",
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
                Count = list.Count,
                Count2=9
            };
            var sstat = new
            {
                Count = list.Count ,
                Count2=8
            };

            var extractor = new ExcelExtractor(@"template\t2.xls");
            var meta = new SheetMeta
            {
                RowHeight = RowHeight.Inherit
            };

            var smeta = new SheetMeta
            {
                Namespace = "S",
            };

            extractor.ForkSheet(0, "lalala")
                .PutData(list2, smeta)
                .PutData(sstat, smeta)
                .PutData(list, meta)
                .PutData(stat);


            extractor.Write("o3.xls");


            #endregion



        }
    }
}
