using System;
using System.Collections.Generic;
using System.IO;
using HaruP.Common;

namespace HaruP.PlayGround
{
    internal class Program
    {
        private static void Main(string[] args)
        {
            
            var metaVertical = new ExcelMeta {Orientation = Orientation.Vertical};



            #region anonymous list
            var wotaMixList = new List<dynamic>
            {
                new
                {
                    Tiger = 1,
                    Fire = "Fire1",
                    Cyber = DateTime.Now,
                    Fiber = (double) 2,
                    Diver = (Single) 3,
                    Viber = "Viber1"
                },
                new
                {
                    Tiger = 4,
                    Cyber = DateTime.Now.AddDays(1),
                    Fiber = (double) 5,
                    Diver = (Single) 6,
                    Viber = "Viber2"
                },
                new
                {
                    Tiger = 7,
                    Fire = "Fire3",
                    Cyber = DateTime.Now.AddDays(2),
                    Fiber = (double) 8,
                    Diver = (Single) 9,
                    Viber = "Viber3"
                }
            };

            var f = new ExcelExtractor();
            using (var fs = new FileStream("Vertical.xls", FileMode.Create))
            {
                f.ExcelWriteStream(wotaMixList, fs, @"template\templateVertical.xls", metaVertical);
            }
            #endregion

            #region strong typed list
            var wotaMixList2 = new List<WotaMix>
            {
                new WotaMix
                {
                    Tiger = 10,
                    Fire = "Fire4",
                    Cyber = DateTime.Now.AddDays(3),
                    Fiber = 11,
                    Diver = 12,
                    Viber = "Viber4"
                },
                new WotaMix
                {
                    Tiger = 13,
                    Fire = "Fire5",
                    Cyber = DateTime.Now.AddDays(4),
                    Fiber = 14,
                    Diver = 15,
                    Viber = "Viber5"
                }
            };

            var g = new ExcelExtractor();
            using (var fs = new FileStream("Horizontal.xls", FileMode.Create))
            {
                g.ExcelWriteStream(wotaMixList2, fs, @"template\templateHorizontal.xls");
            }

            #endregion

        }
    }
}
