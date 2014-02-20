using System.Collections.Generic;
using NPOI.SS.UserModel;

namespace HaruP.Common
{
    public class TemplateMeta
    {
        internal Dictionary<string, ICell> Tags { get; set; }

        public TemplateMeta() 
        {
            Tags=new Dictionary<string,ICell>();
        }
    }
}
