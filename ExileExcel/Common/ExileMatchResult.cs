using System;
using System.Collections.Generic;

namespace ExileExcel.Common
{
    public class ExileMatchResult
    {
        public Dictionary<string, string> HeaderKeyPair { get; set; }
        public Type MatchedType { get; set; }

        /// <summary>
        /// Description of parsed class
        /// the description is defined by ExiliableAttribute
        /// </summary>
        public string TypeDescription { get; set; }

    }
}
