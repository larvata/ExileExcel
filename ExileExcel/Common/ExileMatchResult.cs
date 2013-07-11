namespace ExileExcel.Common
{
    using System;
    using System.Collections.Generic;

    public class ExileMatchResult
    {
        public Dictionary<string, string> HeaderKeyPair { get; set; }
        public Type MatchedType { get; set; }

        /// <summary>
        /// Description of parsed class
        /// the description is defined by ExiliableAttribute
        /// </summary>
        public string TypeDescription { get; set; }

        public ExileMatchResult()
        {
            HeaderKeyPair=new Dictionary<string, string>();
        }

    }
}
