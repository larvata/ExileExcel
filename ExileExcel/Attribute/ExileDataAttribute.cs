using System;
using System.ComponentModel;

namespace ExileExcel.Attribute
{
    [Description("RawData Base Class")]
    [AttributeUsage(AttributeTargets.Property)]
    public class ExilePropertyAttribute : System.Attribute
    {
        private readonly string _headerText;
        private readonly bool _identityColumn;
        private readonly int _headerTextSequence;

        /// <summary>
        /// Excel sheet header text
        /// </summary>
        public string HeaderText
        {
            get { return _headerText; }
        }

        /// <summary>
        /// may never used
        /// </summary>
        public bool IdentityColumn
        {
            get { return _identityColumn; }
        }

        /// <summary>
        /// Output sequence
        /// RULE: HeaderTextSequence!=-1 ASC + HeaderTextSequence==-1 
        /// </summary>
        public int HeaderTextSequence
        {
            get { return _headerTextSequence; }
        }

        public ExilePropertyAttribute(string headerText, int headerTextSequence=-1, bool identityColumn = false)
        {
            _headerTextSequence = headerTextSequence;
            _headerText = headerText;
            _identityColumn = identityColumn;
        }
    }
}
