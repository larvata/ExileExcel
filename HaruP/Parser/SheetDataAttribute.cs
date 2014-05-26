using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;

namespace HaruP.Parser
{
//    [Description("RawData Base Class")]
    public class SheetDataAttribute:Attribute
    {
        private readonly string _headerText;
        private readonly bool _identityColumn;

        public string HeaderText
        {
            get { return _headerText; }
        }

        public bool IdentityColumn
        {
            get { return _identityColumn; }
        }

        public SheetDataAttribute(string headerText, bool identityColumn = false)
        {
            _headerText = headerText;
            _identityColumn = identityColumn;
        }
    }
}
