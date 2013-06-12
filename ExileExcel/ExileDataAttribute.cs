using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;

namespace ExileExcel
{
    [Description("RawData Base Class")]
    public class ExileDataAttribute : Attribute
    {
        private readonly string _headerText;
        private readonly bool _identityColumn;


        public string HeaderText
        {
            get { return this._headerText; }
        }

        public bool IdentityColumn
        {
            get { return _identityColumn; }
        }

        public ExileDataAttribute(string headerText, bool identityColumn = false)
        {
            this._headerText = headerText;
            this._identityColumn = identityColumn;
        }
    }
}
