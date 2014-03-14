using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace HaruP.Common
{
    public class SheetMeta
    {
        internal List<TagMeta> Tags { get; set; }

        public Orientation Orientation { get; set; }

        public RowHeight RowHeight { get; set; }

        private string _namespace = string.Empty;
        public string Namespace
        {
            get
            {
                if (string.IsNullOrEmpty(_namespace)) return string.Empty;
                return _namespace.LastIndexOf('.') == _namespace.Length
                    ? _namespace
                    : _namespace + '.';
            }
            set
            {
                this._namespace = value;
            }
        }

        public SheetMeta()
        {
            Tags=new List<TagMeta>();
        }
    }

    public enum Orientation
    {
        Horizontal,
        Vertical
    }

    public enum RowHeight
    {
        Auto,
        Inherit
    }
}
