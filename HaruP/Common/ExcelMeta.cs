using System.Collections.Generic;
using System.Linq;
using NPOI.SS.UserModel;

namespace HaruP.Common
{
    public class ExcelMeta
    {
        public Orientation Orientation { get; set; }

        private string _namespace=string.Empty;
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

        internal List<TagMeta> Tags { get; set; }

        public ExcelMeta()
        {
            Tags = new List<TagMeta>();
        }

    }

    public enum Orientation
    {
        Horizontal,
        Vertical
    }

    public class TagMeta
    {
        public string TagId { get; set; }

        public string TemplateText { get; set; }
        public ICell Cell { get; set; }

    }

}
