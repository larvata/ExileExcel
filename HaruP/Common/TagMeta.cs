using NPOI.SS.UserModel;

namespace HaruP.Common
{
    public class TagMeta
    {
        public string TagId { get; set; }

        public string TemplateText { get; set; }
        public ICell Cell { get; set; }

        public TagType TagType { get; set; }

    }
}
