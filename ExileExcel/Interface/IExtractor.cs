using System.Collections.Generic;

namespace ExileExcel.Interface
{
    interface IExtractor<T> where T:IExilable
    {
        void FillContent(IList<T> dataList);
    }
}
