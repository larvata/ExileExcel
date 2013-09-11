using System.Collections.Generic;

namespace ExileNPOI.Interface
{
    interface IExtractor<T> where T:IExilable
    {
        void FillContent(IList<T> dataList);
    }
}
