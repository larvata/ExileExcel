namespace ExileNPOI.Interface
{
    using System.Collections.Generic;

    interface IExtractor<T> where T:IExilable
    {
        void FillContent(IList<T> dataList);
    }
}
