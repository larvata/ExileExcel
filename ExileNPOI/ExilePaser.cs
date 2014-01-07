using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExileNPOI
{
    public class ExilePaser<T> where T : IExilable, new()
    {
        public T Parse(string path)
        {
            return new T();
        }

        
       
    }
}
