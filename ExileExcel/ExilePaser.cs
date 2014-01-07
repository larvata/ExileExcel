namespace ExileExcel
{
    public class ExilePaser<T> where T : IExilable, new()
    {
        public T Parse(string path)
        {
            return new T();
        }

        
       
    }
}
