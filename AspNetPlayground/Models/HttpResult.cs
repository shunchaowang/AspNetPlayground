namespace AspNetPlayground.Models
{
    public class HttpResult<T>
    {
        public int Code { get; set; }

        public string Msg { get; set; }

        public T Data { get; set; }

        public static HttpResult<T> GetResult(int code, string msg, T data = default(T))
        {
            return new HttpResult<T>
            {
                Code = code,
                Msg = msg,
                Data = data
            };
        }
    }
}