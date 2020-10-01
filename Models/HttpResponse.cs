namespace AspNetPlayground.Models
{
    public class HttpResponse<T>
    {
        public int Code { get; set; }

        public string Msg { get; set; }

        public T Data { get; set; }

        public static HttpResponse<T> GetResult(int code, string msg, T data = default(T))
        {
            return new HttpResponse<T>
            {
                Code = code,
                Msg = msg,
                Data = data
            };
        }
    }
}