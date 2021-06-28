using System;
using System.Collections.Generic;
using System.Text;

namespace DesafioRPA.classes
{
    public class ResponseResult<T>
    {
        public string erro { get; set; }
        public string mensagem { get; set; }
        public int total { get; set; }
        public List<T> dados { get; set; }
    }
}
