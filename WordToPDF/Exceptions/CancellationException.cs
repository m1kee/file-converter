using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WordToPDF.Exceptions
{
    public class CancellationException : Exception
    {
        public string Method { get; set; }
        private string _message { get; set; }

        public override string Message
        {
            get
            {
                if (string.IsNullOrWhiteSpace(Method))
                    return _message;
                else
                    return $"[{Method}] - {_message}";
            }
        }

        public CancellationException(string method, string message) : base(message) {
            Method = method;
            _message = message;
        }
        public CancellationException(string message) : base(message) {
            _message = message;
        }
        public CancellationException(string message, Exception innerException) : base(message, innerException) {
            _message = message;
        }
    }
}
