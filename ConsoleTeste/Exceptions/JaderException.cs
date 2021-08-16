using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleTeste.Exceptions
{
    [Serializable]
    public class JaderException : Exception
    {
        public JaderException() : base() { }
        public JaderException(string message) : base(message) { }
    }
}
