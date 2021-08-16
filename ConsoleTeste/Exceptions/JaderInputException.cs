using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleTeste.Exceptions
{
    public class JaderInputException : Exception
    {
        public JaderInputException() : base() { }
        public JaderInputException(string message) : base(message) { }
    }
}
