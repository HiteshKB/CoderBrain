using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ServerConsole
{
    class Program
    {
        public static int Main(String[] args)
        {
            SynchronousSocketListener.StartListening();
            return 0;
        }
    }
}
