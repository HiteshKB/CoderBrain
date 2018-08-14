using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ClientConsole
{
    class Program
    {
        public static int Main(String[] args)
        {
            SynchronousSocketClient.StartClient();
            return 0;
        }
    }
}
