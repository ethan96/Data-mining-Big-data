using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ProcessABTest
{
    class Program
    {
        static void Main(string[] args)
        {
            UNICA_API.IntApi ws = new UNICA_API.IntApi();
            ws.ProcessABTestJob();
        }
    }
}
