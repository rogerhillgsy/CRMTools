using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace UpdateOrgsCASLBatch
{
    class Program
    {
        static void Main(string[] args)
        {
            Helper helper = new Helper();
            helper.connectToCRM();

        }
    }
}
