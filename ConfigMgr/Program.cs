using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConfigMgr
{
    class Program
    {
        static void Main(string[] args)
        {
            var sw = Stopwatch.StartNew();
            var helper = new ConfigMgrHelper(args[0], args[1], args[2], args[3]);
            helper.BuildApplicationTree("6000");

            sw.Stop();
            helper.WriteLine($"Total time for loading tree: {sw.Elapsed}");
            Console.WriteLine("\n\n");

            var helper2= new ConfigMgrHelper(args[0], args[1], args[2], args[3]);
            helper2.BuildRecursive();

            Console.Read();
        }
    }
}
