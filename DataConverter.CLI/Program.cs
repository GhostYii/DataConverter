using DataConverter.Core;
using System;

namespace DataConverter.CLI
{
    using Console = System.Console;
    public class CLI
    {
        private static string _Version => "0.0.1";
        private static ExcelConverter _convert = new ExcelConverter();

        // usage: --tojson -f filename -n "sheetName"
        //        --tojson -f filename -i sheetIndex
        //        --tojson -f filename
        //        --tojson -l filename, filename, ...
        //        --tojson -d dir
        //        --tocs -f filename -i sheetIndex -n structName
        //        --tocs -f filename -n structName
        //        --tocs -l filename, filename, ...
        //        --tocs -d dir
        public static void Main(params string[] args)
        {
            Commands.RegisterAllCommandByType(typeof(CLI));
            Console.Title = $"DCT.CLI v{_Version}";

            if (args.Length > 1)
            {
                string cmd = args[0];
                List<string> parms = new List<string>();
                for (int i = 1; i < args.Length; ++i)
                {
                    parms.Add(args[i]);
                }

                Execute(cmd, parms.ToArray());                
            }

            while (true)
            {
                var line = Console.ReadLine();
                if (string.IsNullOrEmpty(line))
                    return;

                args = line.Split(' ', StringSplitOptions.RemoveEmptyEntries);

                if (args.Length < 1)
                    continue;

                List<string> parms = new List<string>();
                for (int i = 1; i < args.Length; ++i)
                {
                    parms.Add(args[i]);
                }

                Execute(args[0], parms.ToArray());
            }
            

        }

        private static void Execute(string name, params string[] args)
        {
            DateTime begin = DateTime.Now;
            
            Commands.Execute(name, args);

            //End(begin);
            var span = DateTime.Now - begin;
            Console.WriteLine($"excute over, cost {span.Milliseconds} ms.");
        }

        private static void End(DateTime begin)
        {
            var span = DateTime.Now - begin;
            Console.WriteLine($"excute over, cost {span.Milliseconds} ms.");
        }

        [CMD("tojson", "convert filename to json and save at savepath")]
        private static void ToJson(string filename, string sheetName, string savePath)
        {

        }

        [CMD]
        private static void ToJson(string filename, int sheetIndex, string savePath)
        {

        }

        [CMD]
        private static void ToJson(string dir)
        {

        }

        [CMD]
        private static void Exit()
        {
            Environment.Exit(0);
        }
    }
}