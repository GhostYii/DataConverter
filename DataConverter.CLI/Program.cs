using DataConverter.Core;
using System.Text;

namespace DataConverter.CLI
{
    using Console = System.Console;
    public class CLI
    {
        private static string _Version => "0.0.1";
        private static ExcelConverter _convert = new ExcelConverter();

        public static void Main(params string[] args)
        {
            Commands.RegisterAllCommandByType(typeof(CLI));
            Console.Title = $"DCT.CLI v{_Version}";

            Core.Console.AddPrintListener(msg => { Console.ResetColor(); Console.WriteLine(msg); });
            Core.Console.AddErrorListener(msg => { Console.ForegroundColor = ConsoleColor.Red; Console.WriteLine(msg); Console.ResetColor(); });
            Core.Console.AddWarningListener(msg => { Console.ForegroundColor = ConsoleColor.Yellow; Console.WriteLine(msg); Console.ResetColor(); });            

            if (args.Length >= 1)
            {
                string cmd = args[0];
                List<string> parms = new List<string>();
                for (int i = 1; i < args.Length; ++i)
                {
                    parms.Add(args[i]);
                }

                Execute(cmd, parms.ToArray());
            }
            else
            {
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
            filename = Path.GetFullPath(filename);

            string jsonStr = _convert.ToJson(filename, sheetName);
            File.WriteAllText(savePath, jsonStr, Encoding.UTF8);
            Console.WriteLine($"{filename}/{sheetName} convert to {savePath}");
        }

        [CMD]
        private static void ToJson(string filename, int sheetIndex, string savePath)
        {
            filename = Path.GetFullPath(filename);

            string jsonStr = _convert.ToJson(filename, sheetIndex);
            File.WriteAllText(savePath, jsonStr, Encoding.UTF8);
            Console.WriteLine($"{filename}/{sheetIndex}(index) convert to {savePath}");
        }

        [CMD]
        private static void ToJson(string dir, string saveDir)
        {
            dir = Path.GetFullPath(dir);

            if (string.IsNullOrEmpty(saveDir))
                saveDir = dir;
            else
                saveDir = Path.GetFullPath(saveDir);

            var files = Directory.GetFiles(dir, "*.xlsx");
            foreach (var file in files)
            {
                var sheets = ExcelHelper.GetWorksheetNames(file);
                foreach (var name in sheets)
                {
                    string savePath = Path.Combine(saveDir, $"{Path.GetFileNameWithoutExtension(file)}.{name}.json");
                    File.WriteAllText(savePath, _convert.ToJson(file, name));
                    Console.WriteLine($"{file}/{name} convert to {savePath}");
                }
            }
        }

        [CMD("dir_to_json")]
        private static void DirToJson(string dir)
        {
            ToJson(dir, string.Empty);
        }

        [CMD("excel_to_json")]
        private static void ExcelToJson(string filename)
        {
            filename = Path.GetFullPath(filename);
            string saveDir = Directory.GetDirectoryRoot(filename);
            var sheets = ExcelHelper.GetWorksheetNames(filename);
            foreach (var name in sheets)
            {
                string savePath = Path.Combine(saveDir, $"{Path.GetFileNameWithoutExtension(filename)}.{name}.json");
                File.WriteAllText(savePath, _convert.ToJson(filename, name));
                Console.WriteLine($"{filename}/{name} convert to {savePath}");
            }
        }

        [CMD("tobson", "convert filename to bson and save at savepath")]
        private static void ToBson(string filename, string sheetName, string savePath)
        {
            filename = Path.GetFullPath(filename);
            savePath = Path.GetFullPath(savePath);

            byte[] bson = BsonConverter.ExcelToBson(filename, sheetName);
            File.WriteAllBytes(savePath, bson);
            Console.WriteLine($"{filename}/{sheetName} convert to {savePath}");
        }

        [CMD]
        private static void ToBson(string dir, string saveDir)
        {
            dir = Path.GetFullPath(dir);
            if (string.IsNullOrEmpty(saveDir))
                saveDir = dir;
            else
                saveDir = Path.GetFullPath(saveDir);

            var files = Directory.GetFiles(dir, "*.xlsx");
            foreach (var file in files)
            {
                var sheets = ExcelHelper.GetWorksheetNames(file);
                foreach (var name in sheets)
                {
                    string savePath = Path.Combine(saveDir, $"{Path.GetFileNameWithoutExtension(file)}.{name}.bin");
                    File.WriteAllBytes(savePath, BsonConverter.ExcelToBson(file, name));
                    Console.WriteLine($"{file}/{name} convert to {savePath}");
                }
            }
        }

        [CMD("excel_to_bson")]
        private static void ExcelToBson(string filename)
        {
            filename = Path.GetFullPath(filename);
            string saveDir = Directory.GetDirectoryRoot(filename);
            var sheets = ExcelHelper.GetWorksheetNames(filename);
            foreach (var name in sheets)
            {
                string savePath = Path.Combine(saveDir, $"{Path.GetFileNameWithoutExtension(filename)}.{name}.bin");
                ToBson(filename, name, savePath);
            }
        }

        [CMD]
        private static void Exit()
        {
            Environment.Exit(0);
        }
    }
}