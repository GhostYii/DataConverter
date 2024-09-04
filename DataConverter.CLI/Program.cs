using DataConverter.Core;
using System.Reflection;
using System.Text;

namespace DataConverter.CLI
{
    using Console = System.Console;
    public class CLI
    {
        private static string _Version => Assembly.GetExecutingAssembly().GetName().Version.ToString();
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
                    string json = _convert.ToJson(file, name);
                    if (string.IsNullOrEmpty(json))
                        continue;

                    File.WriteAllText(savePath, json);
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
                string json = _convert.ToJson(filename, name);
                if (string.IsNullOrEmpty(json))
                    continue;

                File.WriteAllText(savePath, json);
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

        [CMD("to_cs")]
        private static void ToCS(string filename, string nameSpace, string saveDir)
        {
            if (Path.GetExtension(filename) != ".xlsx")
            {
                Console.WriteLine("only .xlsx file support.");
                return;
            }

            if (!File.Exists(filename))
            {
                Console.WriteLine($"{filename} dont exist.");
                return;
            }

            var sheets = ExcelHelper.GetWorksheetNames(filename);
            for (int i = 0; i < sheets.Length; i++)
            {
                string cs = _convert.ToCSharp(filename, i, sheets[i], nameSpace);
                if (string.IsNullOrEmpty(cs))
                    continue;
                string savePath = Path.Combine(saveDir, $"{Path.GetFileNameWithoutExtension(filename)}.{sheets[i]}.cs");
                File.WriteAllText(savePath, cs);
                Console.WriteLine($"{filename}/{sheets[i]} convert to {savePath}");
            }
        }

        [CMD("dir_to_cs")]
        private static void DirToCs(string dir, string nameSpace, string saveDir)
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
                for (int i = 0; i < sheets.Length; i++)
                {
                    string cs = _convert.ToCSharp(file, i, sheets[i], nameSpace);
                    if (string.IsNullOrEmpty(cs))
                        continue;
                    string savePath = Path.Combine(saveDir, $"{Path.GetFileNameWithoutExtension(file)}.{sheets[i]}.cs");
                    File.WriteAllText(savePath, cs);
                    Console.WriteLine($"{file}/{sheets[i]} convert to {savePath}");
                }
            }
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
