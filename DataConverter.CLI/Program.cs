using DataConverter.Core;
using System.Reflection;
using System.Text;

namespace DataConverter.CLI
{
    using Console = System.Console;
    public class CLI
    {
        private static Version? _Version => Assembly.GetExecutingAssembly().GetName().Version;
        private static ExcelConverter _convert = new ExcelConverter();

        public static void Main(params string[] args)
        {
            Commands.RegisterAllCommandByType(typeof(CLI));
            Console.Title = $"DCT.CLI v{_Version}";

            Core.Console.AddPrintListener(msg => ConsoleOutput.WriteLine(msg));
            Core.Console.AddErrorListener(msg => ConsoleOutput.WriteLine(msg, ConsoleColor.Red));
            Core.Console.AddWarningListener(msg => ConsoleOutput.WriteLine(msg, ConsoleColor.Yellow));

            Console.WriteLine($"DataConverter Tool, Version CLI:{_Version}, Core: {typeof(ExcelHelper).Assembly.GetName().Version}");

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

        private static int GetMaxDegreeOfParallelism()
        {
            return Math.Max(1, Math.Min(Environment.ProcessorCount / 2, 4));
        }

        [CMD("tojson", "convert filename to json and save at savepath")]
        private static void ToJson(string filename, string sheetName, string savePath)
        {
            filename = Path.GetFullPath(filename);
            if (!_convert.CheckToJson(filename, sheetName))
                return;

            string jsonStr = _convert.ToJson(filename, sheetName);
            File.WriteAllText(savePath, jsonStr, Encoding.UTF8);
            Console.WriteLine($"{filename}/{sheetName} convert to {savePath}");
        }

        [CMD]
        private static void ToJson(string filename, int sheetIndex, string savePath)
        {
            filename = Path.GetFullPath(filename);
            if (!_convert.CheckToJson(filename, sheetIndex))
                return;

            string jsonStr = _convert.ToJson(filename, sheetIndex);
            File.WriteAllText(savePath, jsonStr, Encoding.UTF8);
            Console.WriteLine($"{filename}/{sheetIndex}(index) convert to {savePath}");
        }

        [CMD]
        private static async Task ToJson(string dir, string saveDir)
        {
            dir = Path.GetFullPath(dir);
            saveDir = string.IsNullOrEmpty(saveDir) ? dir : Path.GetFullPath(saveDir);

            string[] files = Directory.GetFiles(dir, "*.xlsx");
            using ProgressBar progress = new ProgressBar(files.Length, "JSON");
            await Parallel.ForEachAsync(files, new ParallelOptions { MaxDegreeOfParallelism = GetMaxDegreeOfParallelism() }, async (file, _) =>
            {
                try
                {
                    ExcelConverter converter = new ExcelConverter();
                    var sheets = ExcelHelper.GetWorksheetNames(file);
                    if (sheets == null)
                        return;

                    foreach (var name in sheets)
                    {
                        if (!converter.CheckToJson(file, name))
                            continue;

                        string savePath = Path.Combine(saveDir, $"{Path.GetFileNameWithoutExtension(file)}.{name}.json");
                        string json = converter.ToJson(file, name);
                        if (string.IsNullOrEmpty(json))
                            continue;

                        File.WriteAllText(savePath, json);
                    }
                }
                catch (Exception e)
                {
                    ConsoleOutput.WriteLine($"数据表'{Path.GetFileName(file)}'转JSON失败：{e.Message}");
                }
                finally
                {
                    progress.Increment();
                }
            });

            ConsoleOutput.WriteLine($"JSON处理完成：{files.Length} 个文件");
        }

        [CMD("dir_to_json")]
        private static async Task DirToJson(string dir)
        {
            await ToJson(dir, string.Empty);
        }

        [CMD("excel_to_json")]
        private static void ExcelToJson(string filename)
        {
            filename = Path.GetFullPath(filename);
            string saveDir = Path.GetDirectoryName(filename) ?? Directory.GetCurrentDirectory();
            var sheets = ExcelHelper.GetWorksheetNames(filename);
            foreach (var name in sheets)
            {
                if (!_convert.CheckToJson(filename, name))
                    continue;

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

            if (!_convert.CheckToJson(filename, sheetName))
                return;

            byte[] bson = BsonConverter.ExcelToBson(filename, sheetName);
            if (bson.Length == 0)
                return;

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
        private static async Task DirToCs(string dir, string nameSpace, string saveDir)
        {
            dir = Path.GetFullPath(dir);
            saveDir = string.IsNullOrEmpty(saveDir) ? dir : Path.GetFullPath(saveDir);

            string[] files = Directory.GetFiles(dir, "*.xlsx");
            using ProgressBar progress = new ProgressBar(files.Length, "CS");
            await Parallel.ForEachAsync(files, new ParallelOptions { MaxDegreeOfParallelism = GetMaxDegreeOfParallelism() }, async (file, _) =>
            {
                try
                {
                    ExcelConverter converter = new ExcelConverter();
                    var sheets = ExcelHelper.GetWorksheetNames(file);
                    if (sheets == null)
                        return;

                    for (int i = 0; i < sheets.Length; i++)
                    {
                        string cs = converter.ToCSharp(file, i, sheets[i], nameSpace);
                        if (string.IsNullOrEmpty(cs))
                            continue;

                        string savePath = Path.Combine(saveDir, $"{Path.GetFileNameWithoutExtension(file)}.{sheets[i]}.cs");
                        File.WriteAllText(savePath, cs);
                    }
                }
                catch (Exception e)
                {
                    ConsoleOutput.WriteLine($"数据表'{Path.GetFileName(file)}'转CS失败：{e.Message}");
                }
                finally
                {
                    progress.Increment();
                }
            });

            ConsoleOutput.WriteLine($"CS处理完成：{files.Length} 个文件");
        }

        [CMD]
        private static async Task ToBson(string dir, string saveDir)
        {
            dir = Path.GetFullPath(dir);
            saveDir = string.IsNullOrEmpty(saveDir) ? dir : Path.GetFullPath(saveDir);

            string[] files = Directory.GetFiles(dir, "*.xlsx");
            using ProgressBar progress = new ProgressBar(files.Length, "BSON");
            await Parallel.ForEachAsync(files, new ParallelOptions { MaxDegreeOfParallelism = GetMaxDegreeOfParallelism() }, async (file, _) =>
            {
                try
                {
                    ExcelConverter converter = new ExcelConverter();
                    var sheets = ExcelHelper.GetWorksheetNames(file);
                    if (sheets == null)
                        return;

                    foreach (var name in sheets)
                    {
                        if (!converter.CheckToJson(file, name))
                            continue;

                        string savePath = Path.Combine(saveDir, $"{Path.GetFileNameWithoutExtension(file)}.{name}.bin");
                        byte[] bson = BsonConverter.ExcelToBson(file, name);
                        if (bson.Length == 0)
                            continue;

                        File.WriteAllBytes(savePath, bson);
                    }
                }
                catch (Exception e)
                {
                    ConsoleOutput.WriteLine($"数据表'{Path.GetFileName(file)}'转BSON失败：{e.Message}");
                }
                finally
                {
                    progress.Increment();
                }
            });

            ConsoleOutput.WriteLine($"BSON处理完成：{files.Length} 个文件");
        }

        [CMD("excel_to_bson")]
        private static void ExcelToBson(string filename)
        {
            filename = Path.GetFullPath(filename);
            string saveDir = Path.GetDirectoryName(filename) ?? Directory.GetCurrentDirectory();
            var sheets = ExcelHelper.GetWorksheetNames(filename);
            foreach (var name in sheets)
            {
                if (!_convert.CheckToJson(filename, name))
                    continue;

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
