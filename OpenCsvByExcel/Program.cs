using CsvHelper;
using System;
using System.IO;
using System.Threading.Tasks;

namespace OpenCsvByExcel
{
    class Program
    {
        //App.config
        internal static Properties.Settings Settings { get; } = Properties.Settings.Default;

        static void Main(string[] args)
        {
            var opener = new Opener();
            Parallel.ForEach(
                args,
                new ParallelOptions() { MaxDegreeOfParallelism = Settings.ParallelOpen },
                path =>
            {
                Action<string> writeErrorConsole = (message) => Console.Error.WriteLine($"{path}: {message}");

                if (File.Exists(path))
                {
                    try
                    {
                        opener.Open(path);
                    }
                    catch (BadDataException)
                    {
                        writeErrorConsole("Incorrect field count");
                    }
                }
                else
                {
                    writeErrorConsole("File does not exists");
                }
            });
        }
    }
}
