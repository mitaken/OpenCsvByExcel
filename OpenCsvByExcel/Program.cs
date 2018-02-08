using CsvHelper;
using System;
using System.Diagnostics;
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
                void showError(string message) => Trace.Fail($"{path}\n\n{message}");

                if (File.Exists(path))
                {
                    try
                    {
                        opener.Open(path);
                    }
                    catch (BadDataException e)
                    {
                        showError($"Error: {e.Message}\n{e.ReadingContext.RawRecord}");
                    }
                    catch (Exception e)
                    {
                        showError(e.ToString());
                    }
                }
                else
                {
                    showError("File does not exists");
                }
            });
        }
    }
}
