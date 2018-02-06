using ADODB;
using ComInvoker;
using CsvHelper;
using CsvHelper.Configuration;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Ude;

namespace OpenCsvByExcel
{
    internal class Opener
    {
        /// <summary>
        /// Excel Type
        /// </summary>
        private Type excelType;

        /// <summary>
        /// Extension delimiter mapping
        /// </summary>
        private HashSet<string>
            csvExtensions = new HashSet<string>(Program.Settings.CsvExtensions.Cast<string>()),
            tsvExtensions = new HashSet<string>(Program.Settings.TsvExtensions.Cast<string>());

        /// <summary>
        /// Constructor
        /// </summary>
        internal Opener()
        {
            const string ExcelApplication = "Excel.Application";

            excelType = Type.GetTypeFromProgID(ExcelApplication);
            if (excelType == null)
            {
                throw new TypeLoadException("Excel does not found");
            }
        }

        /// <summary>
        /// Open Csv
        /// </summary>
        /// <param name="path">Csv file path</param>
        internal void Open(string path)
        {
            void writeConsole(string key, string value) => Console.WriteLine($"{path}: [{key}] => {value}");

            using (var stream = File.OpenRead(path))
            {
                //Detect file charset
                var charsetDetector = new CharsetDetector();
                charsetDetector.Feed(stream);
                charsetDetector.DataEnd();
                var charset = charsetDetector.Charset ?? Program.Settings.FallbackCharset;
                charsetDetector.Reset();
                stream.Seek(0, SeekOrigin.Begin);
                writeConsole("Charset", charset);

                var csvConfiguration = new Configuration()
                {
                    DetectColumnCountChanges = Program.Settings.DetectColumnCountChanges,
                    HasHeaderRecord = Program.Settings.HasHeaderRecord,
                };
                //Change delimiter by extension
                var extension = Path.GetExtension(path).ToLower();
                if (csvExtensions.Contains(extension))
                {
                    csvConfiguration.Delimiter = ",";
                }
                else if (tsvExtensions.Contains(extension))
                {
                    csvConfiguration.Delimiter = "\t";
                }
                writeConsole("Delimiter", csvConfiguration.Delimiter);

                //Open csv
                using (var reader = new StreamReader(stream, Encoding.GetEncoding(charset)))
                using (var csv = new CsvReader(reader, csvConfiguration))
                using (var invoker = new Invoker())
                {
                    //Read firt row
                    if (!csv.Read())
                    {
                        throw new InvalidDataException("Cannot load first raw");
                    }

                    //Create recordset
                    string[] headers;
                    if (csvConfiguration.HasHeaderRecord)
                    {
                        //Read header and skip header
                        if (csv.ReadHeader() && csv.Read())
                        {
                            headers = csv.Context.HeaderRecord;
                        }
                        else
                        {
                            throw new InvalidDataException("Cannot load header");
                        }
                    }
                    else
                    {
                        headers = Enumerable.Range(0, csv.Context.Record.Length)
                            .Select(x => $"F{x}").ToArray();
                    }
                    var recordset = CreateRecordset(invoker, headers);
                    writeConsole("FieldCount", recordset.Fields.Count.ToString("#,0"));

                    do
                    {
                        //Add recordset
                        var fields = invoker.Invoke<Fields>(recordset.Fields);
                        var csvColumnCount = csv.Context.Record.Length > fields.Count
                            ? fields.Count
                            : csv.Context.Record.Length;
                        recordset.AddNew();
                        for (var i = 0; i < csvColumnCount; i++)
                        {
                            fields[i].Value = csv.Context.Record[i];
                        }
                        recordset.Update();
                        invoker.Release();//release fields
                    } while (csv.Read());
                    writeConsole("RecordCount", recordset.RecordCount.ToString("#,0"));

                    if (recordset?.RecordCount > 0)
                    {
                        //Launch excel
                        var excel = invoker.Invoke<dynamic>(Activator.CreateInstance(excelType));
                        excel.Visible = true;

                        var workbooks = invoker.Invoke<dynamic>(excel.Workbooks);
                        invoker.Invoke<dynamic>(workbooks.Add());//Add Book1
                        var range = invoker.Invoke<dynamic>(excel.ActiveCell);
                        var sheet = invoker.Invoke<dynamic>(excel.ActiveSheet);
                        var queryTables = invoker.Invoke<dynamic>(sheet.QueryTables);
                        var queryTable = invoker.Invoke<dynamic>(queryTables.Add(recordset, range));

                        //App settings
                        queryTable.FieldNames = Program.Settings.HasHeaderRecord;
                        queryTable.AdjustColumnWidth = Program.Settings.AdjustColumnWidth;

                        var result = queryTable.Refresh();
                        writeConsole("RefreshRecordset", result.ToString());
                    }
                    writeConsole("ComStackCount", invoker.StackCount.ToString("#,0"));
                }
            }
        }

        /// <summary>
        /// Create ADODB.Recordset
        /// </summary>
        /// <param name="invoker">ComInvoker</param>
        /// <param name="headers">Header records</param>
        /// <returns>ADODB.Recordset</returns>
        private Recordset CreateRecordset(Invoker invoker, string[] headers)
        {
            //Create recordset
            var recordset = invoker.Invoke<Recordset>(new Recordset());
            var headerFields = invoker.Invoke<Fields>(recordset.Fields);
            foreach (var header in headers)
            {
                headerFields.Append(
                    /*Column Name*/header,
                    /*Column Type*/DataTypeEnum.adVarChar,
                    /*Column Size*/Program.Settings.MaxColumnSize,
                    /*Column Attr*/FieldAttributeEnum.adFldUpdatable);
            }
            recordset.Open();
            invoker.Release();//release headerFields

            return recordset;
        }
    }
}
