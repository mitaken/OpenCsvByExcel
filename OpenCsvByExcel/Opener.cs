﻿using ADODB;
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
            Action<string, string> writeConsole = (key, value) => Console.WriteLine($"{path}: [{key}] => {value}");

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
                    DetectColumnCountChanges = true,
                    HasHeaderRecord = false,
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
                    Recordset recordset = null;
                    while (csv.Read())
                    {
                        if (recordset == null)
                        {
                            recordset = CreateRecordset(invoker, csv.Context.Record.Length);
                            writeConsole("FieldCount", recordset.Fields.Count.ToString("#,0"));
                        }

                        //Add recordset
                        var fields = invoker.Invoke<Fields>(recordset.Fields);
                        recordset.AddNew();
                        for (var i = 0; i < csv.Context.Record.Length; i++)
                        {
                            fields[i].Value = csv.Context.Record[i];
                        }
                        recordset.Update();
                        invoker.Release();//release fields
                    }
                    writeConsole("RecordCount", recordset.RecordCount.ToString("#,0"));

                    if (recordset?.RecordCount > 0)
                    {
                        //Launch excel
                        var excel = invoker.Invoke<dynamic>(Activator.CreateInstance(excelType));
                        excel.Visible = true;

                        var workbooks = invoker.Invoke<dynamic>(excel.Workbooks);
                        var workbook = invoker.Invoke<dynamic>(workbooks.Add());
                        var range = invoker.Invoke<dynamic>(excel.ActiveCell);
                        var loadCount = range.CopyFromRecordset(recordset);
                        writeConsole("ExcelLoadCount", loadCount.ToString("#,0"));
                    }
                    writeConsole("ComStackCount", invoker.StackCount.ToString("#,0"));
                }
            }
        }

        /// <summary>
        /// Create ADODB.Recordset
        /// </summary>
        /// <param name="invoker">ComInvoker</param>
        /// <param name="columnCount">Column count</param>
        /// <returns>ADODB.Recordset</returns>
        private Recordset CreateRecordset(Invoker invoker, int columnCount)
        {
            //Create recordset
            var headers = Enumerable.Range(0, columnCount)
                .Select(x => $"F{x}").ToArray();
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
