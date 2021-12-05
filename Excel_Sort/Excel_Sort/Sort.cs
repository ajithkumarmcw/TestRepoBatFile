using System;
using System.Collections.Generic;
using System.Threading;
using System.Diagnostics;
using System.IO;
using Helper;
using Excel = Microsoft.Office.Interop.Excel;

namespace Excel_Sort
{
    class Sort
    {
        // Sort Validations        
        static void ValidateColumnList(out string[] columnList, Logger.LogData logData, Options opt)
        {
            columnList = opt.ColumnList.Split(',');

            // Check the total number of iterations and columns number 
            if (opt.Iterations == -1)
            {
                opt.Iterations = columnList.Length;
            }           
            else if (opt.Iterations > columnList.Length)
            {
                string[] modColumnList = new string[opt.Iterations];
                for (int i = 0; i < opt.Iterations; i++)
                {                    
                    modColumnList[i] = columnList[(i%columnList.Length)];
                }
                columnList = modColumnList;
                opt.ColumnList = string.Join(",", columnList);
            }
            else if (opt.Iterations < columnList.Length)
            {
                string[] modColumnList = new string[opt.Iterations];
                for (int i = 0; i < opt.Iterations; i++)
                {                    
                    modColumnList[i] = columnList[i];
                }
                columnList = modColumnList;
                opt.ColumnList = string.Join(",", columnList);
            }

        }

        static void ValidateSheetNumber(List<Excel._Workbook> workbooks, List<Excel._Worksheet> worksheets, Logger.LogData logData, int SheetNumber)
        {
            int numSheets = workbooks[0].Sheets.Count;
            if (SheetNumber == 0 || SheetNumber > numSheets)
            {
                Console.WriteLine("Please enter valid sheet number");
                logData.Logging.Add(Logger.LogLogging(LogLevel: "Error", TimeStamp: DateTime.Now.ToString(), Detail: "Please enter valid sheet number"));
                throw new IndexOutOfRangeException("Please enter valid sheet number");
            }

        }

        // Function Description: Sort the Columns in Excel Worksheet 
        static void ExcelSort(Excel.Application app, List<Excel._Workbook> workbooks, List<Excel._Worksheet> worksheets, Logger.Settings set,
                                Options opt, Logger.LogData logData, int repeat)
        {

            for (int rep = 1; rep <= repeat; rep++)
            {
                string operationName = logData.Benchmark.name + "_" + rep.ToString();
                set.repetition = rep;   
                string[] columnList = null;
                ValidateColumnList(out columnList, logData, opt);

                set.workbookFileNames.Add(opt.InputFileName);

                try
                {
                    Utility.ExcelInit(out app, out workbooks, out worksheets, set, opt);
                    ValidateSheetNumber(workbooks, worksheets, logData, opt.SheetNumber);

                    // Adding worksheet to worksheet list and activating
                    worksheets.Add(workbooks[0].Sheets[opt.SheetNumber]);
                    worksheets[0].Activate();

                    Excel.Range range = worksheets[0].Range[opt.Range];
                    range.Select();
                    set.iterationTimings.Add(operationName, new List<double>());

                    Thread.Sleep(opt.SeparationPause);

                    for (int iter = 0; iter < opt.Iterations; iter++)
                    {
                        // Get Column number.
                        int colsId = Int32.Parse(columnList[iter]);
                        Console.WriteLine($"Sorting for Column ID {colsId}");
                        Excel.XlSortOrder sortOrder = 0;
                        Stopwatch watch = new Stopwatch();

                        if (opt.SortOrder == "ASC")
                        {
                            sortOrder = Excel.XlSortOrder.xlAscending;
                        }
                        else if (opt.SortOrder == "DES")
                        {
                            sortOrder = Excel.XlSortOrder.xlDescending;
                        }

                        Logger.IterationEventStart(rep, iter);

                        // Sort the Selected range.  
                        watch.Start();
                        range.Sort(range.Columns[colsId], sortOrder, Header: Excel.XlYesNoGuess.xlYes);
                        watch.Stop();

                        Logger.IterationEventEnd(rep, iter);

                        set.iterationTimings[operationName].Add(watch.ElapsedMilliseconds / 1000.0);
                        Logger.IterationEnd(operationName, iter, logData, set, opt.IterationPause);
                    }

                    Thread.Sleep(opt.SeparationPause);

                    // Save the workbook.
                    
                    workbooks[0].SaveAs(Utility.GetFileName(opt.OutputFileName, rep, opt.Iterations));
                    Thread.Sleep(opt.SeparationPause);

                    Utility.ExcelDeInit(app, workbooks, worksheets, set, opt);

                    set.status = "success";
                    Thread.Sleep(opt.SeparationPause);
                    

                }
                catch (Exception e)
                {
                    Console.WriteLine("Exception raised. Program failed");
                    Console.WriteLine(e);
                    Console.WriteLine(e.Message);

                    Utility.ExcelDeInit(app, workbooks, worksheets, set, opt);
                    set.exception = e;
                    set.status = "failure";
                    Thread.Sleep(opt.SeparationPause);
                    return;

                }
                app = null;
                workbooks = null;
                worksheets = null;
                set.workbookFileNames = new List<string>();
            }
            
            Logger.DeInit(logData, set, set.status, opt, set.exception);

        }

        /// The main entry point for the application.
        [STAThread]
        static void Main(string[] args)
        {
            Options opt;
            opt = Arguments.ParseArgument(args);

            if (args.Length == 0 || opt.InputFileName == null)
            {
                return;
            }

            string benchMarkTestName = "Excel_Sort";
            Logger.LogData logData = Logger.Init(opt, benchMarkTestName, "Excel");
            Console.WriteLine("Sorting in " + opt.SortOrder);

            // Decalre the Excel objects
            Excel.Application app = null;
            List<Excel._Workbook> workbooks = null;
            List<Excel._Worksheet> worksheets = null;

            Logger.Settings set = new Logger.Settings();

            
            ExcelSort(app, workbooks, worksheets, set,  opt, logData, (opt.Repetition));
            
            

        }

    }
}