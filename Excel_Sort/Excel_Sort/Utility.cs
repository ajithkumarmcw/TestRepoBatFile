using System;
using System.Threading;
using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.IO;

namespace Helper
{
    class Utility
    {
        /// <summary>
        /// Minimizes console with the handler
        /// </summary>
        /// <param name="hWnd"></param>
        static void MinimizeConsole(IntPtr hWnd)
        {
            if (hWnd != IntPtr.Zero)
            {
                Logger.ShowWindow(hWnd, Logger.swMinimize);
            }
        }

        /// <summary>
        /// Opens Excel app , minimises console and then opens workbook .Triggers ETW events for opening file
        /// </summary>
        /// <param name="app"></param>
        /// <param name="workbooks"></param>
        /// <param name="worksheets"></param>
        /// <param name="set"></param>
        /// <param name="opt"></param>
        public static void ExcelInit(out Excel.Application app,out List<Excel._Workbook> workbooks,out List<Excel._Worksheet> worksheets, Logger.Settings set, Options opt)
        {
            // Open the Excel Application
            app = new Excel.Application
            {
                Visible = true,
                DisplayAlerts = false
            };

            set.previousFullScreenSetting = Convert.ToInt32(app.DisplayFullScreen);

            app.WindowState = Excel.XlWindowState.xlMinimized;
            app.WindowState = Excel.XlWindowState.xlMaximized;

            // Turning to full screen mode
            app.DisplayFullScreen = true;
            

            MinimizeConsole(set.hWnd);
            Thread.Sleep(opt.SeparationPause);

            workbooks = new List<Excel._Workbook>();
            int numberOfWorkBooks = set.workbookFileNames.Count;

            // ETW logging for opening a file
            //Logging.Log.StartFileOpen(opt.InputFileName, set.repetition, opt.Iterations);
            Logger.StartFileOpen(opt.InputFileName, set.repetition);

            for (int workbookNumber=0; workbookNumber < numberOfWorkBooks; workbookNumber++)
            {
                // Open Input Excel workbook
                workbooks.Add(app.Workbooks.Open(set.workbookFileNames[workbookNumber], UpdateLinks: 0, ReadOnly: false));
                Thread.Sleep(opt.SeparationPause);
            }

            // ETW logging for opening a file
            //Logging.Log.StopFileOpen(opt.InputFileName, set.repetition, opt.Iterations);

            Logger.StopFileOpen(opt.InputFileName, set.repetition);

            // Activate workbook
            workbooks[0].Activate();            
            worksheets = new List<Excel._Worksheet>();         

        }

        /// <summary>
        /// Closes the file step by step. First it restores the original screensettings and then closes the workbook and app
        /// </summary>
        /// <param name="app"></param>
        /// <param name="workbooks"></param>
        /// <param name="worksheets"></param>
        /// <param name="set"></param>
        /// <param name="opt"></param>

        public static void ExcelDeInit(Excel.Application app, List<Excel._Workbook> workbooks, List<Excel._Worksheet> worksheets, Logger.Settings set, Helper.Options opt)
        {
            // Restoring original setting
            if (app != null)
            {
                app.DisplayFullScreen = Convert.ToBoolean(set.previousFullScreenSetting);
            }

            if (worksheets != null)
            {
                for (int sheet = 0; sheet < worksheets.Count; sheet++)
                {
                    if (worksheets[sheet] != null)
                    {
                        Marshal.ReleaseComObject(worksheets[sheet]);
                    }
                }
            }

            if (workbooks != null)
            {
                for (int workbook = 0; workbook < workbooks.Count; workbook++)
                {

                    // Close workbook
                    if (workbooks[workbook] != null)
                    {
                        workbooks[workbook].Close(false, Type.Missing, Type.Missing);
                        Thread.Sleep(opt.SeparationPause);
                        Marshal.ReleaseComObject(workbooks[workbook]);
                    }
                }
            }

            // Close Excel application
            if (app != null)
            {
                app.Quit();
                Marshal.ReleaseComObject(app);
            }           

        }

        /// <summary>
        /// Gets the file name given by user and concatenates repetation number and iteration number with it
        /// Example: BenchmakrResult099910.xlsx, where 0999 -> zero padded repetition number , 10 -> iteration number
        /// </summary>
        /// <param name="outputFileName"></param>
        /// <param name="rep"></param>
        /// <param name="iter"></param>
        /// <returns></returns>
        public static string GetFileName(string outputFileName , int rep, int iter)
        {
            string filename = Path.GetFileNameWithoutExtension(outputFileName);
            string extension = Path.GetExtension(outputFileName);
            string reps = String.Format("{0:0000}", rep);
            string iteration = String.Format("{0:00}", iter);
            string rep_outputFileName = Path.GetDirectoryName(outputFileName) + "\\"+ filename + reps + iteration + extension; 

            return rep_outputFileName;
        }



    }

}
