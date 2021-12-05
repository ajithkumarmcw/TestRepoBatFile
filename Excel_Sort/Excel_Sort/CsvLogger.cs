using System;
using System.IO;
using System.Text;
using System.Linq;
using System.Collections.Generic;

namespace Helper
{
    class CsvLogger
    {
        /// <summary>
        /// Creates a CSV file from the input provided
        /// </summary>
        /// <param name="timing"></param>
        /// <param name="operationName"></param>
        /// <param name="startTime"></param>
        /// <param name="opt"></param>
        public static void GenerateExcel(Dictionary<string, List<Double>> timing, string operationName, string startTime, Options opt)
        {
            var csv = new StringBuilder();
            var writeLine = "";
            string sysPath = Directory.GetCurrentDirectory();
            string excelFileName = sysPath + "\\" + operationName + "_" + startTime + ".csv";            
            string[] columns = opt.ColumnList.Split(',');
            
            writeLine = string.Format("{0},{1},{2},{3},{4},{5},{6}", "Timing", "Repetition", "Iteration", "OperationName", "InputFileName", "SheetNumber", "ColumnID" );
            csv.AppendLine(writeLine);
            foreach (var operation in timing)
            {
                for (int i = 0; i < operation.Value.Count; i++)
                {
                    var iteration = (i + 1).ToString();
                    
                    string[] repetition = operation.Key.Split('_');
                    writeLine = string.Format("{0},{1},{2},{3},{4},{5},{6}", operation.Value[i].ToString("0.000"), repetition[repetition.Length - 1], iteration, operationName, Path.GetFileName(opt.InputFileName), opt.SheetNumber, columns[i]  );
                    csv.AppendLine(writeLine);
                }
            }
            
            File.WriteAllText(excelFileName, csv.ToString());
        }
    }
}