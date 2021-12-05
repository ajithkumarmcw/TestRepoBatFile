using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using CommandLine;

namespace Helper
{
    /// <summary>
    /// Options which user can feed to the program
    /// </summary>
    public class Options
    {
        [Option('i', "InputFileName", Required = true, HelpText = "Path of Input filename or Relative path (In Current directory only)")]
        public string InputFileName { get; set; }

        [Option('p', "IterationPause", Required = false, HelpText = "Pause between iterations", Default = 2000)]
        public int IterationPause { get; set; }

        [Option('s', "SeparationPause", Required = false, HelpText = "Pause before and after iteration", Default = 2000)]
        public int SeparationPause { get; set; }

        [Option("SheetNumber", Required = false, HelpText = "Sheet number", Default = 1)]
        public int SheetNumber { get; set; }

        [Option("SortOrder", Required = false, HelpText = "Sort order : ASC or DES", Default = "ASC")]
        public string SortOrder { get; set; }

        [Option("Range", Required = false, HelpText = "Range in which Sort to be perfromed eg:A1:V400000 ", Default = "A1:V400000")]
        public string Range { get; set; }

        [Option("ColumnList", Required = false, HelpText = "column numbers in the format of col1,col2,col3,col4. Example: --ColumnList 1,2,3", Default = "1,5,18,3,15,7,4,12,9,16")]
        public string ColumnList { get; set; }

        [Option('o', "OutputFileName", Required = true, HelpText = "Path of Output filename or Relative path (In Current directory only)")]
        public string OutputFileName { get; set; }

        [Option('n', "Iterations", Required = false, HelpText = "Number of iterations", Default = -1)]
        public int Iterations { get; set; }

        [Option('r', "Repetition", Required = false, HelpText = "Number of repetition of program", Default = 1)]
        public int Repetition { get; set; }
    }

    /// <summary>
    /// Parses the argument given by user and also sets default case 
    /// </summary>
    class Arguments
    {
        // Command line option for Sort
        public static Options ParseArgument(string[] args)
        {
            Options options = new Options();

            if (args.Length != 0)
            {
                // Invoke Sort default
                if (args[0] == "default")
                {
                    options.InputFileName = "..\\input\\MOCK_Data_Only_for_sorting.xlsx";
                    options.OutputFileName = "SortResult.xlsx";
                    options.Iterations = 10;
                    options.IterationPause = 2000;
                    options.SeparationPause = 2000;
                    options.SheetNumber = 1;
                    options.SortOrder = "ASC";
                    options.Range = "A1:V400000";
                    options.ColumnList = "1,5,18,2,15,7,4,12,9,16";
                    options.Repetition = 2;                    

                    return options;
                }                
            }
            else
            {
                Console.WriteLine("\n\nExcel_Sort.exe default");
                Console.WriteLine("\n-------or---------\n");
            }
            
            Parser.Default.ParseArguments<Options>(args)
                   .WithParsed<Options>(opt =>
                   {
                       options = opt;

                   });
            return options;
        }
    }
}
