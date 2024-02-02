using System;
using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using System.IO;
using NDesk.Options;
using System.Text.RegularExpressions;
using Microsoft.VisualBasic.FileIO;

namespace csv2excel
{
    class Parameters
    {
        internal static bool show_help = false;

        //set default arguments
        internal static string inputFile = "";
        internal static string outputFile = "";
        internal static string columnDelimiter = ",";
        internal static string format = "xlsx";
        internal static bool textOnly = false;
        internal static bool ignoreQuotes = false;
        internal static bool resizeColumns = false;
    }
    class Program
    {
        static int verbosity;

        public static int Main(string[] args)
        {
            var p = new OptionSet() {
                { "i|in=", "the {inputfile} to convert. (REQUIRED)",
                  v => Parameters.inputFile = v },
                { "o|out=", "the path of the {outputfile}.",
                  v => Parameters.outputFile = v },
                { "c|coldel=", "the {delimiter} separating columns of inputfile.",
                  v => Parameters.columnDelimiter = v },
                { "f|format=", "the {format} for the output file [xls|xlsx].",
                  v => Parameters.format = v },
                { "t", "force all cells in output worksheet to be of type Text",
                  v => { if (v != null) Parameters.textOnly = true; } },
                { "q", "ignore double-quotes",
                  v => { if (v != null) Parameters.ignoreQuotes = true; } },
                { "r", "resize width of worksheet columns to fit data",
                  v => { if (v != null) Parameters.resizeColumns = true; } },
                { "v", "increase debug message verbosity",
                  v => { if (v != null) ++verbosity; } },
                { "h|help",  "show this message and exit", 
                  v => Parameters.show_help = v != null },
            };

            try
            {
                p.Parse(args);

                //this is the only required argument
                if (String.IsNullOrWhiteSpace(Parameters.inputFile))
                {
                    Parameters.show_help = true;
                }

                Parameters.format = Parameters.format.ToLower();

                if (Parameters.format != "xlsx" && Parameters.format != "xls")
                {
                    Console.WriteLine("Unrecognized format: {0}", Parameters.format);
                    Parameters.show_help = true;
                }
                

            }
            catch (OptionException e)
            {
                Console.Write("{0}: ", System.AppDomain.CurrentDomain.FriendlyName);
                Console.WriteLine(e.Message);
                Console.WriteLine("Try `{0} --help' for more information.", System.AppDomain.CurrentDomain.FriendlyName);
                return 1;
            }

            if (Parameters.show_help)
            {
                ShowHelp(p);
                return 0;
            }

            Debug("outputFile: \t\t{0}", Parameters.outputFile);
            Debug("inputFile: \t\t{0}", Parameters.inputFile);
            Debug("columnDelimiter: \t{0}", Parameters.columnDelimiter);
            Debug("format: \t\t{0}", Parameters.format);
            Debug("textOnly: \t\t{0}", Parameters.textOnly);
            Debug("ignoreQuotes: \t\t{0}", Parameters.ignoreQuotes);
            Debug("resizeColumns: \t\t{0}", Parameters.resizeColumns);

            Parameters.columnDelimiter = Regex.Unescape(Parameters.columnDelimiter);

            //if outputfile wasn't specified, set it to the same as the input file with the new extension
            if (String.IsNullOrWhiteSpace(Parameters.outputFile))
            {
                Parameters.outputFile = Parameters.inputFile.Replace(Path.GetExtension(Parameters.inputFile), "." + Parameters.format);
            }

            Debug("outputFile (calcd): \t{0}", Parameters.outputFile);
             
            //remove any previous version of the file
            File.Delete(Parameters.outputFile);

            writeToWorkbook();

            return 0;
        }

        static void ShowHelp(OptionSet p)
        {
            Console.WriteLine("Usage: {0} [OPTIONS]", System.AppDomain.CurrentDomain.FriendlyName);
            Console.WriteLine("Converts a given delimited file to an Excel format.");
            Console.WriteLine();
            Console.WriteLine("Options:");
            p.WriteOptionDescriptions(Console.Out);
            Console.WriteLine();
            Console.WriteLine("e.g.: {0} -i input.csv", System.AppDomain.CurrentDomain.FriendlyName);
            Console.WriteLine("e.g.: {0} -i input.csv -c \\t", System.AppDomain.CurrentDomain.FriendlyName);
        }

        static void Debug(string format, params object[] args)
        {
            if (verbosity > 0)
            {
                Console.Write("# ");
                Console.WriteLine(format, args);
            }
        }

        static void writeToWorkbook()
        {
            IWorkbook myWorkbook = null;

            if (Parameters.format == "xlsx")
            {
                myWorkbook = new XSSFWorkbook();
            }
            else if (Parameters.format == "xls")
            {
                myWorkbook = new HSSFWorkbook();
            }
             
            ISheet mySheet = myWorkbook.CreateSheet("Sheet1");

            int rowCount = 0;
            int colCount = 0;

            using (TextFieldParser parser = new TextFieldParser(Parameters.inputFile))
            {
                parser.SetDelimiters(Parameters.columnDelimiter);
                parser.HasFieldsEnclosedInQuotes = !Parameters.ignoreQuotes;

                while (!parser.EndOfData)
                {
                    IRow row = mySheet.CreateRow(rowCount);

                    colCount = 0;
                    string[] fields = parser.ReadFields();

                    foreach (string field in fields)
                    {
                        if (Parameters.textOnly)
                        {
                            row.CreateCell(colCount).SetCellValue(field);
                        }
                        else
                        {
                            double d;

                            if (Double.TryParse(field, out d))
                            {
                                row.CreateCell(colCount).SetCellValue(d);
                            }
                            else //default to string/text
                            {
                                row.CreateCell(colCount).SetCellValue(field);
                            }
                        }

                        colCount++;
                    }

                    rowCount++;
                }
            }

            if(Parameters.resizeColumns)
            {
                for (int i = 0; i <= mySheet.GetRow(0).PhysicalNumberOfCells; i++)
                {
                    mySheet.AutoSizeColumn(i);
                }
            }

            //Write the stream data of workbook
            using (FileStream file = new FileStream(Parameters.outputFile, FileMode.Create))
            {
                myWorkbook.Write(file);
                file.Close();
            }
        }
    }
}
