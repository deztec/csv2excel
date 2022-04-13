using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using System.IO;
using NDesk.Options;
using System.Text.RegularExpressions;
using Microsoft.VisualBasic.FileIO;

namespace csv2excel
{
    class Program
    {
        static int verbosity;

        public static int Main(string[] args)
        {
            bool show_help = false;
            
            //set default arguments
            string inputFile = "";
            string outputFile = "";
            string columnDelimiter = ",";
            string format = "xlsx";
            bool textOnly = false;
            bool ignoreQuotes = false;
            
            var p = new OptionSet() {
                { "i|in=", "the {inputfile} to convert. (REQUIRED)",
                  v => inputFile = v },
                { "o|out=", "the path of the {outputfile}.",
                  v => outputFile = v },
                { "c|coldel=", "the {delimiter} separating columns of inputfile.",
                  v => columnDelimiter = v },
                { "f|format=", "the {format} for the output file [xls|xlsx].",
                  v => format = v },
                { "t", "force all cells in output worksheet to be of type Text",
                  v => { if (v != null) textOnly = true; } },
                { "q", "ignore double-quotes",
                  v => { if (v != null) ignoreQuotes = true; } },
                { "v", "increase debug message verbosity",
                  v => { if (v != null) ++verbosity; } },
                { "h|help",  "show this message and exit", 
                  v => show_help = v != null },
            };

            try
            {
                p.Parse(args);

                //this is the only required argument
                if (String.IsNullOrWhiteSpace(inputFile))
                {
                    show_help = true;
                }

                format = format.ToLower();

                if (format != "xlsx" && format != "xls")
                {
                    Console.WriteLine("Unrecognized format: {0}", format);
                    show_help = true;
                }

            }
            catch (OptionException e)
            {
                Console.Write("{0}: ", System.AppDomain.CurrentDomain.FriendlyName);
                Console.WriteLine(e.Message);
                Console.WriteLine("Try `{0} --help' for more information.", System.AppDomain.CurrentDomain.FriendlyName);
                return 1;
            }

            if (show_help)
            {
                ShowHelp(p);
                return 0;
            }

            Debug("outputFile: \t\t{0}", outputFile);
            Debug("inputFile: \t\t{0}", inputFile);
            Debug("columnDelimiter: \t{0}", columnDelimiter);
            Debug("format: \t\t{0}", format);
            Debug("textOnly: \t\t{0}", textOnly);
            Debug("ignoreQuotes: \t\t{0}", ignoreQuotes);

            columnDelimiter = Regex.Unescape(columnDelimiter);

            //if outputfile wasn't specified, set it to the same as the input file with the new extension
            if (String.IsNullOrWhiteSpace(outputFile))
            {
                outputFile = inputFile.Replace(Path.GetExtension(inputFile), "." + format);
            }

            Debug("outputFile (calcd): \t{0}", outputFile);
             
            //remove any previous version of the file
            File.Delete(outputFile);

            writeToWorkbook(outputFile, inputFile, columnDelimiter, textOnly, ignoreQuotes, format);

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

        static void writeToWorkbook(string outputFile, string inputFile, string columnDelimiter, bool textOnly, bool ignoreQuotes, string format)
        {
            IWorkbook myWorkbook = null;

            if (format == "xlsx")
            {
                myWorkbook = new XSSFWorkbook();
            }
            else if (format == "xls")
            {
                myWorkbook = new HSSFWorkbook();
            }
             
            ISheet mySheet = myWorkbook.CreateSheet("Sheet1");

            int rowCount = 0;
            int colCount = 0;

            using (TextFieldParser parser = new TextFieldParser(inputFile))
            {
                parser.SetDelimiters(columnDelimiter);
                parser.HasFieldsEnclosedInQuotes = !ignoreQuotes;

                while (!parser.EndOfData)
                {
                    IRow row = mySheet.CreateRow(rowCount);

                    colCount = 0;
                    string[] fields = parser.ReadFields();

                    foreach (string field in fields)
                    {
                        if (textOnly)
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

            //Write the stream data of workbook
            using (FileStream file = new FileStream(outputFile, FileMode.Create))
            {
                myWorkbook.Write(file);
                file.Close();
            }
        }
    }
}
