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
            string templateFile = "";
            string columnDelimiter = ",";
            string lineDelimiter = "\\r\\n";
            string format = "xlsx";
            bool textOnly = false;
            
            var p = new OptionSet() {
                { "i|in=", "the {inputfile} to convert.",
                  v => inputFile = v },
                { "o|out=", "the path of the {outputfile}.",
                  v => outputFile = v },
                { "tem|template=", "the path of the {templateFile}.",
                  v => templateFile = v },
                { "c|coldel=", "the {delimiter} separating columns of inputfile.",
                  v => columnDelimiter = v },
                { "l|linedel=", "the {delimiter} separating lines of inputfile.",
                  v => lineDelimiter = v },
                { "f|format=", "the {format} for the output file [xls|xlsx].",
                  v => format = v },
                { "t", "force all cells in output worksheet to be of type Text",
                  v => { if (v != null) textOnly = true; } },
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
            Debug("lineDelimiter: \t{0}", lineDelimiter);
            Debug("format: \t\t{0}", format);
            Debug("textOnly: \t\t{0}", textOnly);

            columnDelimiter = Regex.Unescape(columnDelimiter);
            lineDelimiter = Regex.Unescape(lineDelimiter);

            //if outputfile wasn't specified, set it to the same as the input file with the new extension
            if (String.IsNullOrWhiteSpace(outputFile))
            {
                outputFile = inputFile.Replace(Path.GetExtension(inputFile), "." + format);
            }

            Debug("outputFile (calcd): \t{0}", outputFile);

            string inputData = File.ReadAllText(inputFile);
 
            //remove any previous version of the file
            File.Delete(outputFile);

            writeToWorkbook(outputFile, inputData, columnDelimiter, lineDelimiter, textOnly, format,templateFile);

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
            Console.WriteLine("e.g.: {0} -i input.csv -c \\t -l \\r\\n", System.AppDomain.CurrentDomain.FriendlyName);
        }

        static void Debug(string format, params object[] args)
        {
            if (verbosity > 0)
            {
                Console.Write("# ");
                Console.WriteLine(format, args);
            }
        }

        static void writeToWorkbook(string outputFile, string outputData, string columnDelimiter, string lineDelimiter, bool textOnly, string format, string templateFile)
        {
            IWorkbook myWorkbook = null;
            ISheet mySheet = null;



            IRow copyRow = null;
            TextFieldParser parser = new TextFieldParser(new StringReader(outputData));
            parser.HasFieldsEnclosedInQuotes = true;
            parser.SetDelimiters(columnDelimiter);
            string[] fields;
            int rowCount = 0;
            int colCount = 0;
            if (File.Exists(templateFile))
            {
                myWorkbook = new XSSFWorkbook(templateFile);
                mySheet = myWorkbook.GetSheetAt(0);
                textOnly = true;
                rowCount = 1;
                fields = parser.ReadFields();
                copyRow = mySheet.GetRow(0);
            }
            else
            {
                if (format == "xlsx")
                {
                    myWorkbook = new XSSFWorkbook();
                }
                else if (format == "xls")
                {
                    myWorkbook = new HSSFWorkbook();
                }

                mySheet = myWorkbook.CreateSheet("Sheet1");
            }

            while (!parser.EndOfData)
            {
                IRow row = null;
                if (copyRow == null)
                    row = mySheet.CreateRow(rowCount);
                else
                    row = copyRow.CopyRowTo(rowCount);
                colCount = 0;
                fields = parser.ReadFields();

                foreach (string field in fields)
                {
                    if (copyRow != null)
                        row.GetCell(colCount).SetCellValue(field);
                    else if (textOnly)
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

            
            //Write the stream data of workbook to the root directory
            using (FileStream file = new FileStream(outputFile, FileMode.Create))
            {
                myWorkbook.Write(file);
                file.Close();
            }
        }
    }
}
