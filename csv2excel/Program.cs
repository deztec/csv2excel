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

        public static void Main(string[] args)
        {
            bool show_help = false;

            //set default arguments
            string inputFile = "";
            string outputFile = "";
            string columnDelimiter = ",";
            string lineDelimiter = "\\r\\n";
            string format = "xlsx";

            var p = new OptionSet() {
                { "i|in=", "the {inputfile} to convert.",
                  v => inputFile = v },
                { "o|out=", "the path of the {outputfile}.",
                  v => outputFile = v },
                { "c|coldel=", "the {delimiter} separating columns of inputfile.",
                  v => columnDelimiter = v },
                { "l|linedel=", "the {delimiter} separating lines of inputfile.",
                  v => lineDelimiter = v },
                { "f|format=", "the {format} for the output file [xls|xlsx].",
                  v => format = v },
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
            }
            catch (OptionException e)
            {
                Console.Write("{0}: ", System.AppDomain.CurrentDomain.FriendlyName);
                Console.WriteLine(e.Message);
                Console.WriteLine("Try `{0} --help' for more information.", System.AppDomain.CurrentDomain.FriendlyName);
                return;
            }

            if (show_help)
            {
                ShowHelp(p);
                return;
            }

            Debug("outputFile: \t\t{0}", outputFile);
            Debug("inputFile: \t\t{0}", inputFile);
            Debug("columnDelimiter: \t{0}", columnDelimiter);
            Debug("lineDelimiter: \t{0}", lineDelimiter);
            Debug("format: \t\t{0}", format);

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

            //there are currently two separate methods because creating an IWorkbook interface and then assigning the object was 
            //throwing an exception on the XSSFWorkbook object
            //
            //this will be beneficial because having two separate methods causes the .net JIT compiler to only
            //load the necessary assemblies for that output type
            if (format.ToLower() == "xls")
            {
                writeToXLS(outputFile, inputData, columnDelimiter, lineDelimiter);
            }
            else if (format.ToLower() == "xlsx")
            {
                writeToXLSX(outputFile, inputData, columnDelimiter, lineDelimiter);
            }
            else
            {
                Console.WriteLine("Unrecognized format: {0}", format);
            }
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

        static void writeToXLS(string outputFile, string outputData, string columnDelimiter, string lineDelimiter)
        {
            HSSFWorkbook myWorkbook = new HSSFWorkbook();
            ISheet mySheet = myWorkbook.CreateSheet("Sheet1");

            int rowCount = 0;
            int colCount = 0;

            foreach (string currLine in outputData.Split(new String[] { lineDelimiter }, StringSplitOptions.None))
            {
                IRow row = mySheet.CreateRow(rowCount);

                colCount = 0;

                using (TextFieldParser parser = new TextFieldParser(new StringReader(currLine)))
                {
                    parser.HasFieldsEnclosedInQuotes = true;
                    parser.SetDelimiters(columnDelimiter);

                    string[] fields;

                    while (!parser.EndOfData)
                    {
                        fields = parser.ReadFields();

                        foreach (string field in fields)
                        {
                            row.CreateCell(colCount).SetCellValue(field);
                            colCount++;
                        }
                    } 
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

        static void writeToXLSX(string outputFile, string outputData, string columnDelimiter, string lineDelimiter)
        {
            XSSFWorkbook myWorkbook = new XSSFWorkbook();
            ISheet mySheet = myWorkbook.CreateSheet("Sheet1");

            int rowCount = 0;
            int colCount = 0;

            foreach (string currLine in outputData.Split(new String[] { lineDelimiter }, StringSplitOptions.None))
            {
                IRow row = mySheet.CreateRow(rowCount);

                colCount = 0;

                using (TextFieldParser parser = new TextFieldParser(new StringReader(currLine)))
                {
                    parser.HasFieldsEnclosedInQuotes = true;
                    parser.SetDelimiters(columnDelimiter);

                    string[] fields;

                    while (!parser.EndOfData)
                    {
                        fields = parser.ReadFields();

                        foreach (string field in fields)
                        {
                            row.CreateCell(colCount).SetCellValue(field);
                            colCount++;
                        }
                    }
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
