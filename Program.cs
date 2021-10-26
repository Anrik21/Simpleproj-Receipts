using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;



namespace Receipt_reader
{
    class Program
    {
        private static List<string> logStrings;

        /// <summary>
        /// Function to remove some data for privacy reasons
        /// </summary>
        /// <remarks>Expansion possibilities: Save email data somehow & only check new emails when they arrive instead of doing all of them. 
        /// This function lets us then store emails that are more compressed as well as have less sensitive data</remarks>
        /// <param name="fileName">The email file that needs to be cleaned</param>
        /// <returns>A one dimensional string array that holds the relevant receipt data, where the email header has been removed as well as credit card/membership info</returns>
        private static string[] CleanEmails(string fileName)
        {
            string[] fileInput = System.IO.File.ReadAllLines(@fileName);
            string[] output = Array.Empty<string>();

            for (int i = 0; i < fileInput.Length; i++)
            {
                if (fileInput[i].Contains("------------------------------------------"))
                {
                    for (int j = i; i < fileInput.Length; j++)
                    {
                        if (fileInput[j].Contains("</pre>"))
                        {
                            output = new string[j - i];

                            Array.Copy(fileInput, i, output, 0, j - i);

                            break;
                        }
                    }
                    break;
                }
            }

            for (int i = 0; i < output.Length; i++)
            {
                if (output[i].Contains("Mottaget"))
                    output[i + 1] = "";
                if (output[i].Contains("Medlemsnummer:"))
                {
                    output[i] = "";
                    break;
                }
            }

            return output;
        }

        private static string[,] ReadEmails(string[] fileInput)
        {
            string[,] tempStringArray = new string[1, 1];
            int whereReceiptEnds;
            int lineCounter = 0;
            string tempstringForDate = "";
            bool notDoneWithMainArray = true;

            for (int i = 0; i < fileInput.Length; i++)
            {
                if (fileInput[i].Contains("------------------------------------------") && fileInput[i + 1].Contains("Totalt") && notDoneWithMainArray)
                {
                    whereReceiptEnds = i;

                    tempStringArray = new string[whereReceiptEnds, 2];

                    for (int j = 1; j <= whereReceiptEnds - 1; j++)
                    {
                        // Main logic of finding rows that are rows of valuable info.
                        // The output needed is only 1. Item & 2. Price of item.
                        // So everything else is not needed (such as weight or similar)
                        // If a row has a item & price, save it in array
                        // If a row doesn't have a price, but the next has, save it and skip a row in input
                        // If a row and the next one doesn't, but the 3rd has, save that and skip 2 rows. Could be done in a loop?
                        if (double.TryParse(fileInput[j].Substring(fileInput[j].Length - 6).Trim(), out double price))
                        {
                            tempStringArray[lineCounter, 0] = fileInput[j].Substring(0, fileInput[j].Length - 6).Trim();
                            tempStringArray[lineCounter, 1] = price.ToString();
                            lineCounter++;
                        }
                        else if (double.TryParse(fileInput[j + 1].Substring(fileInput[j + 1].Length - 6).Trim(), out double secondPrice) && !fileInput[j + 1].Contains("Extrapris"))
                        {
                            tempStringArray[lineCounter, 0] = fileInput[j].Substring(0, fileInput[j].Length - 6).Trim();
                            tempStringArray[lineCounter, 1] = secondPrice.ToString();
                            j++;
                            lineCounter++;
                        }
                        else if (double.TryParse(fileInput[j + 2].Substring(fileInput[j + 2].Length - 6).Trim(), out double thirdPrice) && !fileInput[j + 1].Contains("Extrapris"))
                        {
                            tempStringArray[lineCounter, 0] = fileInput[j].Substring(0, fileInput[j].Length - 6).Trim();
                            tempStringArray[lineCounter, 1] = thirdPrice.ToString();
                            j += 2;
                            lineCounter++;
                        }

                        if (fileInput[j + 1].Contains("Extrapris"))
                        {
                            tempStringArray[j, 0] = fileInput[j].Substring(0, fileInput[j].Length - 5).Trim() + " Extrapris";
                        }
                    }
                    i = lineCounter; // Skip some rows just to save a few loops.
                    notDoneWithMainArray = false;
                }

                if (fileInput[i].Contains("AID:")) // Per file, look for date, and append it to the end of the array.
                {
                    tempstringForDate = fileInput[i + 1].Substring(0, 10);
                    break;
                }
            }

            string[,] rowsFromReceipt = new string[lineCounter + 1, 2];

            //Efficiency note: there's probably a more clever way to do this, 
            for (int localvariable = 0; localvariable < lineCounter; localvariable++)
            {
                rowsFromReceipt[localvariable, 0] = tempStringArray[localvariable, 0];
                rowsFromReceipt[localvariable, 1] = tempStringArray[localvariable, 1];
            }

            rowsFromReceipt[lineCounter, 0] = tempstringForDate;

            return rowsFromReceipt;
        }

        static void Main()
        {
            logStrings = new List<string>();
            // foreach file in folder 
            // run method that places sorted information into a thing
            // send thing to below method
            List<string[,]> receiptData = new();

            string[] fileNames = System.IO.Directory.GetFiles(@"C:\Users\anrik\OneDrive\Dokument\Exceltest mail\");

            foreach (string file in fileNames)
            {
                receiptData.Add(ReadEmails(CleanEmails(@file)));
            }

            Console.WriteLine("Finished copying receipts, ready to send it to excel.");

            if (WriteToExcel(receiptData))
            {
                Console.WriteLine("Finished copying to excel, printing logfile");
                foreach (string errors in logStrings)
                    Console.WriteLine(errors);

                Console.WriteLine("Operations finished, thank you for choosing to use my shoddy code. Press enter to exit.");
            }
            else
                Console.WriteLine("The writing did not finish correctly. My apologies! :( \nPress enter to exit");

            Console.ReadLine();
            
        }

        private static bool WriteToExcel(List<string[,]> dataToWriteToExcel)
        {
            if (OperatingSystem.IsWindows())
            {
                Excel.Application excelApp = new();

                if (excelApp == null)
                {
                    Console.WriteLine("Yo waddup, excel didn't work correctly with yo code at the start.");
                    throw new Exception("Excel connection setup broke");
                }

                object misValue = System.Reflection.Missing.Value;

                Excel.Workbook excelWoorkbook = excelApp.Workbooks.Add(misValue);
                Excel.Worksheet excelSheet = (Excel.Worksheet)excelWoorkbook.Worksheets.Item[1];

                int currentRow = 1;
                string dateCopy = "";

                // Utdata jag vill ha: 
                // Datum - Sak - Pris ?
                // 1. hitta datum 2. Spara datum? 3. Skriv: datum, sak, pris

                foreach (string[,] receiptRows in dataToWriteToExcel)
                {

                    dateCopy = receiptRows[receiptRows.GetLength(0) - 1, 0];

                    for (int j = 0; j < receiptRows.GetLength(0); j++) // For each item in the string[], copy a date to column 1, copy first and 2nd value to column 2 & 3, then increment rows.
                    {
                        if (receiptRows[j, 0] != dateCopy)
                        {
                            excelSheet.Cells[currentRow, 1] = dateCopy;
                            excelSheet.Cells[currentRow, 2] = receiptRows[j, 0];
                            excelSheet.Cells[currentRow, 3] = receiptRows[j, 1];
                        }
                        else
                            break;
                        currentRow++;
                    }
                }

                // add writing to excel code here
                try
                {
                    excelWoorkbook.SaveAs(@"C:\Users\anrik\Desktop\Receipter.xlsx", Excel.XlFileFormat.xlWorkbookDefault, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                    excelWoorkbook.Close(true, misValue, misValue);
                    excelApp.Quit();
                }
                catch (COMException)
                {
                    Console.WriteLine("When trying to save file, user declined to rewrite. Will stop attempting to write to excel.");

                    Marshal.ReleaseComObject(excelSheet);
                    Marshal.ReleaseComObject(excelWoorkbook);
                    Marshal.ReleaseComObject(excelApp);
                    return false;
                }

                Marshal.ReleaseComObject(excelSheet);
                Marshal.ReleaseComObject(excelWoorkbook);
                Marshal.ReleaseComObject(excelApp);

                return true;
            }
            return false;
        }
    }
}
