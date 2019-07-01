/*******************************************
 * The purpose of this program is to parse through text in an excel 
 * sheet
 * 
 * Ryan Nicholas
 * Imperium 
 * June 27, 2019
 * 
 *******************************************/

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Text.RegularExpressions;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelParser
{
    class Program
    {
        private static int numFiles;
        private static int passes;

        /*
         * Check for numbers
         */
        public static bool CheckNumbers(string word)
        {
            bool hasNumbers = false;

            foreach (var c in word)
            {
                string input = c.ToString();
                bool isNumber = Regex.IsMatch(input, "[0-9]", RegexOptions.Compiled);

                if (isNumber)
                {
                    hasNumbers = true;
                }
            }

            return hasNumbers;
        }

        /**
         * Determines if a word is arabic
         * @param text          Text
         * @return isArabic         Returns true if the whole word is arabic
         */
        public static bool HasArabicGlyphs(string text)
        {
            char[] glyphs = text.ToCharArray();

            foreach (char glyph in glyphs)
            {
                bool isArabic = false;

                if (glyph >= 0x600 && glyph <= 0x6ff)
                {
                    isArabic = true;
                }
                if (glyph >= 0x750 && glyph <= 0x77f)
                {
                    isArabic = true;
                }
                if (glyph >= 0xfb50 && glyph <= 0xfc3f)
                {
                    isArabic = true;
                }
                if (glyph >= 0xfe70 && glyph <= 0xfefc)
                {
                    isArabic = true;
                }

                if (!isArabic)
                {
                    return false;
                }
            }

            return true;
        }

        /**
         * The purpose of this function is to extract only the words that are arabic
         */
        public static List<string> GetArabicWords(string input)
        {
            string[] delimiters = new string[] { " ", "_", ".", "!", "\"", ":", ";", "#", "(", ")", ",", "'", "{", "}", "-", "%",
            "؟", "،", "”", "“", "‘", "؛", "`", "?", ",", "+", "=", "*", "\\", "\n", "\t"};

            string[] words = input.Split(delimiters, StringSplitOptions.RemoveEmptyEntries);

            List<string> ArabicWords = new List<string>();

            foreach (var word in words)
            {
                if (HasArabicGlyphs(word))
                {
                    if (!CheckNumbers(word))
                    {
                        if (word.Length <= 4000)
                        {
                            ArabicWords.Add(word);
                        }
                    }
                    
                }
            }

            return ArabicWords;
        }

        /*
         * Reverse text
         */
        public static string Reverse(string text)
        {
            if (text == null) return null;
            char[] array = text.ToCharArray();
            Array.Reverse(array);
            return new String(array);
        }

        /*
         * Process files
         * @param sDir          directory
         */
        public static void ProcessInputFiles(string sDir)
        {
            try
            {
                passes = 0;
                numFiles = Directory.GetFiles(sDir, "*.csv", SearchOption.AllDirectories).Length;

                /***************************************************************************************************
                 * SELECT BELOW FOR ALL FILES WITHIN THE DIRECTORY (INCLUDING SUB FILES)
                 * 
                 * WILL NEED TO ADD "SearchOption.AllDirectories" below in for loop
                 * 
                 *numFiles = Directory.GetFiles(sDir, "*", SearchOption.AllDirectories).Length;
                 ****************************************************************************************************/

                foreach (string f in Directory.GetFiles(sDir, "*.csv", SearchOption.AllDirectories))
                {
                    passes++;
                    Console.WriteLine("File: " + passes + " of " + numFiles);
                    Console.WriteLine("File Name: " + f);
                    ParseFile(f);
                }


            }
            catch (System.Exception excpt)
            {
                Console.WriteLine(excpt);
            }
        }
        
        /*
         * Process the data in the file
         */
        public static void ParseFile(string f)
        {
            string text = "";

            Excel.Application excelApp = new Excel.Application
            {
                Visible = false
            };

            Excel.Workbook workbook = excelApp.Workbooks.Open(f, 0, true, 5, "", "", false, Excel.XlPlatform.xlWindows, 
                "", true, false, 0, true, false, false);

            Excel.Sheets excelSheet = workbook.Worksheets;
            string currentSheet = Path.GetFileName(f);
            int getEnd = currentSheet.IndexOf(".csv");
            
            if (getEnd > 0)
            {
                currentSheet = currentSheet.Substring(0, getEnd);
            }

            Console.WriteLine("Sheet name: " + currentSheet);

            Excel.Worksheet excelWorksheet = (Excel.Worksheet)excelSheet[1];

            Excel.Range xlRange = excelWorksheet.UsedRange;
            
            int totalRows = xlRange.Rows.Count;

            int totalColumns = xlRange.Columns.Count;

            
            for (int i = 1; i <= totalRows; i++)
            {
                Console.WriteLine("Row number: " + i + " of " + totalRows);
                var cell = (Excel.Range)excelWorksheet.Cells[i, 1];
                var GetString = (string)cell.Value;
                
                if (GetString == null)
                {
                    GetString = "";
                }

                int index = GetString.IndexOf("_");

                Console.WriteLine("Index number is " + index);

                if (index > 0)
                {
                    GetString = GetString.Substring(index);
                }
                else
                {
                    GetString = "";
                }

                Console.WriteLine("Word: \n" + GetString);

                text += GetString + "\n";
            }
            

            /*
             * Tested individual rows to determine where the data was
            Console.WriteLine("Row number: " + 1 + " of " + totalColumns);
            var cell = xlRange.Cells[2, 1];
            string GetString = (string)cell.Value2;

            int index = GetString.IndexOf("18_40");

            Console.WriteLine("Index number is " + index);

            if (index > 0)
            {
                GetString = GetString.Substring(index);
            }
            else
            {
                GetString = "";
            }
            
            Console.WriteLine("Word: \n" + GetString);

            text += GetString + "\n";
            */

            text = Regex.Replace(text, "<[^>]*>", string.Empty);

            text = Regex.Replace(text, @"^\s*$\n", " ", RegexOptions.Multiline);

            ProcessString(text);
        }

        /**
         * Process the data retrieved from the excel sheet
         */
        public static void ProcessString(string s)
        {
            FileStream fileStream = new FileStream("Verbatim.txt", FileMode.Append);
            StreamWriter writer = new StreamWriter(fileStream);

            List<string> AllArabicWords = GetArabicWords(s);

            Console.WriteLine("Word: " + AllArabicWords.Count);

            int count = 0;

            foreach (string word in AllArabicWords)
            {
                writer.WriteLine(word);
                count++;
            }

            Console.WriteLine("Number of words: " + count);
            writer.Close();
            
        }

        static void Main(string[] args)
        {
            string dir = "C:\\Users\\rnicholas\\Documents\\ArabicVerbatims";

            Console.OutputEncoding = Encoding.Unicode;

            ProcessInputFiles(dir);

            Console.WriteLine("Completed");

            Console.ReadKey();
        }
    }
}
