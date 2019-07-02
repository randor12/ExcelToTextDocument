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
using System.Data.OleDb;
using CsvHelper;

namespace ExcelParser
{
    class Program
    {
        private static int numFiles;
        private static int passes;

        /*
         * Check if object is numeric 
         */
        private static bool IsNumericType(object o)
        {
            switch (Type.GetTypeCode(o.GetType()))
            {
                case TypeCode.Byte:
                case TypeCode.SByte:
                case TypeCode.UInt16:
                case TypeCode.UInt32:
                case TypeCode.UInt64:
                case TypeCode.Int16:
                case TypeCode.Int32:
                case TypeCode.Int64:
                case TypeCode.Decimal:
                case TypeCode.Double:
                case TypeCode.Single:
                    return true;
                default:
                    return false;
            }
        }

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

        /**
         * The purpose of this function is to extract only the words that are arabic
         */
        public static List<string> GetWords(string input)
        {
            string[] delimiters = new string[] { " ", "_", ".", "!", "\"", ":", ";", "#", "(", ")", ",", "'", "{", "}", "-", "%",
            "؟", "،", "”", "“", "‘", "؛", "`", "?", ",", "+", "=", "*", "\\", "\n", "\t", "\r"};

            string[] words = input.Split(delimiters, StringSplitOptions.RemoveEmptyEntries);

            List<string> AllWords = new List<string>();

            foreach (var word in words)
            {
                AllWords.Add(word);
                
            }

            return AllWords;
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
                numFiles = Directory.GetFiles(sDir, "*.xlsx", SearchOption.AllDirectories).Length;

                /***************************************************************************************************
                 * SELECT BELOW FOR ALL FILES WITHIN THE DIRECTORY (INCLUDING SUB FILES)
                 * 
                 * WILL NEED TO ADD "SearchOption.AllDirectories" below in for loop
                 * 
                 *numFiles = Directory.GetFiles(sDir, "*", SearchOption.AllDirectories).Length;
                 ****************************************************************************************************/

                foreach (string f in Directory.GetFiles(sDir, "*.xlsx", SearchOption.AllDirectories))
                {
                    passes++;
                    Console.WriteLine("File: " + passes + " of " + numFiles);
                    Console.WriteLine("File Name: " + f);
                    //ParseCSV(f);
                    ParseFile(f);
                    //ParseExcel(f); // If using OLEDB
                }


            }
            catch (System.Exception excpt)
            {
                Console.WriteLine(excpt);
            }
        }
        

        /*
         * Parses a CSV file
         */
        public static void ParseCSV(string f)
        {
            string text = "";
            using (var reader = new StreamReader(f))
            {
                using (var csvReader = new CsvReader(reader))
                {
                    csvReader.Configuration.BadDataFound = null;

                    while (csvReader.Read())
                    {
                        
                        try
                        { 
                            var GetWords = csvReader.GetField(0);

                            if (GetWords != null)
                            {
                                text += GetWords.ToString() + "\n";
                            }
                        }
                        catch (Exception e)
                        {
                            Console.WriteLine(e.ToString());
                        }
                        
                        
                    }

                    Console.WriteLine("The words are: \n" + text);
                }
            }
        }

        
        /*
         * Process the data in the file
         */
        public static void ParseFile(string f)
        {

            FileStream fileStream = new FileStream("Verbatim.txt", FileMode.Append);
            StreamWriter writer = new StreamWriter(fileStream);

            Excel.Application excelApp = new Excel.Application
            {
                Visible = false
            };
            
            Excel.Workbook workbook = excelApp.Workbooks.Open(f, 0, true, 5, "", "", false, Excel.XlPlatform.xlWindows, 
                "", true, false, 0, true, false, false);

            Excel.Sheets excelSheet = workbook.Worksheets;
            string currentSheet = Path.GetFileName(f);
            int getEnd = currentSheet.IndexOf(".xlsx");
            
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
                string text = "";
                Console.WriteLine("Row number: " + i + " of " + totalRows);

                // Last 26
                for (int j = totalColumns - 26; j <= totalColumns; j++)
                {
                    var cell = (Excel.Range)excelWorksheet.Cells[i, j];
                    string GetString;

                    if (cell.Value == null || IsNumericType(cell.Value))
                    {
                        GetString = "";
                    }
                    else
                    {
                        GetString = (string)cell.Value;

                        List<string> AllWords = GetWords(GetString);

                        Console.WriteLine(AllWords.Count);

                        bool phraseGood = false;

                        foreach (var word in AllWords)
                        {
                            if (HasArabicGlyphs(word))
                            {
                                phraseGood = true;
                            }
                        }

                        Console.WriteLine(phraseGood);

                        if (phraseGood == true)
                        {

                            text += GetString + "\n";
                        }
                    }

                    
                }

                writer.WriteLine(text);
                
                Console.WriteLine("Lines: \n" + text);
                
            }

            writer.Close();
            
        }

        /*
         * attempt two at parsing excel (now using OLEDB instead of Microsoft.Excel)
         */
        public static void ParseExcel(string f)
        {
            string connectionString = "";
            string strFileType = Path.GetExtension(f).ToLower();
            string path = Path.GetFileName(f);

            string text = "";

            if (strFileType == ".xlsx")
            {
                connectionString = "Provider = Microsoft.Jet.OLEDB.4.0; Data Source = " + path + "; Extended Properties = \"text;HDR=Yes;FMT=Delimited\"";
            }

            string query = "SELECT [xxataf] FROM [" + path + "]";
            
            using (OleDbConnection connection = new OleDbConnection(connectionString))
            {
                OleDbCommand command = new OleDbCommand(query, connection);

                connection.Open();

                using (OleDbDataReader reader = command.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        var word = reader[0].ToString();

                        text += word + "\n";
                        Console.WriteLine("The words are: \n" + word);
                    }
                }
            }

            text = Regex.Replace(text, "<[^>]*>", string.Empty);

            text = Regex.Replace(text, @"^\s*$\n", " ", RegexOptions.Multiline);

            ProcessStringByWord(text);
        }

        /**
         * Process the data retrieved from the excel sheet
         */
        public static void ProcessStringByWord(string s)
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
