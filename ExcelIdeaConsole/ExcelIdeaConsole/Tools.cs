/* Project Name: ExcelIdea
 * Copyright (C) Boyan Soubachov, 2014
 * */
using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.Concurrent;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Diagnostics;
using System.Security.Cryptography;

namespace ExcelIdea
{
    [Flags]
    public enum ResultEnum
    {
        //collection types
        Array = 0x01,
        Unit = 0x02,
        Mapped_Range = 0x04,
        Error = 0x08,
        Empty = 0x10,
        //data types
        Float = 0x20,
        String = 0x40,
    }

    public struct ResultStruct
    {
        public ResultEnum Type;
        public int Rows;
        public int Cols;
        public object Value;
    }

    public struct CellRef
    {
        public CellRef(int Sheet_Number, string Cell_Ref)
        {
            this.Sheet_Number = Sheet_Number;
            this.Cell_Ref = Cell_Ref;
        }
        public int Sheet_Number;
        public string Cell_Ref;
    }

    public struct RangeStruct
    {
        public int[] Sorted_Map;
        public object Values;
    }

    public static class RangeRepository
    {
        public static ConcurrentDictionary<string, RangeStruct> RangeRepositoryStore;
    }

    public class ComparisonComparer<T> : IComparer<T>
    {
        private readonly Comparison<T> comparison;

        public ComparisonComparer(Comparison<T> comparison)
        {
            this.comparison = comparison;
        }

        int IComparer<T>.Compare(T x, T y)
        {
            return comparison(x, y);
        }
    }

    class Tools
    {
        public const int RangeMapThreshold = 64;
        //static Regex coordSplitRegex = new Regex(@"([A-Z]+)(\d+)", RegexOptions.Compiled);
        static Regex cellsFromFormulaRegex = new Regex(@"\$?(?:\bXF[A-D]|X[A-E][A-Z]|[A-W][A-Z]{2}|[A-Z]{2}|[A-Z])\$?(?:104857[0-6]|10485[0-6]\d|1048[0-4]\d{2}|104[0-7]\d{3}|10[0-3]\d{4}|[1-9]\d{1,5}|[1-9])d?\b([:\s]\$?(?:\bXF[A-D]|X[A-E][A-Z]|[A-W][A-Z]{2}|[A-Z]{2}|[A-Z])\$?(?:104857[0-6]|10485[0-6]\d|1048[0-4]\d{2}|104[0-7]\d{3}|10[0-3]\d{4}|[1-9]\d{1,5}|[1-9])d?\b)?", RegexOptions.Compiled);
        
        [Serializable]
        public struct ArrayFormulaInfo
        {
            public int Rows;
            public int Cols;
            public string StartCell;
            public string EndCell;
        }

        public static string SHA512String(string input)
        {
            SHA512Managed sha_manager = new SHA512Managed();
            byte[] buff = Encoding.UTF32.GetBytes(input);
            return Convert.ToBase64String(sha_manager.ComputeHash(buff, 0, buff.Length));
        }

        public static List<string> GetCellsFromFormula(string formula)
        {
            List<string> res = new List<string>();
            Match reg_match = Regex.Match(formula, @"\$?(?:\bXF[A-D]|X[A-E][A-Z]|[A-W][A-Z]{2}|[A-Z]{2}|[A-Z])\$?(?:104857[0-6]|10485[0-6]\d|1048[0-4]\d{2}|104[0-7]\d{3}|10[0-3]\d{4}|[1-9]\d{1,5}|[1-9])d?\b([:\s]\$?(?:\bXF[A-D]|X[A-E][A-Z]|[A-W][A-Z]{2}|[A-Z]{2}|[A-Z])\$?(?:104857[0-6]|10485[0-6]\d|1048[0-4]\d{2}|104[0-7]\d{3}|10[0-3]\d{4}|[1-9]\d{1,5}|[1-9])d?\b)?");
            while (reg_match.Success)
            {
                res.Add(reg_match.Value);
                reg_match = reg_match.NextMatch();
            }
            return res;
        }

        /// <summary>
        /// Checks if a string is numeric
        /// </summary>
        /// <param name="input_string">The input string</param>
        /// <returns>A boolean indicating whether the string is purely numerical</returns>
        public static bool CheckNumeric(string input_string)
        {
            for (int i = 0; i < input_string.Length; i++)
                if (!char.IsDigit(input_string[i]) && input_string[i] != '.' && input_string[i] != '-')
                    return false;

            return true;
        }

        /// <summary>
        /// Splits the cell reference into a co-ordinate pair consisting of the column and row in a tuple
        /// </summary>
        /// <param name="reference">The cell reference to split</param>
        /// <returns></returns>
        public static Tuple<string,int> SplitCoordinates(string reference)
        {
            string[] temp = FastCoordSplit(reference);
            return new Tuple<string, int>(temp[0], FastIntParse(temp[1]));
        }

        /// <summary>
        /// Gets the cell address given row and column
        /// </summary>
        /// <param name="row">Row number</param>
        /// <param name="column">Column number</param>
        /// <returns></returns>
        public static string R1C1toA1(int row, int column)
        {
            string result = "";

            while (column > 0)
            {
                result = (char)(65 + ((column - 1) % 26)) + result;
                column = (column - 1) / 26;
            }

            result += row.ToString();
            return result;
        }

        public static int FastIntParse(string value_in)
        {
            int res = 0;
            for (int i = 0; i < value_in.Length; i++)
                res = res * 10 + (value_in[i] - 48);
            return res;
        }

        public static string[] FastCoordSplit(string address)
        {
            int i = 0;
            for (i = 0; i < address.Length; i++)
                if (char.IsDigit(address[i]))
                    break;

            return new string[2] { address.Substring(0, i), address.Substring(i, address.Length - i) };
        }

        public static bool HorizontalIncrement(string first_addr, string second_addr)
        {
            string[] first_split = FastCoordSplit(first_addr);
            string[] second_split = FastCoordSplit(second_addr);
            if (first_split[0] == second_split[0])
                return false;
            else
                return true;
        }

        /// <summary>
        /// Gets the cell row and column (R1C1) given A1 address
        /// </summary>
        /// <param name="address">A1-style address to convert</param>
        /// <returns></returns>
        public static int[] A1toR1C1(string address)
        {
            int[] result = new int[2];

            int start_number, start_col = 0;
            string start_word;
            string[] temp;
            temp = FastCoordSplit(address);
            start_word = temp[0];
            start_number = FastIntParse(temp[1]);
            for (int i = 0; i < start_word.Length; i++)
                start_col += ((int)start_word[i] - 64) * (int)Math.Pow(26, start_word.Length - i - 1);

            result[0] = start_number;
            result[1] = start_col;

            return result;
        }

        /// <summary>
        /// Gets the next cell relative to the current one, x-rows and y-columns away.
        /// </summary>
        /// <param name="cell_start">The base cell</param>
        /// <param name="row">How many rows away</param>
        /// <param name="col">How many columns away</param>
        /// <returns></returns>
        public static string NextCell(string cell_start, int row, int col)
        {
            string res = "";

            if (cell_start.Contains(':'))
            {
                string[] str_range = cell_start.Split(':');

                int[] r1c1_address = A1toR1C1(str_range[0]);
                if (str_range[0][0] !='$')
                    r1c1_address[0] += row;
                if (str_range[0].LastIndexOf('$') <= 0)
                    r1c1_address[1] += col;
                res = R1C1toA1(r1c1_address[0], r1c1_address[1]) + ":";

                r1c1_address = A1toR1C1(str_range[1]);
                if (str_range[1][0] != '$')
                    r1c1_address[0] += row;
                if (str_range[1].LastIndexOf('$') <= 0)
                    r1c1_address[1] += col;
                res += R1C1toA1(r1c1_address[0], r1c1_address[1]);
            }
            else
            {
                int[] r1c1_address = A1toR1C1(cell_start);
                r1c1_address[0] += row;
                r1c1_address[1] += col;
                res = R1C1toA1(r1c1_address[0], r1c1_address[1]);
            }

            return res;
        }

        public static string[] RangeCells(string range_start, string range_end) //NOTE: Possibly memoize these functions. Problem is the already-heavy memory usage.
        {
            range_start = range_start.ToUpper();
            range_end = range_end.ToUpper();

            int start_number, end_number, start_col = 0, end_col = 0;
            string start_word, end_word;
            string[] temp;

            temp = FastCoordSplit(range_start);
            start_word = temp[0];
            start_number = FastIntParse(temp[1]);
            temp = FastCoordSplit(range_end);
            end_number = FastIntParse(temp[1]);
            end_word = temp[0];

            for (int i = 0; i < start_word.Length; i++)
                start_col += ((int)start_word[i] - 64) * (int)Math.Pow(26, start_word.Length - i - 1);
            for (int i = 0; i < end_word.Length; i++)
                end_col += ((int)end_word[i] - 64) * (int)Math.Pow(26, end_word.Length - i - 1);

            int col_width = end_col - start_col + 1;
            int row_length = end_number - start_number + 1;

            string[] output = new string[Math.Max(col_width, 1) * row_length];

            for (int i = start_col; i < end_col + 1; i++) //this is faster single-threaded, slower when using Parallel.For!
            {
                string result = "";
                int z = i;

                while (z > 0)
                {
                    result = (char)(65 + ((z - 1) % 26)) + result;
                    z = (z - 1) / 26;
                }

                for (int j = start_number; j <= end_number; j++) //tried parallelising, no performance gain due to simplicity/granularity of the operation
                    output[(i - start_col) * row_length + (j - start_number)] = result + j.ToString();
            }

            return output;
        }

        /// <summary>
        /// Returns the dimensions of a given cell range
        /// </summary>
        /// <param name="range_start">The start address of the cell range (A1-notation)</param>
        /// <param name="range_end">The end address of the cell range (A1-notation)</param>
        /// <returns>A rows x cols tuple denoting the range dimensions (inclusive)</returns>
        public static Tuple<int,int> RangeSize(string range_start, string range_end)
        {
            range_start = range_start.ToUpper();
            range_end = range_end.ToUpper();

            int[] start_coords = A1toR1C1(range_start);
            int[] end_coords = A1toR1C1(range_end);

            return new Tuple<int, int>(Math.Abs(end_coords[0] - start_coords[0] + 1), Math.Abs(end_coords[1] - start_coords[1] + 1));
        }

        /// <summary>
        /// Determines whether a given string could be in a valid range of cell addresses.
        /// </summary>
        /// <param name="input">Input string to test</param>
        /// <returns></returns>
        public static bool IsCell(string input)
        {
            StringBuilder sb_letters = new StringBuilder();
            StringBuilder sb_numbers = new StringBuilder();

            while (input.Length > 0)
            {
                if (char.IsLetter(input[0]))
                {
                    if (sb_numbers.Length > 0) //if [letter][number][letter], cannot be a cell
                        return false;
                    sb_letters.Append(char.ToUpper(input[0]));
                    if (sb_letters.Length > 3) //if more than 3 letters, cannot be cell
                        return false;
                }
                else if (char.IsDigit(input[0]))
                {
                    if (sb_letters.Length == 0) //if no letter before, cannot be cell
                        return false;
                    sb_numbers.Append(input[0]);
                }
                input = input.Remove(0, 1);
            }

            if (FastIntParse(sb_numbers.ToString()) <= 1048576 && string.CompareOrdinal(sb_letters.ToString(),"XFD") <= 0)
                return true;
            else 
                return false;
        }

        /// <summary>
        /// Binary searches an array with a map for a specific value
        /// </summary>
        /// <typeparam name="T">The array and value type to search for</typeparam>
        /// <param name="search_space">The search space (array) in which to search</param>
        /// <param name="search_value">The value to search for</param>
        /// <param name="sorted_map">The map of the search space which has it sorted ordinally</param>
        /// <returns>The index (before the map) where the result may be found or the bitwise complement of the next largest value</returns>
        public static int BinarySearch<T>(T[] search_space, T search_value, int[] sorted_map)
        {
            //NB: This can be optimised in GPU, will speed it up tremendously.
            int lower, upper, index, comp;

            upper = search_space.Length - 1;
            lower = 0;
            comp = 0;
            index = search_space.Length + 1;

            while (lower <= upper)
            {
                index = (lower + upper) / 2;
                comp = string.CompareOrdinal(search_value.ToString(), search_space[sorted_map[index]].ToString());

                if (comp > 0) //if greater
                    lower = index + 1;
                else if (comp < 0) //if less
                    upper = index - 1;
                else if (comp == 0) //equal
                    break;
            }

            if (comp != 0)
                if (comp > 0)
                    return ~(index + 1);
                else
                    return ~index;
            else
                return index;
        }
    }
}
