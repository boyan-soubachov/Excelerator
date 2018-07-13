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

        public static string[] RangeCells(string range_start, string range_end)
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

        public static string[] FastCoordSplit(string address)
        {
            int i = 0;
            for (i = 0; i < address.Length; i++)
                if (char.IsDigit(address[i]))
                    break;

            return new string[2] { address.Substring(0, i), address.Substring(i, address.Length - i) };
        }

        public static int FastIntParse(string value_in)
        {
            int res = 0;
            for (int i = 0; i < value_in.Length; i++)
                res = res * 10 + (value_in[i] - 48);
            return res;
        }

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

        public static Tuple<int, int> RangeSize(string range_start, string range_end)
        {
            range_start = range_start.ToUpper();
            range_end = range_end.ToUpper();

            int[] start_coords = A1toR1C1(range_start);
            int[] end_coords = A1toR1C1(range_end);

            return new Tuple<int, int>(Math.Abs(end_coords[0] - start_coords[0] + 1), Math.Abs(end_coords[1] - start_coords[1] + 1));
        }

        public static string NextCell(string cell_start, int row, int col)
        {
            string res = "";

            if (cell_start.Contains(':'))
            {
                string[] str_range = cell_start.Split(':');

                int[] r1c1_address = A1toR1C1(str_range[0]);
                if (str_range[0][0] != '$')
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
    }
}
