using System;
using System.Collections.Generic;
using System.Text;

namespace FluentEpplus.Common
{
    public class ExcelUtil
    {
        public static string GetAddress(int row, int column)
        {
            return GetColumnLetter(column) + row.ToString();
        }

        public static int GetColumnNumber(string sCol, bool adjustZeroBasedIndex = true)
        {
            int col = 0;
            int len = sCol.Length - 1;

            for (int i = len; i >=0 ; i--)
            {
                col += (((int) sCol[i]) - 64) * (int) (Math.Pow(26, len - i));
            }

            return adjustZeroBasedIndex ? col - 1 : col;
        }

        public static string GetColumnLetter(int pos)
        {
            if (pos < 1) throw new ArgumentException("Column number is out of range.");

            var sCol = string.Empty;
            do
            {
                sCol = ((char) ('A' + ((pos - 1) % 26))) + sCol;
                pos = (pos - ((pos - 1) % 26)) / 26;
            } while (pos > 0);

            return sCol;

        }
    }
}
