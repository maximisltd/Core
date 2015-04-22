using System;
using System.Collections.Generic;
using System.Linq;

namespace Maximis.Toolkit
{
    public static class ExtensionMethods
    {
        public static IEnumerable<List<T>> InSetsOf<T>(this IEnumerable<T> source, int max)
        {
            List<T> toReturn = new List<T>(max);
            foreach (var item in source)
            {
                toReturn.Add(item);
                if (toReturn.Count == max)
                {
                    yield return toReturn;
                    toReturn = new List<T>(max);
                }
            }
            if (toReturn.Any())
            {
                yield return toReturn;
            }
        }

        public static string LeftOfFirst(this string s, char c)
        {
            int index = s.IndexOf(c);
            return index >= 0 ? s.Substring(0, index) : s;
        }

        public static string LeftOfLast(this string s, char c)
        {
            int index = s.LastIndexOf(c);
            return index >= 0 ? s.Substring(0, index) : s;
        }

        public static string RightOfFirst(this string s, char c)
        {
            int index = s.IndexOf(c);
            return index >= 0 ? s.Substring(index + 1) : s;
        }

        public static string RightOfLast(this string s, char c)
        {
            int index = s.LastIndexOf(c);
            return index >= 0 ? s.Substring(index + 1) : s;
        }

        public static Boolean ToBoolean(this string s, Boolean defVal = false)
        {
            Boolean b = defVal;
            Boolean.TryParse(s, out b);
            return b;
        }

        public static DateTime ToDateTime(this string s, bool defaultMax = false)
        {
            DateTime dt = DateTime.MinValue;
            if (DateTime.TryParse(s, out dt))
                return dt;
            else
                return defaultMax ? DateTime.MaxValue : DateTime.MinValue;
        }

        public static decimal ToDecimal(this string s, decimal defVal = decimal.MinValue)
        {
            decimal d = decimal.MinValue;
            return (decimal.TryParse(s, out d)) ? d : defVal;
        }

        public static double ToDouble(this string s, double defVal = double.MinValue)
        {
            double d = double.MinValue;
            return (double.TryParse(s, out d)) ? d : defVal;
        }

        public static Guid ToGuid(this string s)
        {
            Guid guid = Guid.Empty;
            Guid.TryParse(s, out guid);
            return guid;
        }

        public static int ToInt(this string s, int defVal = int.MinValue)
        {
            int i = int.MinValue;
            return (int.TryParse(s, out i)) ? i : defVal;
        }

        public static string TruncateWithEllipsis(this string s, int maxLength)
        {
            if (maxLength > 0 && s.Length > maxLength - 3) return s.Substring(0, maxLength - 3) + "...";
            return s;
        }
    }
}