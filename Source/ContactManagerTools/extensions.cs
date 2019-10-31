﻿using System;
using System.Collections.Generic;
using System.Data;
using System.Text;
using System.Text.RegularExpressions;

public static class ExtensionMethods {
    //
    //====================================================================================================
    //
    public static bool isBase64String(this string s) {
        s = s.Trim();
        return (s.Length % 4 == 0) && Regex.IsMatch(s, @"^[a-zA-Z0-9\+/]*={0,3}$", RegexOptions.None);

    }
    //
    //====================================================================================================
    /// <summary>
    /// example extention method
    /// </summary>
    /// <param name="value"></param>
    /// <returns></returns>
    public static string UppercaseFirstLetter(this string value) {
        // Uppercase the first letter in the string.
        if (value.Length > 0) {
            char[] array = value.ToCharArray();
            array[0] = char.ToUpper(array[0]);
            return new string(array);
        }
        return value;
    }
    //
    //====================================================================================================
    /// <summary>
    /// like vb Left. Return leftmost characters up to the maxLength (but no error if short)
    /// </summary>
    /// <param name="source"></param>
    /// <param name="maxLength"></param>
    /// <returns></returns>
    public static string Left(this string source, int maxLength) {
        if (string.IsNullOrEmpty(source)) {
            return "";
        } else if (maxLength < 0) {
            throw new ArgumentException("length [" + maxLength + "] must be 0+");
        } else if (source.Length <= maxLength) {
            return source;
        } else {
            return source.Substring(0, maxLength);
        }
    }
    //
    //====================================================================================================
    /// <summary>
    /// like vb Right()
    /// </summary>
    /// <param name="source"></param>
    /// <param name="maxLength"></param>
    /// <returns></returns>
    public static string Right(this string source, int maxLength) {
        if (string.IsNullOrEmpty(source)) {
            return "";
        } else if (maxLength < 0) {
            throw new ArgumentException("length [" + maxLength + "] must be 0+");
        } else if (source.Length <= maxLength) {
            return source;
        } else {
            return source.Substring(source.Length - maxLength, maxLength);
        }
    }
    //
    //====================================================================================================
    /// <summary>
    /// replacement for visual basic isNumeric
    /// </summary>
    /// <param name="expression"></param>
    /// <returns></returns>
    public static bool IsNumeric(this object expression) {
        try {
            if (expression == null) {
                return false;
            } else if (expression is DateTime) {
                return false;
            } else if ((expression is int) || (expression is Int16) || (expression is Int32) || (expression is Int64) || (expression is decimal) || (expression is float) || (expression is double) || (expression is bool)) {
                return true;
            } else if (expression is string) {
                double output = 0;
                return double.TryParse((string)expression, out output);
            } else {
                return false;
            }
        } catch {
            return false;
        }
    }
    //
    //====================================================================================================
    //
    public static bool isOld(this DateTime srcDate) {
        return (srcDate < new DateTime(1900, 1, 1));
    }
    //
    //====================================================================================================
    /// <summary>
    /// if date is invalid, set to minValue
    /// </summary>
    /// <param name="srcDate"></param>
    /// <returns></returns>
    public static DateTime MinValueIfOld(this DateTime srcDate) {
        DateTime returnDate = srcDate;
        if (srcDate < new DateTime(1900, 1, 1)) {
            returnDate = DateTime.MinValue;
        }
        return returnDate;
    }
    //
    //====================================================================================================
    /// <summary>
    /// Case insensitive version of String.Replace().
    /// </summary>
    /// <param name="s">String that contains patterns to replace</param>
    /// <param name="oldValue">Pattern to find</param>
    /// <param name="newValue">New pattern to replaces old</param>
    /// <param name="comparisonType">String comparison type</param>
    /// <returns></returns>
    public static string Replace(this string s, string oldValue, string newValue, StringComparison comparisonType) {
        if (s == null) return null;
        if (String.IsNullOrEmpty(oldValue)) return s;
        StringBuilder result = new StringBuilder(Math.Min(4096, s.Length));
        int pos = 0;
        while (true) {
            int i = s.IndexOf(oldValue, pos, comparisonType);
            if (i < 0) break;
            result.Append(s, pos, i - pos);
            result.Append(newValue);
            pos = i + oldValue.Length;
        }
        result.Append(s, pos, s.Length - pos);
        return result.ToString();
    }
    //
    //====================================================================================================
    /// <summary>
    /// return true/false if the string is a valid base64 encode. Use this to prevent throwing an exception on fail
    /// </summary>
    /// <param name="s"></param>
    /// <returns></returns>
    public static bool IsBase64String(this string s) {
        s = s.Trim();
        return (s.Length % 4 == 0) && Regex.IsMatch(s, @"^[a-zA-Z0-9\+/]*={0,3}$", RegexOptions.None);
    }
    //
    //====================================================================================================
    /// <summary>
    /// swap two entries in a list
    /// </summary>
    /// <typeparam name="T"></typeparam>
    /// <param name="list"></param>
    /// <param name="indexA"></param>
    /// <param name="indexB"></param>
    /// <returns></returns>
    public static IList<T> Swap<T>(this IList<T> list, int indexA, int indexB) {
        if (indexB > -1 && indexB < list.Count) {
            T tmp = list[indexA];
            list[indexA] = list[indexB];
            list[indexB] = tmp;
        }
        return list;
    }
    //
    //====================================================================================================
    /// <summary>
    /// convert a datatable to csv string
    /// </summary>
    /// <param name="dataTable"></param>
    /// <returns></returns>
    public static string ToCsv(this DataTable dataTable) {
        StringBuilder sbData = new StringBuilder();
        //
        // -- Only return Null if there is no structure.
        if (dataTable.Columns.Count == 0) return null;
        //
        // -- append header
        foreach (var col in dataTable.Columns) {
            if (col == null)
                sbData.Append(",");
            else
                sbData.Append("\"" + col.ToString().Replace("\"", "\"\"") + "\",");
        }
        sbData.Replace(",", System.Environment.NewLine, sbData.Length - 1, 1);
        //
        // -- append data
        foreach (DataRow dr in dataTable.Rows) {
            foreach (var column in dr.ItemArray) {
                if (column == null)
                    sbData.Append(",");
                else
                    sbData.Append("\"" + column.ToString().Replace("\"", "\"\"") + "\",");
            }
            sbData.Replace(",", System.Environment.NewLine, sbData.Length - 1, 1);
        }
        return sbData.ToString();
    }
}
