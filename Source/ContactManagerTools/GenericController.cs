
using System;
using System.Diagnostics;
using System.Collections.Generic;
using System.Data;
using System.Net;
using System.Web;
using System.Text;
using Contensive.BaseClasses;
using static Contensive.Addons.ContactManagerTools.Constants;

namespace Contensive.Addons.ContactManagerTools.Controllers {
    //
    //====================================================================================================
    /// <summary>
    /// controller for shared non-specific tasks
    /// </summary>
    public static  class GenericController {
        //
        //=============================================================================
        /// <summary>
        /// Modify the querystring at the end of a link. If there is no, question mark, the link argument is assumed to be a link, not the querysting
        /// </summary>
        /// <param name="link"></param>
        /// <param name="queryName"></param>
        /// <param name="queryValue"></param>
        /// <param name="addIfMissing"></param>
        /// <returns></returns>
        public static string modifyLinkQuery(string link, string queryName, string queryValue, bool addIfMissing = true) {
            string tempmodifyLinkQuery;
            try {
                string[] Element = Array.Empty<string>();
                int ElementCount = 0;
                int ElementPointer = 0;
                string[] NameValue = null;
                string UcaseQueryName = null;
                bool ElementFound = false;
                bool iAddIfMissing = false;
                string QueryString = null;
                //
                iAddIfMissing = addIfMissing;
                if (vbInstr(1, link, "?") != 0) {
                    tempmodifyLinkQuery = link.Left(vbInstr(1, link, "?") - 1);
                    QueryString = link.Substring(tempmodifyLinkQuery.Length + 1);
                } else {
                    tempmodifyLinkQuery = link;
                    QueryString = "";
                }
                UcaseQueryName = vbUCase(encodeRequestVariable(queryName));
                if (!string.IsNullOrEmpty(QueryString)) {
                    Element = QueryString.Split('&');
                    ElementCount = Element.GetUpperBound(0) + 1;
                    for (ElementPointer = 0; ElementPointer < ElementCount; ElementPointer++) {
                        NameValue = Element[ElementPointer].Split('=');
                        if (NameValue.GetUpperBound(0) == 1) {
                            if (vbUCase(NameValue[0]) == UcaseQueryName) {
                                if (string.IsNullOrEmpty(queryValue)) {
                                    Element[ElementPointer] = "";
                                } else {
                                    Element[ElementPointer] = queryName + "=" + queryValue;
                                }
                                ElementFound = true;
                                break;
                            }
                        }
                    }
                }
                if (!ElementFound && (!string.IsNullOrEmpty(queryValue))) {
                    //
                    // element not found, it needs to be added
                    //
                    if (iAddIfMissing) {
                        if (string.IsNullOrEmpty(QueryString)) {
                            QueryString = encodeRequestVariable(queryName) + "=" + encodeRequestVariable(queryValue);
                        } else {
                            QueryString = QueryString + "&" + encodeRequestVariable(queryName) + "=" + encodeRequestVariable(queryValue);
                        }
                    }
                } else {
                    //
                    // element found
                    //
                    QueryString = string.Join("&", Element);
                    if ((!string.IsNullOrEmpty(QueryString)) && (string.IsNullOrEmpty(queryValue))) {
                        //
                        // element found and needs to be removed
                        //
                        QueryString = vbReplace(QueryString, "&&", "&");
                        if (QueryString.Left(1) == "&") {
                            QueryString = QueryString.Substring(1);
                        }
                        if (QueryString.Substring(QueryString.Length - 1) == "&") {
                            QueryString = QueryString.Left(QueryString.Length - 1);
                        }
                    }
                }
                if (!string.IsNullOrEmpty(QueryString)) {
                    tempmodifyLinkQuery = tempmodifyLinkQuery + "?" + QueryString;
                }
            } catch (Exception ex) {
                throw new ApplicationException("Exception in modifyLinkQuery", ex);
            }
            //
            return tempmodifyLinkQuery;
        }
        //
        //=============================================================================
        /// <summary>
        /// Create the part of the sql where clause that is modified by the user, WorkingQuery is the original querystring to change, QueryName is the name part of the name pair to change, If the QueryName is not found in the string
        /// </summary>
        /// <param name="workingQuery"></param>
        /// <param name="queryName"></param>
        /// <param name="queryValue"></param>
        /// <param name="addIfMissing"></param>
        /// <returns></returns>
        //
        //=============================================================================
        //
        public static string modifyQueryString(string workingQuery, string queryName, string queryValue, bool addIfMissing = true) {
            string result;
            //
            if (workingQuery.IndexOf("?") >= 0) {
                return modifyLinkQuery(workingQuery, queryName, queryValue, addIfMissing);
            }
            result = modifyLinkQuery("?" + workingQuery, queryName, queryValue, addIfMissing);
            if (result.Length > 0) { return result.Substring(1); }
            return result;
        }
        //
        //========================================================================
        /// <summary>
        /// encode the name or value part of a querystring, to be parsed with & and = 
        /// </summary>
        /// <param name="Source"></param>
        /// <returns></returns>
        public static string encodeRequestVariable(string Source) {
            if (Source == null) {
                return "";
            }
            return System.Uri.EscapeDataString(Source);
        }
        //
        // ====================================================================================================
        /// <summary>
        /// create a linked string. replace the start <link> tag with the provided anchorTag, and convert the end </link> tag to an </a>
        /// </summary>
        /// <param name="AnchorTag">The anchor tag to be inserted (a complete <a> tag)</param>
        /// <param name="AnchorText">A string containing a <link>toBeAnchored</link> tag set</param>
        /// <returns></returns>

        public static int encodeInteger(object expression) {
            if (expression == null) {
                return 0;
            } else {
                string tmpString = expression.ToString();
                if (string.IsNullOrWhiteSpace(tmpString)) {
                    return 0;
                } else {
                    if (int.TryParse(tmpString, out int result)) {
                        return result;
                    } else {
                        return 0;
                    }
                }
            }

        }
        //
        //====================================================================================================
        //
        public static DateTime encodeDate(object Expression) {
            DateTime result = DateTime.MinValue;
            if (Expression is string) {
                // visual basic - when converting a date to a string, it converts minDate to "12:00:00 AM". 
                // however, Convert.ToDateTime() converts "12:00:00 AM" to the current date.
                // this is a terrible hack, but to be compatible with current software, "#12:00:00 AM#" must return mindate
                if ((String)Expression == "12:00:00 AM") {
                    return result;
                }
            }
            if (DateController.IsDate(Expression)) {
                result = Convert.ToDateTime(Expression);
            }
            return result;
        }
        // todo refactor out vb fpo
        //====================================================================================================
        /// <summary>
        /// temp methods to convert from vb, refactor out
        /// </summary>
        /// <param name="string1"></param>
        /// <param name="string2"></param>
        /// <param name="text1Binary2"></param>
        /// <returns></returns>
        public static int vbInstr(int startBase1, string string1, string string2, int text1Binary2) {
            if (string.IsNullOrEmpty(string1)) {
                return 0;
            } else {
                if (startBase1 < 1) {
                    throw new ArgumentException("Instr() start must be > 0.");
                } else {
                    if (text1Binary2 == 1) {
                        return string1.IndexOf(string2, startBase1 - 1, StringComparison.CurrentCultureIgnoreCase) + 1;
                    } else {
                        return string1.IndexOf(string2, startBase1 - 1, StringComparison.CurrentCulture) + 1;
                    }
                }
            }
        }
        //
        // ====================================================================================================
        //
        public static int vbInstr(int startBase1, string string1, string string2) {
            return vbInstr(startBase1, string1, string2, 2);
        }
        //
        //====================================================================================================
        //
        public static string vbReplace(string expression, string oldValue, string replacement, int compare) {
            if (string.IsNullOrEmpty(expression)) {
                return expression;
            } else if (string.IsNullOrEmpty(oldValue)) {
                return expression;
            } else {
                if (compare == 2) {
                    return expression.Replace(oldValue, replacement);
                } else {

                    StringBuilder sb = new StringBuilder();
                    int previousIndex = 0;
                    int Index = expression.IndexOf(oldValue, StringComparison.CurrentCultureIgnoreCase);
                    while (Index != -1) {
                        sb.Append(expression.Substring(previousIndex, Index - previousIndex));
                        sb.Append(replacement);
                        Index += oldValue.Length;
                        previousIndex = Index;
                        Index = expression.IndexOf(oldValue, Index, StringComparison.CurrentCultureIgnoreCase);
                    }
                    sb.Append(expression.Substring(previousIndex));
                    return sb.ToString();
                }
            }
        }
        //
        // ====================================================================================================
        //
        public static string vbReplace(string expression, string find, string replacement) { return vbReplace(expression, find, replacement, 1); }
        //
        //====================================================================================================
        /// <summary>
        /// Visual Basic UCase
        /// </summary>
        /// <param name="source"></param>
        /// <returns></returns>
        public static string vbUCase(string source) {
            if (string.IsNullOrEmpty(source)) {
                return "";
            } else {
                return source.ToUpper();
            }
        }
        //
        //========================================================================
        /// <summary>
        /// return the standard tablename fieldname path -- always lowercase.
        /// </summary>
        /// <param name="tableName"></param>
        /// <param name="fieldName"></param>
        /// <returns></returns>
        public static string getVirtualTableFieldUnixPath(string tableName, string fieldName) {
            string result = tableName + "/" + fieldName + "/";
            return result.ToLower().Replace(" ", "_").Replace(".", "_");
        }
        //
        //========================================================================
        //
        public static string getVirtualTableFieldIdUnixPath(string tableName, string fieldName, int recordID) {
            return getVirtualTableFieldUnixPath(tableName, fieldName) + recordID.ToString().PadLeft(12, '0') + "/";
        }
        //
        //========================================================================
        /// <summary>
        /// Create a filename for the Virtual Directory for a fieldtypeFile or Image (an uploaded file)
        /// </summary>
        public static string getVirtualRecordUnixPathFilename(string tableName, string fieldName, int recordID, string originalFilename) {
            //string iOriginalFilename = originalFilename.Replace(" ", "_").Replace(".", "_");
            return getVirtualTableFieldIdUnixPath(tableName, fieldName, recordID) + originalFilename;
        }

    }
}
