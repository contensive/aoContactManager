
using Contensive.BaseClasses;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static Contensive.Addons.ContactManager.GenericController;
using static Contensive.Addons.ContactManager.Constants;
using static Contensive.Addons.ContactManager.HtmlController;

namespace Contensive.Addons.ContactManager {
    public class AdminUIController {
        public  enum SortingStateEnum {
            NotSortable = 0,
            SortableSetAZ = 1,
            SortableSetza = 2,
            SortableNotSet = 3
        }
        //
        //====================================================================================================
        //   Report Cell Header
        //       ColumnPtr       :   0 based column number
        //       Title
        //       Width           :   if just a number, assumed to be px in style and an image is added
        //                       :   if 00px, image added with the numberic part
        //                       :   if not number, added to style as is
        //       align           :   style value
        //       ClassStyle      :   class
        //       RQS
        //       SortingState
        //
        public static string getReport_CellHeader(CPBaseClass cp, int ColumnPtr, string Title, string Width, string Align, string ClassStyle, string RefreshQueryString, SortingStateEnum SortingState) {
            string result = "";
            try {
                string Style = null;
                string Copy = null;
                string QS = null;
                int WidthTest = 0;
                //
                if (string.IsNullOrEmpty(Title)) {
                    Copy = "&nbsp;";
                } else {
                    Copy = Title.Replace( " ", "&nbsp;");
                }
                Style = "VERTICAL-ALIGN:bottom;";
                if (string.IsNullOrEmpty(Align)) {
                } else {
                    Style = Style + "TEXT-ALIGN:" + Align + ";";
                }
                //
                switch (SortingState) {
                    case SortingStateEnum.SortableNotSet:
                        QS = GenericController.modifyQueryString(RefreshQueryString, "ColSort", ((int)SortingStateEnum.SortableSetAZ).ToString(), true);
                        QS = GenericController.modifyQueryString(QS, "ColPtr", ColumnPtr.ToString(), true);
                        Copy = "<a href=\"?" + QS + "\" title=\"Sort A-Z\" class=\"ccAdminListCaption\">" + Copy + "</a>";
                        break;
                    case SortingStateEnum.SortableSetza:
                        QS = GenericController.modifyQueryString(RefreshQueryString, "ColSort", ((int)SortingStateEnum.SortableSetAZ).ToString(), true);
                        QS = GenericController.modifyQueryString(QS, "ColPtr", ColumnPtr.ToString(), true);
                        Copy = "<a href=\"?" + QS + "\" title=\"Sort A-Z\" class=\"ccAdminListCaption\">" + Copy + "<img src=\"/ccLib/images/arrowup.gif\" width=8 height=8 border=0></a>";
                        break;
                    case SortingStateEnum.SortableSetAZ:
                        QS = GenericController.modifyQueryString(RefreshQueryString, "ColSort", ((int)SortingStateEnum.SortableSetza).ToString(), true);
                        QS = GenericController.modifyQueryString(QS, "ColPtr", ColumnPtr.ToString(), true);
                        Copy = "<a href=\"?" + QS + "\" title=\"Sort Z-A\" class=\"ccAdminListCaption\">" + Copy + "<img src=\"/ccLib/images/arrowdown.gif\" width=8 height=8 border=0></a>";
                        break;
                }
                //
                if (!string.IsNullOrEmpty(Width)) {
                    Style = Style + "width:" + Width + ";";
                    //WidthTest = GenericController.encodeInteger(Width.ToLower().Replace("px", ""));
                    //if (WidthTest != 0) {
                    //    Style = Style + "width:" + WidthTest + "px;";
                    //    Copy += "<img alt=\"space\" src=\"/ccLib/images/spacer.gif\" width=\"" + WidthTest + "\" height=1 border=0>";
                    //    //Copy = Copy & "<br><img alt=""space"" src=""/ccLib/images/spacer.gif"" width=""" & WidthTest & """ height=1 border=0>"
                    //} else {
                    //}
                }
                result = "\r\n<td style=\"" + Style + "\" class=\"" + ClassStyle + "\">" + Copy + "</td>";
            } catch (Exception ) {
                throw;
            }
            return result;
        }
        //
        //====================================================================================================
        /// <summary>
        /// returns the integer column ptr of the column last selected
        /// </summary>
        /// <param name="core"></param>
        /// <param name="DefaultSortColumnPtr"></param>
        /// <returns></returns>
        public static int getReportSortColumnPtr(CPBaseClass cp, int DefaultSortColumnPtr) {
            int tempGetReportSortColumnPtr = 0;
            string VarText;
            //
            VarText = cp.Doc.GetText("ColPtr");
            tempGetReportSortColumnPtr = GenericController.encodeInteger(VarText);
            if ((tempGetReportSortColumnPtr == 0) && (VarText != "0")) {
                tempGetReportSortColumnPtr = DefaultSortColumnPtr;
            }
            return tempGetReportSortColumnPtr;
        }
        //
        //====================================================================================================
        //   Get Sort Column Type
        //
        //   returns the integer for the type of sorting last requested
        //       0 = nothing selected
        //       1 = sort A-Z
        //       2 = sort Z-A
        //
        //   Orderby is generated by the selection of headers captions by the user
        //   It is up to the calling program to call GetReportOrderBy to get the orderby and use it in the query to generate the cells
        //   This call returns a comma delimited list of integers representing the columns to sort
        //
        public static int getReportSortType(CPBaseClass cp) {
            int tempGetReportSortType = 0;
            string VarText;
            //
            VarText = cp.Doc.GetText("ColPtr");
            if ((GenericController.encodeInteger(VarText) != 0) || (VarText == "0")) {
                //
                // A valid ColPtr was found
                //
                tempGetReportSortType = cp.Doc.GetInteger("ColSort");
            } else {
                tempGetReportSortType = (int)SortingStateEnum.SortableSetAZ;
            }
            return tempGetReportSortType;
        }
        public static string kmaStartTableCell(string widthPercent , int dontKNow , bool isRowEvent , string align ) {
            return "<td width=\"" + widthPercent + "\">";
        }
        //
        //
        //====================================================================================================
        public static string getReport(CPBaseClass cp, int RowCount, string[] ColCaption, string[] ColAlign, string[] ColWidth, string[,] Cells, int PageSize, int PageNumber, string PreTableCopy, string PostTableCopy, int DataRowCount, string ClassStyle) {
            string result = "";
            try {
                int ColCnt = Cells.GetUpperBound(1);
                bool[] ColSortable = new bool[ColCnt + 1];
                for (int Ptr = 0; Ptr < ColCnt; Ptr++) {
                    ColSortable[Ptr] = false;
                }
                //
                result = getReport2(cp, RowCount, ColCaption, ColAlign, ColWidth, Cells, PageSize, PageNumber, PreTableCopy, PostTableCopy, DataRowCount, ClassStyle, ColSortable, 0);
            } catch (Exception) {
                throw;
            }
            return result;
        }
        //
        //====================================================================================================
        public static string getReport2(CPBaseClass cp, int RowCount, string[] ColCaption, string[] ColAlign, string[] ColWidth, string[,] Cells, int PageSize, int PageNumber, string PreTableCopy, string PostTableCopy, int DataRowCount, string ClassStyle, bool[] ColSortable, int DefaultSortColumnPtr) {
            string result = "";
            try {
                string RQS = null;
                int RowBAse = 0;
                var Content = new StringBuilder();
                var Stream = new StringBuilder();
                int ColumnCount = 0;
                int ColumnPtr = 0;
                string ColumnWidth = null;
                int RowPointer = 0;
                string WorkingQS = null;
                //
                int PageCount = 0;
                int PagePointer = 0;
                int LinkCount = 0;
                int ReportPageNumber = 0;
                int ReportPageSize = 0;
                string iClassStyle = null;
                int SortColPtr = 0;
                int SortColType = 0;
                //
                ReportPageNumber = PageNumber;
                if (ReportPageNumber == 0) {
                    ReportPageNumber = 1;
                }
                ReportPageSize = PageSize;
                if (ReportPageSize < 1) {
                    ReportPageSize = 50;
                }
                //
                iClassStyle = ClassStyle;
                if (string.IsNullOrEmpty(iClassStyle)) {
                    iClassStyle = "ccPanel";
                }
                //If IsArray(Cells) Then
                ColumnCount = Cells.GetUpperBound(1);
                //End If
                RQS = cp.Doc.RefreshQueryString;
                //
                SortColPtr = getReportSortColumnPtr(cp, DefaultSortColumnPtr);
                SortColType = getReportSortType(cp);
                //
                // ----- Start the table
                //
                Content.Append(HtmlController.tableStart(3, 1, 0));
                //
                // ----- Header
                //
                Content.Append("\r\n<tr>");
                Content.Append(getReport_CellHeader(cp, 0, "&nbsp", "50px", "Right", "ccAdminListCaption", RQS, SortingStateEnum.NotSortable));
                for (ColumnPtr = 0; ColumnPtr < ColumnCount; ColumnPtr++) {
                    ColumnWidth = ColWidth[ColumnPtr];
                    if (!ColSortable[ColumnPtr]) {
                        //
                        // not sortable column
                        //
                        Content.Append(getReport_CellHeader(cp, ColumnPtr, ColCaption[ColumnPtr], ColumnWidth, ColAlign[ColumnPtr], "ccAdminListCaption", RQS, SortingStateEnum.NotSortable));
                    } else if (ColumnPtr == SortColPtr) {
                        //
                        // This is the current sort column
                        //
                        Content.Append(getReport_CellHeader(cp, ColumnPtr, ColCaption[ColumnPtr], ColumnWidth, ColAlign[ColumnPtr], "ccAdminListCaption", RQS, (SortingStateEnum)SortColType));
                    } else {
                        //
                        // Column is sortable, but not selected
                        //
                        Content.Append(getReport_CellHeader(cp, ColumnPtr, ColCaption[ColumnPtr], ColumnWidth, ColAlign[ColumnPtr], "ccAdminListCaption", RQS, SortingStateEnum.SortableNotSet));
                    }

                    //If ColumnPtr = SortColPtr Then
                    //    '
                    //    ' This column is currently the active sort
                    //    '
                    //    Call Content.Append(GetReport_CellHeader(ColumnPtr, ColCaption(ColumnPtr), ColumnWidth, ColAlign(ColumnPtr), "ccAdminListCaption", RQS, SortColType))
                    //Else
                    //    Call Content.Append(GetReport_CellHeader(ColumnPtr, ColCaption(ColumnPtr), ColumnWidth, ColAlign(ColumnPtr), "ccAdminListCaption", RQS, SortingStateEnum.SortableNotSet))
                    //End If
                }
                Content.Append("\r\n</tr>");
                //
                // ----- Data
                //
                if (RowCount == 0) {
                    Content.Append("\r\n<tr>");
                    Content.Append(getReport_Cell(cp, (RowBAse + RowPointer).ToString(), "right", 1, RowPointer));
                    Content.Append(getReport_Cell(cp, "-- End --", "left", ColumnCount, 0));
                    Content.Append("\r\n</tr>");
                } else {
                    RowBAse = (ReportPageSize * (ReportPageNumber - 1)) + 1;
                    for (RowPointer = 0; RowPointer < RowCount; RowPointer++) {
                        Content.Append("\r\n<tr>");
                        Content.Append(getReport_Cell(cp, (RowBAse + RowPointer).ToString(), "right", 1, RowPointer));
                        for (ColumnPtr = 0; ColumnPtr < ColumnCount; ColumnPtr++) {
                            Content.Append(getReport_Cell(cp, Cells[RowPointer, ColumnPtr], ColAlign[ColumnPtr], 1, RowPointer));
                        }
                        Content.Append("\r\n</tr>");
                    }
                }
                //
                // ----- End Table
                //
                Content.Append(Constants.kmaEndTable);
                result += Content.ToString();
                //
                // ----- Post Table copy
                //
                if ((DataRowCount / (double)ReportPageSize) != Math.Floor((DataRowCount / (double)ReportPageSize))) {
                    PageCount = GenericController.encodeInteger((DataRowCount / (double)ReportPageSize) + 0.5);
                } else {
                    PageCount = GenericController.encodeInteger(DataRowCount / (double)ReportPageSize);
                }
                if (PageCount > 1) {
                    result += "<br>Page " + ReportPageNumber + " (Row " + (RowBAse) + " of " + DataRowCount + ")";
                    if (PageCount > 20) {
                        PagePointer = ReportPageNumber - 10;
                    }
                    if (PagePointer < 1) {
                        PagePointer = 1;
                    }
                    if (PageCount > 1) {
                        result += "<br>Go to Page ";
                        if (PagePointer != 1) {
                            WorkingQS = cp.Doc.RefreshQueryString;
                            WorkingQS = GenericController.modifyQueryString(WorkingQS, "GotoPage", "1", true);
                            result += "<a href=\"" + "?" + WorkingQS + "\">1</A>...&nbsp;";
                        }
                        WorkingQS = cp.Doc.RefreshQueryString;
                        WorkingQS = GenericController.modifyQueryString(WorkingQS, RequestNamePageSize, ReportPageSize.ToString(), true);
                        while ((PagePointer <= PageCount) && (LinkCount < 20)) {
                            if (PagePointer == ReportPageNumber) {
                                result += PagePointer + "&nbsp;";
                            } else {
                                WorkingQS = GenericController.modifyQueryString(WorkingQS, RequestNamePageNumber, PagePointer.ToString(), true);
                                result += "<a href=\"" + "?" + WorkingQS + "\">" + PagePointer + "</A>&nbsp;";
                            }
                            PagePointer = PagePointer + 1;
                            LinkCount = LinkCount + 1;
                        }
                        if (PagePointer < PageCount) {
                            WorkingQS = GenericController.modifyQueryString(WorkingQS, RequestNamePageNumber, PageCount.ToString(), true);
                            result += "...<a href=\"" + "?" + WorkingQS + "\">" + PageCount + "</A>&nbsp;";
                        }
                        if (ReportPageNumber < PageCount) {
                            WorkingQS = GenericController.modifyQueryString(WorkingQS, RequestNamePageNumber, (ReportPageNumber + 1).ToString(), true);
                            result += "...<a href=\"" + "?" + WorkingQS + "\">next</A>&nbsp;";
                        }
                        result += "<br>&nbsp;";
                    }
                }
                //
                result = ""
                + PreTableCopy + "<table border=0 cellpadding=0 cellspacing=0 width=\"100%\"><tr><td style=\"padding:10px;\">"
                + result + "</td></tr></table>"
                + PostTableCopy + "";
            } catch (Exception ex) {
                throw;
            }
            return result;
        }
        //
        //====================================================================================================
        public  static string getReport_Cell(CPBaseClass cp, string Copy, string Align, int Columns, int RowPointer) {
            string tempGetReport_Cell = null;
            string iAlign = null;
            string Style = null;
            string CellCopy = null;
            //
            iAlign = encodeEmpty(Align, "left");
            //
            if ((RowPointer % 2) > 0) {
                Style = "ccAdminListRowOdd";
            } else {
                Style = "ccAdminListRowEven";
            }
            //
            tempGetReport_Cell = "\r\n<td valign=\"middle\" align=\"" + iAlign + "\" class=\"" + Style + "\"";
            if (Columns > 1) {
                tempGetReport_Cell = tempGetReport_Cell + " colspan=\"" + Columns + "\"";
            }
            //
            CellCopy = Copy;
            if (string.IsNullOrEmpty(CellCopy)) {
                CellCopy = "&nbsp;";
            }
            return tempGetReport_Cell + "><span class=\"ccSmall\">" + CellCopy + "</span></td>";
        }
        //
        ////====================================================================================================
        //public static string getReport(CPBaseClass cp, int RowCount, string[] ColCaption, string[] ColAlign, string[] ColWidth, string[,] Cells, int PageSize, int PageNumber, string PreTableCopy, string PostTableCopy, int DataRowCount, string ClassStyle) {
        //    string result = "";
        //    try {
        //        int ColCnt = Cells.GetUpperBound(1);
        //        bool[] ColSortable = new bool[ColCnt + 1];
        //        for (int Ptr = 0; Ptr < ColCnt; Ptr++) {
        //            ColSortable[Ptr] = false;
        //        }
        //        //
        //        result = getReport2(cp, RowCount, ColCaption, ColAlign, ColWidth, Cells, PageSize, PageNumber, PreTableCopy, PostTableCopy, DataRowCount, ClassStyle, ColSortable, 0);
        //    } catch (Exception ex) {

        //    }
        //    return result;
        //}
        ////
        ////====================================================================================================
        //public static string getReport2(CPBaseClass cp, int RowCount, string[] ColCaption, string[] ColAlign, string[] ColWidth, string[,] Cells, int PageSize, int PageNumber, string PreTableCopy, string PostTableCopy, int DataRowCount, string ClassStyle, bool[] ColSortable, int DefaultSortColumnPtr) {
        //    string result = "";
        //    try {
        //        string RQS = null;
        //        int RowBAse = 0;
        //        var Content = new StringBuilder();
        //        var Stream = new StringBuilder();
        //        int ColumnCount = 0;
        //        int ColumnPtr = 0;
        //        string ColumnWidth = null;
        //        int RowPointer = 0;
        //        string WorkingQS = null;
        //        //
        //        int PageCount = 0;
        //        int PagePointer = 0;
        //        int LinkCount = 0;
        //        int ReportPageNumber = 0;
        //        int ReportPageSize = 0;
        //        string iClassStyle = null;
        //        int SortColPtr = 0;
        //        int SortColType = 0;
        //        //
        //        ReportPageNumber = PageNumber;
        //        if (ReportPageNumber == 0) {
        //            ReportPageNumber = 1;
        //        }
        //        ReportPageSize = PageSize;
        //        if (ReportPageSize < 1) {
        //            ReportPageSize = 50;
        //        }
        //        //
        //        iClassStyle = ClassStyle;
        //        if (string.IsNullOrEmpty(iClassStyle)) {
        //            iClassStyle = "ccPanel";
        //        }
        //        //If IsArray(Cells) Then
        //        ColumnCount = Cells.GetUpperBound(1);
        //        //End If
        //        RQS = cp.Doc.RefreshQueryString;
        //        //
        //        SortColPtr = AdminUIController.GetReportSortColumnPtr(cp, DefaultSortColumnPtr);
        //        SortColType = AdminUIController.GetReportSortType(cp);
        //        //
        //        // ----- Start the table
        //        //
        //        Content.Append(HtmlController.tableStart(3, 1, 0));
        //        //
        //        // ----- Header
        //        //
        //        Content.Append("\r\n<tr>");
        //        Content.Append(AdminUIController.getReport_CellHeader(cp, 0, "&nbsp", "50px", "Right", "ccAdminListCaption", RQS, SortingStateEnum.NotSortable));
        //        for (ColumnPtr = 0; ColumnPtr < ColumnCount; ColumnPtr++) {
        //            ColumnWidth = ColWidth[ColumnPtr];
        //            if (!ColSortable[ColumnPtr]) {
        //                //
        //                // not sortable column
        //                //
        //                Content.Append(AdminUIController.getReport_CellHeader(cp, ColumnPtr, ColCaption[ColumnPtr], ColumnWidth, ColAlign[ColumnPtr], "ccAdminListCaption", RQS, SortingStateEnum.NotSortable));
        //            } else if (ColumnPtr == SortColPtr) {
        //                //
        //                // This is the current sort column
        //                //
        //                Content.Append(AdminUIController.getReport_CellHeader(cp, ColumnPtr, ColCaption[ColumnPtr], ColumnWidth, ColAlign[ColumnPtr], "ccAdminListCaption", RQS, (SortingStateEnum)SortColType));
        //            } else {
        //                //
        //                // Column is sortable, but not selected
        //                //
        //                Content.Append(AdminUIController.getReport_CellHeader(cp, ColumnPtr, ColCaption[ColumnPtr], ColumnWidth, ColAlign[ColumnPtr], "ccAdminListCaption", RQS, SortingStateEnum.SortableNotSet));
        //            }

        //            //If ColumnPtr = SortColPtr Then
        //            //    '
        //            //    ' This column is currently the active sort
        //            //    '
        //            //    Call Content.Append(GetReport_CellHeader(ColumnPtr, ColCaption(ColumnPtr), ColumnWidth, ColAlign(ColumnPtr), "ccAdminListCaption", RQS, SortColType))
        //            //Else
        //            //    Call Content.Append(GetReport_CellHeader(ColumnPtr, ColCaption(ColumnPtr), ColumnWidth, ColAlign(ColumnPtr), "ccAdminListCaption", RQS, SortingStateEnum.SortableNotSet))
        //            //End If
        //        }
        //        Content.Append("\r\n</tr>");
        //        //
        //        // ----- Data
        //        //
        //        if (RowCount == 0) {
        //            Content.Append("\r\n<tr>");
        //            Content.Append(getReport_Cell(cp, (RowBAse + RowPointer).ToString(), "right", 1, RowPointer));
        //            Content.Append(getReport_Cell(cp, "-- End --", "left", ColumnCount, 0));
        //            Content.Append("\r\n</tr>");
        //        } else {
        //            RowBAse = (ReportPageSize * (ReportPageNumber - 1)) + 1;
        //            for (RowPointer = 0; RowPointer < RowCount; RowPointer++) {
        //                Content.Append("\r\n<tr>");
        //                Content.Append(getReport_Cell(cp, (RowBAse + RowPointer).ToString(), "right", 1, RowPointer));
        //                for (ColumnPtr = 0; ColumnPtr < ColumnCount; ColumnPtr++) {
        //                    Content.Append(getReport_Cell(cp, Cells[RowPointer, ColumnPtr], ColAlign[ColumnPtr], 1, RowPointer));
        //                }
        //                Content.Append("\r\n</tr>");
        //            }
        //        }
        //        //
        //        // ----- End Table
        //        //
        //        Content.Append(Constants.kmaEndTable);
        //        result += Content.ToString();
        //        //
        //        // ----- Post Table copy
        //        //
        //        if ((DataRowCount / (double)ReportPageSize) != Math.Floor((DataRowCount / (double)ReportPageSize))) {
        //            PageCount = GenericController.encodeInteger((DataRowCount / (double)ReportPageSize) + 0.5);
        //        } else {
        //            PageCount = GenericController.encodeInteger(DataRowCount / (double)ReportPageSize);
        //        }
        //        if (PageCount > 1) {
        //            result += "<br>Page " + ReportPageNumber + " (Row " + (RowBAse) + " of " + DataRowCount + ")";
        //            if (PageCount > 20) {
        //                PagePointer = ReportPageNumber - 10;
        //            }
        //            if (PagePointer < 1) {
        //                PagePointer = 1;
        //            }
        //            if (PageCount > 1) {
        //                result += "<br>Go to Page ";
        //                if (PagePointer != 1) {
        //                    WorkingQS = cp.Doc.RefreshQueryString;
        //                    WorkingQS = GenericController.modifyQueryString(WorkingQS, "GotoPage", "1", true);
        //                    result += "<a href=\"" + "?" + WorkingQS + "\">1</A>...&nbsp;";
        //                }
        //                WorkingQS = cp.Doc.RefreshQueryString;
        //                WorkingQS = GenericController.modifyQueryString(WorkingQS, RequestNamePageSize, ReportPageSize.ToString(), true);
        //                while ((PagePointer <= PageCount) && (LinkCount < 20)) {
        //                    if (PagePointer == ReportPageNumber) {
        //                        result += PagePointer + "&nbsp;";
        //                    } else {
        //                        WorkingQS = GenericController.modifyQueryString(WorkingQS, RequestNamePageNumber, PagePointer.ToString(), true);
        //                        result += "<a href=\"" + "?" + WorkingQS + "\">" + PagePointer + "</A>&nbsp;";
        //                    }
        //                    PagePointer = PagePointer + 1;
        //                    LinkCount = LinkCount + 1;
        //                }
        //                if (PagePointer < PageCount) {
        //                    WorkingQS = GenericController.modifyQueryString(WorkingQS, RequestNamePageNumber, PageCount.ToString(), true);
        //                    result += "...<a href=\"" + "?" + WorkingQS + "\">" + PageCount + "</A>&nbsp;";
        //                }
        //                if (ReportPageNumber < PageCount) {
        //                    WorkingQS = GenericController.modifyQueryString(WorkingQS, RequestNamePageNumber, (ReportPageNumber + 1).ToString(), true);
        //                    result += "...<a href=\"" + "?" + WorkingQS + "\">next</A>&nbsp;";
        //                }
        //                result += "<br>&nbsp;";
        //            }
        //        }
        //        //
        //        result = ""
        //        + PreTableCopy + "<table border=0 cellpadding=0 cellspacing=0 width=\"100%\"><tr><td style=\"padding:10px;\">"
        //        + result + "</td></tr></table>"
        //        + PostTableCopy + "";
        //    } catch (Exception) {
        //        throw;
        //    }
        //    return result;
        //}
        //
        //====================================================================================================
        public static string getBody(CPBaseClass cp, string Caption, string ButtonListLeft, string ButtonListRight, bool AllowAdd, bool AllowDelete, string Description, string ContentSummary, int ContentPadding, string Content) {
            string result = "";
            string ButtonBar = null;
            string LeftButtons = "";
            string RightButtons = "";
            string CellContentSummary = "";
            //
            // Build ButtonBar
            //
            if (!string.IsNullOrEmpty(ButtonListLeft.Trim(' '))) {
                LeftButtons = getButtonsFromList(cp, ButtonListLeft, AllowDelete, AllowAdd, "Button");
            }
            if (!string.IsNullOrEmpty(ButtonListRight.Trim(' '))) {
                RightButtons = getButtonsFromList(cp, ButtonListRight, AllowDelete, AllowAdd, "Button");
            }
            ButtonBar = getButtonBar(cp, LeftButtons, RightButtons);
            if (!string.IsNullOrEmpty(ContentSummary)) {
                CellContentSummary = ""
                    + "\r<div class=\"ccPanelBackground\" style=\"padding:10px;\">"
                    + getPanel(ContentSummary, "ccPanel", "ccPanelShadow", "ccPanelHilite", "100%", 5)
                    + "\r</div>";
            }
            result += ""
                + ButtonBar
                + getTitleBar(cp, Caption, Description)
                + CellContentSummary
                + "<div style=\"padding:" + ContentPadding + "px;\">" + Content + "\r</div>"
                + ButtonBar;
            result = HtmlController.formMultipart(cp, result, cp.Doc.RefreshQueryString, "", "ccForm");
            return result;
        }

    }
}
