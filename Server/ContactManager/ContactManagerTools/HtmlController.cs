
using System;
using System.Collections.Generic;
using System.Net;
using System.Text;
using Contensive.BaseClasses;

namespace Contensive.Addons.ContactManager {
    public class HtmlController {
        //
        internal const string ButtonDelete = "  Delete  ";
        internal const string ButtonFileChange = "   Upload   ";
        internal const string ButtonFileDelete = "    Delete    ";
        internal const string ButtonClose = "  Close   ";
        internal const string ButtonAdd = "   Add    ";
        //
        public class ButtonMetadata {
            public string name = "button";
            public string value = "";
            public string classList = "";
            public bool isDelete = false;
            public bool isClose = false;
            public bool isAdd = false;
        }
        //
        //====================================================================================================
        /// <summary>
        /// create an html block tag like div
        /// </summary>
        /// <param name="TagName"></param>
        /// <param name="InnerHtml"></param>
        /// <param name="HtmlName"></param>
        /// <param name="HtmlClass"></param>
        /// <param name="HtmlId"></param>
        /// <returns></returns>
        public static string genericBlockTag(string TagName, string InnerHtml, string HtmlClass, string HtmlId, string HtmlName) {
            var result = new StringBuilder("<");
            result.Append(TagName.Trim());
            result.Append(!string.IsNullOrEmpty(HtmlName) ? " name=\"" + HtmlName + "\"" : "");
            result.Append(!string.IsNullOrEmpty(HtmlClass) ? " class=\"" + HtmlClass + "\"" : "");
            result.Append(!string.IsNullOrEmpty(HtmlId) ? " id=\"" + HtmlId + "\"" : "");
            result.Append(">");
            result.Append(InnerHtml);
            result.Append("</");
            result.Append(TagName.Trim());
            result.Append(">");
            return result.ToString();
        }
        //
        public static string genericBlockTag(string TagName, string InnerHtml) => genericBlockTag(TagName, InnerHtml, "", "", "");
        //
        public static string genericBlockTag(string TagName, string InnerHtml, string HtmlClass) => genericBlockTag(TagName, InnerHtml, HtmlClass, "", "");
        //
        public static string genericBlockTag(string TagName, string InnerHtml, string HtmlClass, string HtmlId) => genericBlockTag(TagName, InnerHtml, HtmlClass, HtmlId, "");
        //
        // ====================================================================================================
        //
        public static string a(string innerHtml, string href) => "<a href=\"" + WebUtility.HtmlEncode(href) + "\">" + innerHtml + "</a>";
        public static string a(string innerHtml, string href, string htmlClass) => "<a href=\"" + WebUtility.HtmlEncode(href) + "\" class=\"" + htmlClass + "\">" + innerHtml + "</a>";
        //
        //====================================================================================================
        //
        public static string li(string innerHtml) => genericBlockTag("li", innerHtml);
        //
        public static string li(string innerHtml, string htmlClass) => genericBlockTag("li", innerHtml, htmlClass, "");
        //
        public static string li(string innerHtml, string htmlClass, string htmlId) => genericBlockTag("li", innerHtml, htmlClass, htmlId);
        //
        //====================================================================================================
        //
        public static string div(string innerHtml) => genericBlockTag("div", innerHtml);
        //
        public static string div(string innerHtml, string htmlClass) => genericBlockTag("div", innerHtml, htmlClass);
        //
        public static string div(string innerHtml, string htmlClass, string htmlId) => genericBlockTag("div", innerHtml, htmlClass, htmlId);
        //
        //====================================================================================================
        //
        public static string ul(string innerHtml) => genericBlockTag("ul", innerHtml, "", "", "");
        //
        public static string ul(string innerHtml, string htmlClass) => genericBlockTag("ul", innerHtml, htmlClass);
        //
        public static string ul(string innerHtml, string htmlClass, string htmlId) => genericBlockTag("ul", innerHtml, htmlClass, htmlId);
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
        //
        //====================================================================================================
        public static string getButtonsFromList(CPBaseClass cp, List<ButtonMetadata> ButtonList, bool AllowDelete, bool AllowAdd) {
            string s = "";
            foreach (ButtonMetadata button in ButtonList) {

                if (button.isDelete) {
                    s += getButtonDanger(button.value, "if(!DeleteCheck()) return false;", !AllowDelete);
                } else if (button.isAdd) {
                    s += getButtonPrimary(button.value, "return processSubmit(this);", !AllowAdd);
                } else if (button.isClose) {
                    s += getButtonPrimary(button.value, "window.close();");
                } else {
                    s += getButtonPrimary(button.value);
                }

            }
            return s;
        }
        //
        //====================================================================================================
        public static string getButtonsFromList(CPBaseClass cp, string ButtonList, bool AllowDelete, bool AllowAdd, string ButtonName) {
            return getButtonsFromList(cp, buttonStringToButtonList(ButtonList), AllowDelete, AllowAdd);
        }
        //
        // ====================================================================================================
        //
        public static List<ButtonMetadata> buttonStringToButtonList(string ButtonList) {
            var result = new List<ButtonMetadata>();
            string[] Buttons = null;
            if (!string.IsNullOrEmpty(ButtonList.Trim(' '))) {
                Buttons = ButtonList.Split(',');
                foreach (string buttonValue in Buttons) {
                    string buttonValueTrim = buttonValue.Trim();
                    result.Add(new ButtonMetadata() {
                        name = "button",
                        value = buttonValue,
                        isAdd = buttonValueTrim.Equals(ButtonAdd),
                        isClose = buttonValueTrim.Equals(ButtonClose),
                        isDelete = buttonValueTrim.Equals(ButtonDelete)
                    });
                }
            }
            return result;
        }
        //
        //====================================================================================================
        /// <summary>
        /// Return a bootstrap button bar
        /// </summary>
        /// <param name="LeftButtons"></param>
        /// <param name="RightButtons"></param>
        /// <returns></returns>
        public static string getButtonBar(CPBaseClass cp, string LeftButtons, string RightButtons) {
            if (string.IsNullOrWhiteSpace(LeftButtons + RightButtons)) {
                return "";
            } else if (string.IsNullOrWhiteSpace(RightButtons)) {
                return "<div class=\"border bg-white p-2\">" + LeftButtons + "</div>";
            } else {
                return "<div class=\"border bg-white p-2\">" + LeftButtons + "<div class=\"float-right\">" + RightButtons + "</div></div>";
            }
        }
        //
        //====================================================================================================
        /// <summary>
        /// returns a multipart form, required for file uploads
        /// </summary>
        /// <param name="core"></param>
        /// <param name="innerHtml"></param>
        /// <param name="actionQueryString"></param>
        /// <param name="htmlName"></param>
        /// <param name="htmlClass"></param>
        /// <param name="htmlId"></param>
        /// <returns></returns>
        public static string formMultipart(CPBaseClass cp, string innerHtml, string actionQueryString = "", string htmlName = "", string htmlClass = "", string htmlId = "") {
            return formMultipart_start(cp, actionQueryString, htmlName, htmlClass, htmlId) + innerHtml + "</form>";
        }
        //
        //====================================================================================================
        /// <summary>
        /// Starts an HTML form for uploads, Should be closed with main_GetUploadFormEnd
        /// </summary>
        /// <param name="actionQueryString"></param>
        /// <returns></returns>
        public static string formMultipart_start(CPBaseClass cp, string actionQueryString = "", string htmlName = "", string htmlClass = "", string htmlId = "") {
            string result = "<form action=\"?" + ((actionQueryString == "") ? cp.Doc.RefreshQueryString : actionQueryString) + "\" ENCTYPE=\"MULTIPART/FORM-DATA\" method=\"post\" style=\"display: inline;\"";
            if (!string.IsNullOrWhiteSpace(htmlName)) result += " name=\"" + htmlName + "\"";
            if (!string.IsNullOrWhiteSpace(htmlClass)) result += " class=\"" + htmlClass + "\"";
            if (!string.IsNullOrWhiteSpace(htmlId)) result += " id=\"" + htmlId + "\"";
            return result + ">";
        }
        //
        //====================================================================================================
        /// <summary>
        /// Title Bar
        /// </summary>
        /// <param name="core"></param>
        /// <param name="Title"></param>
        /// <param name="Description"></param>
        /// <returns></returns>
        public static string getTitleBar(CPBaseClass cp, string Title, string Description) {
            string result = "";
            result = Title;
            if (!string.IsNullOrEmpty(Description)) {
                result += HtmlController.div(Description);
            }
            result = HtmlController.div(result, "ccAdminTitleBar");
            return result;
        }
        //
        // ====================================================================================================
        //
        public static string getButtonPrimary(string buttonValue, string onclick = "", bool disabled = false, string htmlId = "", string htmlName = "button") {
            return HtmlController.getHtmlInputSubmit(buttonValue, htmlName, htmlId, onclick, disabled, "btn btn-primary mr-1 btn-sm");
        }
        //
        // ====================================================================================================
        //
        public static string getButtonDanger(string buttonValue, string onclick = "", bool disabled = false, string htmlId = "") {
            return HtmlController.getHtmlInputSubmit(buttonValue, "button", htmlId, onclick, disabled, "btn btn-danger mr-1 btn-sm");
        }
        //
        //====================================================================================================
        //
        public static string getHtmlInputSubmit(string htmlValue, string htmlName = "button", string htmlId = "", string onClick = "", bool disabled = false, string htmlClass = "") {
            string attrList = "<input type=submit name=\"" + HtmlController.encodeHtml(htmlName) + "\"";
            attrList += (string.IsNullOrEmpty(htmlName)) ? "" : " name=\"" + htmlName + "\"";
            attrList += (string.IsNullOrEmpty(htmlValue)) ? "" : " value=\"" + htmlValue + "\"";
            attrList += (string.IsNullOrEmpty(htmlId)) ? "" : " id=\"" + htmlId + "\"";
            attrList += (string.IsNullOrEmpty(htmlClass)) ? "" : " class=\"" + htmlClass + "\"";
            attrList += (string.IsNullOrEmpty(onClick)) ? "" : " onclick=\"" + onClick + "\"";
            attrList += (!disabled) ? "" : " disabled";
            return attrList + ">";
        }
        //
        // ====================================================================================================
        /// <summary>
        /// convert html entities in a string
        /// </summary>
        /// <param name="Source"></param>
        /// <returns></returns>
        public static string encodeHtml(string Source) {
            return WebUtility.HtmlEncode(Source);
        }
        //
        //====================================================================================================
        //
        public static string getPanel(string content, string stylePanel, string styleHilite, string styleShadow, string width, int padding, int heightMin) {
            string ContentPanelWidth = "";
            string contentPanelWidthStyle = "";
            if (width.IsNumeric()) {
                ContentPanelWidth = (int.Parse(width) - 2).ToString();
                contentPanelWidthStyle = ContentPanelWidth + "px";
            } else {
                ContentPanelWidth = "100%";
                contentPanelWidthStyle = ContentPanelWidth;
            }
            //
            string s0 = ""
                + "<td style=\"padding:" + padding + "px;vertical-align:top\" class=\"" + stylePanel + "\">"
                + content
                + "</td>"
                + "";
            //
            string s1 = ""
                + "<tr>"
                + s0
                + "</tr>"
                + "";
            string s2 = ""
                + "<table style=\"width:" + contentPanelWidthStyle + ";border:0px;\" class=\"" + stylePanel + "\" cellspacing=\"0\">"
                + s1
                + "</table>"
                + "";
            string s3 = ""
                + "<td colspan=\"3\" width=\"" + ContentPanelWidth + "\" valign=\"top\" align=\"left\" class=\"" + stylePanel + "\">"
                + s2
                + "</td>"
                + "";
            string s4 = ""
                + "<tr>"
                + s3
                + "</tr>"
                + "";
            string result = ""
                + "<table border=\"0\" cellpadding=\"0\" cellspacing=\"0\" width=\"" + width + "\" class=\"" + stylePanel + "\">"
                + s4
                + "</table>"
                + "";
            return result;
        }
        //
        public static string getPanel(string content) => getPanel(content, "ccPanel", "ccPanelHilite", "ccPanelShadow", "100%", 5, 1);
        //
        public static string getPanel(string content, string stylePanel) => getPanel(content, stylePanel, "ccPanelHilite", "ccPanelShadow", "100%", 5, 1);
        //
        public static string getPanel(string content, string stylePanel, string styleHilite) => getPanel(content, stylePanel, styleHilite, "ccPanelShadow", "100%", 5, 1);
        //
        public static string getPanel(string content, string stylePanel, string styleHilite, string styleShadow) => getPanel(content, stylePanel, styleHilite, styleShadow, "100%", 5, 1);
        //
        public static string getPanel(string content, string stylePanel, string styleHilite, string styleShadow, string width) => getPanel(content, stylePanel, styleHilite, styleShadow, width, 5, 1);
        //
        public static string getPanel(string content, string stylePanel, string styleHilite, string styleShadow, string width, int padding) => getPanel(content, stylePanel, styleHilite, styleShadow, width, padding, 1);
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
            } catch (Exception ex) {
                
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
                RQS = core.doc.refreshQueryString;
                //
                SortColPtr = getReportSortColumnPtr(core, DefaultSortColumnPtr);
                SortColType = getReportSortType(core);
                //
                // ----- Start the table
                //
                Content.Add(HtmlController.tableStart(3, 1, 0));
                //
                // ----- Header
                //
                Content.Add("\r\n<tr>");
                Content.Add(getReport_CellHeader(core, 0, "&nbsp", "50px", "Right", "ccAdminListCaption", RQS, SortingStateEnum.NotSortable));
                for (ColumnPtr = 0; ColumnPtr < ColumnCount; ColumnPtr++) {
                    ColumnWidth = ColWidth[ColumnPtr];
                    if (!ColSortable[ColumnPtr]) {
                        //
                        // not sortable column
                        //
                        Content.Add(getReport_CellHeader(core, ColumnPtr, ColCaption[ColumnPtr], ColumnWidth, ColAlign[ColumnPtr], "ccAdminListCaption", RQS, SortingStateEnum.NotSortable));
                    } else if (ColumnPtr == SortColPtr) {
                        //
                        // This is the current sort column
                        //
                        Content.Add(getReport_CellHeader(core, ColumnPtr, ColCaption[ColumnPtr], ColumnWidth, ColAlign[ColumnPtr], "ccAdminListCaption", RQS, (SortingStateEnum)SortColType));
                    } else {
                        //
                        // Column is sortable, but not selected
                        //
                        Content.Add(getReport_CellHeader(core, ColumnPtr, ColCaption[ColumnPtr], ColumnWidth, ColAlign[ColumnPtr], "ccAdminListCaption", RQS, SortingStateEnum.SortableNotSet));
                    }

                    //If ColumnPtr = SortColPtr Then
                    //    '
                    //    ' This column is currently the active sort
                    //    '
                    //    Call Content.Add(GetReport_CellHeader(ColumnPtr, ColCaption(ColumnPtr), ColumnWidth, ColAlign(ColumnPtr), "ccAdminListCaption", RQS, SortColType))
                    //Else
                    //    Call Content.Add(GetReport_CellHeader(ColumnPtr, ColCaption(ColumnPtr), ColumnWidth, ColAlign(ColumnPtr), "ccAdminListCaption", RQS, SortingStateEnum.SortableNotSet))
                    //End If
                }
                Content.Add("\r\n</tr>");
                //
                // ----- Data
                //
                if (RowCount == 0) {
                    Content.Add("\r\n<tr>");
                    Content.Add(getReport_Cell(core, (RowBAse + RowPointer).ToString(), "right", 1, RowPointer));
                    Content.Add(getReport_Cell(core, "-- End --", "left", ColumnCount, 0));
                    Content.Add("\r\n</tr>");
                } else {
                    RowBAse = (ReportPageSize * (ReportPageNumber - 1)) + 1;
                    for (RowPointer = 0; RowPointer < RowCount; RowPointer++) {
                        Content.Add("\r\n<tr>");
                        Content.Add(getReport_Cell(core, (RowBAse + RowPointer).ToString(), "right", 1, RowPointer));
                        for (ColumnPtr = 0; ColumnPtr < ColumnCount; ColumnPtr++) {
                            Content.Add(getReport_Cell(core, Cells[RowPointer, ColumnPtr], ColAlign[ColumnPtr], 1, RowPointer));
                        }
                        Content.Add("\r\n</tr>");
                    }
                }
                //
                // ----- End Table
                //
                Content.Add(kmaEndTable);
                result += Content.Text;
                //
                // ----- Post Table copy
                //
                if ((DataRowCount / (double)ReportPageSize) != Math.Floor((DataRowCount / (double)ReportPageSize))) {
                    PageCount = encodeInteger((DataRowCount / (double)ReportPageSize) + 0.5);
                } else {
                    PageCount = encodeInteger(DataRowCount / (double)ReportPageSize);
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
                            WorkingQS = core.doc.refreshQueryString;
                            WorkingQS = GenericController.modifyQueryString(WorkingQS, "GotoPage", "1", true);
                            result += "<a href=\"" + core.webServer.requestPage + "?" + WorkingQS + "\">1</A>...&nbsp;";
                        }
                        WorkingQS = core.doc.refreshQueryString;
                        WorkingQS = GenericController.modifyQueryString(WorkingQS, RequestNamePageSize, ReportPageSize.ToString(), true);
                        while ((PagePointer <= PageCount) && (LinkCount < 20)) {
                            if (PagePointer == ReportPageNumber) {
                                result += PagePointer + "&nbsp;";
                            } else {
                                WorkingQS = GenericController.modifyQueryString(WorkingQS, RequestNamePageNumber, PagePointer.ToString(), true);
                                result += "<a href=\"" + core.webServer.requestPage + "?" + WorkingQS + "\">" + PagePointer + "</A>&nbsp;";
                            }
                            PagePointer = PagePointer + 1;
                            LinkCount = LinkCount + 1;
                        }
                        if (PagePointer < PageCount) {
                            WorkingQS = GenericController.modifyQueryString(WorkingQS, RequestNamePageNumber, PageCount.ToString(), true);
                            result += "...<a href=\"" + core.webServer.requestPage + "?" + WorkingQS + "\">" + PageCount + "</A>&nbsp;";
                        }
                        if (ReportPageNumber < PageCount) {
                            WorkingQS = GenericController.modifyQueryString(WorkingQS, RequestNamePageNumber, (ReportPageNumber + 1).ToString(), true);
                            result += "...<a href=\"" + core.webServer.requestPage + "?" + WorkingQS + "\">next</A>&nbsp;";
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
                LogController.handleError(core, ex);
            }
            return result;
        }
    }
}
