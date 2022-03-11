
using System;
using System.Collections.Generic;
using System.Net;
using System.Text;
using Contensive.BaseClasses;
using static Contensive.Addons.ContactManagerTools.Constants;

namespace Contensive.Addons.ContactManagerTools {
    public static class HtmlController {
        //
        internal const string ButtonDelete = "  Delete  ";
        internal const string ButtonFileChange = "   Upload   ";
        internal const string ButtonFileDelete = "    Delete    ";
        internal const string ButtonClose = "  Close   ";
        internal const string ButtonAdd = "   Add    ";
        //
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
        public static string getButtonsFromList(List<ButtonMetadata> ButtonList, bool AllowDelete, bool AllowAdd) {
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
        public static string getButtonsFromList( string ButtonList, bool AllowDelete, bool AllowAdd) {
            return getButtonsFromList(buttonStringToButtonList(ButtonList), AllowDelete, AllowAdd);
        }
        //
        // ====================================================================================================
        //
        public static List<ButtonMetadata> buttonStringToButtonList(string ButtonList) {
            var result = new List<ButtonMetadata>();
            if (!string.IsNullOrEmpty(ButtonList.Trim(' '))) {
                string[] Buttons = ButtonList.Split(',');
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
        public static string getButtonBar(string LeftButtons, string RightButtons) {
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
        /// <param name="cp"></param>
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
            string result = "<form action=\"?" + (string.IsNullOrEmpty(actionQueryString ) ? cp.Doc.RefreshQueryString : actionQueryString) + "\" ENCTYPE=\"MULTIPART/FORM-DATA\" method=\"post\" style=\"display: inline;\"";
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
        /// <param name="cp"></param>
        /// <param name="Title"></param>
        /// <param name="Description"></param>
        /// <returns></returns>
        public static string getTitleBar(string Title, string Description) {
            string result =  Title;
            if (!string.IsNullOrEmpty(Description)) {
                return Title + HtmlController.div(Description);
            }
            return HtmlController.div(result, "ccAdminTitleBar");
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
        public static string getPanel(string content, string stylePanel,  string width, int padding) {
            string ContentPanelWidth;
            string contentPanelWidthStyle;
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
        // ====================================================================================================
        /// <summary>
        /// create a table start tag
        /// </summary>
        /// <param name="cellpadding"></param>
        /// <param name="cellspacing"></param>
        /// <param name="border"></param>
        /// <param name="htmlClass"></param>
        /// <returns></returns>
        public static string tableStart(int cellpadding, int cellspacing, int border, string htmlClass = "") => "<table border=\"" + border + "\" cellpadding=\"" + cellpadding + "\" cellspacing=\"" + cellspacing + "\" class=\"" + htmlClass + "\" width=\"100%\">";
        //
        // ====================================================================================================
        /// <summary>
        /// return a row start (tr tag)
        /// </summary>
        /// <returns></returns>
        public static string tableRowStart() => "<tr>";
        //
        // ====================================================================================================
        /// <summary>
        /// create a td tag (without /td)
        /// </summary>
        /// <param name="Width"></param>
        /// <param name="ColSpan"></param>
        /// <param name="EvenRow"></param>
        /// <param name="Align"></param>
        /// <param name="BGColor"></param>
        /// <returns></returns>
        public static string tableCellStart(string Width = "", int ColSpan = 0, bool EvenRow = false, string Align = "", string BGColor = "") {
            string result = "";
            if (!string.IsNullOrEmpty(Width)) {
                result += " width=\"" + Width + "\"";
            }
            if (!string.IsNullOrEmpty(BGColor)) {
                result += " bgcolor=\"" + BGColor + "\"";
            } else if (EvenRow) {
                result += " class=\"ccPanelRowEven\"";
            } else {
                result += " class=\"ccPanelRowOdd\"";
            }
            if (ColSpan != 0) {
                result += " colspan=\"" + ColSpan + "\"";
            }
            if (!string.IsNullOrEmpty(Align)) {
                result += " align=\"" + Align + "\"";
            }
            return "<td" + result + ">";
        }
        //
        // ====================================================================================================
        /// <summary>
        /// create a table cell <td>content</td>
        /// </summary>
        /// <param name="Copy"></param>
        /// <param name="Width"></param>
        /// <param name="ColSpan"></param>
        /// <param name="EvenRow"></param>
        /// <param name="Align"></param>
        /// <param name="BGColor"></param>
        /// <returns></returns>
        public static string td(string Copy, string Width = "", int ColSpan = 0, bool EvenRow = false, string Align = "", string BGColor = "") {
            return tableCellStart(Width, ColSpan, EvenRow, Align, BGColor) + Copy + tableCellEnd;
        }
        //
        // ====================================================================================================
        /// <summary>
        /// create a <tr><td>content</td></tr>
        /// </summary>
        /// <param name="Cell"></param>
        /// <param name="ColSpan"></param>
        /// <param name="EvenRow"></param>
        /// <returns></returns>
        public static string tableRow(string Cell, int ColSpan = 0, bool EvenRow = false) {
            return tableRowStart() + td(Cell, "100%", ColSpan, EvenRow) + Constants.kmaEndTableRow;
        }
    }
    public class ButtonMetadata {
        public string name = "button";
        public string value = "";
        public string classList = "";
        public bool isDelete = false;
        public bool isClose = false;
        public bool isAdd = false;
    }
}
