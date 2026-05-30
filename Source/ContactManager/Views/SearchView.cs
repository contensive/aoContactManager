using Contensive.Addons.ContactManager.Controllers;
using Contensive.BaseClasses;
using System;
using System.Collections.Generic;
using static Contensive.Addons.ContactManager.Controllers.GenericController;

namespace Contensive.Addons.ContactManager.Views {
    /// <summary>
    /// 
    /// </summary>
    public sealed class SearchView {
        // 
        // =================================================================================
        /// <summary>
        /// 
        /// </summary>
        /// <param name="cp"></param>
        /// <param name="ae"></param>
        /// <param name="IsAdminPath"></param>
        /// <returns></returns>
        public static string getResponse(CPBaseClass cp, Controllers.ApplicationController ae, bool IsAdminPath) {
            //
            // 
            //
            try {
                const string guidContactManager = "{57602963-CA46-41EB-8053-8ACB8C47C2F2}";
                //
                // -- detect which tab is selected
                int SubTab = cp.Doc.GetInteger("SubTab");
                if (SubTab == 0) {
                    SubTab = ae.userProperties.selectSubTab;
                    if (SubTab == 0) {
                        SubTab = 1;
                        ae.userProperties.selectSubTab = SubTab;
                    }
                } else {
                    ae.userProperties.selectSubTab = SubTab;
                }
                cp.Doc.AddRefreshQueryString("SubTab", SubTab.ToString());
                //
                //
                // -- create inner layout with contentWithTabs that includes the tabs plus the content of the selected tab
                var layout = cp.AdminUI.CreateLayoutBuilderTabbedBody();



                layout.addTab();
                layout.tabCaption = "Record Fields";
                layout.tabLink = $"?addonGuid={guidContactManager}&formid=3&SubTab=1";
                layout.tabStyleClass = "tab-style-class";
                if (SubTab == 1) {
                    //
                    layout.setActiveTab(layout.tabCaption);
                    var tabLayout = cp.AdminUI.CreateLayoutBuilder();
                    tabLayout.body = getResponse_TabPeople(cp, ae);
                    tabLayout.includeBodyPadding = false;
                    tabLayout.includeForm = false;
                    layout.body = tabLayout.getHtml();
                }
                //
                layout.addTab();
                layout.tabCaption = "Groups";
                layout.tabLink = $"?addonGuid={guidContactManager}&formid=3&SubTab=2";
                layout.tabStyleClass = "tab-style-class";
                if (SubTab == 2) {
                    //
                    layout.setActiveTab(layout.tabCaption);
                    var tabLayout = cp.AdminUI.CreateLayoutBuilder();
                    tabLayout.body = getResponse_TabGroup(cp, ae);
                    tabLayout.includeBodyPadding = false;
                    tabLayout.includeForm = false;
                    layout.body = tabLayout.getHtml();
                }

                //
                layout.addFormHidden("selectionForm", "1");
                layout.addFormHidden(Constants.RequestNameFormID, Convert.ToInt32((int)Constants.FormIdEnum.FormSearch).ToString());

                //string Content = "<div class=\"mt-4\">" + cp.Html.Hidden("SelectionForm", "1") + layout.getHtml() + cp.Html.Hidden(Constants.RequestNameFormID, Convert.ToInt32((int)Constants.FormIdEnum.FormSearch).ToString()) + "</div>";
                // 
                // Assemble page
                //var layout = cp.AdminUI.CreateLayoutBuilder();
                //layout.body = Content;
                layout.description = "Select groups and record fields for search. The search will return only people who satisfy all the selections.";
                layout.failMessage = "";
                layout.includeBodyColor = true;
                layout.includeBodyPadding = true;
                layout.infoMessage = "";
                layout.isOuterContainer = true;
                layout.portalSubNavTitle = "";
                layout.successMessage = "";
                layout.title = "Contact Manager - Create Search";
                layout.warningMessage = "";

                if (IsAdminPath) {
                    layout.addFormButton(Constants.ButtonCancelAll);
                    layout.addFormButton(Constants.ButtonSearch);
                } else {
                    layout.addFormButton(Constants.ButtonSearch);
                }
                string result = layout.getHtml();
                return result;
            } catch (Exception ex) {
                cp.Site.ErrorReport(ex);
                throw;
            }
        }
        // 
        // =================================================================================
        /// <summary>
        /// get the html for the group selection tab
        /// </summary>
        /// <param name="cp"></param>
        /// <param name="ae"></param>
        /// <returns></returns>
        public static string getResponse_TabGroup(CPBaseClass cp, Controllers.ApplicationController ae) {
            string result = "";
            try {
                // 
                if (true) {
                    string RQS = cp.Doc.RefreshQueryString;
                    string ContactGroupCriteria = ae.userProperties.contactGroupCriteria;
                    // 
                    // result = result _
                    // & "<div>Select groups to narrow your results. If any groups are selected, your search will be limited to people in any of the selected groups.</div>" _
                    // & "<div>&nbsp;</div>"
                    // 
                    // Add headers to stream
                    // 
                    result += "<table border=0 width=100% cellspacing=0 cellpadding=4>";
                    result += "<tr>";
                    result += "<td width=30><img src=/cclib/images/spacer.gif width=20 height=1></TD>";
                    result += "<td width=30><img src=/cclib/images/spacer.gif width=20 height=1></TD>";
                    result += "<td width=99%><img src=/cclib/images/spacer.gif width=1 height=1></TD>";
                    result += "</tr>";
                    // 
                    result += "<tr>";
                    result += "<td width=30 align=center class=ccAdminListCaption>Select</TD>";
                    result += "<td width=30 align=center class=ccAdminListCaption>Count</TD>";
                    result += "<td width=99% align=left class=ccAdminListCaption>Group Name</TD>";
                    result += "</tr>";
                    // 
                    string SQL = "SELECT ccGroups.ID as GroupID, ccGroups.Name as GroupName, ccGroups.Caption as GroupCaption, Count(ccMembers.ID) AS CountOfID FROM (ccGroups LEFT JOIN ccMemberRules ON ccGroups.ID = ccMemberRules.GroupID) LEFT JOIN ccMembers ON ccMemberRules.MemberID = ccMembers.ID Where (((ccMemberRules.DateExpires) Is Null Or (ccMemberRules.DateExpires) > " + cp.Db.EncodeSQLDate(DateTime.Now) + ")) GROUP BY ccGroups.ID, ccGroups.Name, ccGroups.Caption ORDER BY ccGroups.Caption;";



                    // 
                    using (var CS = cp.CSNew()) {
                        CS.OpenSQL(SQL);
                        int GroupPointer = 0;
                        while (CS.OK()) {
                            int GroupID = CS.GetInteger("GroupID");
                            string GroupName = CS.GetText("GroupCaption");
                            if (string.IsNullOrEmpty(GroupName)) {
                                GroupName = CS.GetText("GroupName");
                                if (string.IsNullOrEmpty(GroupName)) {
                                    GroupName = "Group " + GroupID;
                                }
                            }
                            string GroupLabel = "Group" + GroupPointer;
                            bool GroupChecked = ContactGroupCriteria.IndexOf($",{GroupID},", StringComparison.OrdinalIgnoreCase) >= 0;

                            string Style;
                            if (GroupPointer % 2 == 0) {
                                Style = "ccAdminListRowEven";
                            } else {
                                Style = "ccAdminListRowOdd";
                            }
                            result += "<TR>";
                            result += "<td class=\"p-1\" width=30 align=center class=\"" + Style + "\">" + cp.Html.CheckBox(GroupLabel, GroupChecked) + cp.Html.Hidden(GroupLabel + ".id", GroupID.ToString()) + "</TD>";
                            result += "<td class=\"p-1\" width=30 align=right class=\"" + Style + "\">" + CS.GetInteger("CountOfID") + "</TD>";
                            result += "<td class=\"p-1\" width=99% align=left class=\"" + Style + "\">" + GroupName + "</TD>";
                            result += "</TR>";
                            GroupPointer += 1;
                            CS.GoNext();
                        }
                        CS.Close();
                        result += "</Table>";
                        // 
                        result += cp.Html.Hidden("GroupCount", GroupPointer.ToString());
                    }
                    result += cp.Html.Hidden("SelectionGroupSubTab", "1");
                    // 
                }
            } catch (Exception ex) {
                cp.Site.ErrorReport(ex);
                throw;
            }
            return result;
        }
        // 
        // =================================================================================
        // 
        public static string getResponse_TabPeople(CPBaseClass cp, Controllers.ApplicationController ae) {
            string result = "";
            try {
                // 
                // prepare visit property to prepopulate form
                string ContactSearchCriteria = ae.userProperties.contactSearchCriteria;
                string[] CriteriaValues = Array.Empty<string>();
                int CriteriaCount = 0;
                if (!string.IsNullOrEmpty(ContactSearchCriteria)) {
                    CriteriaValues = ContactSearchCriteria.Split(new[] { Environment.NewLine }, StringSplitOptions.None);
                    CriteriaCount = CriteriaValues.Length;
                }
                // 
                // Setup fields and capture request changes
                // 
                string Criteria = "(active<>0)and(ContentID=" + cp.Content.GetID("people") + ")and(authorable<>0)";
                // 
                using (var CS = cp.CSNew()) {
                    CS.Open("Content Fields", Criteria, "EditTab,EditSortPriority");
                    var fieldList = new List<FieldMeta>();
                    int FieldPtr = 0;
                    while (CS.OK()) {
                        int fieldType = CS.GetInteger("Type");
                        var @field = new FieldMeta() {
                            currentValue = "",
                            fieldCaption = CS.GetText("Caption"),
                            fieldEditTab = CS.GetText("editTab"),
                            fieldId = CS.GetInteger("ID"),
                            fieldLookupContentName = fieldType != 7 ? "" : cp.Content.GetRecordName("content", CS.GetInteger("LookupContentID")),
                            fieldLookupList = fieldType != 7 ? "" : CS.GetText("LookupList"),
                            fieldName = CS.GetText("name"),
                            fieldOperator = 0,
                            fieldType = fieldType
                        };
                        fieldList.Add(@field);
                        // 
                        // set prepoplate value from visit property
                        // 
                        if (CriteriaCount > 0) {
                            int CriteriaPointer;
                            var loopTo = CriteriaCount - 1;
                            for (CriteriaPointer = 0; CriteriaPointer <= loopTo; CriteriaPointer++) {
                                string[] NameValues;
                                if (CriteriaValues[CriteriaPointer].StartsWith($"{@field.fieldName}=", StringComparison.OrdinalIgnoreCase)) {
                                    NameValues = CriteriaValues[CriteriaPointer].Split('=');
                                    @field.currentValue = NameValues[1];
                                    @field.fieldOperator = 1;
                                } else if (CriteriaValues[CriteriaPointer].StartsWith($"{@field.fieldName}>", StringComparison.OrdinalIgnoreCase)) {
                                    NameValues = CriteriaValues[CriteriaPointer].Split('>');
                                    @field.currentValue = NameValues[1];
                                    @field.fieldOperator = 2;
                                } else if (CriteriaValues[CriteriaPointer].StartsWith($"{@field.fieldName}<", StringComparison.OrdinalIgnoreCase)) {
                                    NameValues = CriteriaValues[CriteriaPointer].Split('<');
                                    @field.currentValue = NameValues[1];
                                    @field.fieldOperator = 3;
                                }
                            }
                        }
                        FieldPtr += 1;
                        CS.GoNext();
                    }
                    CS.Close();
                    int FieldCount = FieldPtr;
                    // 
                    // header
                    // 
                    // result = result _
                    // & "<div>Enter criteria for each field to identify and select your results. The results of a search will have to have all of the criteria you enter.</div>" _
                    // & "<div>&nbsp;</div>"
                    // 
                    // Add headers to stream
                    // 
                    result += "<table border=0 width=100% cellspacing=0 cellpadding=4>";
                    result += "<tr>";
                    bool RowEven;
                    result += AdminUIController.kmaStartTableCell("120") + "<img src=/cclib/images/spacer.gif width=120 height=1></TD>";
                    result += AdminUIController.kmaStartTableCell("99%") + "<img src=/cclib/images/spacer.gif width=1 height=1></TD>";
                    result += "</tr>";
                    // 
                    int RowPointer = 0;
                    string lastEditTab = "-1";
                    string currentEditTab = "";
                    string groupTab = "";
                    foreach (var @field in fieldList) {
                        currentEditTab = @field.fieldEditTab;
                        if ((currentEditTab ?? "") != (lastEditTab ?? "")) {
                            string tabCaption;
                            if (string.IsNullOrEmpty(currentEditTab)) {
                                tabCaption = "Details";
                            } else {
                                tabCaption = currentEditTab;
                            }
                            groupTab = "<tr><td class=\"p-1 cmRowTab\" colspan=\"2\">" + tabCaption + "</TD></TR>";


                            RowPointer += 1;
                            lastEditTab = currentEditTab;
                        }
                        RowEven = RowPointer % 2 == 0;
                        switch (@field.fieldType) {
                            case Constants.FieldTypeDate: {
                                    // 
                                    // Date
                                    // 
                                    result += groupTab + "<TR><td class=\"p-1 cmRowCaption\">" + @field.fieldCaption + "</TD><td class=\"cmRowField\"><table border=0 width=100% cellspacing=0 cellpadding=0><TR><td class=\"p-1\" width=10 align=right>" + getFormInputRadioBox(cp, @field.fieldName + "_D", "0", @field.fieldOperator.ToString(), "") + "</TD><td class=\"p-1\" align=left width=100>ignore</TD><td class=\"p-1\" width=10>&nbsp;&nbsp;</TD><td class=\"p-1\" width=10 align=right>" + getFormInputRadioBox(cp, @field.fieldName + "_D", "1", @field.fieldOperator.ToString(), "") + "</TD><td class=\"p-1\" align=left width=100>=</TD><td class=\"p-1\" width=10>&nbsp;&nbsp;</TD><td class=\"p-1\" width=10 align=right>" + getFormInputRadioBox(cp, @field.fieldName + "_D", "2", @field.fieldOperator.ToString(), "") + "</TD><td class=\"p-1\" align=left width=100>&gt;</TD><td class=\"p-1\" width=10>&nbsp;&nbsp;</TD><td class=\"p-1\" width=10 align=right>" + getFormInputRadioBox(cp, @field.fieldName + "_D", "3", @field.fieldOperator.ToString(), "") + "</TD><td class=\"p-1\" align=left width=100>&lt;</TD><td class=\"p-1\" width=10>&nbsp;&nbsp;</TD><td class=\"p-1\" align=left width=99%>" + getFormInputDate(cp, @field.fieldName, encodeDate(@field.currentValue), "") + "</TD></TR></Table></TD></TR>";











                                    groupTab = "";
                                    RowPointer += 1;
                                    break;
                                }
                            case Constants.FieldTypeCurrency:
                            case Constants.FieldTypeFloat:
                            case Constants.FieldTypeInteger: {
                                    // 
                                    // Numeric
                                    // 
                                    result += groupTab + "<TR><td class=\"p-1 cmRowCaption\">" + @field.fieldCaption + "</TD><td class=\"p-1 cmRowField\"><table border=0 width=100% cellspacing=0 cellpadding=0><TR><td class=\"p-1\" width=10 align=right>" + getFormInputRadioBox(cp, @field.fieldName + "_N", "0", @field.fieldOperator.ToString(), "") + "</TD><td class=\"p-1\" align=left width=100>ignore</TD><td class=\"p-1\" width=10>&nbsp;&nbsp;</TD><td class=\"p-1\" width=10 align=right>" + getFormInputRadioBox(cp, @field.fieldName + "_N", "1", @field.fieldOperator.ToString(), "") + "</TD><td class=\"p-1\" align=left width=100>=</TD><td class=\"p-1\" width=10>&nbsp;&nbsp;</TD><td class=\"p-1\" width=10 align=right>" + getFormInputRadioBox(cp, @field.fieldName + "_N", "2", @field.fieldOperator.ToString(), "") + "</TD><td class=\"p-1\" align=left width=100>&gt;</TD><td class=\"p-1\" width=10>&nbsp;&nbsp;</TD><td class=\"p-1\" width=10 align=right>" + getFormInputRadioBox(cp, @field.fieldName + "_N", "3", @field.fieldOperator.ToString(), "") + "</TD><td class=\"p-1\" align=left width=100>&lt;</TD><td class=\"p-1\" width=10>&nbsp;&nbsp;</TD><td class=\"p-1\" align=left width=99%>" + getFormInputText(cp, @field.fieldName, @field.currentValue, "", "") + "</TD></TR></Table></TD></TR>";











                                    groupTab = "";
                                    RowPointer += 1;
                                    break;
                                }
                            case Constants.FieldTypeBoolean: {
                                    // 
                                    // Boolean
                                    // 
                                    string currentValue = string.IsNullOrWhiteSpace(@field.currentValue) ? "0" : @field.currentValue;
                                    result += groupTab + "<TR><td class=\"p-1 cmRowCaption\">" + @field.fieldCaption + "</TD><td class=\"p-1 cmRowField\"><table border=0 width=100% cellspacing=0 cellpadding=0><TR><td class=\"p-1\" width=10 align=right>" + getFormInputRadioBox(cp, @field.fieldName, "0", currentValue, "") + "</TD><td class=\"p-1\" align=left width=100>  ignore</TD><td class=\"p-1\" width=10>&nbsp;&nbsp;</TD><td class=\"p-1\" width=10 align=right>" + getFormInputRadioBox(cp, @field.fieldName, "1", currentValue, "") + "</TD><td class=\"p-1\" align=left width=100>true</TD><td class=\"p-1\" width=10>&nbsp;&nbsp;</TD><td class=\"p-1\" width=10 align=right>" + getFormInputRadioBox(cp, @field.fieldName, "2", currentValue, "") + "</TD><td class=\"p-1\" align=left width=100>false</TD><td class=\"p-1\" width=99%>&nbsp;</td></TR></Table></TD></TR>";










                                    RowPointer += 1;
                                    groupTab = "";
                                    break;
                                }
                            case Constants.FieldTypeText:
                            case Constants.FieldTypeLongText: {
                                    // 
                                    // Text
                                    // 
                                    result += groupTab + "<TR><td class=\"p-1 cmRowCaption\">" + @field.fieldCaption + "</TD><td class=\"p-1 cmRowField\" valign=absmiddle><table border=0 width=100% cellspacing=0 cellpadding=0><TR><td class=\"p-1\" width=10 align=right>" + getFormInputRadioBox(cp, @field.fieldName + "_T", "", @field.currentValue, "") + "</TD><td class=\"p-1\" align=left width=100>  ignore</TD><td class=\"p-1\" width=10>&nbsp;&nbsp;</TD><td class=\"p-1\" width=10 align=right>" + getFormInputRadioBox(cp, @field.fieldName + "_T", "1", @field.currentValue, "") + "</TD><td class=\"p-1\" align=left width=100>empty</TD><td class=\"p-1\" width=10>&nbsp;&nbsp;</TD><td class=\"p-1\" width=10 align=right>" + getFormInputRadioBox(cp, @field.fieldName + "_T", "2", @field.currentValue, "") + "</TD><td class=\"p-1\" align=left width=100>not&nbsp;empty</TD><td class=\"p-1\" width=10>&nbsp;&nbsp;</TD><td class=\"p-1\" width=10 align=right>" + getFormInputRadioBox(cp, @field.fieldName + "_T", "3", @field.currentValue, @field.fieldName + "_r") + "</TD><td class=\"p-1\" align=center width=100>&nbsp;includes&nbsp;</TD><td class=\"p-1\" align=left width=99%>" + getFormInputText(cp, @field.fieldName, @field.currentValue, "", "cmTextInclude") + "</TD></TR></Table></TD></TR>";










                                    groupTab = "";
                                    RowPointer += 1;
                                    break;
                                }
                            case Constants.FieldTypeLookup: {
                                    // 
                                    // Lookup
                                    // 
                                    result += groupTab + "<TR><td class=\"p-1 cmRowCaption\">" + @field.fieldCaption + "</TD>";

                                    if (!string.IsNullOrEmpty(@field.fieldLookupContentName)) {
                                        result = result + "<td class=\"p-1 cmRowField\">" + cp.Html.SelectContent(@field.fieldName, @field.currentValue, @field.fieldLookupContentName, "", "Any", "form-control select") + "</TD>";

                                    } else {
                                        result = result + "<td class=\"p-1 cmRowField\">" + cp.Html.SelectList(@field.fieldName, @field.currentValue, @field.fieldLookupList, "Any", "form-control select") + "</TD>";

                                    }
                                    result += "</TR>";
                                    groupTab = "";
                                    RowPointer += 1;
                                    break;
                                }
                        }
                    }
                }
                result += "</Table>";
                // 
                result += cp.Html.Hidden("SelectionSearchSubTab", "1");
                // 
            } catch (Exception ex) {
                cp.Site.ErrorReport(ex);
                throw;
            }
            return result;
        }
        // 
        // =================================================================================
        // 
        public static Constants.FormIdEnum processRequest(CPBaseClass cp, Controllers.ApplicationController ae, RequestModel request) {
            var result = Constants.FormIdEnum.FormList;
            try {
                switch (request.Button ?? "") {
                    case Constants.ButtonSearch: {
                            // 
                            if (!string.IsNullOrEmpty(request.SelectionGroupSubTab)) {
                                // 
                                // Save the Form
                                // 
                                int GroupCount = cp.Doc.GetInteger("GroupCount");
                                string ContactGroupCriteria = "";
                                if (GroupCount > 0) {
                                    int GroupPointer;
                                    var loopTo = GroupCount - 1;
                                    for (GroupPointer = 0; GroupPointer <= loopTo; GroupPointer++) {
                                        string GroupLabel = "Group" + GroupPointer;
                                        if (cp.Doc.GetBoolean(GroupLabel)) {
                                            ContactGroupCriteria = ContactGroupCriteria + "," + cp.Doc.GetInteger(GroupLabel + ".id");
                                        }
                                    }
                                }
                                ae.userProperties.contactGroupCriteria = ContactGroupCriteria + ",";
                            }
                            if (!string.IsNullOrEmpty(request.SelectionSearchSubTab)) {
                                // 
                                // SelectionContentSubTab (crlf FieldName tab FieldType tab FieldVAlue tab Operator)
                                // 
                                using (var csField = cp.CSNew()) {
                                    string Criteria = "(active<>0)and(ContentID=" + cp.Content.GetID("people") + ")and(authorable<>0)";
                                    csField.Open("Content Fields", Criteria, "EditSortPriority");
                                    string ContactSearchCriteria = "";
                                    while (csField.OK()) {
                                        string FieldName = csField.GetText("name");
                                        string FieldValue = cp.Doc.GetText(FieldName);
                                        int fieldType = csField.GetInteger("Type");
                                        string NumericOption;
                                        switch (fieldType) {
                                            case Constants.FieldTypeDate: {
                                                    NumericOption = cp.Doc.GetText(FieldName + "_D");
                                                    if (!string.IsNullOrEmpty(NumericOption)) {
                                                        ContactSearchCriteria = ContactSearchCriteria + Environment.NewLine + FieldName + "\t" + fieldType + "\t" + FieldValue + "\t" + NumericOption;




                                                    }

                                                    break;
                                                }
                                            case Constants.FieldTypeCurrency:
                                            case Constants.FieldTypeFloat:
                                            case Constants.FieldTypeInteger: {
                                                    NumericOption = cp.Doc.GetText(FieldName + "_N");
                                                    if (!string.IsNullOrEmpty(NumericOption)) {
                                                        ContactSearchCriteria = ContactSearchCriteria + Environment.NewLine + FieldName + "\t" + fieldType + "\t" + FieldValue + "\t" + NumericOption;




                                                    }

                                                    break;
                                                }
                                            case Constants.FieldTypeBoolean: {
                                                    if (!string.IsNullOrEmpty(FieldValue)) {
                                                        ContactSearchCriteria = ContactSearchCriteria + Environment.NewLine + FieldName + "\t" + fieldType + "\t" + FieldValue + "\t" + "";




                                                    }

                                                    break;
                                                }
                                            case Constants.FieldTypeText: {
                                                    string TextOption = cp.Doc.GetText(FieldName + "_T");
                                                    if (!string.IsNullOrEmpty(TextOption)) {
                                                        ContactSearchCriteria = ContactSearchCriteria + Environment.NewLine + FieldName + "\t" + fieldType.ToString() + "\t" + FieldValue + "\t" + TextOption;




                                                    }

                                                    break;
                                                }
                                            case Constants.FieldTypeLookup: {
                                                    if (!string.IsNullOrEmpty(FieldValue)) {
                                                        ContactSearchCriteria = ContactSearchCriteria + Environment.NewLine + FieldName + "\t" + fieldType + "\t" + FieldValue + "\t" + "";




                                                    }

                                                    break;
                                                }
                                        }
                                        csField.GoNext();
                                    }
                                    csField.Close();
                                    ae.userProperties.contactSearchCriteria = ContactSearchCriteria;
                                }
                            }

                            break;
                        }
                }
            } catch (Exception ex) {
                cp.Site.ErrorReport(ex);
                throw;
            }
            return result;
        }
        // 
        // =================================================================================
        // 
        public static string getFormInputRadioBox(CPBaseClass cp, string ElementName, string ElementValue, string CurrentValue, string htmlId) {
            return cp.Html.RadioBox(ElementName, ElementValue, CurrentValue, "", htmlId);
        }
        // 
        // =================================================================================
        // 
        public static string getFormInputText(CPBaseClass cp, string htmlName, string CurrentValue, string htmlId, string htmlClass) {
            return cp.Html.InputText(htmlName, CurrentValue, 255, htmlClass + " form-control", htmlId);
        }
        // 
        // =================================================================================
        // 
        public static string getFormInputDate(CPBaseClass cp, string ElementName, DateTime CurrentValue, string ElementID) {
            return cp.Html5.InputDate(ElementName, CurrentValue, "", ElementID);
        }
        // 
    }
}