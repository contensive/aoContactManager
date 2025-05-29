using System;
using System.Collections.Generic;
using Contensive.BaseClasses;
using Contensive.Models.Db;
using Microsoft.VisualBasic;

namespace Contensive.Addons.ContactManager.Views {
    public sealed class ListView {
        // 
        // =================================================================================
        // 
        public static constants.FormIdEnum processRequest(CPBaseClass cp, Controllers.ApplicationController ae, RequestModel request) {
            var resultFormId = constants.FormIdEnum.FormList;
            try {
                // 
                cp.Utils.AppendLog("ListFormController.ProcessRequest, enter");
                // 
                switch (request.Button ?? "") {
                    case constants.ButtonNewSearch: {
                            ae.userProperties.contactSearchCriteria = "";
                            ae.userProperties.contactGroupCriteria = "";
                            request.FormID = constants.FormIdEnum.FormSearch;
                            resultFormId = constants.FormIdEnum.FormSearch;
                            break;
                        }
                    case constants.ButtonApply: {
                            // 
                            // Add to or remove from group
                            // 
                            string SQLCriteria = "(1=1)";
                            string SearchCaption = "";
                            buildSearch(cp, ae, ref SQLCriteria, ref SearchCaption);
                            string GroupName;
                            int RowPointer;
                            int memberID;
                            string SQL;
                            switch (request.GroupToolAction) {
                                case constants.GroupToolActionEnum.AddToGroup: {
                                        // 
                                        // ----- Add to Group
                                        // 
                                        if (request.GroupID == 0) {
                                            // 
                                            // Group required and not provided
                                            // 
                                            cp.UserError.Add("Please select a Target Group for this operation");
                                        } else if (request.GroupToolSelect == 0) {
                                            // 
                                            // Add selection to Group
                                            // 
                                            if (request.RowCount > 0) {
                                                GroupName = cp.Group.GetName(request.GroupID.ToString());
                                                var loopTo = request.RowCount - 1;
                                                for (RowPointer = 0; RowPointer <= loopTo; RowPointer++) {
                                                    if (cp.Doc.GetBoolean("M." + RowPointer)) {
                                                        memberID = cp.Doc.GetInteger("MID." + RowPointer);
                                                        cp.Group.AddUser(GroupName, memberID);
                                                    }
                                                }
                                            }
                                        } else {
                                            // 
                                            // Add everyone in search criteria to this group
                                            // 

                                            int CCID = cp.Content.GetID("Member Rules");
                                            SQL = "insert into ccMemberRules (Active,ContentControlID,GroupID,MemberID ) select 1," + CCID + "," + request.GroupID + ",ccMembers.ID from (ccMembers left join ccMemberRules on ccMemberRules.MemberID=ccMembers.ID ) left join ( select MemberID  from ccMemberRules where GroupID in (" + request.GroupID + ")) as InGroups on InGroups.MemberID=ccMembers.ID where " + SQLCriteria + " and InGroups.MemberID is null";






                                            cp.Db.ExecuteNonQuery(SQL);
                                        }
                                        resultFormId = constants.FormIdEnum.FormList;
                                        break;
                                    }
                                case constants.GroupToolActionEnum.RemoveFromGroup: {
                                        // 
                                        // ----- Remove From Group
                                        // 
                                        if (request.GroupID == 0) {
                                            // 
                                            // Group required and not provided
                                            // 
                                            cp.UserError.Add("Please select a Target Group for this operation");
                                        } else if (request.GroupToolSelect == 0) {
                                            // 
                                            // Remove selection from Group
                                            // 
                                            if (request.RowCount > 0) {
                                                GroupName = cp.Group.GetName(request.GroupID.ToString());
                                                var loopTo1 = request.RowCount - 1;
                                                for (RowPointer = 0; RowPointer <= loopTo1; RowPointer++) {
                                                    if (cp.Doc.GetBoolean("M." + RowPointer)) {
                                                        memberID = cp.Doc.GetInteger("MID." + RowPointer);
                                                        cp.Content.Delete("Member Rules", "(GroupID=" + request.GroupID + ")and(MemberID=" + memberID + ")");
                                                    }
                                                }
                                            }
                                        } else {
                                            // 
                                            // Remove everyone in search criteria from this group
                                            // 
                                            SQL = "delete from ccMemberRules where GroupID=" + request.GroupID + " and MemberID in ( select ccMembers.ID from (ccMembers left join ccMemberRules on ccMemberRules.MemberID=ccMembers.ID )  where " + SQLCriteria + ")";




                                            cp.Db.ExecuteNonQuery(SQL);
                                        }

                                        break;
                                    }
                                case constants.GroupToolActionEnum.ExportGroup: {
                                        // 
                                        // ----- Export
                                        // 
                                        cp.Utils.AppendLog("ListFormController.ProcessRequest, export");
                                        // 
                                        bool Aborttask = false;
                                        if (true) {
                                            string ExportName = SearchCaption;
                                            if (request.GroupToolSelect == 0) {
                                                string RowSQL;
                                                // 
                                                // Export selection from Group
                                                // 
                                                ExportName = "Selected rows from " + ExportName;
                                                RowSQL = "";
                                                if (request.RowCount > 0) {
                                                    GroupName = cp.Group.GetName(request.GroupID.ToString());
                                                    var loopTo2 = request.RowCount - 1;
                                                    for (RowPointer = 0; RowPointer <= loopTo2; RowPointer++) {
                                                        if (cp.Doc.GetBoolean("M." + RowPointer)) {
                                                            memberID = cp.Doc.GetInteger("MID." + RowPointer);
                                                            RowSQL = RowSQL + "OR(ccMembers.ID=" + memberID + ")";
                                                        }
                                                    }
                                                    if (string.IsNullOrEmpty(RowSQL)) {
                                                        // 
                                                        // nothing selected, abort export
                                                        // 
                                                        Aborttask = true;
                                                        ae.statusMessage = "<P>You requested to only download the selected entries, and none were selected.<P>";
                                                    } else if (string.IsNullOrEmpty(SQLCriteria)) {
                                                        // 
                                                        // This is the only criteria
                                                        // 
                                                        SQLCriteria = " WHERE(" + Strings.Mid(RowSQL, 3) + ")";
                                                    } else {
                                                        // 
                                                        // Add this criteria to the previous
                                                        // 
                                                        SQLCriteria = SQLCriteria + " And(" + Strings.Mid(RowSQL, 3) + ")";
                                                    }
                                                }
                                            } else {
                                                // 
                                                // Export the search criteria
                                                // 
                                            }
                                            if (!Aborttask) {
                                                string SQLFrom = "ccMembers";
                                                int JoinTableCnt = 0;
                                                if (!string.IsNullOrEmpty(ae.userProperties.contactGroupCriteria)) {
                                                    SQLFrom = "(" + SQLFrom + " left join ccMemberRules on ccMemberRules.MemberID=ccMembers.ID )";
                                                }
                                                int ContentID = ae.userProperties.contactContentID;
                                                var cs = cp.CSNew();
                                                cs.Open("Content Fields", "ContentID=" + ContentID, "EditSortPriority", true, "Name,Caption,Type,LookupContentID");
                                                string SelectList = "";
                                                string FieldNameList = "";
                                                while (cs.OK()) {
                                                    string FieldName = cs.GetText("name");
                                                    switch (cs.GetInteger("type")) {
                                                        case constants.FieldTypeLookup: {
                                                                // 
                                                                // just add the ID into the list
                                                                // 
                                                                SelectList = SelectList + ",ccMembers." + FieldName;
                                                                FieldNameList = FieldNameList + "," + FieldName;
                                                                break;
                                                            }
                                                        case constants.FieldTypeFileText: {
                                                                break;
                                                            }
                                                        // 
                                                        // read file for text - skip field
                                                        // 
                                                        // no field involved, skip it
                                                        case constants.FieldTypeRedirect:
                                                        case constants.FieldTypeManyToMany: {
                                                                break;
                                                            }

                                                        default: {
                                                                // 
                                                                // just add value
                                                                SelectList = SelectList + ",ccMembers." + FieldName;
                                                                FieldNameList = FieldNameList + "," + FieldName;
                                                                break;
                                                            }
                                                    }
                                                    cs.GoNext();
                                                }
                                                cs.Close();
                                                if (string.IsNullOrEmpty(SelectList)) {
                                                    ae.statusMessage = "<P>There was a problem requesting your download.<P>";
                                                } else {
                                                    SelectList = Strings.Mid(SelectList, 2);
                                                    if (!string.IsNullOrEmpty(FieldNameList)) {
                                                        FieldNameList = Strings.Mid(FieldNameList, 2);
                                                    }
                                                    // ExportName = CStr(Now()) & " snapshot of " & LCase(ExportName)
                                                    SQL = "select Distinct " + SelectList + " from " + SQLFrom + " where " + SQLCriteria;
                                                    // 
                                                    // --- tmp only -- need a new api method to cp.addon,executeAsync( addonid, dictionaryofArgs, DownloadName, downloadfilename )
                                                    // 
                                                    var addon = DbBaseModel.create<AddonModel>(cp, constants.addonGuidExportCSV);
                                                    if (addon is null) {
                                                        // 
                                                        ae.statusMessage = "<P>There was a problem requesting your download. The Csv Export Addon is not installed.<P>";
                                                    } else {
                                                        ae.statusMessage = "<P>Your download request has been submitted and will be available on the <a href=" + cp.Site.GetText("adminurl") + "?af=30>Download Requests</a> page shortly.<P>";
                                                        // 
                                                        var download = DbBaseModel.addDefault<DownloadModel>(cp);
                                                        download.name = "Contact Manager Export by " + cp.User.Name + ", " + DateTime.Now.ToString();
                                                        download.requestedBy = cp.User.Id;
                                                        download.dateRequested = DateTime.Now;
                                                        download.filename.content = Constants.vbCrLf;
                                                        download.filename.filename = ContactManagerTools.Controllers.GenericController.getVirtualRecordUnixPathFilename(DownloadModel.tableMetadata.tableNameLower, "filename", download.id, "export.csv");
                                                        download.save(cp);
                                                        // 
                                                        var args = new Dictionary<string, string>() { { "sql", SQL } };
                                                        // 
                                                        var cmdDetail = new TaskModel.CmdDetailClass() {
                                                            addonId = addon.id,
                                                            addonName = addon.name,
                                                            args = args
                                                        };
                                                        // 
                                                        var task = DbBaseModel.addDefault<TaskModel>(cp);
                                                        task.name = "addon [#" + cmdDetail.addonId + "," + cmdDetail.addonName + "]";
                                                        task.cmdDetail = cp.JSON.Serialize(cmdDetail);
                                                        task.resultDownloadId = download.id;
                                                        task.save(cp);
                                                    }
                                                }
                                            }
                                        }
                                        resultFormId = constants.FormIdEnum.FormList;
                                        break;
                                    }
                                case constants.GroupToolActionEnum.SetGroupEmail: {
                                        // 
                                        // ----- Set AllowBulkEmail field
                                        // 
                                        if (request.GroupToolSelect == 0) {
                                            // 
                                            // Just selection
                                            // 
                                            int RecordCnt = 0;
                                            if (request.RowCount > 0) {
                                                GroupName = cp.Group.GetName(request.GroupID.ToString());
                                                var loopTo3 = request.RowCount - 1;
                                                for (RowPointer = 0; RowPointer <= loopTo3; RowPointer++) {
                                                    if (cp.Doc.GetBoolean("M." + RowPointer)) {
                                                        memberID = cp.Doc.GetInteger("MID." + RowPointer);
                                                        cp.Db.ExecuteNonQuery("update ccMembers set AllowBulkEmail=1 where ID=" + memberID);
                                                        RecordCnt += 1;
                                                    }
                                                }
                                            }
                                            ae.statusMessage = "<P>Allow Group Email was set for " + RecordCnt + " people.<P>";
                                        } else {
                                            // 
                                            // Set for everyone in search criteria
                                            // 
                                            SQL = "Update ccMembers set AllowBulkEmail=1 where ID in ( select Distinct ccMembers.ID from (ccMembers left join ccMemberRules on ccMemberRules.MemberID=ccMembers.ID ) " + SQLCriteria + ")";




                                            cp.Db.ExecuteNonQuery(SQL);
                                            ae.statusMessage = "<P>Allow Group Email was set for all people in this selection.<P>";
                                        }

                                        break;
                                    }
                            }

                            break;
                        }
                }
                // 
                cp.Utils.AppendLog("ListFormController.ProcessRequest, exit");
                // 
            } catch (Exception ex) {
                // 
                cp.Utils.AppendLog("ListFormController.ProcessRequest, exception [" + ex.ToString() + "]");
                // 
                cp.Site.ErrorReport(ex);
                throw;
            }
            return resultFormId;
        }
        // 
        // =================================================================================
        // 
        public static string getResponse(CPBaseClass cp, Controllers.ApplicationController ae) {
            string result = "";
            try {
                // 
                if (true) {
                    var layout = cp.AdminUI.CreateLayoutBuilder();
                    // 
                    bool IsAdmin = cp.User.IsAdmin;
                    string TextTest = cp.Doc.GetText(constants.RequestNamePageSize);
                    int pageSize;
                    if (string.IsNullOrEmpty(TextTest)) {
                        pageSize = cp.Utils.EncodeInteger(cp.Visit.GetText("cmPageSize", "50"));
                    } else {
                        pageSize = cp.Utils.EncodeInteger(TextTest);
                        cp.Visit.SetProperty("cmPageSize", pageSize.ToString());
                    }
                    if (pageSize == 0) {
                        pageSize = 50;
                    }
                    int PageNumber = cp.Doc.GetInteger(constants.RequestNamePageNumber);
                    if (PageNumber == 0) {
                        PageNumber = 1;
                    }
                    int GroupID = cp.Doc.GetInteger("GroupID");
                    int GroupToolAction = cp.Doc.GetInteger("GroupToolAction");
                    string AdminURL = cp.Site.GetText("adminurl");
                    int ColumnMax = 5;
                    // 
                    int TopCount = PageNumber * pageSize;
                    string[] ColCaption;
                    // 
                    ColCaption = new string[ColumnMax + 1];
                    string[] ColAlign;
                    ColAlign = new string[ColumnMax + 1];
                    string[] ColWidth;
                    ColWidth = new string[ColumnMax + 1];
                    bool[] ColSortable;
                    ColSortable = new bool[ColumnMax + 1];
                    string[,] Cells;
                    Cells = new string[pageSize + 1, ColumnMax + 1];
                    // 
                    int SortColPtr = 3;
                    TextTest = cp.Visit.GetText("cmSortColumn", "");
                    if (!string.IsNullOrEmpty(TextTest)) {
                        SortColPtr = cp.Utils.EncodeInteger(TextTest);
                    }
                    SortColPtr = ContactManagerTools.AdminUIController.getReportSortColumnPtr(cp, SortColPtr);
                    if ((SortColPtr.ToString() ?? "") != (TextTest ?? "")) {
                        cp.Visit.SetProperty("cmSortColumn", SortColPtr.ToString());
                    }
                    string SQLOrderDir = "";
                    if (ContactManagerTools.AdminUIController.getReportSortType(cp) == 2) {
                        SQLOrderDir = " Desc";
                    }
                    // 
                    // Headers
                    // 
                    int CPtr = 0;
                    ColCaption[CPtr] = "<input type=checkbox id=\"cmSelectAll\">";
                    // ColCaption(CPtr) = "<INPUT TYPE=CheckBox OnClick=""CheckInputs('DelCheck',this.checked);""><BR><img src=/cclib/images/spacer.gif width=10 height=1>"
                    ColAlign[CPtr] = "center";
                    ColWidth[CPtr] = "30px";
                    ColSortable[CPtr] = false;
                    CPtr += 1;
                    // 
                    ColCaption[CPtr] = "Name";
                    ColAlign[CPtr] = "left";
                    ColWidth[CPtr] = "";
                    ColSortable[CPtr] = true;
                    int DefaultSortColumnPtr = CPtr;
                    string SQLOrderBy = "";
                    if (CPtr == SortColPtr) {
                        SQLOrderBy = "Order By ccMembers.Name";
                    }
                    CPtr += 1;
                    // 
                    ColCaption[CPtr] = "Organization";
                    ColAlign[CPtr] = "left";
                    ColWidth[CPtr] = "20%";
                    ColSortable[CPtr] = true;
                    if (CPtr == SortColPtr) {
                        SQLOrderBy = "Order By Organizations.Name";
                    }
                    CPtr += 1;
                    // 
                    ColCaption[CPtr] = "Phone";
                    // ColCaption(CPtr) = "Phone<BR><img src=/cclib/images/spacer.gif width=100 height=1>"
                    ColAlign[CPtr] = "left";
                    ColWidth[CPtr] = "20%";
                    ColSortable[CPtr] = true;
                    if (CPtr == SortColPtr) {
                        SQLOrderBy = "Order By ccMembers.Phone";
                    }
                    CPtr += 1;
                    // 
                    ColCaption[CPtr] = "email";
                    // ColCaption(CPtr) = "email<BR><img src=/cclib/images/spacer.gif width=150 height=1>"
                    ColAlign[CPtr] = "left";
                    ColWidth[CPtr] = "20%";
                    ColSortable[CPtr] = true;
                    if (CPtr == SortColPtr) {
                        SQLOrderBy = "Order By ccMembers.Email";
                    }
                    CPtr += 1;
                    string RQS = cp.Doc.RefreshQueryString;
                    // 
                    // SubTab Menu
                    // 
                    RQS = cp.Doc.RefreshQueryString;
                    RQS = cp.Utils.ModifyQueryString(RQS, "tab", "", false);
                    string Criteria = "(1=1)";
                    string SearchCaption = "";
                    buildSearch(cp, ae, ref Criteria, ref SearchCaption);
                    // 
                    // Get DataRowCount
                    // This had been commented out - but it is needed for the "matches found" caption
                    // 
                    string SQL = "Select distinct ccMembers.ID from ccMembers left join ccMemberRules on ccMemberRules.MemberID=ccMembers.ID where " + Criteria;


                    using (var CS = cp.CSNew()) {
                        CS.OpenSQL(SQL);
                        var DataRowCount = default(int);
                        if (CS.OK()) {
                            DataRowCount = CS.GetRowCount();
                        }
                        CS.Close();
                        // 
                        // Get Data
                        // 
                        string DefaultName = "Guest";
                        SQL = "Select distinct Top " + TopCount + " ccMembers.name,ccMembers.FirstName,ccMembers.LastName,ccMembers.ID, ccMembers.ContentControlID, ccMembers.Visits, ccMembers.LastVisit, ccMembers.Phone, ccMembers.Email,Organizations.Name as OrgName from ((ccMembers left join organizations on Organizations.ID=ccMembers.OrganizationID) left join ccMemberRules on ccMemberRules.MemberID=ccMembers.ID) where " + Criteria + SQLOrderBy + SQLOrderDir;




                        CS.OpenSQL(SQL, "", pageSize, PageNumber);
                        int RowPointer = 0;
                        if (!CS.OK()) {
                            Cells[0, 3] = "This search returned no results";
                            RowPointer = 1;
                        } else {
                            // DataRowCount = Main.GetCSRowCount(CS)
                            while (CS.OK() & RowPointer < pageSize) {
                                CPtr = 0;
                                int RecordID = CS.GetInteger("ID");
                                string CheckBox = cp.Html.CheckBox("M." + RowPointer, false, "cmSelect");
                                Cells[RowPointer, CPtr] = CheckBox + cp.Html.Hidden("MID." + RowPointer, RecordID.ToString());
                                CPtr += 1;
                                string RecordName = CS.GetText("name");
                                if (string.IsNullOrEmpty(RecordName)) {
                                    RecordName = DefaultName + "&nbsp;" + RecordID;
                                }
                                Cells[RowPointer, CPtr] = "<a href=\"?" + RQS + "&" + constants.RequestNameMemberID + "=" + RecordID + "\">" + RecordName + "</a>";
                                CPtr += 1;
                                Cells[RowPointer, CPtr] = CS.GetText("OrgName");
                                CPtr += 1;
                                Cells[RowPointer, CPtr] = CS.GetText("phone");
                                CPtr += 1;
                                Cells[RowPointer, CPtr] = CS.GetText("email");
                                CPtr += 1;
                                RowPointer += 1;
                                CS.GoNext();
                            }
                        }
                        CS.Close();
                        result += cp.Html.Hidden("M.Count", RowPointer.ToString());
                        string BlankPanel = "<div class=\"cmBody ccPanel3DReverse\">x</div>";
                        string RowPageSize = "<TABLE border=0 cellpadding=4 cellspacing=0 width=500><TR><td class=\"p-1\" class=APLeft>Rows Per Page</TD><td class=\"p-1\" class=apright>" + cp.Html5.InputText(constants.RequestNamePageSize, 4, pageSize.ToString()) + "</TD></TR></Table>";




                        string RowGroups = "<TABLE border=0 cellpadding=4 cellspacing=0 width=500><TR><td class=\"p-1\" valign=top class=APLeft>Actions</TD><td class=\"p-1\" class=APRight><div class=\"APRight\">Source Contacts</div><div class=\"form-check\"><div class=\"APRight form-check\">" + cp.Html5.RadioBox("GroupToolSelect", 0, 0, "form-check-input", "js-source-on-page") + "&nbsp;<label for=\"js-source-on-page\">Only those selected on this page</label></div><div class=\"APRight form-check\"><input class=\"form-check-input\" type=radio name=GroupToolSelect value=1 id=\"js-source-everyone\">&nbsp;<label for=\"js-source-everyone\">Everyone in search results</label></div><div style=\"border-top:1px solid black;border-bottom:1px solid white;margin-top:4px;margin-bottom:4px;\"></div><div class=APRight>Perform Action</div><div class=\"APRight form-check\">" + cp.Html5.RadioBox("GroupToolAction", 0, 0, "form-check-input", "js-action-none") + " <label for=\"js-action-none\">No Action</label></div><div class=\"APRight form-check\">" + cp.Html5.RadioBox("GroupToolAction", 1, 0, "form-check-input", "js-action-add") + " <label for=\"js-action-add\">Add to Target Group</label></div><div class=\"APRight form-check\">" + cp.Html5.RadioBox("GroupToolAction", 2, 0, "form-check-input", "js-action-remove") + " <label for=\"js-action-remove\">Remove from Target Group</label></div><div class=\"APRight form-check\">" + cp.Html5.RadioBox("GroupToolAction", 3, 0, "form-check-input", "js-action-export") + " <label for=\"js-action-export\">Export comma delimited file</label></div><div class=\"APRight form-check\">" + cp.Html5.RadioBox("GroupToolAction", 4, 0, "form-check-input", "js-action-allowEmail") + " <label for=\"js-action-allowEmail\">Set Allow Group Email</label></div><div style=\"border-top:1px solid black;border-bottom:1px solid white;margin-top:4px;margin-bottom:4px;\"></label></div><div class=APRight style=\"padding-bottom:6px;\">Target Group</div><div class=\"APRight form-check\">" + cp.Html.SelectContent("GroupID", GroupID.ToString(), "Groups", "", "Select Target Group", "form-control") + "</div></div></TD></TR></Table>";



















                        string ActionPanel = "" + Constants.vbCrLf + "<style>.APLeft{width:100px;text-align:left;}.APRight{text-align:left;}.APRightIndent{text-align:left;padding-left:10px;}</style>" + Constants.vbCrLf + Strings.Replace(BlankPanel, "x", RowPageSize) + Strings.Replace(BlankPanel, "x", RowGroups) + "";









                        string PostTableCopy = "<div class=\"cmBody ccPanel3D\">" + ActionPanel + "</div>" + cp.Html.Hidden("M.Count", RowPointer.ToString()) + cp.Html.Hidden(constants.RequestNameFormID, Convert.ToInt32((int)constants.FormIdEnum.FormList).ToString()) + "";



                        // 
                        // Header
                        // 
                        string Description = "This is a list of people that satisfy the current search criteria. Create a new search to change the criteria. Use the tools at the bottom to make changes to this selection.</p>" +
                            "<p>" + DataRowCount + " Matches found" + ae.statusMessage;


                        string ButtonList = constants.ButtonApply + "," + constants.ButtonNewSearch;
                        string PreTableCopy = "";
                        string Body = ContactManagerTools.AdminUIController.getReport2(cp, RowPointer, ColCaption, ColAlign, ColWidth, Cells, pageSize, PageNumber, PreTableCopy, PostTableCopy, DataRowCount, "ccPanel", ColSortable, SortColPtr);
                        // 
                        // Assemble page
                        layout.body = Body;
                        layout.description = Description;
                        layout.failMessage = "";
                        layout.includeBodyColor = true;
                        layout.includeBodyPadding = true;
                        layout.infoMessage = "";
                        layout.isOuterContainer = true;
                        layout.portalSubNavTitle = "";
                        layout.successMessage = "";
                        layout.title = "Contact Manager";
                        layout.warningMessage = "";
                        layout.addFormButton(constants.ButtonApply);
                        layout.addFormButton(constants.ButtonNewSearch);
                        // 
                        result = layout.getHtml();
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
        /// <summary>
        /// 
        /// </summary>
        /// <param name="cp"></param>
        /// <param name="ae"></param>
        /// <param name="return_Criteria"></param>
        /// <param name="return_SearchCaption"></param>
        public static void buildSearch(CPBaseClass cp, Controllers.ApplicationController ae, ref string return_Criteria, ref string return_SearchCaption) {
            try {
                string LookupName;
                int ContactContentID;
                string ContactGroupCriteria;
                string ContactSearchCriteria;
                string[] CriteriaValues;
                int CriteriaCount;
                int CriteriaPointer;
                string[] FieldParms;
                string SQLValue;
                string[] GroupIDs;
                string GroupDelimiter;
                var GroupPtr = default(int);
                string SQLNow;
                string SQL;
                string FieldCaption;
                var FieldContentLookupID = default(int);
                string FieldContentLookupName;
                // 
                ContactContentID = cp.Content.GetID("people");
                ContactGroupCriteria = ae.userProperties.contactGroupCriteria;
                ContactSearchCriteria = ae.userProperties.contactSearchCriteria;
                return_SearchCaption = "";
                string hint = "";
                // 
                // Search Criteria Fields  (crlf FieldName tab FieldType tab FieldVAlue tab Operator)
                // 
                if (!string.IsNullOrEmpty(ContactSearchCriteria)) {
                    hint += ",110";
                    CriteriaValues = Strings.Split(ContactSearchCriteria, Constants.vbCrLf);
                    CriteriaCount = Information.UBound(CriteriaValues) + 1;
                    var loopTo = CriteriaCount - 1;
                    for (CriteriaPointer = 0; CriteriaPointer <= loopTo; CriteriaPointer++) {
                        hint += ",120";
                        if (!string.IsNullOrEmpty(CriteriaValues[CriteriaPointer])) {
                            hint += ",130";
                            FieldParms = Strings.Split(CriteriaValues[CriteriaPointer], Constants.vbTab);
                            if (Information.UBound(FieldParms) >= 3) {
                                // 
                                // Look up caption
                                // 
                                FieldCaption = "";
                                SQL = "select caption,LookupContentID from ccfields where name=" + cp.Db.EncodeSQLText(FieldParms[0]) + " and contentid=" + ContactContentID;
                                using (var CS = cp.CSNew()) {
                                    CS.OpenSQL(SQL);
                                    if (CS.OK()) {
                                        FieldCaption = CS.GetText("caption");
                                        FieldContentLookupID = CS.GetInteger("LookupContentID");
                                    }
                                    CS.Close();
                                }
                                if (string.IsNullOrEmpty(FieldCaption)) {
                                    FieldCaption = FieldParms[0];
                                    FieldContentLookupID = 0;
                                }
                                switch (cp.Utils.EncodeInteger(FieldParms[1])) {
                                    case constants.FieldTypeLongText:
                                    case constants.FieldTypeText: {
                                            // 
                                            // text
                                            // 
                                            hint += ",140";
                                            switch (FieldParms[3] ?? "") {
                                                case "1": {
                                                        // 
                                                        // is empty
                                                        // 
                                                        hint += ",150";
                                                        return_Criteria = return_Criteria + "AND((ccMembers." + FieldParms[0] + " is null)or(ccMembers." + FieldParms[0] + "=''))";
                                                        return_SearchCaption += ", " + FieldCaption + " is empty";
                                                        break;
                                                    }
                                                case "2": {
                                                        // 
                                                        // is not empty
                                                        // 
                                                        hint += ",160";
                                                        return_Criteria = return_Criteria + "AND(ccMembers." + FieldParms[0] + " is not null)";
                                                        return_SearchCaption += ", " + FieldCaption + " is not empty";
                                                        break;
                                                    }
                                                case "3": {
                                                        // 
                                                        // includes a string
                                                        // 
                                                        hint += ",170";
                                                        if (!string.IsNullOrEmpty(FieldParms[2])) {
                                                            SQLValue = cp.Db.EncodeSQLTextLike(FieldParms[2]);
                                                            return_Criteria = return_Criteria + "AND(ccMembers." + FieldParms[0] + " like " + SQLValue + ")";
                                                            return_SearchCaption += ", " + FieldCaption + " includes '" + FieldParms[2] + "'";
                                                        }

                                                        break;
                                                    }
                                            }

                                            break;
                                        }
                                    case constants.FieldTypeLookup: {
                                            // 
                                            // Lookup
                                            // 
                                            hint += ",180";
                                            if (cp.Utils.EncodeInteger(FieldParms[2]) > 0) {
                                                LookupName = "";
                                                if (FieldContentLookupID != 0) {
                                                    FieldContentLookupName = cp.Content.GetRecordName("Content", FieldContentLookupID);
                                                    if (!string.IsNullOrEmpty(FieldContentLookupName) & FieldParms[2].IsNumeric()) {
                                                        LookupName = cp.Content.GetRecordName(FieldContentLookupName, cp.Utils.EncodeInteger(FieldParms[2]));
                                                    }
                                                }
                                                if (string.IsNullOrEmpty(LookupName)) {
                                                    LookupName = FieldParms[2];
                                                }

                                                return_Criteria = return_Criteria + "AND(ccMembers." + FieldParms[0] + "=" + FieldParms[2] + ")";
                                                return_SearchCaption += ", " + FieldCaption + " = " + LookupName;
                                            }
                                            break;
                                        }
                                    case constants.FieldTypeDate: {
                                            // 
                                            // date
                                            // 
                                            hint += ",185";
                                            switch (cp.Utils.EncodeInteger(FieldParms[3])) {
                                                case 1: {
                                                        return_Criteria = return_Criteria + "AND(ccMembers." + FieldParms[0] + "=" + cp.Db.EncodeSQLDate(cp.Utils.EncodeDate(FieldParms[2])) + ")";
                                                        return_SearchCaption += ", " + FieldCaption + " = " + FieldParms[2];
                                                        break;
                                                    }
                                                case 2: {
                                                        return_Criteria = return_Criteria + "AND(ccMembers." + FieldParms[0] + ">" + cp.Db.EncodeSQLDate(cp.Utils.EncodeDate(FieldParms[2])) + ")";
                                                        return_SearchCaption += ", " + FieldCaption + " > " + FieldParms[2];
                                                        break;
                                                    }
                                                case 3: {
                                                        return_Criteria = return_Criteria + "AND(ccMembers." + FieldParms[0] + "<" + cp.Db.EncodeSQLDate(cp.Utils.EncodeDate(FieldParms[2])) + ")";
                                                        return_SearchCaption += ", " + FieldCaption + " < " + FieldParms[2];
                                                        break;
                                                    }
                                            }

                                            break;
                                        }
                                    case constants.FieldTypeCurrency:
                                    case constants.FieldTypeFloat:
                                    case constants.FieldTypeInteger:
                                    case var @case when @case == constants.FieldTypeLookup: {
                                            // 
                                            // number
                                            // 
                                            hint += ",190";
                                            switch (cp.Utils.EncodeInteger(FieldParms[3])) {
                                                case 1: {
                                                        return_Criteria = return_Criteria + "AND(ccMembers." + FieldParms[0] + "=" + cp.Db.EncodeSQLNumber(cp.Utils.EncodeNumber(FieldParms[2])) + ")";
                                                        return_SearchCaption += ", " + FieldCaption + " = " + FieldParms[2];
                                                        break;
                                                    }
                                                case 2: {
                                                        return_Criteria = return_Criteria + "AND(ccMembers." + FieldParms[0] + ">" + cp.Db.EncodeSQLNumber(cp.Utils.EncodeNumber(FieldParms[2])) + ")";
                                                        return_SearchCaption += ", " + FieldCaption + " > " + FieldParms[2];
                                                        break;
                                                    }
                                                case 3: {
                                                        return_Criteria = return_Criteria + "AND(ccMembers." + FieldParms[0] + "<" + cp.Db.EncodeSQLNumber(cp.Utils.EncodeNumber(FieldParms[2])) + ")";
                                                        return_SearchCaption += ", " + FieldCaption + " < " + FieldParms[2];
                                                        break;
                                                    }
                                            }

                                            break;
                                        }
                                    case constants.FieldTypeBoolean: {
                                            // 
                                            // boolean
                                            // 
                                            hint += ",200";
                                            if (FieldParms[2].Equals("1")) {
                                                // 
                                                // 1 = search for true
                                                // 
                                                return_Criteria = return_Criteria + "AND(ccMembers." + FieldParms[0] + ">0)";
                                                return_SearchCaption += ", " + FieldCaption + " is true";
                                            } else if (FieldParms[2].Equals("2")) {
                                                // 
                                                // 2 = search for false
                                                // 
                                                return_Criteria = return_Criteria + "AND((ccMembers." + FieldParms[0] + "=0)or(ccMembers." + FieldParms[0] + " is null))";
                                                return_SearchCaption += ", " + FieldCaption + " is false";
                                            } else {
                                                // 
                                                // 0 = ignore
                                                // 
                                            }

                                            break;
                                        }

                                    default: {
                                            break;
                                        }
                                }
                            }
                        }
                    }
                }
                if (!string.IsNullOrEmpty(return_SearchCaption)) {
                    hint += ",210";
                    return_SearchCaption = Strings.Mid(return_SearchCaption, 3);
                    return_SearchCaption = " where " + return_SearchCaption;
                }
                // 
                // Group Rules Criteria
                // 
                hint += ",220";
                if (Strings.Left(ContactGroupCriteria, 1) == ",") {
                    ContactGroupCriteria = Strings.Mid(ContactGroupCriteria, 2);
                }
                if (Strings.Right(ContactGroupCriteria, 1) == ",") {
                    ContactGroupCriteria = Strings.Mid(ContactGroupCriteria, 1, Strings.Len(ContactGroupCriteria) - 1);
                }
                if (!string.IsNullOrEmpty(ContactGroupCriteria)) {
                    hint += ",230";
                    GroupIDs = Strings.Split(ContactGroupCriteria, ",");
                    GroupDelimiter = "";
                    SQLNow = cp.Db.EncodeSQLDate(DateTime.Now);
                    if (Information.UBound(GroupIDs) == 0) {
                        hint += ",240";
                        if (string.IsNullOrEmpty(return_SearchCaption)) {
                            return_SearchCaption = " in group '" + getGroupCaption(cp, cp.Utils.EncodeInteger(GroupIDs[GroupPtr])) + "'";
                        } else {
                            return_SearchCaption += ", in group '" + getGroupCaption(cp, cp.Utils.EncodeInteger(GroupIDs[GroupPtr])) + "'";
                        }
                        return_Criteria = return_Criteria + "AND((ccMemberRules.GroupID=" + GroupIDs[0] + ")and((ccMemberRules.DateExpires is null)or(ccMemberRules.DateExpires>" + SQLNow + ")))";
                    } else {
                        hint += ",250";
                        if (!string.IsNullOrEmpty(return_SearchCaption)) {
                            return_SearchCaption += ", in group(s) ";
                        } else {
                            return_SearchCaption = " in group(s) ";
                        }
                        string GroupCriteria = "";
                        var loopTo1 = Information.UBound(GroupIDs);
                        for (GroupPtr = 0; GroupPtr <= loopTo1; GroupPtr++) {
                            hint += ",260";
                            GroupCriteria = GroupCriteria + "OR((ccMemberRules.GroupID=" + GroupIDs[GroupPtr] + ")and((ccMemberRules.DateExpires is null)or(ccMemberRules.DateExpires>" + SQLNow + ")))";
                            // Criteria = Criteria & "AND(ccMemberRules.GroupID=GroupIDs(GroupPtr))"
                            if (GroupPtr == Information.UBound(GroupIDs) & GroupPtr != 0) {
                                hint += ",270";
                                return_SearchCaption += " or '" + getGroupCaption(cp, cp.Utils.EncodeInteger(GroupIDs[GroupPtr])) + "'";
                            } else {
                                hint += ",280";
                                return_SearchCaption += GroupDelimiter + "'" + getGroupCaption(cp, cp.Utils.EncodeInteger(GroupIDs[GroupPtr])) + "'";
                            }
                            GroupDelimiter = ", ";
                        }
                        return_Criteria = return_Criteria + "and(" + Strings.Mid(GroupCriteria, 3) + ")";
                    }
                }
                // 
                // Add Content Criteria
                // 
                hint += ",300";
                // If Not String.IsNullOrEmpty(return_Criteria) Then
                // return_Criteria = " WHERE " & Mid(return_Criteria, 4)
                // End If
                return;
            } catch (Exception ex) {
                cp.Site.ErrorReport(ex);
                throw;
            }

            // 
        }
        // 
        // =================================================================================
        // 
        private static string getGroupCaption(CPBaseClass cp, int groupId) {
            var @group = DbBaseModel.create<GroupModel>(cp, groupId);
            if (group is not null) {
                return string.IsNullOrWhiteSpace(group.name) ? group.caption : group.name;
            }
            return "";
        }
    }
}