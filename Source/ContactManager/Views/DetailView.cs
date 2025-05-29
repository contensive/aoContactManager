using System;
using System.Linq;
using System.Text;
using Contensive.Addons.ContactManagerTools;
using Contensive.BaseClasses;
using Contensive.Models.Db;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace Contensive.Addons.ContactManager.Views {
    public sealed class DetailView {
        // 
        // =================================================================================
        // 
        public static constants.FormIdEnum processRequest(CPBaseClass cp, RequestModel request) {
            request.DetailMemberID = cp.Doc.GetInteger(constants.RequestNameMemberID);
            switch (request.Button ?? "") {
                case constants.ButtonCancel: {
                        return constants.FormIdEnum.FormList;
                    }
                case constants.ButtonSave: {
                        savePersonFromStream(cp, request.DetailMemberID);
                        return constants.FormIdEnum.FormDetails;
                    }
                case constants.ButtonOK: {
                        savePersonFromStream(cp, request.DetailMemberID);
                        return constants.FormIdEnum.FormList;
                    }
                case constants.ButtonNewSearch: {
                        return constants.FormIdEnum.FormSearch;
                    }
            }
            return constants.FormIdEnum.FormDetails;
        }
        // 
        // =================================================================================
        // 
        public static string getResponse(CPBaseClass cp, Controllers.ApplicationController ae, int DetailMemberID) {
            string result = "";
            try {
                using (var csPerson = cp.CSNew()) {
                    csPerson.Open("people", "ID=" + DetailMemberID, "", false);
                    string memberName = csPerson.GetText("name");
                    if (string.IsNullOrEmpty(memberName)) {
                        memberName = Strings.Trim(csPerson.GetText("FirstName") + " " + csPerson.GetText("LastName"));
                        if (string.IsNullOrEmpty(memberName)) {
                            memberName = "Record " + csPerson.GetText("ID");
                        }
                    }
                    // 
                    // Determine current Subtab
                    // 
                    int SubTab = cp.Doc.GetInteger(constants.RequestNameDetailSubtab);
                    if (SubTab == 0) {
                        SubTab = ae.userProperties.subTab;
                        if (SubTab == 0) {
                            SubTab = 1;
                            ae.userProperties.subTab = SubTab;
                        }
                    } else {
                        ae.userProperties.subTab = SubTab;
                    }
                    cp.Doc.AddRefreshQueryString(constants.RequestNameDetailSubtab, SubTab.ToString());
                    // 
                    // SubTab Menu
                    // 
                    cp.Doc.AddRefreshQueryString("tab", "");
                    string ButtonList = constants.ButtonCancel + "," + constants.ButtonSave + "," + constants.ButtonOK + "," + constants.ButtonNewSearch;
                    string Header = "<div>" + memberName + "</div>";
                    // 
                    var Nav = new TabController();
                    Nav.addEntry("Contact", getFormDetail_TabContact(cp, csPerson), "ccAdminTab");
                    Nav.addEntry("Permissions", getFormDetail_TabPermissions(cp, csPerson), "ccAdminTab");
                    Nav.addEntry("Notes", getFormDetail_TabNotes(cp, csPerson), "ccAdminTab");
                    Nav.addEntry("Photos", getFormDetail_TabPhoto(cp, csPerson), "ccAdminTab");
                    Nav.addEntry("Groups", getFormDetail_TabGroup(cp, csPerson), "ccAdminTab");
                    csPerson.Close();
                    // 
                    string Content = Nav.getTabs(cp);
                    Content += cp.Html.Hidden(constants.RequestNameFormID, Convert.ToInt32((int)constants.FormIdEnum.FormDetails).ToString());
                    Content += cp.Html.Hidden(constants.RequestNameMemberID, DetailMemberID.ToString());
                    // 
                    result = AdminUIController.getBody(cp, "Contact Manager &gt;&gt; Contact Details", ButtonList, "", true, true, Header, "", Content);
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
        public static string getFormDetail_TabContact(CPBaseClass cp, CPCSBaseClass csPerson) {
            string result = "";
            try {
                string left = "";
                string right = "";
                // 
                if (!csPerson.OK()) {
                    // 
                    result += "<div>There was a problem retrieving this person's information.</div>";
                } else {
                    // 
                    // Left Side
                    // 
                    left += "<table border=0 width=100% cellspacing=0 cellpadding=0><TR><td width=150><img src=/cclib/images/spacer.gif width=150 height=1></TD><td width=350><img src=/cclib/images/spacer.gif width=350 height=1></TD></TR>";



                    left += getFormDetail_InputTextRow(cp, "Full Name", "Name", csPerson.GetText("name"), false);
                    left += getFormDetail_InputTextRow(cp, "First Name", "FirstName", csPerson.GetText("FirstName"), false);
                    left += getFormDetail_InputTextRow(cp, "Last Name", "LastName", csPerson.GetText("LastName"), false);
                    left += getFormDetail_DividerRow(cp, "Contact");
                    left += getFormDetail_InputTextRow(cp, "Email", "EMAIL", csPerson.GetText("email"), false);
                    left += getFormDetail_InputTextRow(cp, "Phone", "PHONE", csPerson.GetText("PHONE"), false);
                    left += getFormDetail_InputTextRow(cp, "Fax", "Fax", csPerson.GetText("Fax"), false);
                    left += getFormDetail_InputTextRow(cp, "Address", "ADDRESS", csPerson.GetText("ADDRESS"), false);
                    left += getFormDetail_InputTextRow(cp, "Line 2", "ADDRESS2", csPerson.GetText("ADDRESS2"), false);
                    left += getFormDetail_InputTextRow(cp, "City", "City", csPerson.GetText("City"), false);
                    left += getFormDetail_InputTextRow(cp, "State", "State", csPerson.GetText("State"), false);
                    left += getFormDetail_InputTextRow(cp, "Zip", "Zip", csPerson.GetText("Zip"), false);
                    left += getFormDetail_DividerRow(cp, "Company");
                    left += getFormDetail_InputTextRow(cp, "Name", "Company", csPerson.GetText("Company"), false);
                    left += getFormDetail_InputTextRow(cp, "Title", "Title", csPerson.GetText("Title"), false);
                    left += getFormDetail_DividerRow(cp, "Birthday");
                    left += getFormDetail_InputTextRow(cp, "Day", "BirthdayDay", csPerson.GetText("BirthdayDay"), false);
                    left += getFormDetail_InputTextRow(cp, "Month", "BirthdayMonth", csPerson.GetText("BirthdayMonth"), false);
                    left += getFormDetail_InputTextRow(cp, "Year", "BirthdayYear", csPerson.GetText("BirthdayYear"), false);
                    left += "</table>";
                    // 
                    // Right Side
                    // 
                    string copy = cp.Html.InputText("AppendNotes", "", 255);
                    copy = Strings.Replace(copy, " cols=\"100\"", " style=\"width:100%;\"", Compare: Microsoft.VisualBasic.Constants.vbTextCompare);
                    // 
                    right += "<table border=0 width=100% cellspacing=0 cellpadding=0><TR><td width=200><img src=/cclib/images/spacer.gif width=150 height=1></TD><td width=99%><img src=/cclib/images/spacer.gif width=1 height=1></TD></TR>";



                    right += getFormDetail_DividerRow(cp, "Add to Notes");
                    right += "<TR><td colspan=2>" + copy + "</TD></TR>";
                    right += "</table>";
                }
                // 
                result += "<table border=0 width=100% cellspacing=0 cellpadding=0><TR><td width=500 valign=top style=\"border-right:1px solid #808080;padding-right:20px;\">" + left + "</TD><td width=100% valign=top style=\"border-left:1px solid #f0f0f0;padding-left:20px;\">" + right + "</TD></TR></table>";





                return "<div STYLE=\"width:100%;\" class=\"cmBody ccPanel3DReverse\">" + result + "</div>";
            } catch (Exception ex) {
                cp.Site.ErrorReport(ex);
                throw;
            }
        }
        // 
        // =================================================================================
        // 
        public static string getFormDetail_TabPermissions(CPBaseClass cp, CPCSBaseClass CS) {
            string result = "";
            try {
                if (!CS.OK()) {
                    // 
                    result += "<div>There was a problem retrieving this person's information.</div>";
                } else {
                    // 
                    result += "<table border=0 width=100% cellspacing=0 cellpadding=0><TR><td width=200><img src=/cclib/images/spacer.gif width=150 height=1></TD><td width=99%><img src=/cclib/images/spacer.gif width=1 height=1></TD></TR>";



                    result += getFormDetail_DividerRow(cp, "Login");
                    result += getFormDetail_InputTextRow(cp, "Username", "Username", CS.GetText("Username"), false);
                    result += getFormDetail_InputTextRow(cp, "Password", "Password", CS.GetText("Password"), true);
                    result += getFormDetail_InputBooleanRow(cp, "Allow Auto Login", "AutoLogin", CS.GetBoolean("AutoLogin").ToString());
                    result += "</table>";
                }
                result = "<div STYLE=\"width:100%;\" class=\"cmBody ccPanel3DReverse\">" + result + "</div>";
            } catch (Exception ex) {
                cp.Site.ErrorReport(ex);
                throw;
            }
            return result;
        }
        // 
        // =================================================================================
        // 
        public static string getFormDetail_TabNotes(CPBaseClass cp, CPCSBaseClass CS) {
            string getFormDetail_TabNotesRet = default;
            string result = "";
            try {
                if (!CS.OK()) {
                    result += "<div>There was a problem retrieving this person's information.</div>";
                } else {
                    // 
                    result += "<table border=0 width=100% cellspacing=0 cellpadding=0><TR><td width=200><img src=/cclib/images/spacer.gif width=150 height=1></TD><td width=99%><img src=/cclib/images/spacer.gif width=1 height=1></TD></TR>";



                    result += getFormDetail_InputHTMLRow(cp, "Notes", "NotesFilename", CS.GetText("NotesFilename"));
                    result += "</table>";
                }
                // 
                result = "<div STYLE=\"width:100%;\" class=\"cmBody ccPanel3DReverse\">" + result + "</div>";
                getFormDetail_TabNotesRet = result;
            } catch (Exception ex) {
                cp.Site.ErrorReport(ex);
                throw;
            }
            return result;
        }
        // 
        // =================================================================================
        // 
        public static string getFormDetail_TabPhoto(CPBaseClass cp, CPCSBaseClass CS) {
            string result = "";
            try {
                if (!CS.OK()) {
                    result += "<div>There was a problem retrieving this person's information.</div>";
                } else {
                    // 
                    result += "<table border=0 width=100% cellspacing=0 cellpadding=0><TR><td width=200><img src=/cclib/images/spacer.gif width=150 height=1></TD><td width=99%><img src=/cclib/images/spacer.gif width=1 height=1></TD></TR>";



                    result += getFormDetail_InputImageRow(cp, "Thumbnail", "ThumbnailFilename", CS.GetText("ThumbnailFilename"));
                    result += getFormDetail_InputImageRow(cp, "Image", "ImageFilename", CS.GetText("ImageFilename"));
                    result += "</table>";
                }
                // 
                result = "<div STYLE=\"width:100%;\" class=\"cmBody ccPanel3DReverse\">" + result + "</div>";
            } catch (Exception ex) {
                cp.Site.ErrorReport(ex);
                throw;
            }
            return result;
        }
        // 
        // =================================================================================
        // 
        public static string getFormDetail_InputTextRow(CPBaseClass cp, string Caption, string FieldName, string DefaultValue, bool IsPassword) {
            string result;

            try {
                result = "<TR><td style=\"TEXT-ALIGN:left;PADDING-LEFT:20px;\">" + Caption + ":</TD><td style=\"TEXT-ALIGN:left;\">";


                if (IsPassword) {
                    result += "<input type=password name=\"" + FieldName + "\" value=\"" + cp.Utils.EncodeHTML(DefaultValue) + "\" style=\"width:300px;\">";
                } else {
                    result += "<input type=text name=\"" + FieldName + "\" value=\"" + cp.Utils.EncodeHTML(DefaultValue) + "\" style=\"width:350px;\">";
                }
                result += "</TD></TR>";
                return result;
            } catch (Exception ex) {
                cp.Site.ErrorReport(ex);
                throw;
            }
        }
        // 
        // =================================================================================
        // 
        public static string getFormDetail_InputBooleanRow(CPBaseClass cp, string Caption, string FieldName, string DefaultValue) {
            try {
                string result = "";
                result = "<TR><td style=\"TEXT-ALIGN:left;PADDING-LEFT:20px;\">" + Caption + ":</TD><td style=\"TEXT-ALIGN:left;\">";


                result += "<input type=checkbox name=\"" + FieldName + "\" value=\"" + DefaultValue + "\">";
                result += "</TD></TR>";
                return result;
            } catch (Exception ex) {
                cp.Site.ErrorReport(ex);
                throw;
            }
        }
        // 
        // =================================================================================
        // 
        public static string getFormDetail_InputHTMLRow(CPBaseClass cp, string Caption, string FieldName, string DefaultValue) {

            try {
                string result = "<TR><td style=\"TEXT-ALIGN:left;PADDING-LEFT:20px;\">" + Caption + ":</TD><td style=\"TEXT-ALIGN:left;\">";


                result += cp.Html.InputWysiwyg(FieldName, DefaultValue, CPHtmlBaseClass.EditorUserScope.Administrator);
                result += "</TD></TR>";
                return result;
            } catch (Exception ex) {
                cp.Site.ErrorReport(ex);
                throw;
            }
        }
        // 
        // =================================================================================
        // 
        public static string getFormDetail_InputImageRow(CPBaseClass cp, string Caption, string FieldName, string DefaultValue) {
            try {
                string result = "<TR><td style=\"TEXT-ALIGN:left;PADDING-LEFT:20px;\">" + Caption + ":</TD><td style=\"TEXT-ALIGN:left;\">";


                if (string.IsNullOrEmpty(DefaultValue)) {
                    result += cp.Html.InputFile(FieldName);
                    result += "</TD></TR>";
                } else {
                    string Filename = cp.Utils.EncodeHTML(DefaultValue);
                    string EncodedLink = cp.Utils.EncodeUrl(cp.Http.CdnFilePathPrefixAbsolute + DefaultValue);
                    result = "<table border=0 width=100% cellspacing=0 cellpadding=4><TR><td width=200><a href=\"" + EncodedLink + "\" target=\"_blank\"><img src=\"" + EncodedLink + "\" width=200 border=0></a></TD><td width=100% valign=top><div style=\"height:20px;\">Filename:&nbsp;" + Filename + "</div><div style=\"height:20px;\">URL:&nbsp;" + EncodedLink + "</div><div style=\"height:20px;\"><a href=\"" + EncodedLink + "\" target=\"_blank\">Click for Full Size</A></div><div style=\"height:20px;\"><span style=\"width:100px;\">Delete:</span>" + cp.Html.CheckBox(FieldName + ".DeleteFlag", false) + "</div><div style=\"height:20px;\"><span style=\"width:100px;\">Change:</span>" + cp.Html.InputFile(FieldName) + "</div></TD></TR></table></TD></TR>";











                }
                return "";
            } catch (Exception ex) {
                cp.Site.ErrorReport(ex);
                throw;
            }
        }
        // 
        // =================================================================================
        // 
        public static string getFormDetail_DividerRow(CPBaseClass cp, string Caption) {
            try {
                string result = Strings.Replace(Caption, " ", "&nbsp;");
                result = "<TR><td colspan=2 style=\"Padding-top:10px;\"><TABLE border=0 width=100% cellspacing=0 cellpadding=0><TR><td width=1 style=\"white-space:nowrap;\">" + result + "&nbsp;&nbsp;</TD><td width=100%><HR></TD></TABLE></tr>";



                return result;
            } catch (Exception ex) {
                cp.Site.ErrorReport(ex);
                throw;
            }
        }
        // 
        // =================================================================================
        // 
        public static string savePersonFromStream(CPBaseClass cp, int personId) {
            string result = "";
            try {
                using (var csPerson = cp.CSNew()) {
                    if (csPerson.Open("people", "id=" + personId)) {
                        csPerson.SetField("name", cp.Doc.GetText("name"));
                        csPerson.SetField("FirstName", cp.Doc.GetText("FirstName"));
                        csPerson.SetField("LastName", cp.Doc.GetText("LastName"));
                        // 
                        // Contact
                        // 
                        csPerson.SetField("email", cp.Doc.GetText("email"));
                        csPerson.SetField("Phone", cp.Doc.GetText("Phone"));
                        csPerson.SetField("Fax", cp.Doc.GetText("Fax"));
                        csPerson.SetField("Address", cp.Doc.GetText("Address"));
                        csPerson.SetField("Address2", cp.Doc.GetText("Address2"));
                        csPerson.SetField("City", cp.Doc.GetText("City"));
                        csPerson.SetField("State", cp.Doc.GetText("State"));
                        csPerson.SetField("Zip", cp.Doc.GetText("Zip"));
                        // 
                        // Company
                        // 
                        csPerson.SetField("Company", cp.Doc.GetText("Company"));
                        csPerson.SetField("Title", cp.Doc.GetText("Title"));
                        // 
                        // Birthday
                        // 
                        csPerson.SetField("BirthdayDay", cp.Doc.GetInteger("BirthdayDay"));
                        csPerson.SetField("BirthdayMonth", cp.Doc.GetInteger("BirthdayMonth"));
                        csPerson.SetField("BirthdayYear", cp.Doc.GetInteger("BirthdayYear"));
                        // 
                        // Notes
                        // 
                        string Copy = cp.Doc.GetText("AppendNotes");
                        if (!string.IsNullOrEmpty(Copy)) {
                            Copy = "<div style=\"margin-top:10px;border-top:1px dashed black;\">Added " + Conversions.ToString(DateTime.Now) + " by " + cp.Content.GetRecordName("people", cp.User.Id) + "</div><div style=\"margin-left:20px;margin-top:5px;\">" + Copy + "</div>";

                        }
                        csPerson.SetField("NotesFilename", cp.Doc.GetText("NotesFilename") + Copy);
                        // 
                        // Photos
                        // 
                        string thumbnailFieldName = "ThumbnailFilename";
                        if (cp.Doc.GetBoolean(thumbnailFieldName + ".DeleteFlag")) {
                            csPerson.SetField(thumbnailFieldName, "");
                        }
                        string originalFilename = cp.Doc.GetText(thumbnailFieldName);

                        string Path;
                        if (!string.IsNullOrEmpty(originalFilename)) {
                            string Filename = csPerson.GetFilename(thumbnailFieldName, originalFilename);
                            Path = Filename;
                            Path = Strings.Replace(Path, "/", @"\");
                            Path = Strings.Replace(Path, @"\" + originalFilename, "");
                            csPerson.SetField(thumbnailFieldName, Filename);
                            // Call CS.SetFile(FieldName, Path)
                        }
                        // 
                        string imageFieldName = "ImageFilename";
                        if (cp.Doc.GetBoolean(imageFieldName + ".DeleteFlag")) {
                            csPerson.SetField(imageFieldName, "");
                        }
                        originalFilename = cp.Doc.GetText(imageFieldName);
                        if (!string.IsNullOrEmpty(originalFilename)) {
                            string Filename = csPerson.GetFilename(imageFieldName, originalFilename);
                            Path = Filename;
                            Path = Strings.Replace(Path, "/", @"\");
                            Path = Strings.Replace(Path, @"\" + originalFilename, "");
                            csPerson.SetField(imageFieldName, Filename);
                            // Call CS.SetFile(FieldName, Path)
                        }
                        // 
                        // Permissions
                        // 
                        csPerson.SetField("Username", cp.Doc.GetText("Username"));
                        csPerson.SetField("Password", cp.Doc.GetText("Password"));
                        csPerson.SetField("AutoLogin", cp.Doc.GetBoolean("AutoLogin"));
                        // 
                        // Groups
                        // 
                        saveMemberRules(cp, csPerson);
                    }
                    csPerson.Close();

                }
            } catch (Exception ex) {
                cp.Site.ErrorReport(ex);
                throw;
            }
            return result;
        }
        // 
        // ========================================================================
        // 
        public static string getFormDetail_TabGroup(CPBaseClass cp, CPCSBaseClass CSMember) {
            try {
                string result = "";

                // 
                // ----- Gather all the SecondaryContent that associates to the PrimaryContent
                int PrimaryContentID = cp.Content.GetID("People");
                int SecondaryContentID = cp.Content.GetID("Groups");
                var sb = new StringBuilder();
                sb.Append(Microsoft.VisualBasic.Constants.vbCrLf + "<!-- GroupRule Table --><table border=0 width=100% cellspacing=0 cellpadding=0><TR><td width=150><img src=/cclib/images/spacer.gif width=150 height=1></TD><td width=99%><img src=/cclib/images/spacer.gif width=1 height=1></TD></TR>");



                int DetailMemberID = CSMember.GetInteger("ID");
                var GroupCount = default(int);
                if (DetailMemberID != 0) {
                    // 
                    // ----- read in the groups that this member has subscribed (exclude new member records)
                    // 
                    int MembershipCount = 0;
                    int[] Membership = Array.Empty<int>();
                    DateTime[] DateExpires = Array.Empty<DateTime>();
                    bool[] active = Array.Empty<bool>();
                    using (var csRules = cp.CSNew()) {
                        string SQL = " SELECT Active,GroupID,DateExpires FROM ccMemberRules WHERE MemberID=" + DetailMemberID;


                        csRules.OpenSQL(SQL);
                        while (csRules.OK()) {
                            int MembershipSize = 0;
                            if (MembershipCount >= MembershipSize) {
                                MembershipSize += 10;
                                Array.Resize(ref Membership, MembershipSize + 1);
                                Array.Resize(ref active, MembershipSize + 1);
                                Array.Resize(ref DateExpires, MembershipSize + 1);
                            }
                            Membership[MembershipCount] = csRules.GetInteger("GroupID");
                            DateExpires[MembershipCount] = csRules.GetDate("DateExpires");
                            active[MembershipCount] = csRules.GetBoolean("Active");
                            MembershipCount += 1;
                            csRules.GoNext();
                        }
                        csRules.Close();
                    }
                    // 
                    // ----- read in all the groups, sorted by ContentName
                    // 
                    using (var CS = cp.CSNew()) {
                        string Sql = " SELECT ccGroups.ID AS ID, ccContent.Name AS SectionName, ccGroups.Caption AS GroupCaption, ccGroups.name AS GroupName, ccGroups.SortOrder FROM ccGroups LEFT JOIN ccContent ON ccGroups.ContentControlID = ccContent.ID Where (((ccGroups.Active) >0) And ((ccContent.Active) >0))";


                        Sql = Sql + " GROUP BY ccGroups.ID, ccContent.Name, ccGroups.Caption, ccGroups.name, ccGroups.SortOrder ORDER BY ccContent.Name, ccGroups.Caption";

                        CS.OpenSQL(Sql);
                        // 
                        // Output all the groups, with the active and dateexpires from those joined
                        // 
                        string SectionName = "";
                        GroupCount = 0;
                        bool CanSeeHiddenGroups = cp.User.IsDeveloper;
                        while (CS.OK()) {
                            string GroupName = CS.GetText("GroupName");
                            if (Strings.Mid(GroupName, 1, 1) != "_" | CanSeeHiddenGroups) {
                                string GroupCaption = CS.GetText("GroupCaption");
                                int GroupID = CS.GetInteger("ID");
                                if (string.IsNullOrEmpty(GroupCaption)) {
                                    GroupCaption = GroupName;
                                    if (string.IsNullOrEmpty(GroupCaption)) {
                                        GroupCaption = "Group&nbsp;" + GroupID;
                                    }
                                }
                                bool GroupActive = false;
                                string DateExpireValue = "";
                                if (MembershipCount != 0) {
                                    int MembershipPointer;
                                    var loopTo = MembershipCount - 1;
                                    for (MembershipPointer = 0; MembershipPointer <= loopTo; MembershipPointer++) {
                                        if (Membership[MembershipPointer] == GroupID) {
                                            GroupActive = active[MembershipPointer];
                                            if (!DateExpires[MembershipPointer].Equals(DateTime.MinValue)) {
                                                DateExpireValue = DateExpires[MembershipPointer].ToString();
                                            }
                                            break;
                                        }
                                    }
                                }
                                string ReportLink;
                                if (GroupID > 0) {
                                    ReportLink = "<a href=\"?af=12&rid=35&recordid=" + GroupID + "\" target=_blank>Group&nbsp;Report</a>";
                                } else {
                                    ReportLink = "&nbsp;";
                                }
                                string Caption;
                                // 
                                if (GroupCount == 0) {
                                    Caption = "Groups";
                                } else {
                                    Caption = "&nbsp;";
                                }
                                sb.Append(cp.Html.Hidden("Memberrules." + GroupCount + ".ID", GroupID.ToString()));
                                GroupCaption = Strings.Replace(GroupCaption, " ", "&nbsp;");
                                sb.Append(Microsoft.VisualBasic.Constants.vbCrLf + "<!-- GroupRule Row --><TR><td style=\"TEXT-ALIGN:left;PADDING-LEFT:20px;border-top:1px solid white;\">" + cp.Html.CheckBox("MemberRules." + GroupCount, GroupActive) + GroupCaption + "</TD><td style=\"TEXT-ALIGN:left;PADDING-LEFT:10px;border-top:1px solid white;\">Expires " + cp.Html.InputText("MemberRules." + GroupCount + ".DateExpires", DateExpireValue, 255) + "</TD></TR>");



                                GroupCount += 1;
                            }
                            CS.GoNext();
                        }
                        CS.Close();
                    }
                }
                if (DetailMemberID == 0) {
                    sb.Append("<TR><td valign=middle align=right><span>Groups</span></TD><td><span>Groups will be available after this record is saved</SPAN></TD><TR>");


                } else if (GroupCount == 0) {
                    sb.Append("<TR><td valign=middle align=right><span>Groups</span></TD><td><span>There are currently no groups defined</SPAN></TD><TR>");


                } else {
                    sb.Append("<input type=\"hidden\" name=\"MemberRules.RowCount\" value=\"" + GroupCount + "\">");
                }
                sb.Append(Microsoft.VisualBasic.Constants.vbCrLf + "<!-- GroupRule Table End --></table>");
                result = "<div STYLE=\"width:100%;\" class=\"cmBody ccPanel3DReverse\">" + sb.ToString() + "</div>";
                return result;
            } catch (Exception ex) {
                cp.Site.ErrorReport(ex);
                throw;
            }
        }
        // 
        // ========================================================================
        // 
        private static void saveMemberRules(CPBaseClass cp, CPCSBaseClass CSMember) {
            try {
                // 
                // --- create MemberRule records for all selected
                int PeopleID = CSMember.GetInteger("ID");
                int GroupCount = cp.Doc.GetInteger("MemberRules.RowCount");
                if (GroupCount > 0) {
                    int GroupPointer;
                    var loopTo = GroupCount - 1;
                    for (GroupPointer = 0; GroupPointer <= loopTo; GroupPointer++) {
                        // 
                        // ----- Read Response
                        // 
                        int GroupID = cp.Doc.GetInteger("MemberRules." + GroupPointer + ".ID");
                        bool RuleNeeded = cp.Doc.GetBoolean("MemberRules." + GroupPointer);
                        var DateExpires = cp.Doc.GetDate("MemberRules." + GroupPointer + ".DateExpires");
                        object DateExpiresVariant;
                        if (DateExpires == DateTime.MinValue) {
                            DateExpiresVariant = DateTime.MinValue;
                        } else {
                            DateExpiresVariant = DateExpires;
                        }
                        // 
                        // ----- Update Record
                        // 
                        var ruleList = DbBaseModel.createList<MemberRuleModel>(cp, "(MemberID=" + PeopleID + ")and(GroupID=" + GroupID + ")");
                        if (ruleList.Count == 0) {
                            // 
                            // No record exists
                            if (RuleNeeded) {
                                // 
                                // No record, Rule needed, add it
                                var newRule = DbBaseModel.addDefault<MemberRuleModel>(cp);
                                newRule.active = true;
                                newRule.memberId = PeopleID;
                                newRule.groupId = GroupID;
                                newRule.dateExpires = DateExpires;
                                newRule.save(cp);
                            }
                        } else {
                            // 
                            // Record exists
                            MemberRuleModel keepRule = null;
                            if (RuleNeeded) {
                                // 
                                // record exists, and it is needed, update the DateExpires if changed
                                keepRule = ruleList.First();
                                if (!keepRule.active | !keepRule.dateExpires.Equals(DateExpires)) {
                                    keepRule.active = true;
                                    keepRule.dateExpires = DateExpires;
                                    keepRule.save(cp);
                                }
                            }
                            foreach (var rule in ruleList) {
                                if (!ReferenceEquals(rule, keepRule))
                                    DbBaseModel.delete<MemberRuleModel>(cp, rule.id);
                            }
                        }
                    }
                }
            } catch (Exception ex) {
                cp.Site.ErrorReport(ex);
                throw;
                throw;
            }
        }
        // '
        // '========================================================================
        // '
        // Public Shared Function getFormDetail_TabHomes(cp As CPBaseClass, CSMember As CPCSBaseClass) As String
        // Dim result As String

        // Try

        // Dim DetailMemberID As Integer

        // Dim ColumnCount As Integer
        // Dim CS As CPCSBaseClass = cp.CSNew()
        // Dim SQL As String
        // Dim SQLOrderBy As String
        // Dim DateRangeID As Integer
        // Dim AllowNonEmail As Boolean
        // Dim RQS As String
        // Dim PageSize As Integer
        // Dim PageNumber As Integer
        // Dim TopCount As Integer
        // Dim RowCnt As Integer
        // Dim DataRowCount As Integer
        // Dim ColumnPtr As Integer
        // Dim ColCaption() As String
        // Dim ColAlign() As String
        // Dim ColWidth() As String
        // Dim Cells(,) As String
        // Dim SQLWhere As String
        // Dim SQLFrom As String
        // Dim SortField As String
        // Dim SortDirection As Integer
        // Dim Propertyid As Integer
        // Dim Copy As String
        // Dim RowClass As String
        // Dim PeopleCID As Integer
        // Dim RecordID As Integer

        // Dim DateCompleted As Date
        // Dim IsGrgOK As Boolean
        // '

        // Const DateRangeToday = 10
        // Const DateRangeYesterday = 20
        // Const DateRangePastWeek = 30
        // Const DateRangePastMonth = 40
        // Const ColumnCnt = 4
        // '
        // IsGrgOK = True
        // Dim Stream As String = ""
        // '
        // If Not IsGrgOK Then
        // '
        // ' The RealEstate datasource is not available
        // '
        // Stream = "<P>This site is not configured to display the Homes Viewed tab. This tab requires the grgRealEstate Datasource.</P>"
        // Else
        // SortField = cp.Doc.GetText("SortField")
        // SortDirection = cp.Doc.GetInteger("SortDirection")
        // If String.IsNullOrEmpty(SortField) Then
        // SortField = "DateAdded"
        // SortDirection = -1
        // End If
        // '
        // RQS = cp.Doc.RefreshQueryString
        // PageSize = cp.Doc.GetInteger(RequestNamePageSize)
        // If PageSize = 0 Then
        // PageSize = 50
        // End If
        // PageNumber = cp.Doc.GetInteger(RequestNamePageNumber)
        // If PageNumber = 0 Then
        // PageNumber = 1
        // End If
        // TopCount = PageNumber * PageSize
        // '
        // ' Setup Headings
        // '
        // ReDim ColCaption(ColumnCnt)
        // ReDim ColAlign(ColumnCnt)
        // ReDim ColWidth(ColumnCnt)
        // ReDim Cells(PageSize, ColumnCnt)
        // '
        // ColCaption(ColumnPtr) = "Date"
        // ColAlign(ColumnPtr) = "center"
        // ColWidth(ColumnPtr) = "150"
        // ColumnPtr += 1
        // '
        // ColCaption(ColumnPtr) = "Emailed"
        // ColAlign(ColumnPtr) = "left"
        // ColWidth(ColumnPtr) = "50"
        // ColumnPtr += 1
        // '
        // ColCaption(ColumnPtr) = "Property"
        // ColAlign(ColumnPtr) = "Left"
        // ColWidth(ColumnPtr) = "300"
        // ColumnPtr += 1
        // '
        // ColCaption(ColumnPtr) = "Search"
        // ColAlign(ColumnPtr) = "left"
        // ColWidth(ColumnPtr) = "100%"
        // ColumnPtr += 1
        // '
        // ' Build Query
        // '
        // SQLFrom = "" _
        // & " From grgPropertyLog L" _
        // & " left join grgPropertySearches S on S.ID=L.PropertySearchID"
        // SQLWhere = " Where" _
        // & "(S.ID is not null)" _
        // & "and(L.MemberID=" & CSMember.GetInteger("ID") & ")" _
        // & ""
        // If Not AllowNonEmail Then
        // SQLWhere &= "and(L.IsEmailSearch is not null)and(L.IsEmailSearch<>0)"
        // End If
        // Dim rightNow As Date = Now()
        // Dim rightNowDate As Date = rightNow.Date
        // Select Case DateRangeID
        // Case DateRangeToday
        // SQLWhere &= "and(L.DateAdded>=" & cp.Db.EncodeSQLDate(rightNowDate) & ")"
        // Case DateRangeYesterday
        // SQLWhere &= "and(L.DateAdded<" & cp.Db.EncodeSQLDate(rightNowDate) & ")and(L.DateAdded>=" & cp.Db.EncodeSQLDate(rightNowDate.AddDays(-1)) & ")"
        // Case DateRangePastWeek
        // SQLWhere &= "and(L.DateAdded>" & cp.Db.EncodeSQLDate(rightNowDate.AddDays(-7)) & ")"
        // Case DateRangePastMonth
        // SQLWhere &= "and(L.DateAdded>" & cp.Db.EncodeSQLDate(rightNowDate.AddDays(-30)) & ")"
        // End Select
        // SQLOrderBy = " ORDER BY L." & SortField
        // If SortDirection <> 0 Then
        // SQLOrderBy &= " Desc"
        // End If
        // If UCase(SortField) <> "PROPERTYSEARCHID" Then
        // SQLOrderBy &= ",L.PropertySearchID"
        // End If
        // '
        // ' Get DataRowCount
        // '
        // SQL = "select count(*) as Cnt " & SQLFrom & SQLWhere
        // CS.OpenSQL(SQL)
        // If CS.OK() Then
        // DataRowCount = CS.GetInteger("cnt")
        // End If
        // Call CS.Close()
        // '
        // ' Get Data
        // '
        // SQL = "select S.*" _
        // & ",L.PropertyID as PropertyID" _
        // & ",L.PropertyName as PropertyName" _
        // & ",L.DateAdded as ViewingDateAdded" _
        // & ",L.IsEmailSearch as ViewingIsEmailSearch" _
        // & ",L.VisitID as ViewingVisitID" _
        // & " " & SQLFrom & SQLWhere & SQLOrderBy

        // If CS.OpenSQL(SQL, "", PageSize, PageNumber) Then
        // '
        // ' No Searchs saved
        // '
        // Stream &= "<tr class=D0><td class=D0 colspan=""" & ColumnCount & """ width="" 100%"">No Auto Agent records were found.</td></tr>"
        // Else
        // '
        // ' List out the AutoAgents
        // '
        // RowCnt = 0
        // RowClass = ""
        // PeopleCID = cp.Content.GetID("People")
        // If Not CS.OK() Then
        // '
        // ' EMPTY
        // '
        // Else
        // Do While CS.OK() And (RowCnt < PageSize)
        // RecordID = CS.GetInteger("ID")
        // DateCompleted = CS.GetDate("DateCompleted")
        // Cells(RowCnt, 0) = CS.GetDate("DateAdded") & cp.Html.Hidden("RowID" & RowCnt, RecordID.ToString())
        // Cells(RowCnt, 1) = CS.GetText("name")
        // Cells(RowCnt, 2) = If(CS.GetBoolean("ViewingIsEmailSearch"), "Yes", "No")
        // Propertyid = CS.GetInteger("PropertyID")
        // Copy = CS.GetText("PropertyName")
        // If String.IsNullOrEmpty(Copy) Then
        // Copy = "&nbsp;"
        // Else
        // Copy = "<a href=/index.asp?grg_pid=" & Propertyid & " target=_blank>" & Copy & "</a>"
        // End If
        // Cells(RowCnt, 4) = Copy
        // RowCnt += 1
        // Call CS.GoNext()
        // Loop
        // End If
        // Stream &= cp.Html.Hidden("RowCount", RowCnt.ToString())
        // End If
        // Call CS.Close()
        // '
        // DetailMemberID = CSMember.GetInteger("ID")
        // Dim PreTableCopy As String = ""
        // Dim PostTableCopy As String = ""
        // Stream = AdminUIController.getReport(cp, RowCnt, ColCaption, ColAlign, ColWidth, Cells, PageSize, PageNumber, PreTableCopy, PostTableCopy, DataRowCount, "")
        // End If
        // result = "<div STYLE=""width:100%;"" class=""cmBody ccPanel3DReverse"">" & Stream & "</div>"
        // Catch ex As Exception
        // cp.Site.ErrorReport(ex)
        // Throw
        // End Try
        // Return result
        // End Function
        // 


    }
}