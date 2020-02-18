
Option Strict On
Option Explicit On

Imports Contensive.BaseClasses
Imports Contensive.Models.Db
Imports System.Text
Imports System.Linq
Imports Contensive.Addons.ContactManagerTools

Namespace Views
    Public Class DetailFormController
        '
        '=================================================================================
        '
        Public Shared Function ProcessRequest(cp As CPBaseClass, ae As Controllers.ApplicationController, request As Views.CMngrClass.RequestClass) As FormIdEnum
            request.DetailMemberID = cp.Doc.GetInteger(RequestNameMemberID)
            Select Case request.Button
                Case ButtonCancel
                    Return FormIdEnum.FormList
                Case ButtonSave
                    Call savePersonFromStream(cp, request.DetailMemberID)
                    Return FormIdEnum.FormDetails
                Case ButtonOK
                    Call savePersonFromStream(cp, request.DetailMemberID)
                    Return FormIdEnum.FormList
                Case ButtonNewSearch
                    Return FormIdEnum.FormSearch
            End Select
            Return FormIdEnum.FormDetails
        End Function
        '
        '=================================================================================
        '
        Public Shared Function getEventsTab() As String
            Return "events"
        End Function
        '
        '=================================================================================
        '
        Public Shared Function getResponse(cp As CPBaseClass, ae As Controllers.ApplicationController, DetailMemberID As Integer) As String
            Dim result As String = ""
            Try
                Using csPerson As CPCSBaseClass = cp.CSNew()
                    csPerson.Open("people", "ID=" & DetailMemberID, "", False)
                    Dim memberName As String = csPerson.GetText("name")
                    If memberName = "" Then
                        memberName = Trim(csPerson.GetText("FirstName") & " " & csPerson.GetText("LastName"))
                        If memberName = "" Then
                            memberName = "Record " & csPerson.GetText("ID")
                        End If
                    End If
                    '
                    ' Determine current Subtab
                    '
                    Dim SubTab As Integer = cp.Doc.GetInteger(RequestNameDetailSubtab)
                    If SubTab = 0 Then
                        SubTab = ae.userProperties.subTab
                        If SubTab = 0 Then
                            SubTab = 1
                            ae.userProperties.subTab = SubTab
                        End If
                    Else
                        ae.userProperties.subTab = SubTab
                    End If
                    Call cp.Doc.AddRefreshQueryString(RequestNameDetailSubtab, SubTab.ToString())
                    '
                    ' SubTab Menu
                    '
                    Call cp.Doc.AddRefreshQueryString("tab", "")
                    Dim ButtonList As String = ButtonCancel & "," & ButtonSave & "," & ButtonOK & "," & ButtonNewSearch
                    Dim Header As String = "<div>" & memberName & "</div>"
                    '
                    Dim Nav As New TabController()
                    Call Nav.addEntry("Contact", getFormDetail_TabContact(cp, csPerson), "ccAdminTab")
                    Call Nav.addEntry("Permissions", GetFormDetail_TabPermissions(cp, csPerson), "ccAdminTab")
                    Call Nav.addEntry("Notes", GetFormDetail_TabNotes(cp, csPerson), "ccAdminTab")
                    Call Nav.addEntry("Photos", GetFormDetail_TabPhoto(cp, csPerson), "ccAdminTab")
                    Call Nav.addEntry("Groups", getFormDetail_TabGroup(cp, csPerson), "ccAdminTab")
                    Call csPerson.Close()
                    '
                    Dim Content As String = Nav.getTabs(cp)
                    Content &= cp.Html.Hidden(RequestNameFormID, Convert.ToInt32(FormIdEnum.FormDetails).ToString())
                    Content &= cp.Html.Hidden(RequestNameMemberID, DetailMemberID.ToString())
                    '
                    result = AdminUIController.getBody(cp, "Contact Manager &gt;&gt; Contact Details", ButtonList, "", True, True, Header, "", 0, Content)
                End Using
            Catch ex As Exception
                cp.Site.ErrorReport(ex)
            End Try
            Return result
        End Function
        '
        '=================================================================================
        '
        Public Shared Function getFormDetail_TabContact(cp As CPBaseClass, csPerson As CPCSBaseClass) As String
            Dim result As String = ""
            Try
                Dim left As String = ""
                Dim right As String = ""
                '
                If Not csPerson.OK() Then
                    '
                    result &= "<div>There was a problem retrieving this person's information.</div>"
                Else
                    '
                    ' Left Side
                    '
                    left &= "<table border=0 width=100% cellspacing=0 cellpadding=0>" _
                    & "<TR>" _
                    & "<td width=150><img src=/cclib/images/spacer.gif width=150 height=1></TD>" _
                    & "<td width=350><img src=/cclib/images/spacer.gif width=350 height=1></TD>" _
                    & "</TR>"
                    left &= GetFormDetail_InputTextRow(cp, "Full Name", "Name", csPerson.GetText("name"), False)
                    left &= GetFormDetail_InputTextRow(cp, "First Name", "FirstName", csPerson.GetText("FirstName"), False)
                    left &= GetFormDetail_InputTextRow(cp, "Last Name", "LastName", csPerson.GetText("LastName"), False)
                    left &= GetFormDetail_DividerRow(cp, "Contact")
                    left &= GetFormDetail_InputTextRow(cp, "Email", "EMAIL", csPerson.GetText("email"), False)
                    left &= GetFormDetail_InputTextRow(cp, "Phone", "PHONE", csPerson.GetText("PHONE"), False)
                    left &= GetFormDetail_InputTextRow(cp, "Fax", "Fax", csPerson.GetText("Fax"), False)
                    left &= GetFormDetail_InputTextRow(cp, "Address", "ADDRESS", csPerson.GetText("ADDRESS"), False)
                    left &= GetFormDetail_InputTextRow(cp, "Line 2", "ADDRESS2", csPerson.GetText("ADDRESS2"), False)
                    left &= GetFormDetail_InputTextRow(cp, "City", "City", csPerson.GetText("City"), False)
                    left &= GetFormDetail_InputTextRow(cp, "State", "State", csPerson.GetText("State"), False)
                    left &= GetFormDetail_InputTextRow(cp, "Zip", "Zip", csPerson.GetText("Zip"), False)
                    left &= GetFormDetail_DividerRow(cp, "Company")
                    left &= GetFormDetail_InputTextRow(cp, "Name", "Company", csPerson.GetText("Company"), False)
                    left &= GetFormDetail_InputTextRow(cp, "Title", "Title", csPerson.GetText("Title"), False)
                    left &= GetFormDetail_DividerRow(cp, "Birthday")
                    left &= GetFormDetail_InputTextRow(cp, "Day", "BirthdayDay", csPerson.GetText("BirthdayDay"), False)
                    left &= GetFormDetail_InputTextRow(cp, "Month", "BirthdayMonth", csPerson.GetText("BirthdayMonth"), False)
                    left &= GetFormDetail_InputTextRow(cp, "Year", "BirthdayYear", csPerson.GetText("BirthdayYear"), False)
                    left &= "</table>"
                    '
                    ' Right Side
                    '
                    Dim copy As String = cp.Html.InputText("AppendNotes", "", 255)
                    copy = Replace(copy, " cols=""100""", " style=""width:100%;""", , , vbTextCompare)
                    '
                    right &= "<table border=0 width=100% cellspacing=0 cellpadding=0>" _
                        & "<TR>" _
                        & "<td width=200><img src=/cclib/images/spacer.gif width=150 height=1></TD>" _
                        & "<td width=99%><img src=/cclib/images/spacer.gif width=1 height=1></TD>" _
                        & "</TR>"
                    right &= GetFormDetail_DividerRow(cp, "Add to Notes")
                    right &= "<TR><td colspan=2>" & copy & "</TD></TR>"
                    right &= "</table>"
                End If
                '
                result &= "" _
                    & "<table border=0 width=100% cellspacing=0 cellpadding=0>" _
                    & "<TR>" _
                    & "<td width=500 valign=top style=""border-right:1px solid #808080;padding-right:20px;"">" & left & "</TD>" _
                    & "<td width=100% valign=top style=""border-left:1px solid #f0f0f0;padding-left:20px;"">" & right & "</TD>" _
                    & "</TR>" _
                    & "</table>"
                Return "<div STYLE=""width:100%;"" class=""cmBody ccPanel3DReverse"">" & result & "</div>"
            Catch ex As Exception
                cp.Site.ErrorReport(ex)
                Return String.Empty
            End Try
        End Function
        '
        '=================================================================================
        '
        Public Shared Function GetFormDetail_TabPermissions(cp As CPBaseClass, CS As CPCSBaseClass) As String
            Dim result As String = ""
            Try
                If Not CS.OK() Then
                    '
                    result &= "<div>There was a problem retrieving this person's information.</div>"
                Else
                    '
                    result &= "<table border=0 width=100% cellspacing=0 cellpadding=0>" _
                        & "<TR>" _
                        & "<td width=200><img src=/cclib/images/spacer.gif width=150 height=1></TD>" _
                        & "<td width=99%><img src=/cclib/images/spacer.gif width=1 height=1></TD>" _
                        & "</TR>"
                    result &= GetFormDetail_DividerRow(cp, "Login")
                    result &= GetFormDetail_InputTextRow(cp, "Username", "Username", CS.GetText("Username"), False)
                    result &= GetFormDetail_InputTextRow(cp, "Password", "Password", CS.GetText("Password"), True)
                    result &= GetFormDetail_InputBooleanRow(cp, "Allow Auto Login", "AutoLogin", CS.GetBoolean("AutoLogin").ToString())
                    result &= "</table>"
                End If
                result = "<div STYLE=""width:100%;"" class=""cmBody ccPanel3DReverse"">" & result & "</div>"
            Catch ex As Exception
                cp.Site.ErrorReport(ex)
            End Try
            Return result
        End Function
        '
        '=================================================================================
        '
        Public Shared Function GetFormDetail_TabNotes(cp As CPBaseClass, CS As CPCSBaseClass) As String
            Dim result As String = ""
            Try
                If Not CS.OK() Then
                    result &= "<div>There was a problem retrieving this person's information.</div>"
                Else
                    '
                    result &= "<table border=0 width=100% cellspacing=0 cellpadding=0>" _
                        & "<TR>" _
                        & "<td width=200><img src=/cclib/images/spacer.gif width=150 height=1></TD>" _
                        & "<td width=99%><img src=/cclib/images/spacer.gif width=1 height=1></TD>" _
                        & "</TR>"
                    result &= GetFormDetail_InputHTMLRow(cp, "Notes", "NotesFilename", CS.GetText("NotesFilename"))
                    result &= "</table>"
                End If
                '
                result = "<div STYLE=""width:100%;"" class=""cmBody ccPanel3DReverse"">" & result & "</div>"
                GetFormDetail_TabNotes = result
            Catch ex As Exception
                cp.Site.ErrorReport(ex)
            End Try
            Return result
        End Function
        '
        '=================================================================================
        '
        Public Shared Function GetFormDetail_TabPhoto(cp As CPBaseClass, CS As CPCSBaseClass) As String
            Dim result As String = ""
            Try
                If Not CS.OK() Then
                    result &= "<div>There was a problem retrieving this person's information.</div>"
                Else
                    '
                    result &= "<table border=0 width=100% cellspacing=0 cellpadding=0>" _
                        & "<TR>" _
                        & "<td width=200><img src=/cclib/images/spacer.gif width=150 height=1></TD>" _
                        & "<td width=99%><img src=/cclib/images/spacer.gif width=1 height=1></TD>" _
                        & "</TR>"
                    result &= GetFormDetail_InputImageRow(cp, "Thumbnail", "ThumbnailFilename", CS.GetText("ThumbnailFilename"))
                    result &= GetFormDetail_InputImageRow(cp, "Image", "ImageFilename", CS.GetText("ImageFilename"))
                    result &= "</table>"
                End If
                '
                result = "<div STYLE=""width:100%;"" class=""cmBody ccPanel3DReverse"">" & result & "</div>"
            Catch ex As Exception
                cp.Site.ErrorReport(ex)
            End Try
            Return result
        End Function
        '
        '=================================================================================
        '
        Public Shared Function GetFormDetail_InputTextRow(cp As CPBaseClass, Caption As String, FieldName As String, DefaultValue As String, IsPassword As Boolean) As String
            Dim result As String = ""
            Try
                result = "" _
                    & "<TR><td style=""TEXT-ALIGN:left;PADDING-LEFT:20px;"">" _
                    & Caption & ":" _
                    & "</TD><td style=""TEXT-ALIGN:left;"">"
                If IsPassword Then
                    result &= "<input type=password name=""" & FieldName & """ value=""" & cp.Utils.EncodeHTML(DefaultValue) & """ style=""width:300px;"">"
                Else
                    result &= "<input type=text name=""" & FieldName & """ value=""" & cp.Utils.EncodeHTML(DefaultValue) & """ style=""width:350px;"">"
                End If
                result &= "</TD></TR>"
            Catch ex As Exception
                cp.Site.ErrorReport(ex)
            End Try
            Return result
        End Function
        '
        '=================================================================================
        '
        Public Shared Function GetFormDetail_InputBooleanRow(cp As CPBaseClass, Caption As String, FieldName As String, DefaultValue As String) As String
            Dim result As String = ""
            Try
                result = "" _
                    & "<TR><td style=""TEXT-ALIGN:left;PADDING-LEFT:20px;"">" _
                    & Caption & ":" _
                    & "</TD><td style=""TEXT-ALIGN:left;"">"
                result &= "<input type=checkbox name=""" & FieldName & """ value=""" & DefaultValue & """>"
                result &= "</TD></TR>"
            Catch ex As Exception
                cp.Site.ErrorReport(ex)
            End Try
            Return result
        End Function
        '
        '=================================================================================
        '
        Public Shared Function GetFormDetail_InputHTMLRow(cp As CPBaseClass, Caption As String, FieldName As String, DefaultValue As String) As String
            Dim result As String = ""
            Try
                result = "" _
                    & "<TR><td style=""TEXT-ALIGN:left;PADDING-LEFT:20px;"">" _
                    & Caption & ":" _
                    & "</TD><td style=""TEXT-ALIGN:left;"">"
                result &= cp.Html.InputWysiwyg(FieldName, DefaultValue, CPHtmlBaseClass.EditorUserScope.Administrator)
                result &= "</TD></TR>"
            Catch ex As Exception
                cp.Site.ErrorReport(ex)
            End Try
            Return result
        End Function
        '
        '=================================================================================
        '
        Public Shared Function GetFormDetail_InputImageRow(cp As CPBaseClass, Caption As String, FieldName As String, DefaultValue As String) As String
            Dim result As String = ""
            Try
                Dim EncodedLink As String
                Dim Filename As String
                '
                result = "" _
                    & "<TR><td style=""TEXT-ALIGN:left;PADDING-LEFT:20px;"">" _
                    & Caption & ":" _
                    & "</TD><td style=""TEXT-ALIGN:left;"">"
                If DefaultValue = "" Then
                    result &= cp.Html.InputFile(FieldName)
                    result &= "</TD></TR>"
                Else
                    Filename = cp.Utils.EncodeHTML(DefaultValue)
                    EncodedLink = cp.Utils.EncodeUrl("http://" & cp.Request.Host & cp.Site.FilePath & DefaultValue)
                    result = result _
                        & "<table border=0 width=100% cellspacing=0 cellpadding=4><TR>" _
                        & "<td width=200><a href=""" & EncodedLink & """ target=""_blank""><img src=""" & EncodedLink & """ width=200 border=0></a></TD>" _
                        & "<td width=100% valign=top>" _
                        & "<div style=""height:20px;"">Filename:&nbsp;" & Filename & "</div>" _
                        & "<div style=""height:20px;"">URL:&nbsp;" & EncodedLink & "</div>" _
                        & "<div style=""height:20px;""><a href=""" & EncodedLink & """ target=""_blank"">Click for Full Size</A></div>" _
                        & "<div style=""height:20px;""><span style=""width:100px;"">Delete:</span>" & cp.Html.CheckBox(FieldName & ".DeleteFlag", False) & "</div>" _
                        & "<div style=""height:20px;""><span style=""width:100px;"">Change:</span>" & cp.Html.InputFile(FieldName) & "</div>" _
                        & "</TD></TR>" _
                        & "</table>" _
                        & "" _
                        & "</TD></TR>"
                End If
            Catch ex As Exception
                cp.Site.ErrorReport(ex)
            End Try
            Return result
        End Function
        '
        '=================================================================================
        '
        Public Shared Function GetFormDetail_DividerRow(cp As CPBaseClass, Caption As String) As String
            Dim result As String = ""
            Try
                result = Replace(Caption, " ", "&nbsp;")
                result = "<TR><td colspan=2 style=""Padding-top:10px;"">" _
                    & "<TABLE border=0 width=100% cellspacing=0 cellpadding=0>" _
                    & "<TR><td width=1 style=""white-space:nowrap;"">" & result & "&nbsp;&nbsp;</TD><td width=100%><HR></TD>" _
                    & "</TABLE>" _
                    & "</tr>"
            Catch ex As Exception
                cp.Site.ErrorReport(ex)
            End Try
            Return result
        End Function
        '
        '=================================================================================
        '
        Public Shared Function savePersonFromStream(cp As CPBaseClass, personId As Integer) As String
            Dim result As String = ""
            Try
                Using csPerson As CPCSBaseClass = cp.CSNew()
                    If csPerson.Open("people", "id=" & personId) Then
                        Call csPerson.SetField("name", cp.Doc.GetText("name"))
                        Call csPerson.SetField("FirstName", cp.Doc.GetText("FirstName"))
                        Call csPerson.SetField("LastName", cp.Doc.GetText("LastName"))
                        '
                        ' Contact
                        '
                        Call csPerson.SetField("email", cp.Doc.GetText("email"))
                        Call csPerson.SetField("Phone", cp.Doc.GetText("Phone"))
                        Call csPerson.SetField("Fax", cp.Doc.GetText("Fax"))
                        Call csPerson.SetField("Address", cp.Doc.GetText("Address"))
                        Call csPerson.SetField("Address2", cp.Doc.GetText("Address2"))
                        Call csPerson.SetField("City", cp.Doc.GetText("City"))
                        Call csPerson.SetField("State", cp.Doc.GetText("State"))
                        Call csPerson.SetField("Zip", cp.Doc.GetText("Zip"))
                        '
                        ' Company
                        '
                        Call csPerson.SetField("Company", cp.Doc.GetText("Company"))
                        Call csPerson.SetField("Title", cp.Doc.GetText("Title"))
                        '
                        ' Birthday
                        '
                        Call csPerson.SetField("BirthdayDay", cp.Doc.GetInteger("BirthdayDay"))
                        Call csPerson.SetField("BirthdayMonth", cp.Doc.GetInteger("BirthdayMonth"))
                        Call csPerson.SetField("BirthdayYear", cp.Doc.GetInteger("BirthdayYear"))
                        '
                        ' Notes
                        '
                        Dim Copy As String = cp.Doc.GetText("AppendNotes")
                        If Copy <> "" Then
                            Copy = "" _
                                & "<div style=""margin-top:10px;border-top:1px dashed black;"">Added " & Now & " by " & cp.Content.GetRecordName("people", cp.User.Id) & "</div>" _
                                & "<div style=""margin-left:20px;margin-top:5px;"">" & Copy & "</div>"
                        End If
                        Call csPerson.SetField("NotesFilename", cp.Doc.GetText("NotesFilename") & Copy)
                        '
                        ' Photos
                        '
                        Dim thumbnailFieldName As String = "ThumbnailFilename"
                        If cp.Doc.GetBoolean(thumbnailFieldName & ".DeleteFlag") Then
                            Call csPerson.SetField(thumbnailFieldName, "")
                        End If
                        Dim originalFilename As String = cp.Doc.GetText(thumbnailFieldName)

                        Dim Path As String
                        If originalFilename <> "" Then
                            Dim Filename As String = csPerson.GetFilename(thumbnailFieldName, originalFilename)
                            Path = Filename
                            Path = Replace(Path, "/", "\")
                            Path = Replace(Path, "\" & originalFilename, "")
                            Call csPerson.SetField(thumbnailFieldName, Filename)
                            'Call CS.SetFile(FieldName, Path)
                        End If
                        '
                        Dim imageFieldName As String = "ImageFilename"
                        If cp.Doc.GetBoolean(imageFieldName & ".DeleteFlag") Then
                            Call csPerson.SetField(imageFieldName, "")
                        End If
                        originalFilename = cp.Doc.GetText(imageFieldName)
                        If originalFilename <> "" Then
                            Dim Filename As String = csPerson.GetFilename(imageFieldName, originalFilename)
                            Path = Filename
                            Path = Replace(Path, "/", "\")
                            Path = Replace(Path, "\" & originalFilename, "")
                            Call csPerson.SetField(imageFieldName, Filename)
                            'Call CS.SetFile(FieldName, Path)
                        End If
                        '
                        ' Permissions
                        '
                        Call csPerson.SetField("Username", cp.Doc.GetText("Username"))
                        Call csPerson.SetField("Password", cp.Doc.GetText("Password"))
                        Call csPerson.SetField("AutoLogin", cp.Doc.GetBoolean("AutoLogin"))
                        '
                        ' Groups
                        '
                        Call saveMemberRules(cp, csPerson)
                    End If
                    Call csPerson.Close()

                End Using
            Catch ex As Exception
                cp.Site.ErrorReport(ex)
            End Try
            Return result
        End Function
        '
        '========================================================================
        '
        Public Shared Function getFormDetail_TabGroup(cp As CPBaseClass, CSMember As CPCSBaseClass) As String
            Dim result As String = ""
            Try

                '
                ' ----- Gather all the SecondaryContent that associates to the PrimaryContent
                Dim PrimaryContentID As Integer = cp.Content.GetID("People")
                Dim SecondaryContentID As Integer = cp.Content.GetID("Groups")
                Dim sb As New StringBuilder()
                sb.Append(vbCrLf & "<!-- GroupRule Table --><table border=0 width=100% cellspacing=0 cellpadding=0>" _
                        & "<TR>" _
                        & "<td width=150><img src=/cclib/images/spacer.gif width=150 height=1></TD>" _
                        & "<td width=99%><img src=/cclib/images/spacer.gif width=1 height=1></TD>" _
                        & "</TR>")
                Dim DetailMemberID As Integer = CSMember.GetInteger("ID")
                Dim GroupCount As Integer
                If DetailMemberID <> 0 Then
                    '
                    ' ----- read in the groups that this member has subscribed (exclude new member records)
                    '
                    Dim MembershipCount As Integer = 0
                    Dim Membership() As Integer = Array.Empty(Of Integer)
                    Dim DateExpires() As Date = Array.Empty(Of Date)
                    Dim active() As Boolean = Array.Empty(Of Boolean)
                    Using csRules As CPCSBaseClass = cp.CSNew()
                        Dim SQL As String = "" _
                                & " SELECT Active,GroupID,DateExpires" _
                                & " FROM ccMemberRules" _
                                & " WHERE MemberID=" & DetailMemberID
                        csRules.OpenSQL(SQL)
                        Do While csRules.OK()
                            Dim MembershipSize As Integer = 0
                            If MembershipCount >= MembershipSize Then
                                MembershipSize += 10
                                ReDim Preserve Membership(MembershipSize)
                                ReDim Preserve active(MembershipSize)
                                ReDim Preserve DateExpires(MembershipSize)
                            End If
                            Membership(MembershipCount) = csRules.GetInteger("GroupID")
                            DateExpires(MembershipCount) = csRules.GetDate("DateExpires")
                            active(MembershipCount) = csRules.GetBoolean("Active")
                            MembershipCount += 1
                            csRules.GoNext()
                        Loop
                        Call csRules.Close()
                    End Using
                    '
                    ' ----- read in all the groups, sorted by ContentName
                    '
                    Using CS As CPCSBaseClass = cp.CSNew()
                        Dim Sql As String = "" _
                                & " SELECT ccGroups.ID AS ID, ccContent.Name AS SectionName, ccGroups.Caption AS GroupCaption, ccGroups.name AS GroupName, ccGroups.SortOrder" _
                                & " FROM ccGroups LEFT JOIN ccContent ON ccGroups.ContentControlID = ccContent.ID" _
                                & " Where (((ccGroups.Active) >0) And ((ccContent.Active) >0))"
                        Sql = Sql _
                                & " GROUP BY ccGroups.ID, ccContent.Name, ccGroups.Caption, ccGroups.name, ccGroups.SortOrder" _
                                & " ORDER BY ccContent.Name, ccGroups.Caption"
                        CS.OpenSQL(Sql)
                        '
                        ' Output all the groups, with the active and dateexpires from those joined
                        '
                        Dim SectionName As String = ""
                        GroupCount = 0
                        Dim CanSeeHiddenGroups As Boolean = cp.User.IsDeveloper
                        Do While CS.OK()
                            Dim GroupName As String = CS.GetText("GroupName")
                            If (Mid(GroupName, 1, 1) <> "_") Or CanSeeHiddenGroups Then
                                Dim GroupCaption As String = CS.GetText("GroupCaption")
                                Dim GroupID As Integer = CS.GetInteger("ID")
                                If GroupCaption = "" Then
                                    GroupCaption = GroupName
                                    If GroupCaption = "" Then
                                        GroupCaption = "Group&nbsp;" & GroupID
                                    End If
                                End If
                                Dim GroupActive As Boolean = False
                                Dim DateExpireValue As String = ""
                                If MembershipCount <> 0 Then
                                    Dim MembershipPointer As Integer
                                    For MembershipPointer = 0 To MembershipCount - 1
                                        If Membership(MembershipPointer) = GroupID Then
                                            GroupActive = active(MembershipPointer)
                                            If (Not DateExpires(MembershipPointer).Equals(Date.MinValue)) Then
                                                DateExpireValue = DateExpires(MembershipPointer).ToString()
                                            End If
                                            Exit For
                                        End If
                                    Next
                                End If
                                Dim ReportLink As String
                                If GroupID > 0 Then
                                    ReportLink = "<a href=""?af=12&rid=35&recordid=" & GroupID & """ target=_blank>Group&nbsp;Report</a>"
                                Else
                                    ReportLink = "&nbsp;"
                                End If
                                Dim Caption As String
                                '
                                If GroupCount = 0 Then
                                    Caption = "Groups"
                                Else
                                    Caption = "&nbsp;"
                                End If
                                sb.Append(cp.Html.Hidden("Memberrules." & GroupCount & ".ID", GroupID.ToString()))
                                GroupCaption = Replace(GroupCaption, " ", "&nbsp;")
                                sb.Append(vbCrLf & "<!-- GroupRule Row -->" _
                                        & "<TR>" _
                                        & "<td style=""TEXT-ALIGN:left;PADDING-LEFT:20px;border-top:1px solid white;"">" & cp.Html.CheckBox("MemberRules." & GroupCount, GroupActive) & GroupCaption & "</TD>" _
                                        & "<td style=""TEXT-ALIGN:left;PADDING-LEFT:10px;border-top:1px solid white;"">Expires " & cp.Html.InputText("MemberRules." & GroupCount & ".DateExpires", DateExpireValue, 255) & "</TD>" _
                                        & "</TR>")
                                GroupCount = GroupCount + 1
                            End If
                            CS.GoNext()
                        Loop
                        CS.Close()
                    End Using
                End If
                If DetailMemberID = 0 Then
                    sb.Append("<TR>" _
                            & "<td valign=middle align=right><span>Groups</span></TD>" _
                            & "<td><span>Groups will be available after this record is saved</SPAN></TD>" _
                            & "<TR>")
                ElseIf GroupCount = 0 Then
                    sb.Append("<TR>" _
                            & "<td valign=middle align=right><span>Groups</span></TD>" _
                            & "<td><span>There are currently no groups defined</SPAN></TD>" _
                            & "<TR>")
                Else
                    sb.Append("<input type=""hidden"" name=""MemberRules.RowCount"" value=""" & GroupCount & """>")
                End If
                sb.Append(vbCrLf & "<!-- GroupRule Table End --></table>")
                result = "<div STYLE=""width:100%;"" class=""cmBody ccPanel3DReverse"">" & sb.ToString() & "</div>"
            Catch ex As Exception
                cp.Site.ErrorReport(ex)
            End Try
            Return result
        End Function
        '
        '========================================================================
        '
        Private Shared Sub saveMemberRules(cp As CPBaseClass, CSMember As CPCSBaseClass)
            Try
                '
                ' --- create MemberRule records for all selected
                Dim PeopleID As Integer = CSMember.GetInteger("ID")
                Dim GroupCount As Integer = cp.Doc.GetInteger("MemberRules" & ".RowCount")
                If GroupCount > 0 Then
                    Dim GroupPointer As Integer
                    For GroupPointer = 0 To GroupCount - 1
                        '
                        ' ----- Read Response
                        '
                        Dim GroupID As Integer = cp.Doc.GetInteger("MemberRules" & "." & GroupPointer & "." & "ID")
                        Dim RuleNeeded As Boolean = cp.Doc.GetBoolean("MemberRules" & "." & GroupPointer)
                        Dim DateExpires As Date = cp.Doc.GetDate("MemberRules" & "." & GroupPointer & "." & "DateExpires")
                        Dim DateExpiresVariant As Object
                        If DateExpires = Date.MinValue Then
                            DateExpiresVariant = Date.MinValue
                        Else
                            DateExpiresVariant = DateExpires
                        End If
                        '
                        ' ----- Update Record
                        '
                        Dim ruleList As List(Of MemberRuleModel) = MemberRuleModel.createList(Of MemberRuleModel)(cp, "(MemberID=" & PeopleID & ")and(GroupID=" & GroupID & ")")
                        If ruleList.Count = 0 Then
                            '
                            ' No record exists
                            If RuleNeeded Then
                                '
                                ' No record, Rule needed, add it
                                Dim newRule As MemberRuleModel = MemberRuleModel.addDefault(Of MemberRuleModel)(cp)
                                newRule.active = True
                                newRule.memberId = PeopleID
                                newRule.groupId = GroupID
                                newRule.dateExpires = DateExpires
                                newRule.save(cp)
                            End If
                        Else
                            '
                            ' Record exists
                            Dim keepRule As MemberRuleModel = Nothing
                            If RuleNeeded Then
                                '
                                ' record exists, and it is needed, update the DateExpires if changed
                                keepRule = ruleList.First()
                                If (Not keepRule.active) Or (Not keepRule.dateExpires.Equals(DateExpires)) Then
                                    keepRule.active = True
                                    keepRule.dateExpires = DateExpires
                                    keepRule.save(cp)
                                End If
                            End If
                            For Each rule In ruleList
                                If (rule IsNot keepRule) Then MemberRuleModel.delete(Of MemberRuleModel)(cp, rule.id)
                            Next
                        End If
                    Next
                End If
            Catch ex As Exception
                cp.Site.ErrorReport(ex)
            End Try
        End Sub
        '
        '========================================================================
        '
        Public Shared Function GetFormDetail_TabHomes(cp As CPBaseClass, CSMember As CPCSBaseClass) As String
            Dim result As String = ""
            Try
                Dim Cell As String
                Dim DetailMemberID As Integer
                Dim RowClassOdd As String
                Dim RowClassEven As String
                Dim ColumnCount As Integer
                Dim CS As CPCSBaseClass = cp.CSNew()
                Dim SQL As String
                Dim SQLOrderBy As String
                Dim DateRangeID As Integer
                Dim AllowNonEmail As Boolean
                Dim RQS As String
                Dim PageSize As Integer
                Dim PageNumber As Integer
                Dim TopCount As Integer
                Dim RowCnt As Integer
                Dim DataRowCount As Integer
                Dim PreTableCopy As String
                Dim PostTableCopy As String
                Dim ColumnPtr As Integer
                Dim ColCaption() As String
                Dim ColAlign() As String
                Dim ColWidth() As String
                Dim Cells(,) As String
                Dim Stream As String
                Dim SQLWhere As String
                Dim SQLFrom As String
                Dim SortField As String
                Dim SortDirection As Integer
                Dim Propertyid As Integer
                Dim Copy As String
                Dim RowClass As String
                Dim PeopleCID As Integer
                Dim RecordID As Integer
                Dim RecordName As String
                Dim DateCompleted As Date
                Dim IsGrgOK As Boolean
                '
                Const DateRangeAll = 0
                Const DateRangeToday = 10
                Const DateRangeYesterday = 20
                Const DateRangePastWeek = 30
                Const DateRangePastMonth = 40
                Const ColumnCnt = 4
                '
                IsGrgOK = True
                '
                If Not IsGrgOK Then
                    '
                    ' The RealEstate datasource is not available
                    '
                    Stream = "<P>This site is not configured to display the Homes Viewed tab. This tab requires the grgRealEstate Datasource.</P>"
                Else
                    SortField = cp.Doc.GetText("SortField")
                    SortDirection = cp.Doc.GetInteger("SortDirection")
                    If SortField = "" Then
                        SortField = "DateAdded"
                        SortDirection = -1
                    End If
                    '
                    RQS = cp.Doc.RefreshQueryString
                    PageSize = cp.Doc.GetInteger(RequestNamePageSize)
                    If PageSize = 0 Then
                        PageSize = 50
                    End If
                    PageNumber = cp.Doc.GetInteger(RequestNamePageNumber)
                    If PageNumber = 0 Then
                        PageNumber = 1
                    End If
                    TopCount = PageNumber * PageSize
                    '
                    ' Setup Headings
                    '
                    ReDim ColCaption(ColumnCnt)
                    ReDim ColAlign(ColumnCnt)
                    ReDim ColWidth(ColumnCnt)
                    ReDim Cells(PageSize, ColumnCnt)
                    '
                    ColCaption(ColumnPtr) = "Date"
                    ColAlign(ColumnPtr) = "center"
                    ColWidth(ColumnPtr) = "150"
                    ColumnPtr = ColumnPtr + 1
                    '
                    ColCaption(ColumnPtr) = "Emailed"
                    ColAlign(ColumnPtr) = "left"
                    ColWidth(ColumnPtr) = "50"
                    ColumnPtr = ColumnPtr + 1
                    '
                    ColCaption(ColumnPtr) = "Property"
                    ColAlign(ColumnPtr) = "Left"
                    ColWidth(ColumnPtr) = "300"
                    ColumnPtr = ColumnPtr + 1
                    '
                    ColCaption(ColumnPtr) = "Search"
                    ColAlign(ColumnPtr) = "left"
                    ColWidth(ColumnPtr) = "100%"
                    ColumnPtr = ColumnPtr + 1
                    '
                    ' Build Query
                    '
                    SQLFrom = "" _
                        & " From grgPropertyLog L" _
                        & " left join grgPropertySearches S on S.ID=L.PropertySearchID"
                    SQLWhere = " Where" _
                        & "(S.ID is not null)" _
                        & "and(L.MemberID=" & CSMember.GetInteger("ID") & ")" _
                        & ""
                    If Not AllowNonEmail Then
                        SQLWhere = SQLWhere & "and(L.IsEmailSearch is not null)and(L.IsEmailSearch<>0)"
                    End If
                    Dim rightNow As Date = Now()
                    Dim rightNowDate As Date = rightNow.Date
                    Select Case DateRangeID
                        Case DateRangeToday
                            SQLWhere = SQLWhere & "and(L.DateAdded>=" & cp.Db.EncodeSQLDate(rightNowDate) & ")"
                        Case DateRangeYesterday
                            SQLWhere = SQLWhere & "and(L.DateAdded<" & cp.Db.EncodeSQLDate(rightNowDate) & ")and(L.DateAdded>=" & cp.Db.EncodeSQLDate(rightNowDate.AddDays(-1)) & ")"
                        Case DateRangePastWeek
                            SQLWhere = SQLWhere & "and(L.DateAdded>" & cp.Db.EncodeSQLDate(rightNowDate.AddDays(-7)) & ")"
                        Case DateRangePastMonth
                            SQLWhere = SQLWhere & "and(L.DateAdded>" & cp.Db.EncodeSQLDate(rightNowDate.AddDays(-30)) & ")"
                    End Select
                    SQLOrderBy = " ORDER BY L." & SortField
                    If SortDirection <> 0 Then
                        SQLOrderBy = SQLOrderBy & " Desc"
                    End If
                    If UCase(SortField) <> "PROPERTYSEARCHID" Then
                        SQLOrderBy = SQLOrderBy & ",L.PropertySearchID"
                    End If
                    '
                    ' Get DataRowCount
                    '
                    SQL = "select count(*) as Cnt " & SQLFrom & SQLWhere
                    CS.OpenSQL(SQL)
                    If CS.OK() Then
                        DataRowCount = CS.GetInteger("cnt")
                    End If
                    Call CS.Close()
                    '
                    ' Get Data
                    '
                    SQL = "select S.*" _
                        & ",L.PropertyID as PropertyID" _
                        & ",L.PropertyName as PropertyName" _
                        & ",L.DateAdded as ViewingDateAdded" _
                        & ",L.IsEmailSearch as ViewingIsEmailSearch" _
                        & ",L.VisitID as ViewingVisitID" _
                        & " " & SQLFrom & SQLWhere & SQLOrderBy
                    CS.OpenSQL(SQL, "", PageSize, PageNumber)
                    If Not CS.OK() Then
                        '
                        ' No Searchs saved
                        '
                        Stream = Stream & "<tr class=D0><td class=D0 colspan=""" & ColumnCount & """ width="" 100%"">No Auto Agent records were found.</td></tr>"
                    Else
                        '
                        ' List out the AutoAgents
                        '
                        RowCnt = 0
                        RowClass = RowClassEven
                        PeopleCID = cp.Content.GetID("People")
                        If Not CS.OK() Then
                            '
                            ' EMPTY
                            '
                        Else
                            Do While CS.OK() And (RowCnt < PageSize)
                                RecordID = CS.GetInteger("ID")
                                DateCompleted = CS.GetDate("DateCompleted")
                                Cells(RowCnt, 0) = CS.GetDate("DateAdded") & cp.Html.Hidden("RowID" & RowCnt, RecordID.ToString())
                                Cells(RowCnt, 1) = CS.GetText("name")
                                Cells(RowCnt, 2) = If(CS.GetBoolean("ViewingIsEmailSearch"), "Yes", "No")
                                Propertyid = CS.GetInteger("PropertyID")
                                Copy = CS.GetText("PropertyName")
                                If Copy = "" Then
                                    Copy = "&nbsp;"
                                Else
                                    Copy = "<a href=/index.asp?grg_pid=" & Propertyid & " target=_blank>" & Copy & "</a>"
                                End If
                                Cells(RowCnt, 4) = Copy
                                RowCnt = RowCnt + 1
                                Call CS.GoNext()
                            Loop
                        End If
                        Stream = Stream & cp.Html.Hidden("RowCount", RowCnt.ToString())
                    End If
                    Call CS.Close()
                    '
                    DetailMemberID = CSMember.GetInteger("ID")
                    Stream = AdminUIController.getReport(cp, RowCnt, ColCaption, ColAlign, ColWidth, Cells, PageSize, PageNumber, PreTableCopy, PostTableCopy, DataRowCount, "")
                End If
                result = "<div STYLE=""width:100%;"" class=""cmBody ccPanel3DReverse"">" & Stream & "</div>"
            Catch ex As Exception
                cp.Site.ErrorReport(ex)
            End Try
            Return result
        End Function
        '


    End Class
End Namespace

