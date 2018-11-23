﻿
Option Strict On
Option Explicit On

Imports Contensive.BaseClasses
Imports Contensive.Addons.ContactManager.GenericController
Imports System.Text
Imports System.Linq

Namespace Views
    Public Class DetailFormController
        '
        '=================================================================================
        '
        Public Shared Function ProcessRequest(cp As CPBaseClass, ae As Controllers.ApplicationController, request As Contensive.Addons.ContactManager.Views.CMngrClass.RequestClass) As FormIdEnum
            request.DetailMemberID = cp.Doc.GetInteger(RequestNameMemberID)
            Select Case request.Button
                Case ButtonCancel
                    Return FormIdEnum.FormList
                Case ButtonSave
                    Call SaveContactFromStream(cp, request.DetailMemberID)
                    Return FormIdEnum.FormDetails
                Case ButtonOK
                    Call SaveContactFromStream(cp, request.DetailMemberID)
                    Return FormIdEnum.FormList
                Case ButtonNewSearch
                    Return FormIdEnum.FormSearch
            End Select
            Return FormIdEnum.FormDetails
        End Function
        '
        '=================================================================================
        '
        Public Shared Function GetEventsTab() As String
            Return "events"
        End Function
        '
        '=================================================================================
        '
        Public Shared Function getResponse(cp As CPBaseClass, DetailMemberID As Integer) As String
            Dim result As String = ""
            Try
                '
                Dim SubTab As Integer
                Dim RQS As String
                Dim Nav As New TabController()
                Dim Header As String
                Dim AdminUI As Object
                Dim Content As String
                Dim ButtonList As String
                Dim Button As String
                Dim CS As CPCSBaseClass = cp.CSNew()
                Dim MemberName As String
                Dim hint As String
                '
                If True Then
                    CS.Open("people", "ID=" & DetailMemberID, "", False)
                    MemberName = CS.GetText("name")
                    If MemberName = "" Then
                        MemberName = Trim(CS.GetText("FirstName") & " " & CS.GetText("LastName"))
                        If MemberName = "" Then
                            MemberName = "Record " & CS.GetText("ID")
                        End If
                    End If
                    '
                    ' Determine current Subtab
                    '
                    SubTab = cp.Doc.GetInteger(RequestNameDetailSubtab)
                    If SubTab = 0 Then
                        SubTab = cp.Utils.EncodeInteger(cp.User.GetText(RequestNameDetailSubtab, "1"))
                        If SubTab = 0 Then
                            SubTab = 1
                            Call cp.User.SetProperty(RequestNameDetailSubtab, CStr(SubTab))
                        End If
                    Else
                        Call cp.User.SetProperty(RequestNameDetailSubtab, CStr(SubTab))
                    End If
                    Call cp.Doc.AddRefreshQueryString(RequestNameDetailSubtab, SubTab.ToString())
                    '
                    ' SubTab Menu
                    '
                    Call cp.Doc.AddRefreshQueryString("tab", "")
                    ButtonList = ButtonCancel & "," & ButtonSave & "," & ButtonOK & "," & ButtonNewSearch
                    Header = "<div>" & MemberName & "</div>"
                    '
                    Call Nav.addEntry("Contact", GetFormDetail_TabContact(cp, CS), "ccAdminTab")
                    Call Nav.addEntry("Permissions", GetFormDetail_TabPermissions(cp, CS), "ccAdminTab")
                    Call Nav.addEntry("Notes", GetFormDetail_TabNotes(cp, CS), "ccAdminTab")
                    Call Nav.addEntry("Photos", GetFormDetail_TabPhoto(cp, CS), "ccAdminTab")
                    Call Nav.addEntry("Groups", GetFormDetail_TabGroup(cp, CS), "ccAdminTab")
                    Call CS.Close()
                    '
                    Content = Nav.getTabs(cp)
                    Content = Content & cp.Html.Hidden(RequestNameFormID, Convert.ToInt32(FormIdEnum.FormDetails).ToString())
                    Content = Content & cp.Html.Hidden(RequestNameMemberID, DetailMemberID.ToString())
                    '
                    getResponse = AdminUIController.getBody(cp, "Contact Manager &gt;&gt; Contact Details", ButtonList, "", True, True, Header, "", 0, Content)
                End If
            Catch ex As Exception
                cp.Site.ErrorReport(ex)
            End Try
            Return result
        End Function
        '
        '=================================================================================
        '
        Public Shared Function GetFormDetail_TabContact(cp As CPBaseClass, CS As CPCSBaseClass) As String
            Dim s As String = ""
            Try
                Dim Left As String = ""
                Dim Right As String
                Dim Copy As String
                '
                If Not CS.OK() Then
                    '
                    s = s & "<div>There was a problem retrieving this person's information.</div>"
                Else
                    '
                    ' Left Side
                    '
                    Left = Left & "<table border=0 width=100% cellspacing=0 cellpadding=0>" _
                    & "<TR>" _
                    & "<TD width=150><img src=/cclib/images/spacer.gif width=150 height=1></TD>" _
                    & "<TD width=350><img src=/cclib/images/spacer.gif width=350 height=1></TD>" _
                    & "</TR>"
                    Left = Left & GetFormDetail_InputTextRow(cp, "Full Name", "Name", CS.GetText("name"), False)
                    Left = Left & GetFormDetail_InputTextRow(cp, "First Name", "FirstName", CS.GetText("FirstName"), False)
                    Left = Left & GetFormDetail_InputTextRow(cp, "Last Name", "LastName", CS.GetText("LastName"), False)
                    Left = Left & GetFormDetail_DividerRow(cp, "Contact")
                    Left = Left & GetFormDetail_InputTextRow(cp, "Email", "EMAIL", CS.GetText("email"), False)
                    Left = Left & GetFormDetail_InputTextRow(cp, "Phone", "PHONE", CS.GetText("PHONE"), False)
                    Left = Left & GetFormDetail_InputTextRow(cp, "Fax", "Fax", CS.GetText("Fax"), False)
                    Left = Left & GetFormDetail_InputTextRow(cp, "Address", "ADDRESS", CS.GetText("ADDRESS"), False)
                    Left = Left & GetFormDetail_InputTextRow(cp, "Line 2", "ADDRESS2", CS.GetText("ADDRESS2"), False)
                    Left = Left & GetFormDetail_InputTextRow(cp, "City", "City", CS.GetText("City"), False)
                    Left = Left & GetFormDetail_InputTextRow(cp, "State", "State", CS.GetText("State"), False)
                    Left = Left & GetFormDetail_InputTextRow(cp, "Zip", "Zip", CS.GetText("Zip"), False)
                    Left = Left & GetFormDetail_DividerRow(cp, "Company")
                    Left = Left & GetFormDetail_InputTextRow(cp, "Name", "Company", CS.GetText("Company"), False)
                    Left = Left & GetFormDetail_InputTextRow(cp, "Title", "Title", CS.GetText("Title"), False)
                    Left = Left & GetFormDetail_DividerRow(cp, "Birthday")
                    Left = Left & GetFormDetail_InputTextRow(cp, "Day", "BirthdayDay", CS.GetText("BirthdayDay"), False)
                    Left = Left & GetFormDetail_InputTextRow(cp, "Month", "BirthdayMonth", CS.GetText("BirthdayMonth"), False)
                    Left = Left & GetFormDetail_InputTextRow(cp, "Year", "BirthdayYear", CS.GetText("BirthdayYear"), False)
                    Left = Left & "</table>"
                    '
                    ' Right Side
                    '
                    Copy = cp.Html.InputText("AppendNotes", "", "20", "100")
                    Copy = Replace(Copy, " cols=""100""", " style=""width:100%;""", , , vbTextCompare)
                    '
                    Right = Right & "<table border=0 width=100% cellspacing=0 cellpadding=0>" _
                    & "<TR>" _
                    & "<TD width=200><img src=/cclib/images/spacer.gif width=150 height=1></TD>" _
                    & "<TD width=99%><img src=/cclib/images/spacer.gif width=1 height=1></TD>" _
                    & "</TR>"
                    Right = Right & GetFormDetail_DividerRow(cp, "Add to Notes")
                    Right = Right & "<TR><TD colspan=2>" & Copy & "</TD></TR>"
                    Right = Right & "</table>"
                End If
                '
                s = s & "<table border=0 width=100% cellspacing=0 cellpadding=0>" _
                & "<TR>" _
                & "<TD width=500 valign=top style=""border-right:1px solid #808080;padding-right:20px;"">" & Left & "</TD>" _
                & "<TD width=100% valign=top style=""border-left:1px solid #f0f0f0;padding-left:20px;"">" & Right & "</TD>" _
                & "</TR>" _
                & "</table>"
                s = "<div STYLE=""width:100%;"" class=""cmBody ccPanel3DReverse"">" & s & "</div>"
                s = s
            Catch ex As Exception
                cp.Site.ErrorReport(ex)
            End Try
            Return s
        End Function
        '
        '=================================================================================
        '
        Public Shared Function GetFormDetail_TabPermissions(cp As CPBaseClass, CS As CPCSBaseClass) As String
            Dim s As String = ""
            Try
                If Not CS.OK() Then
                    '
                    s = s & "<div>There was a problem retrieving this person's information.</div>"
                Else
                    '
                    s = s & "<table border=0 width=100% cellspacing=0 cellpadding=0>" _
                        & "<TR>" _
                        & "<TD width=200><img src=/cclib/images/spacer.gif width=150 height=1></TD>" _
                        & "<TD width=99%><img src=/cclib/images/spacer.gif width=1 height=1></TD>" _
                        & "</TR>"
                    s = s & GetFormDetail_DividerRow(cp, "Login")
                    s = s & GetFormDetail_InputTextRow(cp, "Username", "Username", CS.GetText("Username"), False)
                    s = s & GetFormDetail_InputTextRow(cp, "Password", "Password", CS.GetText("Password"), True)
                    s = s & GetFormDetail_InputBooleanRow(cp, "Allow Auto Login", "AutoLogin", CS.GetBoolean("AutoLogin").ToString())
                    s = s & "</table>"
                End If
                s = "<div STYLE=""width:100%;"" class=""cmBody ccPanel3DReverse"">" & s & "</div>"
            Catch ex As Exception
                cp.Site.ErrorReport(ex)
            End Try
            Return s
        End Function
        '
        '=================================================================================
        '
        Public Shared Function GetFormDetail_TabNotes(cp As CPBaseClass, CS As CPCSBaseClass) As String
            Dim s As String = ""
            Try
                If Not CS.OK() Then
                    s = s & "<div>There was a problem retrieving this person's information.</div>"
                Else
                    '
                    s = s & "<table border=0 width=100% cellspacing=0 cellpadding=0>" _
                        & "<TR>" _
                        & "<TD width=200><img src=/cclib/images/spacer.gif width=150 height=1></TD>" _
                        & "<TD width=99%><img src=/cclib/images/spacer.gif width=1 height=1></TD>" _
                        & "</TR>"
                    s = s & GetFormDetail_InputHTMLRow(cp, "Notes", "NotesFilename", CS.GetText("NotesFilename"))
                    s = s & "</table>"
                End If
                '
                s = "<div STYLE=""width:100%;"" class=""cmBody ccPanel3DReverse"">" & s & "</div>"
                GetFormDetail_TabNotes = s
            Catch ex As Exception
                cp.Site.ErrorReport(ex)
            End Try
            Return s
        End Function
        '
        '=================================================================================
        '
        Public Shared Function GetFormDetail_TabPhoto(cp As CPBaseClass, CS As CPCSBaseClass) As String
            Dim s As String = ""
            Try
                If Not CS.OK() Then
                    s = s & "<div>There was a problem retrieving this person's information.</div>"
                Else
                    '
                    s = s & "<table border=0 width=100% cellspacing=0 cellpadding=0>" _
                        & "<TR>" _
                        & "<TD width=200><img src=/cclib/images/spacer.gif width=150 height=1></TD>" _
                        & "<TD width=99%><img src=/cclib/images/spacer.gif width=1 height=1></TD>" _
                        & "</TR>"
                    s = s & GetFormDetail_InputImageRow(cp, "Thumbnail", "ThumbnailFilename", CS.GetText("ThumbnailFilename"))
                    s = s & GetFormDetail_InputImageRow(cp, "Image", "ImageFilename", CS.GetText("ImageFilename"))
                    s = s & "</table>"
                End If
                '
                s = "<div STYLE=""width:100%;"" class=""cmBody ccPanel3DReverse"">" & s & "</div>"
            Catch ex As Exception
                cp.Site.ErrorReport(ex)
            End Try
            Return s
        End Function
        '
        '=================================================================================
        '
        Public Shared Function GetFormDetail_InputTextRow(cp As CPBaseClass, Caption As String, FieldName As String, DefaultValue As String, IsPassword As Boolean) As String
            Dim s As String = ""
            Try
                s = "" _
                    & "<TR><TD style=""TEXT-ALIGN:left;PADDING-LEFT:20px;"">" _
                    & Caption & ":" _
                    & "</TD><TD style=""TEXT-ALIGN:left;"">"
                If IsPassword Then
                    s = s & "<input type=password name=""" & FieldName & """ value=""" & cp.Utils.EncodeHTML(DefaultValue) & """ style=""width:300px;"">"
                Else
                    s = s & "<input type=text name=""" & FieldName & """ value=""" & cp.Utils.EncodeHTML(DefaultValue) & """ style=""width:350px;"">"
                End If
                s = s & "</TD></TR>"
            Catch ex As Exception
                cp.Site.ErrorReport(ex)
            End Try
            Return s
        End Function
        '
        '=================================================================================
        '
        Public Shared Function GetFormDetail_InputBooleanRow(cp As CPBaseClass, Caption As String, FieldName As String, DefaultValue As String) As String
            Dim s As String = ""
            Try
                s = "" _
                    & "<TR><TD style=""TEXT-ALIGN:left;PADDING-LEFT:20px;"">" _
                    & Caption & ":" _
                    & "</TD><TD style=""TEXT-ALIGN:left;"">"
                s = s & "<input type=checkbox name=""" & FieldName & """ value=""" & DefaultValue & """>"
                s = s & "</TD></TR>"
            Catch ex As Exception
                cp.Site.ErrorReport(ex)
            End Try
            Return s
        End Function
        '
        '=================================================================================
        '
        Public Shared Function GetFormDetail_InputHTMLRow(cp As CPBaseClass, Caption As String, FieldName As String, DefaultValue As String) As String
            Dim s As String = ""
            Try
                s = "" _
                    & "<TR><TD style=""TEXT-ALIGN:left;PADDING-LEFT:20px;"">" _
                    & Caption & ":" _
                    & "</TD><TD style=""TEXT-ALIGN:left;"">"
                s = s & cp.Html.InputWysiwyg(FieldName, DefaultValue)
                s = s & "</TD></TR>"
            Catch ex As Exception
                cp.Site.ErrorReport(ex)
            End Try
            Return s
        End Function
        '
        '=================================================================================
        '
        Public Shared Function GetFormDetail_InputImageRow(cp As CPBaseClass, Caption As String, FieldName As String, DefaultValue As String) As String
            Dim s As String = ""
            Try
                Dim EncodedLink As String
                Dim Filename As String
                '
                s = "" _
                    & "<TR><TD style=""TEXT-ALIGN:left;PADDING-LEFT:20px;"">" _
                    & Caption & ":" _
                    & "</TD><TD style=""TEXT-ALIGN:left;"">"
                If DefaultValue = "" Then
                    s = s & cp.Html.InputFile(FieldName)
                    s = s & "</TD></TR>"
                Else
                    Filename = cp.Utils.EncodeHTML(DefaultValue)
                    EncodedLink = cp.Utils.EncodeUrl("http://" & cp.Request.Host & cp.Site.FilePath & DefaultValue)
                    s = s _
                        & "<table border=0 width=100% cellspacing=0 cellpadding=4><TR>" _
                        & "<TD width=200><a href=""" & EncodedLink & """ target=""_blank""><img src=""" & EncodedLink & """ width=200 border=0></a></TD>" _
                        & "<TD width=100% valign=top>" _
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
            Return s
        End Function
        '
        '=================================================================================
        '
        Public Shared Function GetFormDetail_DividerRow(cp As CPBaseClass, Caption As String) As String
            Dim s As String = ""
            Try
                s = Replace(Caption, " ", "&nbsp;")
                s = "<TR><TD colspan=2 style=""Padding-top:10px;"">" _
                    & "<TABLE border=0 width=100% cellspacing=0 cellpadding=0>" _
                    & "<TR><TD width=1 style=""white-space:nowrap;"">" & s & "&nbsp;&nbsp;</TD><TD width=100%><HR></TD>" _
                    & "</TABLE>" _
                    & "</tr>"
            Catch ex As Exception
                cp.Site.ErrorReport(ex)
            End Try
            Return s
        End Function
        '
        '=================================================================================
        '
        Public Shared Function SaveContactFromStream(cp As CPBaseClass, DetailMemberID As Integer) As String
            Dim result As String = ""
            Try
                Dim CS As CPCSBaseClass = cp.CSNew()
                Dim Filename As String
                Dim OriginalFilename As String
                Dim Path As String
                Dim FieldName As String
                Dim Copy As String
                '
                If CS.Open("people", "id=" & DetailMemberID) Then
                    Call CS.SetField("name", cp.Doc.GetText("name"))
                    Call CS.SetField("FirstName", cp.Doc.GetText("FirstName"))
                    Call CS.SetField("LastName", cp.Doc.GetText("LastName"))
                    '
                    ' Contact
                    '
                    Call CS.SetField("email", cp.Doc.GetText("email"))
                    Call CS.SetField("Phone", cp.Doc.GetText("Phone"))
                    Call CS.SetField("Fax", cp.Doc.GetText("Fax"))
                    Call CS.SetField("Address", cp.Doc.GetText("Address"))
                    Call CS.SetField("Address2", cp.Doc.GetText("Address2"))
                    Call CS.SetField("City", cp.Doc.GetText("City"))
                    Call CS.SetField("State", cp.Doc.GetText("State"))
                    Call CS.SetField("Zip", cp.Doc.GetText("Zip"))
                    '
                    ' Company
                    '
                    Call CS.SetField("Company", cp.Doc.GetText("Company"))
                    Call CS.SetField("Title", cp.Doc.GetText("Title"))
                    '
                    ' Birthday
                    '
                    Call CS.SetField("BirthdayDay", cp.Doc.GetInteger("BirthdayDay"))
                    Call CS.SetField("BirthdayMonth", cp.Doc.GetInteger("BirthdayMonth"))
                    Call CS.SetField("BirthdayYear", cp.Doc.GetInteger("BirthdayYear"))
                    '
                    ' Notes
                    '
                    Copy = cp.Doc.GetText("AppendNotes")
                    If Copy <> "" Then
                        Copy = "" _
                            & "<div style=""margin-top:10px;border-top:1px dashed black;"">Added " & Now & " by " & cp.Content.GetRecordName("people", cp.User.Id) & "</div>" _
                            & "<div style=""margin-left:20px;margin-top:5px;"">" & Copy & "</div>"
                    End If
                    Call CS.SetField("NotesFilename", cp.Doc.GetText("NotesFilename") & Copy)
                    '
                    ' Photos
                    '
                    FieldName = "ThumbnailFilename"
                    If cp.Doc.GetBoolean(FieldName & ".DeleteFlag") Then
                        Call CS.SetField(FieldName, "")
                    End If
                    OriginalFilename = cp.Doc.GetText(FieldName)
                    If OriginalFilename <> "" Then
                        Filename = CS.GetFilename(FieldName, OriginalFilename)
                        Path = Filename
                        Path = Replace(Path, "/", "\")
                        Path = Replace(Path, "\" & OriginalFilename, "")
                        Call CS.SetField(FieldName, Filename)
                        'Call CS.SetFile(FieldName, Path)
                    End If
                    '
                    FieldName = "ImageFilename"
                    If cp.Doc.GetBoolean(FieldName & ".DeleteFlag") Then
                        Call CS.SetField(FieldName, "")
                    End If
                    OriginalFilename = cp.Doc.GetText(FieldName)
                    If OriginalFilename <> "" Then
                        Filename = CS.GetFilename(FieldName, OriginalFilename)
                        Path = Filename
                        Path = Replace(Path, "/", "\")
                        Path = Replace(Path, "\" & OriginalFilename, "")
                        Call CS.SetField(FieldName, Filename)
                        'Call CS.SetFile(FieldName, Path)
                    End If
                    '
                    ' Permissions
                    '
                    Call CS.SetField("Username", cp.Doc.GetText("Username"))
                    Call CS.SetField("Password", cp.Doc.GetText("Password"))
                    Call CS.SetField("AutoLogin", cp.Doc.GetBoolean("AutoLogin"))
                    '
                    ' Groups
                    '
                    Call SaveMemberRules(cp, CS)
                End If
                Call CS.Close()
            Catch ex As Exception
                cp.Site.ErrorReport(ex)
            End Try
            Return result
        End Function
        '
        '========================================================================
        '
        Public Shared Function GetFormDetail_TabGroup(cp As CPBaseClass, CSMember As CPCSBaseClass) As String
            Dim result As String = ""
            Try
                Dim sb As New StringBuilder()
                Dim Copy As String
                Dim SQL As String
                Dim CS As CPCSBaseClass = cp.CSNew()
                Dim MembershipCount As Integer
                Dim MembershipSize As Integer
                Dim MembershipPointer As Integer
                Dim SectionName As String
                Dim PrimaryContentID As Integer
                Dim SecondaryContentID As Integer
                Dim CanSeeHiddenGroups As Boolean
                Dim DateExpireValue As String
                Dim GroupCount As Integer
                Dim GroupID As Integer
                Dim GroupName As String
                Dim GroupCaption As String
                Dim GroupActive As Boolean
                Dim Membership() As Integer
                Dim DateExpires() As Date
                Dim active() As Boolean
                Dim Caption As String
                Dim MethodName As String
                Dim ReportLink As String
                Dim AdminUI As Object
                Dim DetailMemberID As Integer
                '
                DetailMemberID = CSMember.GetInteger("ID")
                '
                ' ----- Gather all the SecondaryContent that associates to the PrimaryContent
                '
                PrimaryContentID = cp.Content.GetID("People")
                SecondaryContentID = cp.Content.GetID("Groups")
                '
                MembershipCount = 0
                MembershipSize = 0
                sb.Append(vbCrLf & "<!-- GroupRule Table --><table border=0 width=100% cellspacing=0 cellpadding=0>" _
                    & "<TR>" _
                    & "<TD width=150><img src=/cclib/images/spacer.gif width=150 height=1></TD>" _
                    & "<TD width=99%><img src=/cclib/images/spacer.gif width=1 height=1></TD>" _
                    & "</TR>")
                If DetailMemberID <> 0 Then
                    '
                    ' ----- read in the groups that this member has subscribed (exclude new member records)
                    '
                    SQL = "SELECT Active,GroupID,DateExpires" _
                        & " FROM ccMemberRules" _
                        & " WHERE MemberID=" & DetailMemberID
                    CS.OpenSQL(SQL)
                    Do While CS.OK()
                        If MembershipCount >= MembershipSize Then
                            MembershipSize = MembershipSize + 10
                            ReDim Preserve Membership(MembershipSize)
                            ReDim Preserve active(MembershipSize)
                            ReDim Preserve DateExpires(MembershipSize)
                        End If
                        Membership(MembershipCount) = CS.GetInteger("GroupID")
                        DateExpires(MembershipCount) = CS.GetDate("DateExpires")
                        active(MembershipCount) = CS.GetBoolean("Active")
                        MembershipCount = MembershipCount + 1
                        CS.GoNext()
                    Loop
                    Call CS.Close()
                    '
                    ' ----- read in all the groups, sorted by ContentName
                    '
                    SQL = "SELECT ccGroups.ID AS ID, ccContent.Name AS SectionName, ccGroups.Caption AS GroupCaption, ccGroups.name AS GroupName, ccGroups.SortOrder" _
                        & " FROM ccGroups LEFT JOIN ccContent ON ccGroups.ContentControlID = ccContent.ID" _
                        & " Where (((ccGroups.Active) >0) And ((ccContent.Active) >0))"
                    SQL = SQL _
                        & " GROUP BY ccGroups.ID, ccContent.Name, ccGroups.Caption, ccGroups.name, ccGroups.SortOrder" _
                        & " ORDER BY ccContent.Name, ccGroups.Caption"
                    CS.OpenSQL(SQL)
                    '
                    ' Output all the groups, with the active and dateexpires from those joined
                    '
                    'F.Add Controllers.AdminUIController.EditTableOpen
                    SectionName = ""
                    GroupCount = 0
                    CanSeeHiddenGroups = cp.User.IsDeveloper
                    Do While CS.OK()
                        GroupName = CS.GetText("GroupName")
                        If (Mid(GroupName, 1, 1) <> "_") Or CanSeeHiddenGroups Then
                            GroupCaption = CS.GetText("GroupCaption")
                            GroupID = CS.GetInteger("ID")
                            If GroupCaption = "" Then
                                GroupCaption = GroupName
                                If GroupCaption = "" Then
                                    GroupCaption = "Group&nbsp;" & GroupID
                                End If
                            End If
                            GroupActive = False
                            DateExpireValue = ""
                            If MembershipCount <> 0 Then
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
                            If GroupID > 0 Then
                                ReportLink = "<a href=""?af=12&rid=35&recordid=" & GroupID & """ target=_blank>Group&nbsp;Report</a>"
                            Else
                                ReportLink = "&nbsp;"
                            End If
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
                                & "<TD style=""TEXT-ALIGN:left;PADDING-LEFT:20px;border-top:1px solid white;"">" & cp.Html.CheckBox("MemberRules." & GroupCount, GroupActive) & GroupCaption & "</TD>" _
                                & "<TD style=""TEXT-ALIGN:left;PADDING-LEFT:10px;border-top:1px solid white;"">Expires " & cp.Html.InputText("MemberRules." & GroupCount & ".DateExpires", DateExpireValue, "1", "20") & "</TD>" _
                                & "</TR>")
                            GroupCount = GroupCount + 1
                        End If
                        CS.GoNext()
                    Loop
                    CS.Close()
                End If
                If DetailMemberID = 0 Then
                    sb.Append("<TR>" _
                        & "<TD valign=middle align=right><span>Groups</span></TD>" _
                        & "<TD><span>Groups will be available after this record is saved</SPAN></TD>" _
                        & "<TR>")
                ElseIf GroupCount = 0 Then
                    sb.Append("<TR>" _
                        & "<TD valign=middle align=right><span>Groups</span></TD>" _
                        & "<TD><span>There are currently no groups defined</SPAN></TD>" _
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
        Private Shared Sub SaveMemberRules(cp As CPBaseClass, CSMember As CPCSBaseClass)
            Try
                Dim GroupCount As Integer
                Dim GroupPointer As Integer
                Dim CSPointer As Integer
                Dim MethodName As String
                Dim GroupID As Integer
                Dim RuleNeeded As Boolean
                Dim CSRule As Integer
                Dim DateExpires As Date
                Dim DateExpiresVariant As Object
                Dim RuleActive As Boolean
                Dim RuleDateExpires As Date
                Dim MemberRuleID As Integer
                Dim PeopleID As Integer
                '
                ' --- create MemberRule records for all selected
                PeopleID = CSMember.GetInteger("ID")
                GroupCount = cp.Doc.GetInteger("MemberRules" & ".RowCount")
                If GroupCount > 0 Then
                    For GroupPointer = 0 To GroupCount - 1
                        '
                        ' ----- Read Response
                        '
                        GroupID = cp.Doc.GetInteger("MemberRules" & "." & GroupPointer & "." & "ID")
                        RuleNeeded = cp.Doc.GetBoolean("MemberRules" & "." & GroupPointer)
                        DateExpires = cp.Doc.GetDate("MemberRules" & "." & GroupPointer & "." & "DateExpires")
                        If DateExpires = Date.MinValue Then
                            DateExpiresVariant = Date.MinValue
                        Else
                            DateExpiresVariant = DateExpires
                        End If
                        '
                        ' ----- Update Record
                        '
                        Dim ruleList As List(Of Models.MemberRuleModel) = Models.MemberRuleModel.createList(cp, "(MemberID=" & PeopleID & ")and(GroupID=" & GroupID & ")")
                        If ruleList.Count = 0 Then
                            '
                            ' No record exists
                            If RuleNeeded Then
                                '
                                ' No record, Rule needed, add it
                                Dim newRule As Models.MemberRuleModel = Models.MemberRuleModel.add(cp)
                                newRule.Active = True
                                newRule.MemberID = PeopleID
                                newRule.GroupID = GroupID
                                newRule.DateExpires = DateExpires
                                newRule.save(cp)
                            End If
                        Else
                            '
                            ' Record exists
                            Dim keepRule As Models.MemberRuleModel = Nothing
                            If RuleNeeded Then
                                '
                                ' record exists, and it is needed, update the DateExpires if changed
                                keepRule = ruleList.First()
                                If (Not keepRule.Active) Or (Not keepRule.DateExpires.Equals(DateExpires)) Then
                                    keepRule.Active = True
                                    keepRule.DateExpires = DateExpires
                                    keepRule.save(cp)
                                End If
                            End If
                            For Each rule In ruleList
                                If (rule IsNot keepRule) Then Models.MemberRuleModel.delete(cp, rule.id)
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
                        Stream = Stream & "<tr class=D0><td class=D0 colspan=""" & ColumnCount & """ width=""100%"">No Auto Agent records were found.</td></tr>"
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

