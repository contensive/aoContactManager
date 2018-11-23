
Option Strict On
Option Explicit On

Imports Contensive.BaseClasses
Imports Contensive.Addons.ContactManager.GenericController
Imports System.Text
Imports System.Linq

Namespace Views
    Public Class ListFormController
        '
        '=================================================================================
        '
        Public Shared Function ProcessRequest(cp As CPBaseClass, ae As Controllers.ApplicationController, request As ContactManager.Views.CMngrClass.RequestClass) As FormIdEnum
            '
            Select Case request.Button
                Case ButtonNewSearch
                    ae.userProperties.contactSearchCriteria = ""
                    ae.userProperties.contactGroupCriteria = ""
                    request.FormID = FormIdEnum.FormSearch
                Case ButtonApply
                    '
                    ' Add to or remove from group
                    '
                    Dim SQLCriteria As String = ""
                    Dim SearchCaption As String = ""
                    Call BuildSearch(cp, ae, SQLCriteria, SearchCaption)
                    Dim GroupName As String
                    Dim RowPointer As Integer
                    Dim memberID As Integer
                    Dim SQL As String
                    Select Case request.GroupToolAction
                        Case GroupToolActionEnum.AddToGroup
                            '
                            ' ----- Add to Group
                            '
                            If (request.GroupID = 0) Then
                                '
                                ' Group required and not provided
                                '
                                Call cp.UserError.Add("Please select a Target Group for this operation")
                            ElseIf request.GroupToolSelect = 0 Then
                                '
                                ' Add selection to Group
                                '
                                If (request.RowCount > 0) Then
                                    GroupName = cp.Group.GetName(request.GroupID.ToString())
                                    For RowPointer = 0 To request.RowCount - 1
                                        If cp.Doc.GetBoolean("M." & RowPointer) Then
                                            memberID = cp.Doc.GetInteger("MID." & RowPointer)
                                            Call cp.Group.AddUser(GroupName, memberID)
                                        End If
                                    Next
                                End If
                            Else
                                '
                                ' Add everyone in search criteria to this group
                                '

                                Dim CCID As Integer = cp.Content.GetID("Member Rules")
                                SQL = "insert into ccMemberRules (Active,ContentControlID,GroupID,MemberID )" _
                                                            & " select 1," & CCID & "," & request.GroupID & ",ccMembers.ID" _
                                                            & " from (ccMembers" _
                                                            & " left join ccMemberRules on ccMemberRules.MemberID=ccMembers.ID )" _
                                                            & " left join ( select MemberID  from ccMemberRules where GroupID in (" & request.GroupID & ")) as InGroups on InGroups.MemberID=ccMembers.ID" _
                                                            & " " & SQLCriteria _
                                                            & " and InGroups.MemberID is null" _
                                                            & ""
                                Call cp.Db.ExecuteNonQuery(SQL)
                            End If
                        Case GroupToolActionEnum.RemoveFromGroup
                            '
                            ' ----- Remove From Group
                            '
                            If (request.GroupID = 0) Then
                                '
                                ' Group required and not provided
                                '
                                Call cp.UserError.Add("Please select a Target Group for this operation")
                            ElseIf request.GroupToolSelect = 0 Then
                                '
                                ' Remove selection from Group
                                '
                                If (request.RowCount > 0) Then
                                    GroupName = cp.Group.GetName(request.GroupID.ToString())
                                    For RowPointer = 0 To request.RowCount - 1
                                        If cp.Doc.GetBoolean("M." & RowPointer) Then
                                            memberID = cp.Doc.GetInteger("MID." & RowPointer)
                                            Call cp.Content.Delete("Member Rules", "(GroupID=" & request.GroupID & ")and(MemberID=" & memberID & ")")
                                        End If
                                    Next
                                End If
                            Else
                                '
                                ' Remove everyone in search criteria from this group
                                '
                                SQL = "delete from ccMemberRules where GroupID=" & request.GroupID & " and MemberID in (" _
                                                            & " select ccMembers.ID" _
                                                            & " from (ccMembers" _
                                                            & " left join ccMemberRules on ccMemberRules.MemberID=ccMembers.ID )" _
                                                            & " " & SQLCriteria _
                                                            & ")"
                                Call cp.Db.ExecuteNonQuery(SQL)
                            End If
                        Case GroupToolActionEnum.ExportGroup
                            '
                            ' ----- Export
                            '
                            Dim Aborttask As Boolean = False
                            If True Then
                                Dim ExportName As String = SearchCaption
                                If request.GroupToolSelect = 0 Then
                                    Dim RowSQL As String
                                    '
                                    ' Export selection from Group
                                    '
                                    ExportName = "Selected rows from " & ExportName
                                    RowSQL = ""
                                    If (request.RowCount > 0) Then
                                        GroupName = cp.Group.GetName(request.GroupID.ToString())
                                        For RowPointer = 0 To request.RowCount - 1
                                            If cp.Doc.GetBoolean("M." & RowPointer) Then
                                                memberID = cp.Doc.GetInteger("MID." & RowPointer)
                                                RowSQL = RowSQL & "OR(ccMembers.ID=" & memberID & ")"
                                            End If
                                        Next
                                        If RowSQL = "" Then
                                            '
                                            ' nothing selected, abort export
                                            '
                                            Aborttask = True
                                            ae.StatusMessage = "<P>You requested to only download the selected entries, and none were selected.<P>"
                                        ElseIf SQLCriteria = "" Then
                                            '
                                            ' This is the only criteria
                                            '
                                            SQLCriteria = " WHERE(" & Mid(RowSQL, 3) & ")"
                                        Else
                                            '
                                            ' Add this criteria to the previous
                                            '
                                            SQLCriteria = SQLCriteria & " And(" & Mid(RowSQL, 3) & ")"
                                        End If
                                    End If
                                Else
                                    '
                                    ' Export the search criteria
                                    '
                                End If
                                If Not Aborttask Then
                                    Dim SQLFrom As String = "ccMembers"
                                    Dim JoinTableCnt As Integer = 0
                                    If cp.User.GetText("ContactGroupCriteria", "") <> "" Then
                                        SQLFrom = "(" & SQLFrom & " left join ccMemberRules on ccMemberRules.MemberID=ccMembers.ID )"
                                    End If
                                    Dim ContentID As Integer = cp.User.GetInteger("ContactContentID", cp.Content.GetID("people").ToString())
                                    Dim cs As CPCSBaseClass = cp.CSNew()
                                    cs.Open("Content Fields", "ContentID=" & ContentID, "EditSortPriority", True, "Name,Caption,Type,LookupContentID")
                                    Dim SelectList As String = ""
                                    Dim FieldNameList As String = ""
                                    Do While cs.OK()
                                        Dim FieldName As String = cs.GetText("name")
                                        Select Case cs.GetInteger("type")
                                            Case FieldTypeLookup
                                                '
                                                ' just add the ID into the list
                                                '
                                                SelectList = SelectList & ",ccMembers." & FieldName
                                                FieldNameList = FieldNameList & "," & FieldName
                                            Case FieldTypeFileText
                                                                    '
                                                                    ' read file for text - skip field
                                            Case FieldTypeRedirect, FieldTypeManyToMany
                                                '
                                                ' no field involved, skip it
                                            Case Else
                                                '
                                                ' just add value
                                                SelectList = SelectList & ",ccMembers." & FieldName
                                                FieldNameList = FieldNameList & "," & FieldName
                                        End Select
                                        Call cs.GoNext()
                                    Loop
                                    Call cs.Close()
                                    If SelectList = "" Then
                                        ae.StatusMessage = "<P>There was a problem requesting your download.<P>"
                                    Else
                                        SelectList = Mid(SelectList, 2)
                                        If FieldNameList <> "" Then
                                            FieldNameList = Mid(FieldNameList, 2)
                                        End If
                                        'ExportName = CStr(Now()) & " snapshot of " & LCase(ExportName)
                                        SQL = "select Distinct " & SelectList & " from " & SQLFrom & SQLCriteria
                                        '
                                        ' -- skip for now
                                        Throw New NotImplementedException()
                                        ' Call main.RequestTask("BuildCSV", SQL, ExportName, "MemberExport-" & CStr(main.GetRandomLong) & ".csv")
                                        ' StatusMessage = "<P>Your download request has been submitted and will be available on the <a href=" & main.SiteProperty_AdminURL & "?af=30>Download Requests</a> page shortly.<P>"
                                    End If
                                End If
                            End If
                        Case GroupToolActionEnum.SetGroupEmail
                            '
                            ' ----- Set AllowBulkEmail field
                            '
                            If request.GroupToolSelect = 0 Then
                                '
                                ' Just selection
                                '
                                Dim RecordCnt As Integer = 0
                                If (request.RowCount > 0) Then
                                    GroupName = cp.Group.GetName(request.GroupID.ToString())
                                    For RowPointer = 0 To request.RowCount - 1
                                        If cp.Doc.GetBoolean("M." & RowPointer) Then
                                            memberID = cp.Doc.GetInteger("MID." & RowPointer)
                                            Call cp.Db.ExecuteNonQuery("update ccMembers set AllowBulkEmail=1 where ID=" & memberID)
                                            RecordCnt = RecordCnt + 1
                                        End If
                                    Next
                                End If
                                ae.StatusMessage = "<P>Allow Group Email was set for " & RecordCnt & " people.<P>"
                            Else
                                '
                                ' Set for everyone in search criteria
                                '
                                SQL = "Update ccMembers set AllowBulkEmail=1 where ID in (" _
                                                            & " select Distinct ccMembers.ID" _
                                                            & " from (ccMembers" _
                                                            & " left join ccMemberRules on ccMemberRules.MemberID=ccMembers.ID )" _
                                                            & " " & SQLCriteria _
                                                            & ")"
                                Call cp.Db.ExecuteNonQuery(SQL)
                                ae.StatusMessage = "<P>Allow Group Email was set for all people in this selection.<P>"
                            End If
                    End Select
            End Select
        End Function
        '
        '=================================================================================
        '
        Public Shared Function getResponse(cp As CPBaseClass, ae As Controllers.ApplicationController, IsAdminPath As Boolean) As String
            Dim result As String = ""
            Try
                Dim SQLOrderDir As String
                Dim SQLOrderBy As String
                Dim LastVisit As Date
                Dim CheckBox As String
                Dim CS As CPCSBaseClass = cp.CSNew()
                Dim RecordName As String
                Dim RecordID As Integer
                Dim SQL As String
                Dim RQS As String
                Dim Criteria As String
                Dim PageSize As Integer
                Dim PageNumber As Integer
                Dim TopCount As Integer
                Dim RowPointer As Integer
                Dim DataRowCount As Integer
                Dim PreTableCopy As String
                Dim PostTableCopy As String
                Dim ColumnCount As Integer
                Dim CPtr As Integer
                Dim ColCaption() As String
                Dim ColAlign() As String
                Dim ColWidth() As String
                Dim Cells(,) As String
                Dim ColSortable() As Boolean
                Dim DefaultSortColumnPtr As Integer
                Dim GroupID As Integer
                Dim GroupToolAction As Integer
                Dim ActionPanel As String
                Dim VisitsCell As String
                Dim VisitCount As Integer
                Dim AdminURL As String
                Dim CCID As Integer
                Dim DefaultName As String
                Dim SearchCaption As String
                Dim BlankPanel As String
                Dim RowPageSize As String
                Dim RowGroups As String
                Dim ButtonList As String
                Dim Header As String
                Dim Body As String
                Dim Description As String
                Dim IsAdmin As Boolean
                Dim SortColPtr As Integer
                Dim TextTest As String
                '
                If True Then
                    IsAdmin = cp.User.IsAdmin()
                    RQS = cp.Doc.RefreshQueryString
                    TextTest = cp.Doc.GetText(RequestNamePageSize)
                    If TextTest = "" Then
                        PageSize = cp.Utils.EncodeInteger(cp.Visit.GetText("cmPageSize", "50"))
                    Else
                        PageSize = cp.Utils.EncodeInteger(TextTest)
                        Call cp.Visit.SetProperty("cmPageSize", CStr(PageSize))
                    End If
                    'PageSize = cp.doc.getInteger(RequestNamePageSize)
                    If PageSize = 0 Then
                        PageSize = 50
                    End If
                    PageNumber = cp.Doc.GetInteger(RequestNamePageNumber)
                    If PageNumber = 0 Then
                        PageNumber = 1
                    End If
                    GroupID = cp.Doc.GetInteger("GroupID")
                    GroupToolAction = cp.Doc.GetInteger("GroupToolAction")
                    AdminURL = cp.Site.GetText("adminurl")
                    ColumnCount = 11
                    If IsAdmin Then
                        '
                        ' If admin, give them an edit column
                        '
                        ColumnCount = ColumnCount + 1
                    End If
                    '
                    TopCount = PageNumber * PageSize
                    '
                    ReDim ColCaption(ColumnCount)
                    ReDim ColAlign(ColumnCount)
                    ReDim ColWidth(ColumnCount)
                    ReDim ColSortable(ColumnCount)
                    ReDim Cells(PageSize, ColumnCount)
                    '
                    SortColPtr = 3
                    TextTest = cp.Visit.GetText("cmSortColumn", "")
                    If TextTest <> "" Then
                        SortColPtr = cp.Utils.EncodeInteger(TextTest)
                    End If
                    SortColPtr = AdminUIController.getReportSortColumnPtr(cp, SortColPtr)
                    If CStr(SortColPtr) <> TextTest Then
                        Call cp.Visit.SetProperty("cmSortColumn", CStr(SortColPtr))
                    End If
                    If AdminUIController.getReportSortType(cp) = 2 Then
                        SQLOrderDir = " Desc"
                    End If
                    '
                    ' Headers
                    '
                    CPtr = 0
                    ColCaption(CPtr) = "<INPUT TYPE=CheckBox OnClick=""CheckInputs('DelCheck',this.checked);"">"
                    'ColCaption(CPtr) = "<INPUT TYPE=CheckBox OnClick=""CheckInputs('DelCheck',this.checked);""><BR><img src=/cclib/images/spacer.gif width=10 height=1>"
                    ColAlign(CPtr) = "center"
                    ColWidth(CPtr) = "10"
                    ColSortable(CPtr) = False
                    CPtr = CPtr + 1
                    '
                    If IsAdmin Then
                        ColCaption(CPtr) = "Edit"
                        'ColCaption(CPtr) = "Edit<BR><img src=/cclib/images/spacer.gif width=10 height=1>"
                        ColAlign(CPtr) = "center"
                        ColWidth(CPtr) = "10"
                        ColSortable(CPtr) = False
                        CPtr = CPtr + 1
                    End If
                    '
                    ColCaption(CPtr) = "Details"
                    'ColCaption(CPtr) = "Details<BR><img src=/cclib/images/spacer.gif width=10 height=1>"
                    ColAlign(CPtr) = "center"
                    ColWidth(CPtr) = "10"
                    ColSortable(CPtr) = False
                    CPtr = CPtr + 1
                    '
                    ColCaption(CPtr) = "Full Name"
                    ColAlign(CPtr) = "left"
                    ColWidth(CPtr) = "20%"
                    ColSortable(CPtr) = True
                    DefaultSortColumnPtr = CPtr
                    If CPtr = SortColPtr Then
                        SQLOrderBy = "Order By ccMembers.Name"
                    End If
                    CPtr = CPtr + 1
                    '
                    ColCaption(CPtr) = "First Name"
                    ColAlign(CPtr) = "left"
                    ColWidth(CPtr) = "20%"
                    ColSortable(CPtr) = True
                    If CPtr = SortColPtr Then
                        SQLOrderBy = "Order By ccMembers.FirstName"
                    End If
                    CPtr = CPtr + 1
                    '
                    ColCaption(CPtr) = "Last Name"
                    ColAlign(CPtr) = "left"
                    ColWidth(CPtr) = "20%"
                    ColSortable(CPtr) = True
                    If CPtr = SortColPtr Then
                        SQLOrderBy = "Order By ccMembers.LastName"
                    End If
                    CPtr = CPtr + 1
                    '
                    ColCaption(CPtr) = "Organization"
                    ColAlign(CPtr) = "left"
                    ColWidth(CPtr) = "20%"
                    ColSortable(CPtr) = True
                    If CPtr = SortColPtr Then
                        SQLOrderBy = "Order By Organizations.Name"
                    End If
                    CPtr = CPtr + 1
                    '
                    ColCaption(CPtr) = "Phone"
                    'ColCaption(CPtr) = "Phone<BR><img src=/cclib/images/spacer.gif width=100 height=1>"
                    ColAlign(CPtr) = "left"
                    ColWidth(CPtr) = "100"
                    ColSortable(CPtr) = True
                    If CPtr = SortColPtr Then
                        SQLOrderBy = "Order By ccMembers.Phone"
                    End If
                    CPtr = CPtr + 1
                    '
                    ColCaption(CPtr) = "email"
                    'ColCaption(CPtr) = "email<BR><img src=/cclib/images/spacer.gif width=150 height=1>"
                    ColAlign(CPtr) = "left"
                    ColWidth(CPtr) = "150"
                    ColSortable(CPtr) = True
                    If CPtr = SortColPtr Then
                        SQLOrderBy = "Order By ccMembers.Email"
                    End If
                    CPtr = CPtr + 1
                    '
                    ColCaption(CPtr) = "Visits"
                    'ColCaption(CPtr) = "Visits<BR><img src=/cclib/images/spacer.gif width=40 height=1>"
                    ColAlign(CPtr) = "right"
                    ColWidth(CPtr) = "40"
                    ColSortable(CPtr) = True
                    If CPtr = SortColPtr Then
                        SQLOrderBy = "Order By ccMembers.Visits"
                    End If
                    CPtr = CPtr + 1
                    '
                    ColCaption(CPtr) = "Last Visit"
                    'ColCaption(CPtr) = "Last Visit<BR><img src=/cclib/images/spacer.gif width=80 height=1>"
                    ColAlign(CPtr) = "right"
                    ColWidth(CPtr) = "80"
                    ColSortable(CPtr) = True
                    If CPtr = SortColPtr Then
                        SQLOrderBy = "Order By ccMembers.LastVisit"
                    End If
                    CPtr = CPtr + 1
                    '
                    ColCaption(CPtr) = "&nbsp;"
                    'ColCaption(CPtr) = "&nbsp;<BR><img src=/cclib/images/spacer.gif width=1 height=1>"
                    ColAlign(CPtr) = "right"
                    ColWidth(CPtr) = "1"
                    ColSortable(CPtr) = False
                    CPtr = CPtr + 1
                    '
                    ' SubTab Menu
                    '
                    RQS = cp.Doc.RefreshQueryString
                    RQS = cp.Utils.ModifyQueryString(RQS, "tab", "", False)
                    Call BuildSearch(cp, ae, Criteria, SearchCaption)
                    '
                    ' purge duplicate in ccMemberRules
                    '
                    SQL = "Delete from ccMemberRules where ID in (" _
                        & " SELECT DISTINCT B.ID" _
                        & " FROM ccMemberRules A, ccMemberRules B" _
                        & " WHERE (((A.MemberID)=[B].[MemberID]) AND ((A.GroupID)=[B].[GroupID]) AND ((A.ID)<[B].[ID]))" _
                        & ")"
                    Call cp.Db.ExecuteNonQuery(SQL)
                    '
                    '   Get DataRowCount
                    '       This had been commented out - but it is needed for the "matches found" caption
                    '
                    SQL = "Select distinct ccMembers.ID" _
                        & " from ccMembers" _
                        & " left join ccMemberRules on ccMemberRules.MemberID=ccMembers.ID" _
                        & Criteria
                    CS.OpenSQL(SQL)
                    If CS.OK() Then
                        DataRowCount = CS.GetRowCount()
                    End If
                    Call CS.Close()
                    '
                    '   Get Data
                    '
                    DefaultName = "Guest"
                    SQL = "Select distinct Top " & TopCount & " ccMembers.name,ccMembers.FirstName,ccMembers.LastName,ccMembers.ID, ccMembers.ContentControlID, ccMembers.Visits, ccMembers.LastVisit, ccMembers.Phone, ccMembers.Email,Organizations.Name as OrgName" _
                        & " from ((ccMembers" _
                        & " left join organizations on Organizations.ID=ccMembers.OrganizationID)" _
                        & " left join ccMemberRules on ccMemberRules.MemberID=ccMembers.ID)" _
                        & Criteria _
                        & SQLOrderBy & SQLOrderDir
                    CS.OpenSQL(SQL, "", PageSize, PageNumber)
                    RowPointer = 0
                    If Not CS.OK() Then
                        Cells(0, 3) = "This search returned no results"
                        RowPointer = 1
                    Else
                        'DataRowCount = Main.GetCSRowCount(CS)
                        Do While CS.OK() And (RowPointer < PageSize)
                            CPtr = 0
                            RecordID = CS.GetInteger("ID")
                            CCID = CS.GetInteger("ContentControlID")
                            VisitCount = CS.GetInteger("Visits")
                            VisitsCell = CStr(VisitCount)
                            CheckBox = cp.Html.CheckBox("M." & RowPointer, , "DelCheck")
                            Cells(RowPointer, CPtr) = CheckBox & cp.Html.Hidden("MID." & RowPointer, RecordID.ToString())
                            CPtr = CPtr + 1
                            RecordName = CS.GetText("name")
                            If RecordName = "" Then
                                RecordName = DefaultName
                            End If
                            If RecordName = DefaultName Then
                                RecordName = DefaultName & "&nbsp;" & RecordID
                            End If
                            If IsAdmin Then
                                '
                                ' Edit row
                                '
                                Cells(RowPointer, CPtr) = "<a href=""" & AdminURL & "?cid=" & CCID & "&af=4&id=" & RecordID & """><img src=/cclib/images/IconContentEdit.gif width=18 height=22 border=0></a>"
                                CPtr = CPtr + 1
                            End If
                            Cells(RowPointer, CPtr) = "<a href=""?" & RQS & "&" & RequestNameMemberID & "=" & RecordID & """><img border=0 src=""/cclib/images/icons/ContactDetails.gif"" width=16 height=16></a>"
                            CPtr = CPtr + 1
                            Cells(RowPointer, CPtr) = RecordName
                            CPtr = CPtr + 1
                            Cells(RowPointer, CPtr) = CS.GetText("FirstName")
                            CPtr = CPtr + 1
                            Cells(RowPointer, CPtr) = CS.GetText("LastName")
                            CPtr = CPtr + 1
                            Cells(RowPointer, CPtr) = CS.GetText("OrgName")
                            CPtr = CPtr + 1
                            Cells(RowPointer, CPtr) = CS.GetText("phone")
                            CPtr = CPtr + 1
                            Cells(RowPointer, CPtr) = CS.GetText("email")
                            CPtr = CPtr + 1
                            Cells(RowPointer, CPtr) = VisitsCell
                            CPtr = CPtr + 1
                            LastVisit = CS.GetDate("LastVisit").Date
                            If LastVisit = Date.MinValue Then
                                Cells(RowPointer, CPtr) = "&nbsp;"
                            Else
                                Cells(RowPointer, CPtr) = LastVisit.ToString()
                            End If
                            CPtr = CPtr + 1
                            Cells(RowPointer, CPtr) = "&nbsp;"
                            CPtr = CPtr + 1
                            RowPointer = RowPointer + 1
                            Call CS.GoNext()
                        Loop
                    End If
                    Call CS.Close()
                    result = result & cp.Html.Hidden("M.Count", RowPointer.ToString())
                    '
                    BlankPanel = "<div class=""cmBody ccPanel3DReverse"">x</div>"
                    'Main.GetPanel("x", "ccPanelInput", "ccPanelShadow", "ccPanelHilite")
                    RowPageSize = "<TABLE border=0 cellpadding=4 cellspacing=0 width=500>" _
                & "<TR>" _
                & "<TD class=APLeft>Rows Per Page</TD>" _
                & "<TD class=apright>" & cp.Html.InputText(RequestNamePageSize, PageSize.ToString(), "1", "10") & "</TD>" _
                & "</TR>" _
                & "</Table>"
                    RowGroups = "<TABLE border=0 cellpadding=4 cellspacing=0 width=500><TR>" _
                & "<TD valign=top class=APLeft>Actions</TD>" _
                & "<TD class=APRight>" _
                    & "" _
                    & "<div class=APRight>Source Contacts</div>" _
                    & "<div class=APRightIndent>" & cp.Html.RadioBox("GroupToolSelect", 0, 0) & "&nbsp;Only those selected on this page</div>" _
                    & "<div class=APRightIndent><input type=radio name=GroupToolSelect value=1>&nbsp;Everyone in search results</div>" _
                    & "<div style=""border-top:1px solid black;border-bottom:1px solid white;margin-top:4px;margin-bottom:4px;""></div>" _
                    & "<div class=APRight>Perform Action</div>" _
                    & "<div class=APRightIndent>" & cp.Html.RadioBox("GroupToolAction", 0, 0) & " No Action</div>" _
                    & "<div class=APRightIndent>" & cp.Html.RadioBox("GroupToolAction", 1, 0) & " Add to Target Group</div>" _
                    & "<div class=APRightIndent>" & cp.Html.RadioBox("GroupToolAction", 2, 0) & " Remove from Target Group</div>" _
                    & "<div class=APRightIndent>" & cp.Html.RadioBox("GroupToolAction", 3, 0) & " Export comma delimited file</div>" _
                    & "<div class=APRightIndent>" & cp.Html.RadioBox("GroupToolAction", 4, 0) & " Set Allow Group Email</div>" _
                    & "<div style=""border-top:1px solid black;border-bottom:1px solid white;margin-top:4px;margin-bottom:4px;""></div>" _
                    & "<div class=APRight style=""padding-bottom:6px;"">Target Group</div>" _
                    & "<div class=APRightIndent>" & cp.Html.SelectContent("GroupID", GroupID.ToString(), "Groups") & "</div>" _
                    & "</TD>" _
                & "</TR></Table>"

                    '& "<div class=APRightIndent>" & cp.Html.RadioBox("GroupToolSelect", 1, 0) & " Everyone in search results</div>" _

                    ActionPanel = "" _
                & vbCrLf _
                & "<style>" _
                & ".APLeft{width:100px;text-align:left;}" _
                & ".APRight{text-align:left;}" _
                & ".APRightIndent{text-align:left;padding-left:10px;}" _
                & "</style>" _
                & vbCrLf _
                & Replace(BlankPanel, "x", RowPageSize) _
                & Replace(BlankPanel, "x", RowGroups) _
                & ""
                    PostTableCopy = "" _
                & "<div class=""cmBody ccPanel3D"">" & ActionPanel & "</div>" _
                & cp.Html.Hidden("M.Count", RowPointer.ToString()) _
                & cp.Html.Hidden(RequestNameFormID, Convert.ToInt32(FormIdEnum.FormList).ToString()) _
                & ""
                    '
                    ' Header
                    '
                    Description = Description _
                & SearchCaption & "<BR>" & DataRowCount & " Matches found" _
                & ae.StatusMessage _
                & vbCrLf
                    Header = HtmlController.getPanel("<P>" & Description & "</P>", "ccPanel", "ccPanelShadow", "ccPanelHilite", "100%", 20)
                    Header = "<div class=ccPanelBackground style=""padding:10px;"">" & Header & "</div>"
                    ButtonList = ButtonApply & "," & ButtonNewSearch
                    Body = AdminUIController.getReport2(cp, RowPointer, ColCaption, ColAlign, ColWidth, Cells, PageSize, PageNumber, PreTableCopy, PostTableCopy, DataRowCount, "ccPanel", ColSortable, SortColPtr)
                    'Body = Controllers.AdminUIController.GetReport2(Main, RowPointer, ColCaption, ColAlign, ColWidth, Cells, PageSize, PageNumber, PreTableCopy, PostTableCopy, DataRowCount, "ccPanel", ColSortable, DefaultSortColumnPtr)
                    'Body = "<div style=""Background-color:white;"">" & Body & "</div>"
                    '
                    ' Assemble page
                    '
                    result = AdminUIController.getBody(cp, "Contact Manager &gt;&gt; List People", ButtonList, "", True, True, Description, "", 0, Body)
                    '
                End If
            Catch ex As Exception
                cp.Site.ErrorReport(ex)
            End Try
            Return result
        End Function
        '
        '=================================================================================
        '
        Public Shared Sub BuildSearch(cp As CPBaseClass, ae As Controllers.ApplicationController, ByRef return_Criteria As String, ByRef return_SearchCaption As String)
            Try
                Dim LookupName As String
                Dim ContactContentID As Integer
                Dim ContactGroupCriteria As String
                Dim ContactSearchCriteria As String
                Dim CriteriaValues() As String
                Dim CriteriaCount As Integer
                Dim CriteriaPointer As Integer
                Dim FieldParms() As String
                Dim SQLValue As String
                Dim GroupIDs() As String
                Dim GroupDelimiter As String
                Dim GroupPtr As Integer
                Dim ContentName As String
                Dim GroupCriteria As String
                Dim SQLNow As String
                Dim SQL As String
                Dim CS As CPCSBaseClass = cp.CSNew()
                Dim FieldCaption As String
                Dim FieldContentLookupID As Integer
                Dim FieldContentLookupName As String
                Dim hint As String
                '
                ContactContentID = cp.Content.GetID("people") ' cp.Utils.EncodeInteger(cp.User.GetText("ContactContentID", cp.Content.GetID("people").ToString()))
                ContactGroupCriteria = ae.userProperties.contactGroupCriteria
                ContactSearchCriteria = ae.userProperties.contactSearchCriteria
                return_SearchCaption = ""
                '
                ' Search Criteria Fields  (crlf FieldName tab FieldType tab FieldVAlue tab Operator)
                '
                If ContactSearchCriteria <> "" Then
                    hint = hint & ",110"
                    CriteriaValues = Split(ContactSearchCriteria, vbCrLf)
                    CriteriaCount = UBound(CriteriaValues) + 1
                    For CriteriaPointer = 0 To CriteriaCount - 1
                        hint = hint & ",120"
                        If CriteriaValues(CriteriaPointer) <> "" Then
                            hint = hint & ",130"
                            FieldParms = Split(CriteriaValues(CriteriaPointer), vbTab)
                            If UBound(FieldParms) >= 3 Then
                                '
                                ' Look up caption
                                '
                                FieldCaption = ""
                                SQL = "select caption,LookupContentID from ccfields where name=" & cp.Db.EncodeSQLText(FieldParms(0)) & " and contentid=" & ContactContentID
                                CS.OpenSQL(SQL)
                                If CS.OK() Then
                                    FieldCaption = CS.GetText("caption")
                                    FieldContentLookupID = CS.GetInteger("LookupContentID")
                                End If
                                Call CS.Close()
                                If FieldCaption = "" Then
                                    FieldCaption = FieldParms(0)
                                    FieldContentLookupID = 0
                                End If
                                Select Case cp.Utils.EncodeInteger(FieldParms(1))
                                    Case FieldTypeLongText, FieldTypeText
                                        '
                                        ' text
                                        '
                                        hint = hint & ",140"
                                        Select Case FieldParms(3)
                                            Case "1"
                                                '
                                                ' is empty
                                                '
                                                hint = hint & ",150"
                                                return_Criteria = return_Criteria & "AND((ccMembers." & FieldParms(0) & " is null)or(ccMembers." & FieldParms(0) & "=''))"
                                                return_SearchCaption = return_SearchCaption & ", " & FieldCaption & " is empty"
                                            Case "2"
                                                '
                                                ' is not empty
                                                '
                                                hint = hint & ",160"
                                                return_Criteria = return_Criteria & "AND(ccMembers." & FieldParms(0) & " is not null)"
                                                return_SearchCaption = return_SearchCaption & ", " & FieldCaption & " is not empty"
                                            Case "3"
                                                '
                                                ' includes a string
                                                '
                                                hint = hint & ",170"
                                                If FieldParms(2) <> "" Then
                                                    SQLValue = cp.Db.EncodeSQLText(FieldParms(2))
                                                    SQLValue = "'%" & Mid(SQLValue, 2, Len(SQLValue) - 2) & "%'"
                                                    return_Criteria = return_Criteria & "AND(ccMembers." & FieldParms(0) & " like " & SQLValue & ")"
                                                    return_SearchCaption = return_SearchCaption & ", " & FieldCaption & " includes '" & FieldParms(2) & "'"
                                                End If
                                        End Select
                                    Case FieldTypeLookup
                                        '
                                        ' Lookup
                                        '
                                        hint = hint & ",180"
                                        LookupName = ""
                                        If FieldContentLookupID <> 0 Then
                                            FieldContentLookupName = cp.Content.GetRecordName("Content", FieldContentLookupID)
                                            If (FieldContentLookupName <> "") And (FieldParms(2).IsNumeric) Then
                                                LookupName = cp.Content.GetRecordName(FieldContentLookupName, cp.Utils.EncodeInteger(FieldParms(2)))
                                            End If
                                        End If
                                        If LookupName = "" Then
                                            LookupName = FieldParms(2)
                                        End If

                                        return_Criteria = return_Criteria & "AND(ccMembers." & FieldParms(0) & "=" & FieldParms(2) & ")"
                                        return_SearchCaption = return_SearchCaption & ", " & FieldCaption & " = " & LookupName
                                    Case FieldTypeDate
                                        '
                                        ' date
                                        '
                                        hint = hint & ",185"
                                        Select Case cp.Utils.EncodeInteger(FieldParms(3))
                                            Case 1
                                                return_Criteria = return_Criteria & "AND(ccMembers." & FieldParms(0) & "=" & cp.Db.EncodeSQLDate(cp.Utils.EncodeDate(FieldParms(2))) & ")"
                                                return_SearchCaption = return_SearchCaption & ", " & FieldCaption & " = " & FieldParms(2)
                                            Case 2
                                                return_Criteria = return_Criteria & "AND(ccMembers." & FieldParms(0) & ">" & cp.Db.EncodeSQLDate(cp.Utils.EncodeDate(FieldParms(2))) & ")"
                                                return_SearchCaption = return_SearchCaption & ", " & FieldCaption & " > " & FieldParms(2)
                                            Case 3
                                                return_Criteria = return_Criteria & "AND(ccMembers." & FieldParms(0) & "<" & cp.Db.EncodeSQLDate(cp.Utils.EncodeDate(FieldParms(2))) & ")"
                                                return_SearchCaption = return_SearchCaption & ", " & FieldCaption & " < " & FieldParms(2)
                                        End Select
                                    Case FieldTypeCurrency, FieldTypeFloat, FieldTypeInteger, FieldTypeLookup
                                        '
                                        ' number
                                        '
                                        hint = hint & ",190"
                                        Select Case cp.Utils.EncodeInteger(FieldParms(3))
                                            Case 1
                                                return_Criteria = return_Criteria & "AND(ccMembers." & FieldParms(0) & "=" & cp.Db.EncodeSQLNumber(cp.Utils.EncodeNumber(FieldParms(2))) & ")"
                                                return_SearchCaption = return_SearchCaption & ", " & FieldCaption & " = " & FieldParms(2)
                                            Case 2
                                                return_Criteria = return_Criteria & "AND(ccMembers." & FieldParms(0) & ">" & cp.Db.EncodeSQLNumber(cp.Utils.EncodeNumber(FieldParms(2))) & ")"
                                                return_SearchCaption = return_SearchCaption & ", " & FieldCaption & " > " & FieldParms(2)
                                            Case 3
                                                return_Criteria = return_Criteria & "AND(ccMembers." & FieldParms(0) & "<" & cp.Db.EncodeSQLNumber(cp.Utils.EncodeNumber(FieldParms(2))) & ")"
                                                return_SearchCaption = return_SearchCaption & ", " & FieldCaption & " < " & FieldParms(2)
                                        End Select
                                    Case FieldTypeBoolean
                                        '
                                        ' boolean
                                        '
                                        hint = hint & ",200"
                                        If cp.Utils.EncodeBoolean(FieldParms(2)) Then
                                            '
                                            ' search for true
                                            '
                                            return_Criteria = return_Criteria & "AND(ccMembers." & FieldParms(0) & "<>0)AND(ccMembers." & FieldParms(0) & " is not null)"
                                        Else
                                            '
                                            ' search for false
                                            '
                                            return_Criteria = return_Criteria & "AND((ccMembers." & FieldParms(0) & "=0)or(ccMembers." & FieldParms(0) & " is null))"
                                        End If
                                        return_SearchCaption = return_SearchCaption & ", " & FieldCaption & " is " & cp.Utils.EncodeBoolean(FieldParms(2))
                                    Case Else
                                End Select
                            End If
                        End If
                    Next
                End If
                If return_SearchCaption <> "" Then
                    hint = hint & ",210"
                    return_SearchCaption = Mid(return_SearchCaption, 3)
                    return_SearchCaption = " where " & return_SearchCaption
                End If
                '
                ' Group Rules Criteria
                '
                hint = hint & ",220"
                If Left(ContactGroupCriteria, 1) = "," Then
                    ContactGroupCriteria = Mid(ContactGroupCriteria, 2)
                End If
                If Right(ContactGroupCriteria, 1) = "," Then
                    ContactGroupCriteria = Mid(ContactGroupCriteria, 1, Len(ContactGroupCriteria) - 1)
                End If
                If ContactGroupCriteria <> "" Then
                    hint = hint & ",230"
                    GroupIDs = Split(ContactGroupCriteria, ",")
                    GroupDelimiter = ""
                    SQLNow = cp.Db.EncodeSQLDate(Now)
                    If UBound(GroupIDs) = 0 Then
                        hint = hint & ",240"
                        If return_SearchCaption = "" Then
                            return_SearchCaption = " in group " & "'" & getGroupCaption(cp, GroupIDs(GroupPtr)) & "'"
                        Else
                            return_SearchCaption = return_SearchCaption & ", in group " & "'" & getGroupCaption(cp, GroupIDs(GroupPtr)) & "'"
                        End If
                        return_Criteria = return_Criteria & "AND((ccMemberRules.GroupID=" & GroupIDs(0) & ")and((ccMemberRules.DateExpires is null)or(ccMemberRules.DateExpires>" & SQLNow & ")))"
                    Else
                        hint = hint & ",250"
                        If return_SearchCaption <> "" Then
                            return_SearchCaption = return_SearchCaption & ", in group(s) "
                        Else
                            return_SearchCaption = " in group(s) "
                        End If
                        For GroupPtr = 0 To UBound(GroupIDs)
                            hint = hint & ",260"
                            GroupCriteria = GroupCriteria & "OR((ccMemberRules.GroupID=" & GroupIDs(GroupPtr) & ")and((ccMemberRules.DateExpires is null)or(ccMemberRules.DateExpires>" & SQLNow & ")))"
                            'Criteria = Criteria & "AND(ccMemberRules.GroupID=GroupIDs(GroupPtr))"
                            If GroupPtr = UBound(GroupIDs) And GroupPtr <> 0 Then
                                hint = hint & ",270"
                                return_SearchCaption = return_SearchCaption & " or " & "'" & getGroupCaption(cp, GroupIDs(GroupPtr)) & "'"
                            Else
                                hint = hint & ",280"
                                return_SearchCaption = return_SearchCaption & GroupDelimiter & "'" & getGroupCaption(cp, GroupIDs(GroupPtr)) & "'"
                            End If
                            GroupDelimiter = ", "
                        Next
                        return_Criteria = return_Criteria & "and(" & Mid(GroupCriteria, 3) & ")"
                    End If
                End If
                '
                ' Add Content Criteria
                '
                hint = hint & ",300"
                If return_Criteria <> "" Then
                    return_Criteria = " WHERE " & Mid(return_Criteria, 4)
                End If
                Exit Sub
            Catch ex As Exception
                cp.Site.ErrorReport(ex)
            End Try

            '
        End Sub
        '
        '=================================================================================
        '
        Private Shared Function getGroupCaption(cp As CPBaseClass, GroupID As String) As String
            Dim group As Models.GroupModel = Models.GroupModel.create(cp, GroupID)
            If (group IsNot Nothing) Then
                Return If(String.IsNullOrWhiteSpace(group.name), group.Caption, group.name)
            End If
            Return ""
        End Function
    End Class
End Namespace





