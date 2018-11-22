
Option Strict On
Option Explicit On

Imports Contensive.BaseClasses
Imports Contensive.Addons.ContactManager.GenericController
Imports System.Text
Imports System.Linq

Namespace Views
    Public Class CMngrClass
        Inherits AddonBaseClass
        '
        '=====================================================================================
        ''' <summary>
        ''' AddonDescription
        ''' </summary>
        ''' <param name="CP"></param>
        ''' <returns></returns>
        Public Overrides Function Execute(ByVal CP As CPBaseClass) As Object
            Dim result As String = ""
            Dim sw As New Stopwatch : sw.Start()
            Try
                '
                ' -- initialize application. If authentication needed and not login page, pass true
                Using ae As New Controllers.applicationController(CP, False)
                    '
                    ' -- your code
                    result = GetContent(CP)
                    If ae.packageErrorList.Count > 0 Then
                        result = "Hey user, this happened - " & Join(ae.packageErrorList.ToArray, "<br>")
                    End If
                End Using
            Catch ex As Exception
                CP.Site.ErrorReport(ex)
            End Try
            Return result
        End Function
        '
        '=================================================================================
        '
        Public Function GetContent(cp As CPBaseClass) As String
            Dim result As String = ""
            Try
                Dim IsAdminPath As Boolean = True
                Dim RQS As String = cp.Doc.RefreshQueryString
                Dim TabNumber As Integer = cp.Doc.GetInteger("tab")
                Dim DetailMemberID As Integer = cp.Doc.GetInteger(RequestNameMemberID)
                Dim ContactSearchCriteria As String = cp.User.GetText("ContactSearchCriteria")
                Dim ContactGroupCriteria As String = cp.User.GetText("ContactGroupCriteria")
                '
                Call cp.Doc.AddRefreshQueryString("tab", TabNumber.ToString())
                '
                Dim FormID As Integer = cp.Doc.GetInteger(RequestNameFormID)
                Dim Button As String = cp.Doc.GetText("Button")
                Dim StatusMessage As String
                If FormID = 0 Then
                    '
                    ' ----- Set default form
                    '
                    If (DetailMemberID <> 0) Then
                        '
                        ' Detail form
                        '
                        FormID = FormDetails
                    ElseIf (ContactSearchCriteria <> "") Or (ContactGroupCriteria <> "") Then
                        '
                        ' List form
                        '
                        FormID = FormList
                    Else
                        '
                        ' Search Form
                        '
                        FormID = FormSearch
                    End If
                ElseIf Button <> "" Then
                    '
                    ' ----- Process Previous Forms
                    '
                    Select Case FormID
                        Case FormSearch
                            '
                            ' Search Form
                            '
                            Select Case Button
                                Case ButtonSearch
                                    FormID = ProcessSearchForm(cp)
                                Case ButtonCancel
                                    Exit Function
                            End Select
                        Case FormList
                            '
                            ' List Form
                            '
                            Select Case Button
                                Case ButtonNewSearch
                                    ContactSearchCriteria = ""
                                    ContactGroupCriteria = ""
                                    Call cp.User.SetProperty("ContactSearchCriteria", ContactSearchCriteria)
                                    Call cp.User.SetProperty("ContactGroupCriteria", ContactGroupCriteria)
                                    FormID = FormSearch
                                Case ButtonApply
                                    '
                                    ' Add to or remove from group
                                    '
                                    Dim RowCount As Integer = cp.Doc.GetInteger("M.Count")
                                    Dim GroupID As Integer = cp.Doc.GetInteger("GroupID")
                                    Dim GroupToolAction As Integer = cp.Doc.GetInteger("GroupToolAction")
                                    Dim GroupToolSelect As Integer = cp.Doc.GetInteger("GroupToolSelect")
                                    Dim SQLCriteria As String = ""
                                    Dim SearchCaption As String
                                    Call BuildSearch(cp, SQLCriteria, SearchCaption)
                                    Dim GroupName As String
                                    Dim RowPointer As Integer
                                    Dim memberID As Integer
                                    Dim SQL As String
                                    Select Case GroupToolAction
                                        Case 1
                                            '
                                            ' ----- Add to Group
                                            '
                                            If (GroupID = 0) Then
                                                '
                                                ' Group required and not provided
                                                '
                                                Call cp.UserError.Add("Please select a Target Group for this operation")
                                            ElseIf GroupToolSelect = 0 Then
                                                '
                                                ' Add selection to Group
                                                '
                                                If (RowCount > 0) Then
                                                    GroupName = cp.Group.GetName(GroupID.ToString())
                                                    For RowPointer = 0 To RowCount - 1
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
                                                            & " select 1," & CCID & "," & GroupID & ",ccMembers.ID" _
                                                            & " from (ccMembers" _
                                                            & " left join ccMemberRules on ccMemberRules.MemberID=ccMembers.ID )" _
                                                            & " left join ( select MemberID  from ccMemberRules where GroupID in (" & GroupID & ")) as InGroups on InGroups.MemberID=ccMembers.ID" _
                                                            & " " & SQLCriteria _
                                                            & " and InGroups.MemberID is null" _
                                                            & ""
                                                Call cp.Db.ExecuteNonQuery(SQL)
                                            End If
                                        Case 2
                                            '
                                            ' ----- Remove From Group
                                            '
                                            If (GroupID = 0) Then
                                                '
                                                ' Group required and not provided
                                                '
                                                Call cp.UserError.Add("Please select a Target Group for this operation")
                                            ElseIf GroupToolSelect = 0 Then
                                                '
                                                ' Remove selection from Group
                                                '
                                                If (RowCount > 0) Then
                                                    GroupName = cp.Group.GetName(GroupID.ToString())
                                                    For RowPointer = 0 To RowCount - 1
                                                        If cp.Doc.GetBoolean("M." & RowPointer) Then
                                                            memberID = cp.Doc.GetInteger("MID." & RowPointer)
                                                            Call cp.Content.Delete("Member Rules", "(GroupID=" & GroupID & ")and(MemberID=" & memberID & ")")
                                                        End If
                                                    Next
                                                End If
                                            Else
                                                '
                                                ' Remove everyone in search criteria from this group
                                                '
                                                SQL = "delete from ccMemberRules where GroupID=" & GroupID & " and MemberID in (" _
                                                            & " select ccMembers.ID" _
                                                            & " from (ccMembers" _
                                                            & " left join ccMemberRules on ccMemberRules.MemberID=ccMembers.ID )" _
                                                            & " " & SQLCriteria _
                                                            & ")"
                                                Call cp.Db.ExecuteNonQuery(SQL)
                                            End If
                                        Case 3
                                            '
                                            ' ----- Export
                                            '
                                            Dim Aborttask As Boolean = False
                                            If True Then
                                                Dim ExportName As String = SearchCaption
                                                If GroupToolSelect = 0 Then
                                                    Dim RowSQL As String
                                                    '
                                                    ' Export selection from Group
                                                    '
                                                    ExportName = "Selected rows from " & ExportName
                                                    RowSQL = ""
                                                    If (RowCount > 0) Then
                                                        GroupName = cp.Group.GetName(GroupID.ToString())
                                                        For RowPointer = 0 To RowCount - 1
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
                                                            StatusMessage = "<P>You requested to only download the selected entries, and none were selected.<P>"
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
                                                        StatusMessage = "<P>There was a problem requesting your download.<P>"
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
                                        Case 4
                                            '
                                            ' ----- Set AllowBulkEmail field
                                            '
                                            If GroupToolSelect = 0 Then
                                                '
                                                ' Just selection
                                                '
                                                Dim RecordCnt As Integer = 0
                                                If (RowCount > 0) Then
                                                    GroupName = cp.Group.GetName(GroupID.ToString())
                                                    For RowPointer = 0 To RowCount - 1
                                                        If cp.Doc.GetBoolean("M." & RowPointer) Then
                                                            memberID = cp.Doc.GetInteger("MID." & RowPointer)
                                                            Call cp.Db.ExecuteNonQuery("update ccMembers set AllowBulkEmail=1 where ID=" & memberID)
                                                            RecordCnt = RecordCnt + 1
                                                        End If
                                                    Next
                                                End If
                                                StatusMessage = "<P>Allow Group Email was set for " & RecordCnt & " people.<P>"
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
                                                StatusMessage = "<P>Allow Group Email was set for all people in this selection.<P>"
                                            End If
                                    End Select
                            End Select
                        Case FormDetails
                            DetailMemberID = cp.Doc.GetInteger(RequestNameMemberID)
                            Select Case Button
                                Case ButtonCancel
                                    FormID = FormList
                                Case ButtonSave
                                    Call SaveContactFromStream(cp, DetailMemberID)
                                    FormID = FormDetails
                                Case ButtonOK
                                    Call SaveContactFromStream(cp, DetailMemberID)
                                    FormID = FormList
                                Case ButtonNewSearch
                                    FormID = FormSearch
                            End Select
                    End Select
                End If
                '
                ' ----- Output the next form
                Select Case FormID
                    Case FormDetails
                        '
                        ' Detail form
                        '
                        result = result & GetFormDetail(cp, DetailMemberID, StatusMessage)
                    Case FormList
                        '
                        ' List form
                        '
                        result = result & GetFormList(cp, StatusMessage, IsAdminPath)
                    Case Else
                        '
                        ' Search Form
                        '
                        result = result & GetFormSearch(cp, StatusMessage, IsAdminPath)
                End Select
                result = "<div class=ccbodyadmin>" & result & "</div>"
                '
                ' wrapper for style strength
                result = "" _
                    & vbCrLf & vbTab & "<div class=""contactManager"">" _
                    & nop(result) _
                    & vbCrLf & vbTab & "</div>" _
                    & ""
            Catch ex As Exception
                cp.Site.ErrorReport(ex)
            End Try
            Return result
        End Function
        '
        '=================================================================================
        '
        Private Function GetFormSearch(cp As CPBaseClass, StatusMessage As String, IsAdminPath As Boolean) As String
            Dim result As String = ""
            Try
                Dim SubTab As Integer
                Dim Nav As New TabController()
                Dim Header As String
                Dim Content As String
                Dim ButtonList As String
                '
                If True Then
                    '
                    ' Determine current Subtab
                    '
                    SubTab = cp.Doc.GetInteger("SubTab")
                    If SubTab = 0 Then
                        SubTab = cp.Utils.EncodeInteger(cp.User.GetText("SelectSubTab", "1"))
                        If SubTab = 0 Then
                            SubTab = 1
                            Call cp.User.SetProperty("SelectSubTab", CStr(SubTab))
                        End If
                    Else
                        Call cp.User.SetProperty("SelectSubTab", CStr(SubTab))
                    End If
                    Call cp.Doc.AddRefreshQueryString("SubTab", SubTab.ToString)
                    '
                    ' SubTab Menu
                    '
                    Call cp.Doc.AddRefreshQueryString("tab", "")
                    If IsAdminPath Then
                        ButtonList = ButtonCancelAll & "," & ButtonSearch
                    Else
                        ButtonList = ButtonSearch
                    End If
                    '
                    Header = "<div>Use the selections in each tab below to create a criteria for your search and hit the Search button.</div>"
                    '
                    Call Nav.addEntry("Record&nbsp;Fields", GetFormSearch_TabPeople(cp), "ccAdminTab")
                    Call Nav.addEntry("Groups", GetFormSearch_TabGroup(cp), "ccAdminTab")
                    '
                    Content = "" _
                        & cp.Html.Hidden("SelectionForm", "1") _
                        & Nav.getTabs(cp) _
                        & cp.Html.Hidden(RequestNameFormID, FormSearch.ToString()) _
                        & ""
                    GetFormSearch = AdminUIController.getBody(cp, "Contact Manager &gt;&gt; Selection Criteria", ButtonList, "", True, True, Header, "", 0, Content)
                End If
                result = GetFormSearch
            Catch ex As Exception
                cp.Site.ErrorReport(ex)
            End Try
            Return result
        End Function
        '
        '=================================================================================
        '
        Private Function GetFormSearch_TabGroup(cp As CPBaseClass) As String
            Dim s As String = ""
            Try
                '
                Dim Argument1 As String
                Dim CS As CPCSBaseClass = cp.CSNew()
                Dim ContactGroupCriteria As String
                Dim GroupCount As Integer
                Dim GroupPointer As Integer
                Dim GroupChecked As Boolean
                Dim GroupLabel As String
                Dim GroupID As Integer
                Dim GroupName As String
                Dim RowEven As Boolean
                Dim Button As String
                Dim SQL As String
                Dim RQS As String
                Dim SubTab As Integer
                Dim FormSave As Boolean
                Dim FormClear As Boolean
                Dim Style As String
                '
                If True Then
                    Button = "GroupSelect"
                    RQS = cp.Doc.RefreshQueryString
                    ContactGroupCriteria = cp.User.GetText("ContactGroupCriteria", "")
                    '
                    s = s _
                        & "<div>Select groups to narrow your results. If any groups are selected, your search will be limited to people in any of the selected groups.</div>" _
                        & "<div>&nbsp;</div>"
                    '
                    ' Add headers to stream
                    '
                    s = s & "<table border=0 width=100% cellspacing=0 cellpadding=4>"
                    s = s & "<tr>"
                    s = s & "<TD width=30><img src=/cclib/images/spacer.gif width=20 height=1></TD>"
                    s = s & "<TD width=30><img src=/cclib/images/spacer.gif width=20 height=1></TD>"
                    s = s & "<TD width=99%><img src=/cclib/images/spacer.gif width=1 height=1></TD>"
                    s = s & "</tr>"
                    '
                    s = s & "<tr>"
                    s = s & "<TD width=30 align=center class=ccAdminListCaption>Select</TD>"
                    s = s & "<TD width=30 align=center class=ccAdminListCaption>Count</TD>"
                    s = s & "<TD width=99% align=left class=ccAdminListCaption>Group Name</TD>"
                    s = s & "</tr>"
                    '
                    SQL = "SELECT ccGroups.ID as GroupID, ccGroups.Name as GroupName, ccGroups.Caption as GroupCaption, Count(ccMembers.ID) AS CountOfID" _
                        & " FROM (ccGroups LEFT JOIN ccMemberRules ON ccGroups.ID = ccMemberRules.GroupID) LEFT JOIN ccMembers ON ccMemberRules.MemberID = ccMembers.ID" _
                        & " Where (((ccMemberRules.DateExpires) Is Null Or (ccMemberRules.DateExpires) > " & cp.Db.EncodeSQLDate(Now()) & "))" _
                        & " GROUP BY ccGroups.ID, ccGroups.Name, ccGroups.Caption" _
                        & " ORDER BY ccGroups.Caption;"
                    CS.OpenSQL(SQL)
                    GroupPointer = 0
                    Do While CS.OK()
                        GroupID = CS.GetInteger("GroupID")
                        GroupName = CS.GetText("GroupCaption")
                        If GroupName = "" Then
                            GroupName = CS.GetText("GroupName")
                            If GroupName = "" Then
                                GroupName = "Group " & GroupID
                            End If
                        End If
                        GroupLabel = "Group" & GroupPointer
                        GroupChecked = (InStr(1, ContactGroupCriteria, "," & GroupID & ",") <> 0)

                        If ((GroupPointer Mod 2) = 0) Then
                            Style = "ccAdminListRowEven"
                        Else
                            Style = "ccAdminListRowOdd"
                        End If
                        s = s & "<TR>"
                        s = s & "<TD width=30 align=center class=""" & Style & """>" & cp.Html.CheckBox(GroupLabel, GroupChecked) & cp.Html.Hidden(GroupLabel & ".id", GroupID.ToString()) & "</TD>"
                        s = s & "<TD width=30 align=right class=""" & Style & """>" & CS.GetInteger("CountOfID") & "</TD>"
                        s = s & "<TD width=99% align=left class=""" & Style & """>" & GroupName & "</TD>"
                        s = s & "</TR>"
                        GroupPointer = GroupPointer + 1
                        Call CS.GoNext()
                    Loop
                    Call CS.Close()
                    s = s & "</Table>"
                    '
                    s = s & cp.Html.Hidden("GroupCount", GroupPointer.ToString())
                    GetFormSearch_TabGroup = s & cp.Html.Hidden("SelectionGroupSubTab", "1")
                    '
                End If
            Catch ex As Exception
                cp.Site.ErrorReport(ex)
            End Try
            Return s
        End Function
        '
        '=================================================================================
        '
        Private Function GetFormSearch_TabPeople(cp As CPBaseClass) As String
            Dim s As String = ""
            Try
                '
                Dim Argument1 As String
                Dim CS As CPCSBaseClass = cp.CSNew()
                Dim ContactSearchCriteria As String
                Dim FieldPtr As Integer
                Dim GroupChecked As Boolean
                Dim GroupLabel As String
                Dim GroupID As Integer
                Dim RowEven As Boolean
                Dim Button As String
                Dim SQL As String
                Dim RQS As String
                Dim SubTab As Integer
                Dim Criteria As String
                Dim VisitSearchCriteria As String
                Dim CriteriaNames() As String
                Dim CriteriaValues() As String
                Dim CriteriaCount As Integer
                Dim CriteriaPointer As Integer
                Dim FieldName() As String
                Dim FieldCaption() As String
                Dim fieldId() As Integer
                Dim fieldType() As Integer
                Dim FieldValue() As String
                Dim FieldOperator() As Integer
                Dim FieldLookupContentName() As String
                Dim FieldLookupList() As String
                Dim fieldEditTab() As String
                Dim ContentID As Integer
                Dim currentEditTab As String
                Dim lastEditTab As String
                Dim FieldCount As Integer
                Dim FieldSize As Integer
                Dim FormSave As Boolean
                Dim FormClear As Boolean
                Dim NameValues() As String
                Dim RowPointer As Integer
                Dim tabCaption As String
                Dim groupTab As String
                '
                If True Then
                    Button = "CriteriaSelect"
                    RQS = cp.Doc.RefreshQueryString
                    '
                    ' prepare visit property to prepopulate form
                    '
                    ContactSearchCriteria = cp.User.GetText("ContactSearchCriteria", "")
                    CriteriaValues = {}
                    If ContactSearchCriteria <> "" Then
                        CriteriaValues = Split(ContactSearchCriteria, vbCrLf)
                        CriteriaCount = UBound(CriteriaValues) + 1
                    End If
                    '
                    ' Setup fields and capture request changes
                    '
                    Criteria = "(active<>0)and(ContentID=" & cp.Content.GetID("people") & ")and(authorable<>0)"
                    CS.Open("Content Fields", Criteria, "EditTab,EditSortPriority")
                    FieldPtr = 0
                    ReDim Preserve FieldName(FieldSize)
                    ReDim Preserve FieldCaption(FieldSize)
                    ReDim Preserve fieldId(FieldSize)
                    ReDim Preserve fieldType(FieldSize)
                    ReDim Preserve FieldValue(FieldSize)
                    ReDim Preserve FieldOperator(FieldSize)
                    ReDim Preserve FieldLookupContentName(FieldSize)
                    ReDim Preserve FieldLookupList(FieldSize)
                    ReDim Preserve fieldEditTab(FieldSize)
                    Do While CS.OK()
                        If FieldPtr >= FieldSize Then
                            FieldSize = FieldSize + 100
                            ReDim Preserve FieldName(FieldSize)
                            ReDim Preserve FieldCaption(FieldSize)
                            ReDim Preserve fieldId(FieldSize)
                            ReDim Preserve fieldType(FieldSize)
                            ReDim Preserve FieldValue(FieldSize)
                            ReDim Preserve FieldOperator(FieldSize)
                            ReDim Preserve FieldLookupContentName(FieldSize)
                            ReDim Preserve FieldLookupList(FieldSize)
                            ReDim Preserve fieldEditTab(FieldSize)
                        End If
                        FieldName(FieldPtr) = CS.GetText("name")
                        FieldCaption(FieldPtr) = CS.GetText("Caption")
                        fieldId(FieldPtr) = CS.GetInteger("ID")
                        fieldType(FieldPtr) = CS.GetInteger("Type")
                        If fieldType(FieldPtr) = 7 Then
                            ContentID = CS.GetInteger("LookupContentID")
                            If ContentID > 0 Then
                                FieldLookupContentName(FieldPtr) = cp.Content.GetRecordName("content", ContentID)
                            End If
                            FieldLookupList(FieldPtr) = CS.GetText("LookupList")
                        End If
                        fieldEditTab(FieldPtr) = CS.GetText("editTab")
                        '
                        ' set prepoplate value from visit property
                        '
                        If CriteriaCount > 0 Then
                            For CriteriaPointer = 0 To CriteriaCount - 1
                                FieldOperator(FieldPtr) = 0
                                If InStr(1, CriteriaValues(CriteriaPointer), FieldName(FieldPtr) & "=", vbTextCompare) = 1 Then
                                    NameValues = Split(CriteriaValues(CriteriaPointer), "=")
                                    FieldValue(FieldPtr) = NameValues(1)
                                    FieldOperator(FieldPtr) = 1
                                ElseIf InStr(1, CriteriaValues(CriteriaPointer), FieldName(FieldPtr) & ">", vbTextCompare) = 1 Then
                                    NameValues = Split(CriteriaValues(CriteriaPointer), ">")
                                    FieldValue(FieldPtr) = NameValues(1)
                                    FieldOperator(FieldPtr) = 2
                                ElseIf InStr(1, CriteriaValues(CriteriaPointer), FieldName(FieldPtr) & "<", vbTextCompare) = 1 Then
                                    NameValues = Split(CriteriaValues(CriteriaPointer), "<")
                                    FieldValue(FieldPtr) = NameValues(1)
                                    FieldOperator(FieldPtr) = 3
                                End If
                            Next
                        End If
                        FieldPtr = FieldPtr + 1
                        Call CS.GoNext()
                    Loop
                    Call CS.Close()
                    FieldCount = FieldPtr
                    '
                    ' header
                    '
                    s = s _
                        & "<div>Enter criteria for each field to identify and select your results. The results of a search will have to have all of the criteria you enter.</div>" _
                        & "<div>&nbsp;</div>"
                    '
                    ' Add headers to stream
                    '
                    s = s & "<table border=0 width=100% cellspacing=0 cellpadding=4>"
                    s = s & "<tr>"
                    s = s & AdminUIController.kmaStartTableCell("120", 1, RowEven, "right") & "<img src=/cclib/images/spacer.gif width=120 height=1></TD>"
                    s = s & AdminUIController.kmaStartTableCell("99%", 1, RowEven, "left") & "<img src=/cclib/images/spacer.gif width=1 height=1></TD>"
                    s = s & "</tr>"
                    '
                    RowPointer = 0
                    lastEditTab = "-1"
                    currentEditTab = ""
                    groupTab = ""
                    For FieldPtr = 0 To FieldCount - 1
                        currentEditTab = fieldEditTab(FieldPtr)
                        If currentEditTab <> lastEditTab Then
                            If currentEditTab = "" Then
                                tabCaption = "Details"
                            Else
                                tabCaption = currentEditTab
                            End If
                            groupTab = "" _
                        & "<tr>" _
                        & "<td colspan=""2"" class=""cmRowTab"">" & tabCaption & "</TD>" _
                        & "</TR>"
                            RowPointer = RowPointer + 1
                            lastEditTab = currentEditTab
                        End If
                        RowEven = ((RowPointer Mod 2) = 0)
                        Select Case fieldType(FieldPtr)
                            Case FieldTypeDate
                                '
                                ' Date
                                '
                                s = s & groupTab _
                            & "<TR>" _
                            & "<td class=""cmRowCaption"">" & FieldCaption(FieldPtr) & "</TD>" _
                            & "<TD class=""cmRowField"">" _
                            & "<table border=0 width=100% cellspacing=0 cellpadding=0><TR>" _
                                & "<TD width=10 align=right>" & GetFormInputRadioBox(cp, FieldName(FieldPtr) & "_D", "0", FieldOperator(FieldPtr).ToString(), "") & "</TD><TD align=left width=100>ignore</TD>" _
                                & "<TD width=10>&nbsp;&nbsp;</TD><TD width=10 align=right>" & GetFormInputRadioBox(cp, FieldName(FieldPtr) & "_D", "1", FieldOperator(FieldPtr).ToString(), "") & "</TD><TD align=left width=100>=</TD>" _
                                & "<TD width=10>&nbsp;&nbsp;</TD><TD width=10 align=right>" & GetFormInputRadioBox(cp, FieldName(FieldPtr) & "_D", "2", FieldOperator(FieldPtr).ToString(), "") & "</TD><TD align=left width=100>&gt;</TD>" _
                                & "<TD width=10>&nbsp;&nbsp;</TD><TD width=10 align=right>" & GetFormInputRadioBox(cp, FieldName(FieldPtr) & "_D", "3", FieldOperator(FieldPtr).ToString(), "") & "</TD><TD align=left width=100>&lt;</TD>" _
                                & "<TD width=10>&nbsp;&nbsp;</TD><TD align=left width=99%>" & GetFormInputDate(cp, FieldName(FieldPtr), FieldValue(FieldPtr), "10", "", "") & "</TD>" _
                            & "</TR></Table>" _
                            & "</TD>" _
                            & "</TR>"
                                groupTab = ""
                                RowPointer = RowPointer + 1
                            Case FieldTypeCurrency, FieldTypeFloat, FieldTypeInteger
                                '
                                ' Numeric
                                '
                                s = s & groupTab _
                            & "<TR>" _
                            & "<td class=""cmRowCaption"">" & FieldCaption(FieldPtr) & "</TD>" _
                            & "<TD class=""cmRowField"">" _
                            & "<table border=0 width=100% cellspacing=0 cellpadding=0><TR>" _
                                & "<TD width=10 align=right>" & GetFormInputRadioBox(cp, FieldName(FieldPtr) & "_N", "0", FieldOperator(FieldPtr).ToString(), "") & "</TD><TD align=left width=100>ignore</TD>" _
                                & "<TD width=10>&nbsp;&nbsp;</TD><TD width=10 align=right>" & GetFormInputRadioBox(cp, FieldName(FieldPtr) & "_N", "1", FieldOperator(FieldPtr).ToString(), "") & "</TD><TD align=left width=100>=</TD>" _
                                & "<TD width=10>&nbsp;&nbsp;</TD><TD width=10 align=right>" & GetFormInputRadioBox(cp, FieldName(FieldPtr) & "_N", "2", FieldOperator(FieldPtr).ToString(), "") & "</TD><TD align=left width=100>&gt;</TD>" _
                                & "<TD width=10>&nbsp;&nbsp;</TD><TD width=10 align=right>" & GetFormInputRadioBox(cp, FieldName(FieldPtr) & "_N", "3", FieldOperator(FieldPtr).ToString(), "") & "</TD><TD align=left width=100>&lt;</TD>" _
                                & "<TD width=10>&nbsp;&nbsp;</TD><TD align=left width=99%>" & GetFormInputText(cp, FieldName(FieldPtr), FieldValue(FieldPtr), "1", "5", "", "") & "</TD>" _
                            & "</TR></Table>" _
                            & "</TD>" _
                            & "</TR>"
                                groupTab = ""
                                RowPointer = RowPointer + 1
                            Case FieldTypeBoolean
                                '
                                ' Boolean
                                '
                                s = s & groupTab _
                            & "<TR>" _
                            & "<td class=""cmRowCaption"">" & FieldCaption(FieldPtr) & "</TD>" _
                            & "<TD class=""cmRowField"">" _
                            & "<table border=0 width=100% cellspacing=0 cellpadding=0><TR>" _
                                & "<TD width=10 align=right>" & GetFormInputRadioBox(cp, FieldName(FieldPtr), "", FieldValue(FieldPtr), "") & "</TD><TD align=left width=100>" & "  ignore</TD>" _
                                & "<TD width=10>&nbsp;&nbsp;</TD><TD width=10 align=right>" & GetFormInputRadioBox(cp, FieldName(FieldPtr), "1", FieldValue(FieldPtr), "") & "</TD><TD align=left width=100>" & "true</TD>" _
                                & "<TD width=10>&nbsp;&nbsp;</TD><TD width=10 align=right>" & GetFormInputRadioBox(cp, FieldName(FieldPtr), "0", FieldValue(FieldPtr), "") & "</TD><TD align=left width=100>" & "false</TD>" _
                                & "<TD width=99%>&nbsp;</td>" _
                            & "</TR></Table>" _
                            & "</TD>" _
                            & "</TR>"
                                '& "[" & cp.Html.RadioBox(FieldName(FieldPtr), "", FieldValue(FieldPtr)) & "ignore]" _
                                '& "[" & cp.Html.RadioBox(FieldName(FieldPtr), "1", FieldValue(FieldPtr)) & "true]" _
                                '& "[" & cp.Html.RadioBox(FieldName(FieldPtr), "0", FieldValue(FieldPtr)) & "false]" _
                                RowPointer = RowPointer + 1
                                groupTab = ""
                            Case FieldTypeText, FieldTypeLongText
                                '
                                ' Text
                                '
                                s = s & groupTab _
                                    & "<TR>" _
                                    & "<td class=""cmRowCaption"">" & FieldCaption(FieldPtr) & "</TD>" _
                                    & "<TD class=""cmRowField"" valign=absmiddle>" _
                                    & "<table border=0 width=100% cellspacing=0 cellpadding=0><TR>" _
                                & "<TD width=10 align=right>" & GetFormInputRadioBox(cp, FieldName(FieldPtr) & "_T", "", FieldValue(FieldPtr), "") & "</TD><TD align=left width=100>" & "  ignore</TD>" _
                                & "<TD width=10>&nbsp;&nbsp;</TD><TD width=10 align=right>" & GetFormInputRadioBox(cp, FieldName(FieldPtr) & "_T", "1", FieldValue(FieldPtr), "") & "</TD><TD align=left width=100>" & "empty</TD>" _
                                & "<TD width=10>&nbsp;&nbsp;</TD><TD width=10 align=right>" & GetFormInputRadioBox(cp, FieldName(FieldPtr) & "_T", "2", FieldValue(FieldPtr), "") & "</TD><TD align=left width=100>" & "not&nbsp;empty</TD>" _
                                & "<TD width=10>&nbsp;&nbsp;</TD><TD width=10 align=right>" & GetFormInputRadioBox(cp, FieldName(FieldPtr) & "_T", "3", FieldValue(FieldPtr), "t" & FieldPtr) & "</TD><TD align=center width=100>" & "&nbsp;includes&nbsp;" & "</TD><TD align=left width=99%>" & GetFormInputText(cp, FieldName(FieldPtr), FieldValue(FieldPtr), "1", "20", "", "var e=getElementById('t" & FieldPtr & "');e.checked=1;") & "</TD>" _
                                    & "</TR></Table>" _
                                    & "</TD>" _
                                    & "</TR>"
                                groupTab = ""
                                RowPointer = RowPointer + 1
                            Case FieldTypeLookup
                                '
                                ' Lookup
                                '
                                s = s & groupTab _
                                    & "<TR><td class=""cmRowCaption"">" & FieldCaption(FieldPtr) & "</TD>" _
                                    & ""
                                If FieldLookupContentName(FieldPtr) <> "" Then
                                    s = s _
                                        & "<TD class=""cmRowField"">" _
                                        & cp.Html.SelectContent(FieldName(FieldPtr), FieldValue(FieldPtr), FieldLookupContentName(FieldPtr), , "Any") & "</TD>"
                                Else
                                    s = s _
                                        & "<TD class=""cmRowField"">" _
                                        & cp.Html.SelectList(FieldName(FieldPtr), FieldValue(FieldPtr), FieldLookupList(FieldPtr), , "Any") & "</TD>"
                                End If
                                s = s & "</TR>"
                                groupTab = ""
                                RowPointer = RowPointer + 1
                        End Select
                    Next
                    s = s & "</Table>"
                    '
                    GetFormSearch_TabPeople = s & cp.Html.Hidden("SelectionSearchSubTab", "1")
                    '
                End If
            Catch ex As Exception
                cp.Site.ErrorReport(ex)
            End Try
            Return s
        End Function
        '
        '=================================================================================
        '
        Private Function GetFormList(cp As CPBaseClass, StatusMessage As String, IsAdminPath As Boolean) As String
            Dim result As String = ""
            Try
                Dim SQLOrderDir As String
                Dim SQLOrderBy As String
                Dim LastVisit As Date
                Dim CheckBox As String
                Dim Argument1 As String
                Dim CS As CPCSBaseClass = cp.CSNew()
                Dim ContactGroupCriteria As String
                Dim GroupCount As Integer
                Dim GroupPointer As Integer
                Dim GroupChecked As Boolean
                Dim RecordName As String
                Dim ContentName As String
                Dim RecordID As Integer
                Dim RowEven As Boolean
                Dim Button As String
                Dim SQL As String
                Dim RQS As String
                Dim SubTab As Integer
                Dim FormSave As Boolean
                Dim FormClear As Boolean
                Dim ContactContentID As Integer
                '
                Dim Criteria As String
                Dim ContentGorupCriteria As String
                Dim ContactSearchCriteria As String
                Dim FieldParms() As String
                Dim CriteriaValues As Object
                Dim CriteriaCount As Integer
                Dim CriteriaPointer As Integer
                Dim PageSize As Integer
                Dim PageNumber As Integer
                Dim TopCount As Integer
                Dim RowPointer As Integer
                Dim DataRowCount As Integer
                '
                Dim PreTableCopy As String
                Dim PostTableCopy As String
                '
                Dim ColumnCount As Integer
                Dim CPtr As Integer
                Dim ColCaption() As String
                Dim ColAlign() As String
                Dim ColWidth() As String
                Dim Cells(,) As String
                Dim ColSortable() As Boolean
                Dim DefaultSortColumnPtr As Integer
                '
                Dim GroupID As Integer
                Dim GroupToolAction As Integer
                Dim ActionPanel As String
                Dim RowCount As Integer
                Dim GroupName As String
                Dim memberID As Integer
                Dim QS As String
                Dim VisitsCell As String
                Dim VisitCount As Integer
                Dim AdminURL As String
                Dim CCID As Integer
                Dim SQLValue As String
                Dim DefaultName As String
                Dim Copy As String
                Dim SearchCaption As String
                Dim BlankPanel As String
                Dim RowPageSize As String
                Dim RowGroups As String
                Dim GroupIDs() As String
                Dim GroupPtr As Integer
                Dim GroupDelimiter As String
                Dim ButtonList As String
                Dim ButtonBar As String
                Dim Header As String
                Dim Body As String
                Dim AdminUI As Object
                Dim Description As String
                Dim IsAdmin As Boolean
                Dim SortColPtr As Integer
                Dim SortColType As Integer
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
                    Call BuildSearch(cp, Criteria, SearchCaption)
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
                    GetFormList = GetFormList & cp.Html.Hidden("M.Count", RowPointer.ToString())
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
                & cp.Html.Hidden(RequestNameFormID, FormList.ToString()) _
                & ""
                    '
                    ' Header
                    '
                    Description = Description _
                & SearchCaption & "<BR>" & DataRowCount & " Matches found" _
                & StatusMessage _
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
                    GetFormList = AdminUIController.getBody(cp, "Contact Manager &gt;&gt; List People", ButtonList, "", True, True, Description, "", 0, Body)
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
        Private Function GetEventsTab() As String
            Return "events"
        End Function
        '
        '=================================================================================
        '
        Private Function ProcessSearchForm(cp As CPBaseClass) As Integer
            Dim result As Integer = 0
            Try
                Dim TextOption As String
                Dim NumericOption As String
                Dim Argument1 As String
                Dim CS As CPCSBaseClass = cp.CSNew()
                Dim SelectedGroupIDList As String
                Dim GroupCount As Integer
                Dim GroupPointer As Integer
                Dim GroupLabel As String
                Dim ContactGroupCriteria As String
                '
                Dim ContactSearchCriteria As String
                Dim FieldPointer As Integer
                Dim GroupChecked As Boolean
                Dim GroupID As Integer
                Dim RowEven As Boolean
                Dim Button As String
                Dim SQL As String
                Dim RQS As String
                Dim SubTab As Integer
                Dim Criteria As String
                Dim VisitSearchCriteria As String
                Dim CriteriaNames() As String
                Dim CriteriaValues() As String
                Dim CriteriaCount As Integer
                Dim CriteriaPointer As Integer
                Dim FieldName As String
                Dim FieldValue As String
                Dim FieldCount As Integer
                Dim FieldSize As Integer
                Dim FormSave As Boolean
                Dim FormClear As Boolean
                Dim NameValues As Object
                Dim RowPointer As Integer
                Dim FieldOperator As String
                Dim fieldType As Integer

                '
                Dim ContactContentID As Integer
                '
                If True Then
                    If cp.Doc.GetText("SelectionGroupSubTab") <> "" Then
                        '
                        ' Save the Form
                        '
                        GroupCount = cp.Doc.GetInteger("GroupCount")
                        If GroupCount > 0 Then
                            For GroupPointer = 0 To GroupCount - 1
                                GroupLabel = "Group" & GroupPointer
                                If cp.Doc.GetBoolean(GroupLabel) Then
                                    ContactGroupCriteria = ContactGroupCriteria & "," & cp.Doc.GetInteger(GroupLabel & ".id")
                                End If
                            Next
                        End If
                        ContactGroupCriteria = ContactGroupCriteria & ","
                        Call cp.User.SetProperty("ContactGroupCriteria", ContactGroupCriteria)
                    End If
                    If cp.Doc.GetText("SelectionContentSubTab") <> "" Then
                        '
                        ' SelectionContentSubTab
                        '
                        ContactContentID = cp.Doc.GetInteger("ContentID")
                        Call cp.User.SetProperty("ContactContentID", ContactContentID.ToString())
                    End If
                    If cp.Doc.GetText("SelectionSearchSubTab") <> "" Then
                        '
                        ' SelectionContentSubTab (crlf FieldName tab FieldType tab FieldVAlue tab Operator)
                        '
                        Criteria = "(active<>0)and(ContentID=" & cp.Content.GetID("people") & ")and(authorable<>0)"
                        CS.Open("Content Fields", Criteria, "EditSortPriority")
                        Do While CS.OK()
                            FieldName = CS.GetText("name")
                            FieldValue = cp.Doc.GetText(FieldName)
                            fieldType = CS.GetInteger("Type")
                            Select Case fieldType
                                Case FieldTypeDate
                                    NumericOption = cp.Doc.GetText(FieldName & "_D")
                                    If NumericOption <> "" Then
                                        ContactSearchCriteria = ContactSearchCriteria _
                                    & vbCrLf _
                                    & FieldName & vbTab _
                                    & fieldType & vbTab _
                                    & FieldValue & vbTab _
                                    & NumericOption
                                    End If
                                Case FieldTypeCurrency, FieldTypeFloat, FieldTypeInteger
                                    NumericOption = cp.Doc.GetText(FieldName & "_N")
                                    If NumericOption <> "" Then
                                        ContactSearchCriteria = ContactSearchCriteria _
                                    & vbCrLf _
                                    & FieldName & vbTab _
                                    & fieldType & vbTab _
                                    & FieldValue & vbTab _
                                    & NumericOption
                                    End If
                                Case FieldTypeBoolean
                                    If FieldValue <> "" Then
                                        ContactSearchCriteria = ContactSearchCriteria _
                                    & vbCrLf _
                                    & FieldName & vbTab _
                                    & fieldType & vbTab _
                                    & FieldValue & vbTab _
                                    & ""
                                    End If
                                Case FieldTypeText
                                    TextOption = cp.Doc.GetText(FieldName & "_T")
                                    If TextOption <> "" Then
                                        ContactSearchCriteria = ContactSearchCriteria _
                                    & vbCrLf _
                                    & FieldName & vbTab _
                                    & CStr(fieldType) & vbTab _
                                    & FieldValue & vbTab _
                                    & TextOption
                                    End If
                                Case FieldTypeLookup
                                    If FieldValue <> "" Then
                                        ContactSearchCriteria = ContactSearchCriteria _
                                    & vbCrLf _
                                    & FieldName & vbTab _
                                    & fieldType & vbTab _
                                    & FieldValue & vbTab _
                                    & ""
                                    End If
                            End Select
                            Call CS.GoNext()
                        Loop
                        Call CS.Close()
                        Call cp.User.SetProperty("ContactSearchCriteria", ContactSearchCriteria)
                    End If
                    ProcessSearchForm = FormList
                End If
            Catch ex As Exception
                cp.Site.ErrorReport(ex)
            End Try
            Return result
        End Function
        '
        '=================================================================================
        '
        'Private Function GetReportQueryString(cp As CPBaseClass, ReportPageNumber As Integer, ReportPageSize As Integer, memberID As Integer, ReportID As Integer) As String
        '    Dim result As String = ""
        '    Try
        '        GetReportQueryString = cp.Doc.RefreshQueryString
        '        GetReportQueryString = cp.Utils.ModifyQueryString(GetReportQueryString, "aa", "0", False)
        '        GetReportQueryString = cp.Utils.ModifyQueryString(GetReportQueryString, "af", AdminFormReports, True)
        '        GetReportQueryString = cp.Utils.ModifyQueryString(GetReportQueryString, "asf", AdminFormReports, True)
        '        GetReportQueryString = cp.Utils.ModifyQueryString(GetReportQueryString, "rid", CStr(ReportID), True)
        '        GetReportQueryString = cp.Utils.ModifyQueryString(GetReportQueryString, "DateTo", CStr(Int(Now + 1)), True)
        '        GetReportQueryString = cp.Utils.ModifyQueryString(GetReportQueryString, "DateFrom", "1/1/2000", True)
        '        GetReportQueryString = cp.Utils.ModifyQueryString(GetReportQueryString, "MemberID", CStr(memberID), True)
        '        GetReportQueryString = cp.Utils.ModifyQueryString(GetReportQueryString, "ExcludeBrowsers", "0", True)
        '        GetReportQueryString = cp.Utils.ModifyQueryString(GetReportQueryString, "ExcludeIP", "0", True)
        '        GetReportQueryString = cp.Utils.ModifyQueryString(GetReportQueryString, "ExcludeNewVisitors", "0", True)
        '        GetReportQueryString = cp.Utils.ModifyQueryString(GetReportQueryString, "ExcludeOldVisitors", "0", True)
        '        GetReportQueryString = cp.Utils.ModifyQueryString(GetReportQueryString, "PageNumber", CStr(ReportPageNumber), True)
        '        GetReportQueryString = cp.Utils.ModifyQueryString(GetReportQueryString, "PageSize", CStr(ReportPageSize), True)
        '    Catch ex As Exception
        '        cp.Site.ErrorReport(ex)
        '    End Try
        '    Return result
        'End Function
        '
        '=================================================================================
        '
        Private Function GetGroupCaption(cp As CPBaseClass, GroupID As String) As String
            Dim result As String = ""
            Try
                Dim CS As CPCSBaseClass = cp.CSNew()
                '
                CS.OpenRecord("groups", cp.Utils.EncodeInteger(GroupID), "Caption,name")
                If CS.OK() Then
                    GetGroupCaption = CS.GetText("caption")
                    If GetGroupCaption = "" Then
                        GetGroupCaption = CS.GetText("name")
                        If GetGroupCaption = "" Then
                            GetGroupCaption = "Group " & GroupID
                        End If
                    End If
                End If
                Call CS.Close()
                result = GetGroupCaption
            Catch ex As Exception
                cp.Site.ErrorReport(ex)
            End Try
            Return result
        End Function
        '
        '=================================================================================
        '
        Private Sub BuildSearch(cp As CPBaseClass, Criteria As String, SearchCaption As String)
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
                ContactContentID = cp.Utils.EncodeInteger(cp.User.GetText("ContactContentID", cp.Content.GetID("people").ToString()))
                ContactGroupCriteria = cp.User.GetText("ContactGroupCriteria", "")
                ContactSearchCriteria = cp.User.GetText("ContactSearchCriteria", "")
                SearchCaption = ""
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
                                                Criteria = Criteria & "AND((ccMembers." & FieldParms(0) & " is null)or(ccMembers." & FieldParms(0) & "=''))"
                                                SearchCaption = SearchCaption & ", " & FieldCaption & " is empty"
                                            Case "2"
                                                '
                                                ' is not empty
                                                '
                                                hint = hint & ",160"
                                                Criteria = Criteria & "AND(ccMembers." & FieldParms(0) & " is not null)"
                                                SearchCaption = SearchCaption & ", " & FieldCaption & " is not empty"
                                            Case "3"
                                                '
                                                ' includes a string
                                                '
                                                hint = hint & ",170"
                                                If FieldParms(2) <> "" Then
                                                    SQLValue = cp.Db.EncodeSQLText(FieldParms(2))
                                                    SQLValue = "'%" & Mid(SQLValue, 2, Len(SQLValue) - 2) & "%'"
                                                    Criteria = Criteria & "AND(ccMembers." & FieldParms(0) & " like " & SQLValue & ")"
                                                    SearchCaption = SearchCaption & ", " & FieldCaption & " includes '" & FieldParms(2) & "'"
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

                                        Criteria = Criteria & "AND(ccMembers." & FieldParms(0) & "=" & FieldParms(2) & ")"
                                        SearchCaption = SearchCaption & ", " & FieldCaption & " = " & LookupName
                                    Case FieldTypeDate
                                        '
                                        ' date
                                        '
                                        hint = hint & ",185"
                                        Select Case cp.Utils.EncodeInteger(FieldParms(3))
                                            Case 1
                                                Criteria = Criteria & "AND(ccMembers." & FieldParms(0) & "=" & cp.Db.EncodeSQLDate(cp.Utils.EncodeDate(FieldParms(2))) & ")"
                                                SearchCaption = SearchCaption & ", " & FieldCaption & " = " & FieldParms(2)
                                            Case 2
                                                Criteria = Criteria & "AND(ccMembers." & FieldParms(0) & ">" & cp.Db.EncodeSQLDate(cp.Utils.EncodeDate(FieldParms(2))) & ")"
                                                SearchCaption = SearchCaption & ", " & FieldCaption & " > " & FieldParms(2)
                                            Case 3
                                                Criteria = Criteria & "AND(ccMembers." & FieldParms(0) & "<" & cp.Db.EncodeSQLDate(cp.Utils.EncodeDate(FieldParms(2))) & ")"
                                                SearchCaption = SearchCaption & ", " & FieldCaption & " < " & FieldParms(2)
                                        End Select
                                    Case FieldTypeCurrency, FieldTypeFloat, FieldTypeInteger, FieldTypeLookup
                                        '
                                        ' number
                                        '
                                        hint = hint & ",190"
                                        Select Case cp.Utils.EncodeInteger(FieldParms(3))
                                            Case 1
                                                Criteria = Criteria & "AND(ccMembers." & FieldParms(0) & "=" & cp.Db.EncodeSQLNumber(cp.Utils.EncodeNumber(FieldParms(2))) & ")"
                                                SearchCaption = SearchCaption & ", " & FieldCaption & " = " & FieldParms(2)
                                            Case 2
                                                Criteria = Criteria & "AND(ccMembers." & FieldParms(0) & ">" & cp.Db.EncodeSQLNumber(cp.Utils.EncodeNumber(FieldParms(2))) & ")"
                                                SearchCaption = SearchCaption & ", " & FieldCaption & " > " & FieldParms(2)
                                            Case 3
                                                Criteria = Criteria & "AND(ccMembers." & FieldParms(0) & "<" & cp.Db.EncodeSQLNumber(cp.Utils.EncodeNumber(FieldParms(2))) & ")"
                                                SearchCaption = SearchCaption & ", " & FieldCaption & " < " & FieldParms(2)
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
                                            Criteria = Criteria & "AND(ccMembers." & FieldParms(0) & "<>0)AND(ccMembers." & FieldParms(0) & " is not null)"
                                        Else
                                            '
                                            ' search for false
                                            '
                                            Criteria = Criteria & "AND((ccMembers." & FieldParms(0) & "=0)or(ccMembers." & FieldParms(0) & " is null))"
                                        End If
                                        SearchCaption = SearchCaption & ", " & FieldCaption & " is " & cp.Utils.EncodeBoolean(FieldParms(2))
                                    Case Else
                                End Select
                            End If
                        End If
                    Next
                End If
                If SearchCaption <> "" Then
                    hint = hint & ",210"
                    SearchCaption = Mid(SearchCaption, 3)
                    SearchCaption = " where " & SearchCaption
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
                        If SearchCaption = "" Then
                            SearchCaption = " in group " & "'" & GetGroupCaption(cp, GroupIDs(GroupPtr)) & "'"
                        Else
                            SearchCaption = SearchCaption & ", in group " & "'" & GetGroupCaption(cp, GroupIDs(GroupPtr)) & "'"
                        End If
                        Criteria = Criteria & "AND((ccMemberRules.GroupID=" & GroupIDs(0) & ")and((ccMemberRules.DateExpires is null)or(ccMemberRules.DateExpires>" & SQLNow & ")))"
                    Else
                        hint = hint & ",250"
                        If SearchCaption <> "" Then
                            SearchCaption = SearchCaption & ", in group(s) "
                        Else
                            SearchCaption = " in group(s) "
                        End If
                        For GroupPtr = 0 To UBound(GroupIDs)
                            hint = hint & ",260"
                            GroupCriteria = GroupCriteria & "OR((ccMemberRules.GroupID=" & GroupIDs(GroupPtr) & ")and((ccMemberRules.DateExpires is null)or(ccMemberRules.DateExpires>" & SQLNow & ")))"
                            'Criteria = Criteria & "AND(ccMemberRules.GroupID=GroupIDs(GroupPtr))"
                            If GroupPtr = UBound(GroupIDs) And GroupPtr <> 0 Then
                                hint = hint & ",270"
                                SearchCaption = SearchCaption & " or " & "'" & GetGroupCaption(cp, GroupIDs(GroupPtr)) & "'"
                            Else
                                hint = hint & ",280"
                                SearchCaption = SearchCaption & GroupDelimiter & "'" & GetGroupCaption(cp, GroupIDs(GroupPtr)) & "'"
                            End If
                            GroupDelimiter = ", "
                        Next
                        Criteria = Criteria & "and(" & Mid(GroupCriteria, 3) & ")"
                    End If
                End If
                '
                ' Add Content Criteria
                '
                hint = hint & ",300"
                If Criteria <> "" Then
                    Criteria = " WHERE " & Mid(Criteria, 4)
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
        Private Function GetFormDetail(cp As CPBaseClass, DetailMemberID As Integer, StatusMessage As String) As String
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
                    Call Nav.AddEntry("Contact", GetFormDetail_TabContact(cp, CS), "ccAdminTab")
                    Call Nav.AddEntry("Permissions", GetFormDetail_TabPermissions(cp, CS), "ccAdminTab")
                    Call Nav.AddEntry("Notes", GetFormDetail_TabNotes(cp, CS), "ccAdminTab")
                    Call Nav.AddEntry("Photos", GetFormDetail_TabPhoto(cp, CS), "ccAdminTab")
                    Call Nav.AddEntry("Groups", GetFormDetail_TabGroup(cp, CS), "ccAdminTab")
                    Call CS.Close()
                    '
                    Content = Nav.getTabs(cp)
                    Content = Content & cp.Html.Hidden(RequestNameFormID, FormDetails.ToString())
                    Content = Content & cp.Html.Hidden(RequestNameMemberID, DetailMemberID.ToString())
                    '
                    GetFormDetail = AdminUIController.GetBody(cp, "Contact Manager &gt;&gt; Contact Details", ButtonList, "", True, True, Header, "", 0, Content)
                End If
            Catch ex As Exception
                cp.Site.ErrorReport(ex)
            End Try
            Return result
        End Function
        '
        '=================================================================================
        '
        Private Function GetFormDetail_TabContact(cp As CPBaseClass, CS As CPCSBaseClass) As String
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
        Private Function GetFormDetail_TabPermissions(cp As CPBaseClass, CS As CPCSBaseClass) As String
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
        Private Function GetFormDetail_TabNotes(cp As CPBaseClass, CS As CPCSBaseClass) As String
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
        Private Function GetFormDetail_TabPhoto(cp As CPBaseClass, CS As CPCSBaseClass) As String
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
        Private Function GetFormDetail_InputTextRow(cp As CPBaseClass, Caption As String, FieldName As String, DefaultValue As String, IsPassword As Boolean) As String
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
        Private Function GetFormDetail_InputBooleanRow(cp As CPBaseClass, Caption As String, FieldName As String, DefaultValue As String) As String
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
        Private Function GetFormDetail_InputHTMLRow(cp As CPBaseClass, Caption As String, FieldName As String, DefaultValue As String) As String
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
        Private Function GetFormDetail_InputImageRow(cp As CPBaseClass, Caption As String, FieldName As String, DefaultValue As String) As String
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
        Private Function GetFormDetail_DividerRow(cp As CPBaseClass, Caption As String) As String
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
        Private Function SaveContactFromStream(cp As CPBaseClass, DetailMemberID As Integer) As String
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
        Private Function GetFormDetail_TabGroup(cp As CPBaseClass, CSMember As CPCSBaseClass) As String
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
        Private Sub SaveMemberRules(cp As CPBaseClass, CSMember As CPCSBaseClass)
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
        Private Function GetFormDetail_TabHomes(cp As CPBaseClass, CSMember As CPCSBaseClass) As String
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
                    Stream = AdminUIController.GetReport(cp, RowCnt, ColCaption, ColAlign, ColWidth, Cells, PageSize, PageNumber, PreTableCopy, PostTableCopy, DataRowCount, "")
                End If
                result = "<div STYLE=""width:100%;"" class=""cmBody ccPanel3DReverse"">" & Stream & "</div>"
            Catch ex As Exception
                cp.Site.ErrorReport(ex)
            End Try
            Return result
        End Function
        '
        '========================================================================
        '
        Function GetFormInputRadioBox(cp As CPBaseClass, ElementName As String, ElementValue As String, CurrentValue As String, ElementID As String) As String
            Return cp.Html.RadioBox(ElementName, ElementValue, CurrentValue, ElementID)
        End Function
        '
        '========================================================================
        '
        Function GetFormInputText(cp As CPBaseClass, ElementName As String, CurrentValue As String, Height As String, Width As String, ElementID As String, OnFocus As String) As String
            Return cp.Html.InputText(ElementName, CurrentValue, Height, Width, False, "", ElementID)
        End Function
        '
        '========================================================================
        '
        Function GetFormInputDate(cp As CPBaseClass, ElementName As String, CurrentValue As String, Width As String, ElementID As String, OnFocus As String) As String
            Return cp.Html.InputDate(ElementName, CurrentValue, Width, "", ElementID)
        End Function
    End Class
End Namespace
