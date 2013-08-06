VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CMngrClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit
'
Const RequestNamePageSize = "PageSize"
Const RequestNameMemberID = "MemberID"
Const RequestNameDetailSubtab = "CMDetailTab"
'
Const RequestNameFormID = "fid"
Const FormSearch = 1
Const FormList = 2
Const FormDetails = 3
'
Private Csv As Object
Private Main As Object
'Private Main As MainClass
'
'========================================================================
'   v3.3 Compatibility
'       To make an Add-on that works the same in v3.3 and v3.4, use this adapter instead of the execute above
'========================================================================
'
Public Function Execute(CsvObject As Object, MainObject As Object, OptionString As String, FilterInput As String) As String
    Set Csv = CsvObject
    Call Init(MainObject)
    Execute = GetContent(OptionString)
End Function
'
'========================================================================
'   Init()
'========================================================================
'
Public Sub Init(MainObject As Object)
'Public Sub Init(MainObject As ccWeb3.MainClass)
    Set Main = MainObject
    Exit Sub
    '
    ' ----- Error Trap
    '
ErrorTrap:
    Call HandleError("CMngrClass", "Init", Err.Number, Err.Source, Err.Description, True, False)
    End Sub
'
'=================================================================================
'   Aggregate Object Interface
'=================================================================================
'
Public Function GetContent(OptionString As String) As String
    On Error GoTo ErrorTrap
    '
    Dim PeopleCID As Long
    Dim AdminUI As Object
    'Dim AdminUI As New ccWeb3.AdminUIClass
    Dim RecordCnt As Long
    Dim Argument1 As String
    Dim Argument2 As String
    Dim TabNumber As Long
    Dim SourceFilename As String
    Dim CutPosition As Long
    Dim InBuffer As String
    Dim RQS As String
    Dim ContactSearchCriteria As String
    Dim Button As String
    Dim RowCount As Long
    Dim GroupID As Long
    Dim GroupToolAction As Long
    Dim GroupName As String
    Dim RowPointer As Long
    Dim memberID As Long
    Dim GroupToolSelect As Long
    Dim ContactContentCriteria As String
    Dim ContactGroupCriteria As String
    Dim Criteria As String
    Dim SearchCaption As String
    Dim CCID As Long
    Dim SQL As String
    Dim StatusMessage As String
    Dim FieldName As String
    Dim SelectList As String
    Dim FieldNameList As String
    Dim SQLFrom As String
    Dim SQLCriteria As String
    Dim LookupContentID As Long
    Dim LookupTable As String
    Dim ContentID As Long
    Dim CS As Long
    Dim JoinTableCnt As Long
    Dim JoinTableAlias As String
    Dim Filename As String
    Dim Filename2 As String
    Dim ExportName As String
    Dim DetailMemberID As Long
    Dim FormID As Long
    Dim IsGoMethodEnabled As Boolean
    Dim IsAdminPath As Boolean
    '
    If Not (Main Is Nothing) Then
        Set AdminUI = CreateObject("ccWeb3.AdminUIClass")
        IsAdminPath = (UCase(Main.getAddonOption("source", OptionString)) = "ADMIN")
        If Not Main.IsAdmin Then
            '
            ' You must be admin to use this feature
            '
            If IsAdminPath Then
                '
                ' Admin site, but not an admin
                '
                GetContent = AdminUI.GetFormBodyAdminOnly(Main)
                GetContent = AdminUI.GetBody(Main, "Contact Manager", ButtonCancelAll, "", False, False, GetContent, "", 0, "")
            Else
                '
                ' Public site, not an admin
                '
                GetContent = AdminUI.GetFormBodyAdminOnly(Main)
                GetContent = AdminUI.GetBody(Main, "Contact Manager", "", "", False, False, GetContent, "", 0, "")
            End If
        Else
            '
            PeopleCID = Main.GetContentID("people")
            RQS = Main.RefreshQueryString
            TabNumber = Main.GetStreamInteger("tab")
            DetailMemberID = Main.GetStreamInteger(RequestNameMemberID)
            ContactSearchCriteria = Main.GetMemberProperty("ContactSearchCriteria", ContactSearchCriteria)
            ContactGroupCriteria = Main.GetMemberProperty("ContactGroupCriteria", "")
            ContactContentCriteria = Main.GetMemberProperty("ContactContentID", PeopleCID)
            IsGoMethodEnabled = kmaEncodeBoolean(Main.GetSiteProperty("EnableExtensionGoMethod", "0"))
            '
            Call Main.AddRefreshQueryString("tab", TabNumber)
            '
            FormID = Main.GetStreamInteger(RequestNameFormID)
            Button = Main.GetStreamText("Button")
            If FormID = 0 Then
                '
                ' ----- Set default form
                '
                If (DetailMemberID <> 0) Then
                    '
                    ' Detail form
                    '
                    FormID = FormDetails
                ElseIf (ContactSearchCriteria <> "") Or (ContactGroupCriteria <> "") Or (ContactContentCriteria <> CStr(PeopleCID)) Then
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
                                FormID = ProcessSearchForm
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
                                ContactContentCriteria = ""
                                Call Main.SetMemberProperty("ContactSearchCriteria", ContactSearchCriteria)
                                Call Main.SetMemberProperty("ContactGroupCriteria", ContactGroupCriteria)
                                Call Main.SetMemberProperty("ContactContentID", ContactContentCriteria)
                                FormID = FormSearch
                            Case ButtonApply
                                '
                                ' Add to or remove from group
                                '
                                RowCount = Main.GetStreamInteger("M.Count")
                                GroupID = Main.GetStreamInteger("GroupID")
                                GroupToolAction = Main.GetStreamInteger("GroupToolAction")
                                GroupToolSelect = Main.GetStreamInteger("GroupToolSelect")
                                SQLCriteria = ""
                                Call BuildSearch(SQLCriteria, SearchCaption)
                                Select Case GroupToolAction
                                    Case 1
                                        '
                                        ' ----- Add to Group
                                        '
                                        If (GroupID = 0) Then
                                            '
                                            ' Group required and not provided
                                            '
                                            Call Main.AddUserError("Please select a Target Group for this operation")
                                        ElseIf GroupToolSelect = 0 Then
                                            '
                                            ' Add selection to Group
                                            '
                                            If (RowCount > 0) Then
                                                GroupName = Main.GetGroupByID(GroupID)
                                                For RowPointer = 0 To RowCount - 1
                                                    If Main.GetStreamBoolean("M." & RowPointer) Then
                                                        memberID = Main.GetStreamInteger("MID." & RowPointer)
                                                        Call Main.AddGroupMember(GroupName, memberID)
                                                    End If
                                                Next
                                            End If
                                        Else
                                            '
                                            ' Add everyone in search criteria to this group
                                            '
                                            CCID = Main.GetContentID("Member Rules")
                                            SQL = "insert into ccMemberRules (Active,ContentControlID,GroupID,MemberID )" _
                                                & " select 1," & CCID & "," & GroupID & ",ccMembers.ID" _
                                                & " from (ccMembers" _
                                                & " left join ccMemberRules on ccMemberRules.MemberID=ccMembers.ID )" _
                                                & " left join ( select MemberID  from ccMemberRules where GroupID in (" & GroupID & ")) as InGroups on InGroups.MemberID=ccMembers.ID" _
                                                & " " & SQLCriteria _
                                                & " and InGroups.MemberID is null" _
                                                & ""
                                            Call Main.ExecuteSQL("default", SQL)
                                        End If
                                    Case 2
                                        '
                                        ' ----- Remove From Group
                                        '
                                        If (GroupID = 0) Then
                                            '
                                            ' Group required and not provided
                                            '
                                            Call Main.AddUserError("Please select a Target Group for this operation")
                                        ElseIf GroupToolSelect = 0 Then
                                            '
                                            ' Remove selection from Group
                                            '
                                            If (RowCount > 0) Then
                                                GroupName = Main.GetGroupByID(GroupID)
                                                For RowPointer = 0 To RowCount - 1
                                                    If Main.GetStreamBoolean("M." & RowPointer) Then
                                                        memberID = Main.GetStreamInteger("MID." & RowPointer)
                                                        Call Main.DeleteContentRecords("Member Rules", "(GroupID=" & GroupID & ")and(MemberID=" & memberID & ")")
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
                                            Call Main.ExecuteSQL("default", SQL)
                                        End If
                                    Case 3
                                        '
                                        ' ----- Export
                                        '
                                        If Main.SiteProperty_BuildVersion < "3.3.452" Then
                                            Call Main.AddUserError("This operation requires a database upgrade for this site. Use the Contensive Application Manager to upgrade your site.")
                                            StatusMessage = "<P>There was a problem requesting your download. Your Contensive database needs to be updated.<P>"
                                        Else
                                            ExportName = SearchCaption
                                            If GroupToolSelect = 0 Then
Dim RowSQL As String
Dim Aborttask As Boolean
                                                '
                                                ' Export selection from Group
                                                '
                                                ExportName = "Selected rows from " & ExportName
                                                RowSQL = ""
                                                If (RowCount > 0) Then
                                                    GroupName = Main.GetGroupByID(GroupID)
                                                    For RowPointer = 0 To RowCount - 1
                                                        If Main.GetStreamBoolean("M." & RowPointer) Then
                                                            memberID = Main.GetStreamInteger("MID." & RowPointer)
                                                            RowSQL = RowSQL & "OR(ccMembers.ID=" & memberID & ")"
                                                        End If
                                                    Next
                                                    If RowSQL = "" Then
                                                        '
                                                        ' nothing selected, abort export
                                                        '
                                                        Aborttask = True
                                                        Call Main.TestPoint("No contacts to export")
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
                                            SQLFrom = "ccMembers"
                                            JoinTableCnt = 0
                                            If Main.GetMemberProperty("ContactGroupCriteria", "") <> "" Then
                                                SQLFrom = "(" & SQLFrom & " left join ccMemberRules on ccMemberRules.MemberID=ccMembers.ID )"
                                            End If
                                            ContentID = kmaEncodeInteger(Main.GetMemberProperty("ContactContentID", Main.GetContentID("people")))
                                            CS = Main.OpenCSContent("Content Fields", "ContentID=" & ContentID, "EditSortPriority", True, False, False, "Name,Caption,Type,LookupContentID")
                                            Do While Main.IsCSOK(CS)
                                                FieldName = Main.GetCStext(CS, "name")
                                                Select Case Main.GetCSInteger(CS, "type")
                                                    Case FieldTypeLookup
                                                        '
                                                        ' just add the ID into the list
                                                        '
                                                        SelectList = SelectList & ",ccMembers." & FieldName
                                                        FieldNameList = FieldNameList & "," & FieldName
                                                    ' Skip it
                                                    '    '
                                                    '    ' add table join for secondary name
                                                    '    '
                                                    '    JoinTableAlias = "J" & JoinTableCnt
                                                    '    LookupContentID = Main.GetCSInteger(CS, "LookupContentID")
                                                    '    If LookupContentID <> 0 Then
                                                    '        LookupTable = Main.GetContentTablename(Main.GetContentNameByID(LookupContentID))
                                                    '        SelectList = SelectList & ",ccMembers." & FieldName & "," & JoinTableAlias & ".Name as NameOf" & FieldName
                                                    '        SQLFrom = "(" & SQLFrom & " left join " & LookupTable & " " & JoinTableAlias & " on " & JoinTableAlias & ".ID=ccMembers." & FieldName & ")"
                                                    '        FieldNameList = FieldNameList & "," & FieldName & ",NameOf" & FieldName
                                                    '        JoinTableCnt = JoinTableCnt + 1
                                                    '    End If
                                                    Case FieldTypeTextFile
                                                        '
                                                        ' read file for text - skip field
                                                        '
                                                    Case FieldTypeRedirect, FieldTypeManyToMany
                                                        '
                                                        ' no field involved, skip it
                                                        '
                                                    Case Else
                                                        '
                                                        ' just add value
                                                        '
                                                        SelectList = SelectList & ",ccMembers." & FieldName
                                                        FieldNameList = FieldNameList & "," & FieldName
                                                End Select
                                                Call Main.NextCSRecord(CS)
                                            Loop
                                            Call Main.CloseCS(CS)
                                            If SelectList = "" Then
                                                Call Main.TestPoint("Export All Contacts, but no fields were found to export")
                                                StatusMessage = "<P>There was a problem requesting your download.<P>"
                                            Else
                                                SelectList = Mid(SelectList, 2)
                                                If FieldNameList <> "" Then
                                                    FieldNameList = Mid(FieldNameList, 2)
                                                End If
                                                'ExportName = CStr(Now()) & " snapshot of " & LCase(ExportName)
                                                SQL = "select Distinct " & SelectList & " from " & SQLFrom & SQLCriteria
                                                Call Main.RequestTask("BuildCSV", SQL, ExportName, "MemberExport-" & CStr(Main.GetRandomLong) & ".csv")
                                                Call Main.TestPoint("Export All Contacts, SQL=" & SQL)
                                                StatusMessage = "<P>Your download request has been submitted and will be available on the <a href=" & Main.SiteProperty_AdminURL & "?af=30>Download Requests</a> page shortly.<P>"
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
                                            RecordCnt = 0
                                            If (RowCount > 0) Then
                                                GroupName = Main.GetGroupByID(GroupID)
                                                For RowPointer = 0 To RowCount - 1
                                                    If Main.GetStreamBoolean("M." & RowPointer) Then
                                                        memberID = Main.GetStreamInteger("MID." & RowPointer)
                                                        Call Main.ExecuteSQL("default", "update ccMembers set AllowBulkEmail=1 where ID=" & memberID)
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
                                            Call Main.ExecuteSQL("default", SQL)
                                            StatusMessage = "<P>Allow Group Email was set for all people in this selection.<P>"
                                        End If
                                End Select
                        End Select
                    Case FormDetails
                        DetailMemberID = Main.GetStreamInteger(RequestNameMemberID)
                        Select Case Button
                            Case ButtonCancel
                                FormID = FormList
                            Case ButtonSave
                                Call SaveContactFromStream(DetailMemberID)
                                FormID = FormDetails
                            Case ButtonOK
                                Call SaveContactFromStream(DetailMemberID)
                                FormID = FormList
                            Case ButtonNewSearch
                                FormID = FormSearch
                        End Select
                End Select
            End If
            '
            ' ----- Output the next form
            '
            '
            'GetContent = "" _
            '    & "<link href=""/ccLib/Styles/ccTabs.css"" type=""text/css"" rel=""stylesheet"">" _
            '    & "<script type=text/javascript src=""/ccLib/ClientSide/ccDynamicTab.js""></script>"
            Select Case FormID
                Case FormDetails
                    '
                    ' Detail form
                    '
                    GetContent = GetContent & GetFormDetail(DetailMemberID, StatusMessage, IsGoMethodEnabled)
                Case FormList
                    '
                    ' List form
                    '
                    GetContent = GetContent & GetFormList(StatusMessage, IsAdminPath)
                Case Else
                    '
                    ' Search Form
                    '
                    GetContent = GetContent & GetFormSearch(StatusMessage, IsAdminPath)
            End Select
            GetContent = "<div class=ccbodyadmin>" & GetContent & "</div>"
            'GetContent = GetContent & "<script type=text/javascript>TabInit();</script>"
        End If
    End If
    '
    ' wrapper for style strength
    '
    GetContent = "" _
        & vbCrLf & vbTab & "<div class=""contactManager"">" _
        & kmaIndent(GetContent) _
        & vbCrLf & vbTab & "</div>" _
        & ""
    '
    Exit Function
ErrorTrap:
    Call HandleError("CMngrClass", "GetContent", Err.Number, Err.Source, Err.Description, True, False)
    End Function
'
'=================================================================================
'
'=================================================================================
'
Private Function GetFormSearch(StatusMessage As String, IsAdminPath As Boolean) As String
    On Error GoTo ErrorTrap
    '
    Dim SubTab As Long
    Dim RQS As String
    Dim Nav As New MenuLiveTabClass
    Dim Header As String
    Dim AdminUI As Object
    Dim Content As String
    Dim ButtonList As String
    '
    If Not (Main Is Nothing) Then
        '
        ' Determine current Subtab
        '
        Set AdminUI = CreateObject("ccWeb3.AdminUIClass")
        SubTab = Main.GetStreamInteger("SubTab")
        If SubTab = 0 Then
            SubTab = kmaEncodeInteger(Main.GetMemberProperty("SelectSubTab", "1"))
            If SubTab = 0 Then
                SubTab = 1
                Call Main.SetMemberProperty("SelectSubTab", CStr(SubTab))
            End If
        Else
            Call Main.SetMemberProperty("SelectSubTab", CStr(SubTab))
        End If
        Call Main.AddRefreshQueryString("SubTab", SubTab)
        '
        ' SubTab Menu
        '
        Call Main.AddRefreshQueryString("tab", "")
        If IsAdminPath Then
            ButtonList = ButtonCancelAll & "," & ButtonSearch
        Else
            ButtonList = ButtonSearch
        End If
        '
        Header = "<div>Use the selections in each tab below to create a criteria for your search and hit the Search button.</div>"
        '
        Call Nav.AddEntry("Record&nbsp;Fields", GetFormSearch_TabPeople(), "ccAdminTab")
        Call Nav.AddEntry("Groups", GetFormSearch_TabGroup(), "ccAdminTab")
        Call Nav.AddEntry("Content", GetFormSearch_TabContent(), "ccAdminTab")
        '
        Content = "" _
            & Main.GetFormInputHidden("SelectionForm", "1") _
            & Nav.GetTabs() _
            & Main.GetFormInputHidden(RequestNameFormID, FormSearch) _
            & ""
        GetFormSearch = AdminUI.GetBody(Main, "Contact Manager &gt;&gt; Selection Criteria", ButtonList, "", True, True, Header, "", 0, Content)
    End If
    '
    Exit Function
ErrorTrap:
    Call HandleError("CMngrClass", "GetFormSearch", Err.Number, Err.Source, Err.Description, True, False)
    End Function
'
'=================================================================================
'
'=================================================================================
'
Private Function GetFormSearch_TabContent() As String
    On Error GoTo ErrorTrap
    '
    Dim Argument1 As String
    Dim CS As Long
    Dim ContactGroupCriteria As String
    Dim GroupCount As Long
    Dim GroupPointer As Long
    Dim GroupChecked As Boolean
    Dim ContentName As String
    Dim ContentID As Long
    Dim RowEven As Boolean
    Dim Button As String
    Dim SQL As String
    Dim RQS As String
    Dim SubTab As Long
    Dim FormSave As Boolean
    Dim FormClear As Boolean
    Dim ContactContentID As Long
    Dim ContentControlCriteria As String
    Dim s As String
    Dim Style As String
    '
    If Not (Main Is Nothing) Then
        Button = "ContentSelect"
        RQS = Main.RefreshQueryString
        FormSave = (Main.GetStreamText(Button) = "Save")
        '
        ' Prepopulate the form
        '
        ContactContentID = kmaEncodeInteger(Main.GetMemberProperty("ContactContentID", 0))
        If ContactContentID = 0 Then
            ContactContentID = Main.GetContentID("people")
        End If
        '
        s = s _
            & "<div>All Contensive data are stored in folders called Content Definitions. Select the Content Definitions which contain your the records you wish to select.</div>" _
            & "<div>&nbsp;</div>"
        '
        ' Add headers to stream
        '
        s = s & "<table border=0 width=100% cellspacing=0 cellpadding=4>"
        s = s & kmaStartTableRow
        s = s & "<TD width=60><img src=/cclib/images/spacer.gif width=60 height=1></TD>"
        s = s & "<TD width=99%><img src=/cclib/images/spacer.gif width=1 height=1></TD>"
        s = s & kmaEndTableRow
        '
        s = s & kmaStartTableRow
        s = s & "<TD width=60 align=center class=ccAdminListCaption>Select</TD>"
        s = s & "<TD width=99% align=left class=ccAdminListCaption>Content</TD>"
        s = s & kmaEndTableRow
        
        s = s & kmaStartTableRow
        s = s & kmaStartTableCell("60", 1, RowEven, "center") & "<img src=/cclib/images/spacer.gif width=60 height=1></TD>"
        s = s & kmaStartTableCell("99%", 1, RowEven, "left") & "<img src=/cclib/images/spacer.gif width=1 height=1></TD>"
        s = s & kmaEndTableRow
        '
        ContentControlCriteria = Main.GetContentControlCriteria("people")
        ContentControlCriteria = Replace(ContentControlCriteria, "ccMembers", "ccContent")
        ContentControlCriteria = Replace(ContentControlCriteria, "ContentControlID", "ID")
        CS = Main.OpenCSContent("Content", ContentControlCriteria, "name", , , , "name,ID")
        GroupPointer = 0
        Do While Main.IsCSOK(CS)
            ContentName = Main.GetCStext(CS, "name")
            ContentID = Main.GetCSInteger(CS, "ID")
            If ((GroupPointer Mod 2) = 0) Then
                Style = "ccAdminListRowEven"
            Else
                Style = "ccAdminListRowOdd"
            End If
            s = s & "<TR>"
            s = s & "<TD width=30 align=center class=""" & Style & """>" & Main.GetFormInputRadioBox("ContentID", ContentID, ContactContentID) & "</TD>"
            s = s & "<TD width=99% align=left class=""" & Style & """>" & ContentName & "</TD>"
            s = s & "</TR>"
            GroupPointer = GroupPointer + 1
            Call Main.NextCSRecord(CS)
        Loop
        Call Main.CloseCS(CS)
        s = s & "</Table>"
        '
        GetFormSearch_TabContent = s & Main.GetFormInputHidden("SelectionContentSubTab", "1")
        '
    End If
    '
    Exit Function
ErrorTrap:
    Call HandleError("CMngrClass", "GetFormSearch_TabContent", Err.Number, Err.Source, Err.Description, True, False)
    End Function
'
'=================================================================================
'
'=================================================================================
'
Private Function GetFormSearch_TabGroup() As String
    On Error GoTo ErrorTrap
    '
    Dim Argument1 As String
    Dim CS As Long
    Dim ContactGroupCriteria As String
    Dim GroupCount As Long
    Dim GroupPointer As Long
    Dim GroupChecked As Boolean
    Dim GroupLabel As String
    Dim GroupID As Long
    Dim GroupName As String
    Dim RowEven As Boolean
    Dim Button As String
    Dim SQL As String
    Dim RQS As String
    Dim SubTab As Long
    Dim FormSave As Boolean
    Dim FormClear As Boolean
    Dim s As String
    Dim Style As String
    '
    If Not (Main Is Nothing) Then
        Button = "GroupSelect"
        RQS = Main.RefreshQueryString
        ContactGroupCriteria = Main.GetMemberProperty("ContactGroupCriteria", "")
        '
        s = s _
            & "<div>Select groups to narrow your results. If any groups are selected, your search will be limited to people in any of the selected groups.</div>" _
            & "<div>&nbsp;</div>"
        '
        ' Add headers to stream
        '
        s = s & "<table border=0 width=100% cellspacing=0 cellpadding=4>"
        s = s & kmaStartTableRow
        s = s & "<TD width=30><img src=/cclib/images/spacer.gif width=20 height=1></TD>"
        s = s & "<TD width=30><img src=/cclib/images/spacer.gif width=20 height=1></TD>"
        s = s & "<TD width=99%><img src=/cclib/images/spacer.gif width=1 height=1></TD>"
        s = s & kmaEndTableRow
        '
        s = s & kmaStartTableRow
        s = s & "<TD width=30 align=center class=ccAdminListCaption>Select</TD>"
        s = s & "<TD width=30 align=center class=ccAdminListCaption>Count</TD>"
        s = s & "<TD width=99% align=left class=ccAdminListCaption>Group Name</TD>"
        s = s & kmaEndTableRow
        '
        'S = S & Main.GetFormStart
        '
        SQL = "SELECT ccGroups.ID as GroupID, ccGroups.Name as GroupName, ccGroups.Caption as GroupCaption, Count(ccMembers.ID) AS CountOfID" _
            & " FROM (ccGroups LEFT JOIN ccMemberRules ON ccGroups.ID = ccMemberRules.GroupID) LEFT JOIN ccMembers ON ccMemberRules.MemberID = ccMembers.ID" _
            & " Where (((ccMemberRules.DateExpires) Is Null Or (ccMemberRules.DateExpires) > " & KmaEncodeSQLDate(Now()) & "))" _
            & " GROUP BY ccGroups.ID, ccGroups.Name, ccGroups.Caption" _
            & " ORDER BY ccGroups.Caption;"
        'SQL = "SELECT ccGroups.ID as GroupID, ccGroups.Name as GroupName, ccGroups.Caption as GroupCaption, Count(ccMembers.ID) AS CountOfID" _
        '    & " FROM (ccGroups LEFT JOIN ccMemberRules ON ccGroups.ID = ccMemberRules.GroupID) LEFT JOIN ccMembers ON ccMemberRules.MemberID = ccMembers.ID" _
        '    & " Where (((ccMemberRules.DateExpires) Is Null Or (ccMemberRules.DateExpires) > " & KmaEncodeSQLDate(Now()) & "))" _
        '    & " GROUP BY ccGroups.ID, ccGroups.Name, ccGroups.Caption" _
        '    & " HAVING (((Count(ccMembers.ID)) Is Not Null))" _
        '    & " ORDER BY ccGroups.Caption;"
        CS = Main.OpenCSSQL("Default", SQL)
        GroupPointer = 0
        Do While Main.IsCSOK(CS)
            GroupID = Main.GetCSInteger(CS, "GroupID")
            GroupName = Main.GetCStext(CS, "GroupCaption")
            If GroupName = "" Then
                GroupName = Main.GetCStext(CS, "GroupName")
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
            s = s & "<TD width=30 align=center class=""" & Style & """>" & Main.GetFormInputCheckBox(GroupLabel, GroupChecked) & Main.GetFormInputHidden(GroupLabel & ".id", GroupID) & "</TD>"
            s = s & "<TD width=30 align=right class=""" & Style & """>" & Main.GetCSInteger(CS, "CountOfID") & "</TD>"
            s = s & "<TD width=99% align=left class=""" & Style & """>" & GroupName & "</TD>"
            s = s & "</TR>"
            'S = S & kmaStartTableRow
            'S = S & kmaStartTableCell("30", 1, RowEven, "center") & Main.GetFormInputCheckBox(GroupLabel, GroupChecked) & Main.GetFormInputHidden(GroupLabel & ".id", GroupID) & "</TD>"
            'S = S & kmaStartTableCell("30", 1, RowEven, "right") & Main.GetCSInteger(CS, "CountOfID") & "&nbsp;&nbsp;&nbsp;</TD>"
            'S = S & kmaStartTableCell("100%", 1, RowEven, "left") & "&nbsp;" & GroupName & "</TD>"
            'S = S & kmaEndTableRow
            GroupPointer = GroupPointer + 1
            Call Main.NextCSRecord(CS)
            Loop
        Call Main.CloseCS(CS)
        s = s & "</Table>"
        '
        s = s & Main.GetFormInputHidden("GroupCount", GroupPointer)
        GetFormSearch_TabGroup = s & Main.GetFormInputHidden("SelectionGroupSubTab", "1")
        '
        End If
    '
    Exit Function
ErrorTrap:
    Call HandleError("CMngrClass", "GetFormSearch_TabGroup", Err.Number, Err.Source, Err.Description, True, False)
    End Function
'
'=================================================================================
'
'=================================================================================
'
Private Function GetFormSearch_TabPeople() As String
    On Error GoTo ErrorTrap
    '
    Dim Argument1 As String
    Dim CS As Long
    Dim ContactSearchCriteria As String
    Dim FieldPtr As Long
    Dim GroupChecked As Boolean
    Dim GroupLabel As String
    Dim GroupID As Long
    Dim RowEven As Boolean
    Dim Button As String
    Dim SQL As String
    Dim RQS As String
    Dim SubTab As Long
    Dim Criteria As String
    Dim VisitSearchCriteria As String
    Dim CriteriaNames() As String
    Dim CriteriaValues() As String
    Dim CriteriaCount As Long
    Dim CriteriaPointer As Long
    Dim FieldName() As String
    Dim FieldCaption() As String
    Dim fieldId() As Long
    Dim fieldType() As Long
    Dim FieldValue() As String
    Dim FieldOperator() As Long
    Dim FieldLookupContentName() As String
    Dim FieldLookupList() As String
    Dim fieldEditTab() As String
    Dim ContentID As Long
    Dim currentEditTab As String
    Dim lastEditTab As String
    Dim FieldCount As Long
    Dim FieldSize As Long
    Dim FormSave As Boolean
    Dim FormClear As Boolean
    Dim NameValues As Variant
    Dim RowPointer As Long
    Dim s As String
    Dim tabCaption As String
    Dim groupTab As String
    '
    If Not (Main Is Nothing) Then
        Button = "CriteriaSelect"
        RQS = Main.RefreshQueryString
        '
        ' prepare visit property to prepopulate form
        '
        ContactSearchCriteria = Main.GetMemberProperty("ContactSearchCriteria", "")
        If ContactSearchCriteria <> "" Then
            CriteriaValues = Split(ContactSearchCriteria, vbCrLf)
            CriteriaCount = UBound(CriteriaValues) + 1
        End If
        '
        ' Setup fields and capture request changes
        '
        Criteria = "(active<>0)and(ContentID=" & Main.GetContentID("people") & ")and(authorable<>0)"
        CS = Main.OpenCSContent("Content Fields", Criteria, "EditTab,EditSortPriority")
        FieldPtr = 0
        Do While Main.IsCSOK(CS)
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
            FieldName(FieldPtr) = Main.GetCStext(CS, "name")
            FieldCaption(FieldPtr) = Main.GetCStext(CS, "Caption")
            fieldId(FieldPtr) = Main.GetCSInteger(CS, "ID")
            fieldType(FieldPtr) = Main.GetCSInteger(CS, "Type")
            If fieldType(FieldPtr) = 7 Then
                ContentID = Main.GetCSInteger(CS, "LookupContentID")
                If ContentID > 0 Then
                    FieldLookupContentName(FieldPtr) = Main.GetContentNameByID(ContentID)
                End If
                FieldLookupList(FieldPtr) = Main.GetCStext(CS, "LookupList")
            End If
            fieldEditTab(FieldPtr) = Main.GetCStext(CS, "editTab")
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
            Call Main.NextCSRecord(CS)
        Loop
        Call Main.CloseCS(CS)
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
        s = s & kmaStartTableRow
        s = s & kmaStartTableCell("120", 1, RowEven, "right") & "<img src=/cclib/images/spacer.gif width=120 height=1></TD>"
        s = s & kmaStartTableCell("99%", 1, RowEven, "left") & "<img src=/cclib/images/spacer.gif width=1 height=1></TD>"
        s = s & kmaEndTableRow
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
                            & "<TD width=10 align=right>" & GetFormInputRadioBox(FieldName(FieldPtr) & "_D", 0, FieldOperator(FieldPtr), "") & "</TD><TD align=left width=100>ignore</TD>" _
                            & "<TD width=10>&nbsp;&nbsp;</TD><TD width=10 align=right>" & GetFormInputRadioBox(FieldName(FieldPtr) & "_D", 1, FieldOperator(FieldPtr), "") & "</TD><TD align=left width=100>=</TD>" _
                            & "<TD width=10>&nbsp;&nbsp;</TD><TD width=10 align=right>" & GetFormInputRadioBox(FieldName(FieldPtr) & "_D", 2, FieldOperator(FieldPtr), "") & "</TD><TD align=left width=100>&gt;</TD>" _
                            & "<TD width=10>&nbsp;&nbsp;</TD><TD width=10 align=right>" & GetFormInputRadioBox(FieldName(FieldPtr) & "_D", 3, FieldOperator(FieldPtr), "") & "</TD><TD align=left width=100>&lt;</TD>" _
                            & "<TD width=10>&nbsp;&nbsp;</TD><TD align=left width=99%>" & GetFormInputDate(FieldName(FieldPtr), FieldValue(FieldPtr), 10, "", "") & "</TD>" _
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
                            & "<TD width=10 align=right>" & GetFormInputRadioBox(FieldName(FieldPtr) & "_N", 0, FieldOperator(FieldPtr), "") & "</TD><TD align=left width=100>ignore</TD>" _
                            & "<TD width=10>&nbsp;&nbsp;</TD><TD width=10 align=right>" & GetFormInputRadioBox(FieldName(FieldPtr) & "_N", 1, FieldOperator(FieldPtr), "") & "</TD><TD align=left width=100>=</TD>" _
                            & "<TD width=10>&nbsp;&nbsp;</TD><TD width=10 align=right>" & GetFormInputRadioBox(FieldName(FieldPtr) & "_N", 2, FieldOperator(FieldPtr), "") & "</TD><TD align=left width=100>&gt;</TD>" _
                            & "<TD width=10>&nbsp;&nbsp;</TD><TD width=10 align=right>" & GetFormInputRadioBox(FieldName(FieldPtr) & "_N", 3, FieldOperator(FieldPtr), "") & "</TD><TD align=left width=100>&lt;</TD>" _
                            & "<TD width=10>&nbsp;&nbsp;</TD><TD align=left width=99%>" & GetFormInputText(FieldName(FieldPtr), FieldValue(FieldPtr), 1, 5, "", "") & "</TD>" _
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
                            & "<TD width=10 align=right>" & GetFormInputRadioBox(FieldName(FieldPtr), "", FieldValue(FieldPtr), "") & "</TD><TD align=left width=100>" & "  ignore</TD>" _
                            & "<TD width=10>&nbsp;&nbsp;</TD><TD width=10 align=right>" & GetFormInputRadioBox(FieldName(FieldPtr), "1", FieldValue(FieldPtr), "") & "</TD><TD align=left width=100>" & "true</TD>" _
                            & "<TD width=10>&nbsp;&nbsp;</TD><TD width=10 align=right>" & GetFormInputRadioBox(FieldName(FieldPtr), "0", FieldValue(FieldPtr), "") & "</TD><TD align=left width=100>" & "false</TD>" _
                            & "<TD width=99%>&nbsp;</td>" _
                        & "</TR></Table>" _
                        & "</TD>" _
                        & "</TR>"
                        '& "[" & Main.GetFormInputRadioBox(FieldName(FieldPtr), "", FieldValue(FieldPtr)) & "ignore]" _
                        '& "[" & Main.GetFormInputRadioBox(FieldName(FieldPtr), "1", FieldValue(FieldPtr)) & "true]" _
                        '& "[" & Main.GetFormInputRadioBox(FieldName(FieldPtr), "0", FieldValue(FieldPtr)) & "false]" _
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
                            & "<TD width=10 align=right>" & GetFormInputRadioBox(FieldName(FieldPtr) & "_T", "", FieldValue(FieldPtr), "") & "</TD><TD align=left width=100>" & "  ignore</TD>" _
                            & "<TD width=10>&nbsp;&nbsp;</TD><TD width=10 align=right>" & GetFormInputRadioBox(FieldName(FieldPtr) & "_T", "1", FieldValue(FieldPtr), "") & "</TD><TD align=left width=100>" & "empty</TD>" _
                            & "<TD width=10>&nbsp;&nbsp;</TD><TD width=10 align=right>" & GetFormInputRadioBox(FieldName(FieldPtr) & "_T", "2", FieldValue(FieldPtr), "") & "</TD><TD align=left width=100>" & "not&nbsp;empty</TD>" _
                            & "<TD width=10>&nbsp;&nbsp;</TD><TD width=10 align=right>" & GetFormInputRadioBox(FieldName(FieldPtr) & "_T", "3", FieldValue(FieldPtr), "t" & FieldPtr) & "</TD><TD align=center width=100>" & "&nbsp;includes&nbsp;" & "</TD><TD align=left width=99%>" & GetFormInputText(FieldName(FieldPtr), FieldValue(FieldPtr), 1, 20, "", "var e=getElementById('t" & FieldPtr & "');e.checked=1;") & "</TD>" _
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
                            & Main.GetFormInputSelect(FieldName(FieldPtr), FieldValue(FieldPtr), FieldLookupContentName(FieldPtr), , "Any") & "</TD>"
                    Else
                        s = s _
                            & "<TD class=""cmRowField"">" _
                            & Main.GetFormInputSelectList(FieldName(FieldPtr), FieldValue(FieldPtr), FieldLookupList(FieldPtr), , "Any") & "</TD>"
                    End If
                    s = s & "</TR>"
                    groupTab = ""
                    RowPointer = RowPointer + 1
                End Select
            Next
        s = s & "</Table>"
        '
        GetFormSearch_TabPeople = s & Main.GetFormInputHidden("SelectionSearchSubTab", "1")
        '
        End If
    '
    Exit Function
ErrorTrap:
    Call HandleError("CMngrClass", "GetFormSearch_TabPeople", Err.Number, Err.Source, Err.Description, True, False)
    End Function
'
'=================================================================================
'
'=================================================================================
'
Private Function GetFormList(StatusMessage As String, IsAdminPath As Boolean) As String
    On Error GoTo ErrorTrap
    '
    Dim SQLOrderDir As String
    Dim SQLOrderBy As String
    Dim LastVisit As Date
    Dim CheckBox As String
    Dim Argument1 As String
    Dim CS As Long
    Dim ContactGroupCriteria As String
    Dim GroupCount As Long
    Dim GroupPointer As Long
    Dim GroupChecked As Boolean
    Dim RecordName As String
    Dim ContentName As String
    Dim RecordID As Long
    Dim RowEven As Boolean
    Dim Button As String
    Dim SQL As String
    Dim RQS As String
    Dim SubTab As Long
    Dim FormSave As Boolean
    Dim FormClear As Boolean
    Dim ContactContentID As Long
    '
    Dim Criteria As String
    Dim ContentGorupCriteria As String
    Dim ContactSearchCriteria As String
    Dim FieldParms() As String
    Dim CriteriaValues As Variant
    Dim CriteriaCount As Long
    Dim CriteriaPointer As Long
    Dim PageSize As Long
    Dim PageNumber As Long
    Dim TopCount As Long
    Dim RowPointer As Long
    Dim DataRowCount As Long
    '
    Dim PreTableCopy As String
    Dim PostTableCopy As String
    '
    Dim ColumnCount As Long
    Dim CPtr As Long
    Dim ColCaption() As String
    Dim ColAlign() As String
    Dim ColWidth() As String
    Dim Cells() As String
    Dim ColSortable() As Boolean
    Dim DefaultSortColumnPtr As Long
    '
    Dim GroupID As Long
    Dim GroupToolAction As Long
    Dim ActionPanel As String
    Dim RowCount As Long
    Dim GroupName As String
    Dim memberID As Long
    Dim QS As String
    Dim VisitsCell As String
    Dim VisitCount As Long
    Dim AdminURL As String
    Dim CCID As Long
    Dim SQLValue As String
    Dim DefaultName As String
    Dim Copy As String
    Dim SearchCaption As String
    Dim BlankPanel As String
    Dim RowPageSize As String
    Dim RowGroups As String
    Dim GroupIDs() As String
    Dim GroupPtr As Long
    Dim GroupDelimiter As String
    Dim ButtonList As String
    Dim ButtonBar As String
    Dim Header As String
    Dim Body As String
    Dim AdminUI As Object
    Dim Description As String
    Dim IsAdmin As Boolean
    Dim SortColPtr As Long
    Dim SortColType As Long
    Dim TextTest As String
    '
    If Not (Main Is Nothing) Then
        Set AdminUI = CreateObject("ccWeb3.AdminUIClass")
        IsAdmin = Main.IsAdmin()
        RQS = Main.RefreshQueryString
        TextTest = Main.GetStreamText(RequestNamePageSize)
        If TextTest = "" Then
            PageSize = kmaEncodeInteger(Main.getvisitproperty("cmPageSize", "50"))
        Else
            PageSize = kmaEncodeInteger(TextTest)
            Call Main.setvisitproperty("cmPageSize", CStr(PageSize))
        End If
        'PageSize = Main.GetStreamInteger(RequestNamePageSize)
        If PageSize = 0 Then
            PageSize = 50
        End If
        PageNumber = Main.GetStreamInteger(RequestNamePageNumber)
        If PageNumber = 0 Then
            PageNumber = 1
        End If
        GroupID = Main.GetStreamInteger("GroupID")
        GroupToolAction = Main.GetStreamInteger("GroupToolAction")
        AdminURL = Main.SiteProperty_AdminURL
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
        TextTest = Main.getvisitproperty("cmSortColumn", "")
        If TextTest <> "" Then
            SortColPtr = kmaEncodeInteger(TextTest)
        End If
        SortColPtr = AdminUI.GetReportSortColumnPtr(Main, SortColPtr)
        If CStr(SortColPtr) <> TextTest Then
            Call Main.setvisitproperty("cmSortColumn", CStr(SortColPtr))
        End If
        If AdminUI.GetReportSortType(Main) = 2 Then
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
        RQS = Main.RefreshQueryString
        RQS = ModifyQueryString(RQS, "tab", "", False)
        Call BuildSearch(Criteria, SearchCaption)
        '
        ' purge duplicate in ccMemberRules
        '
        SQL = "Delete from ccMemberRules where ID in (" _
            & " SELECT DISTINCT B.ID" _
            & " FROM ccMemberRules A, ccMemberRules B" _
            & " WHERE (((A.MemberID)=[B].[MemberID]) AND ((A.GroupID)=[B].[GroupID]) AND ((A.ID)<[B].[ID]))" _
            & ")"
        Main.TestPoint ("Purge all ccMemberRule duplicates SQL = " & SQL)
        Call Main.ExecuteSQL("default", SQL)
        '
        '   Get DataRowCount
        '       This had been commented out - but it is needed for the "matches found" caption
        '
        SQL = "Select distinct ccMembers.ID" _
            & " from ccMembers" _
            & " left join ccMemberRules on ccMemberRules.MemberID=ccMembers.ID" _
            & Criteria
        Call Main.TestPoint("Selection SQL=" & SQL)
        CS = Main.OpenCSSQL("default", SQL)
        If Main.IsCSOK(CS) Then
            DataRowCount = Main.GetCSRowCount(CS)
        End If
        Call Main.CloseCS(CS)
        '
        '   Get Data
        '
        DefaultName = Main.GetContentFieldProperty("people", "name", "default")
        SQL = "Select distinct Top " & TopCount & " ccMembers.name,ccMembers.FirstName,ccMembers.LastName,ccMembers.ID, ccMembers.ContentControlID, ccMembers.Visits, ccMembers.LastVisit, ccMembers.Phone, ccMembers.Email,Organizations.Name as OrgName" _
            & " from ((ccMembers" _
            & " left join organizations on Organizations.ID=ccMembers.OrganizationID)" _
            & " left join ccMemberRules on ccMemberRules.MemberID=ccMembers.ID)" _
            & Criteria _
            & SQLOrderBy & SQLOrderDir
        Call Main.TestPoint("Selection SQL=" & SQL)
        CS = Main.OpenCSSQL("default", SQL, PageSize, PageNumber)
        RowPointer = 0
        If Not Main.IsCSOK(CS) Then
            Cells(0, 3) = "This search returned no results"
            RowPointer = 1
        Else
            'DataRowCount = Main.GetCSRowCount(CS)
            Do While Main.IsCSOK(CS) And (RowPointer < PageSize)
                CPtr = 0
                RecordID = Main.GetCSInteger(CS, "ID")
                CCID = Main.GetCSInteger(CS, "ContentControlID")
                VisitCount = Main.GetCSInteger(CS, "Visits")
                If IsAdminPath And (VisitCount > 0) Then
                    QS = GetReportQueryString(PageNumber, PageSize, RecordID, ReportFormVisitList)
                    VisitsCell = "<A href=""" & Main.ServerPage & "?" & QS & """ Target=""_blank"">" & VisitCount & "</A>"
                Else
                    VisitsCell = CStr(VisitCount)
                End If
                CheckBox = Main.GetFormInputCheckBox("M." & RowPointer, , "DelCheck")
                Cells(RowPointer, CPtr) = CheckBox & Main.GetFormInputHidden("MID." & RowPointer, RecordID)
                CPtr = CPtr + 1
                RecordName = Main.GetCStext(CS, "name")
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
                Cells(RowPointer, CPtr) = Main.GetCStext(CS, "FirstName")
                CPtr = CPtr + 1
                Cells(RowPointer, CPtr) = Main.GetCStext(CS, "LastName")
                CPtr = CPtr + 1
                Cells(RowPointer, CPtr) = Main.GetCStext(CS, "OrgName")
                CPtr = CPtr + 1
                Cells(RowPointer, CPtr) = Main.GetCStext(CS, "phone")
                CPtr = CPtr + 1
                Cells(RowPointer, CPtr) = Main.GetCStext(CS, "email")
                CPtr = CPtr + 1
                Cells(RowPointer, CPtr) = VisitsCell
                CPtr = CPtr + 1
                LastVisit = Int(Main.GetCSDate(CS, "LastVisit"))
                If LastVisit = CDate(0) Then
                    Cells(RowPointer, CPtr) = "&nbsp;"
                Else
                    Cells(RowPointer, CPtr) = LastVisit
                End If
                CPtr = CPtr + 1
                Cells(RowPointer, CPtr) = "&nbsp;"
                CPtr = CPtr + 1
                RowPointer = RowPointer + 1
                Call Main.NextCSRecord(CS)
            Loop
        End If
        Call Main.CloseCS(CS)
        GetFormList = GetFormList & Main.GetFormInputHidden("M.Count", RowPointer)
        '
        BlankPanel = "<div class=""cmBody ccPanel3DReverse"">x</div>"
        'Main.GetPanel("x", "ccPanelInput", "ccPanelShadow", "ccPanelHilite")
        RowPageSize = "<TABLE border=0 cellpadding=4 cellspacing=0 width=500>" _
            & "<TR>" _
            & "<TD class=APLeft>Rows Per Page</TD>" _
            & "<TD class=apright>" & Main.GetFormInputText(RequestNamePageSize, PageSize, 1, 10) & "</TD>" _
            & "</TR>" _
            & "</Table>"
        RowGroups = "<TABLE border=0 cellpadding=4 cellspacing=0 width=500><TR>" _
            & "<TD valign=top class=APLeft>Actions</TD>" _
            & "<TD class=APRight>" _
                & "" _
                & "<div class=APRight>Source Contacts</div>" _
                & "<div class=APRightIndent>" & Main.GetFormInputRadioBox("GroupToolSelect", 0, 0) & "&nbsp;Only those selected on this page</div>" _
                & "<div class=APRightIndent><input type=radio name=GroupToolSelect value=1>&nbsp;Everyone in search results</div>" _
                & "<div style=""border-top:1px solid black;border-bottom:1px solid white;margin-top:4px;margin-bottom:4px;""></div>" _
                & "<div class=APRight>Perform Action</div>" _
                & "<div class=APRightIndent>" & Main.GetFormInputRadioBox("GroupToolAction", 0, 0) & " No Action</div>" _
                & "<div class=APRightIndent>" & Main.GetFormInputRadioBox("GroupToolAction", 1, 0) & " Add to Target Group</div>" _
                & "<div class=APRightIndent>" & Main.GetFormInputRadioBox("GroupToolAction", 2, 0) & " Remove from Target Group</div>" _
                & "<div class=APRightIndent>" & Main.GetFormInputRadioBox("GroupToolAction", 3, 0) & " Export comma delimited file</div>" _
                & "<div class=APRightIndent>" & Main.GetFormInputRadioBox("GroupToolAction", 4, 0) & " Set Allow Group Email</div>" _
                & "<div style=""border-top:1px solid black;border-bottom:1px solid white;margin-top:4px;margin-bottom:4px;""></div>" _
                & "<div class=APRight style=""padding-bottom:6px;"">Target Group</div>" _
                & "<div class=APRightIndent>" & Main.GetFormInputSelect("GroupID", GroupID, "Groups") & "</div>" _
                & "</TD>" _
            & "</TR></Table>"
        
'& "<div class=APRightIndent>" & Main.GetFormInputRadioBox("GroupToolSelect", 1, 0) & " Everyone in search results</div>" _

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
            & Main.GetFormInputHidden("M.Count", RowPointer) _
            & Main.GetFormInputHidden(RequestNameFormID, FormList) _
            & ""
        '
        ' Header
        '
        Description = Description _
            & SearchCaption & "<BR>" & DataRowCount & " Matches found" _
            & StatusMessage _
            & vbCrLf
        Header = Main.GetPanel("<P>" & Description & "</P>", "ccPanel", "ccPanelShadow", "ccPanelHilite", "100%", 20)
        Header = "<div class=ccPanelBackground style=""padding:10px;"">" & Header & "</div>"
        ButtonList = ButtonApply & "," & ButtonNewSearch
        Body = AdminUI.GetReport2(Main, RowPointer, ColCaption, ColAlign, ColWidth, Cells, PageSize, PageNumber, PreTableCopy, PostTableCopy, DataRowCount, "ccPanel", ColSortable, SortColPtr)
        'Body = AdminUI.GetReport2(Main, RowPointer, ColCaption, ColAlign, ColWidth, Cells, PageSize, PageNumber, PreTableCopy, PostTableCopy, DataRowCount, "ccPanel", ColSortable, DefaultSortColumnPtr)
        'Body = "<div style=""Background-color:white;"">" & Body & "</div>"
        '
        ' Assemble page
        '
        GetFormList = AdminUI.GetBody(Main, "Contact Manager &gt;&gt; List People", ButtonList, "", True, True, Description, "", 0, Body)
        '
        End If
    '
    Exit Function
ErrorTrap:
    Call HandleError("CMngrClass", "GetFormList", Err.Number, Err.Source, Err.Description, True, False)
End Function
'
'=================================================================================
'
'=================================================================================
'
Private Function GetEventsTab() As String
    On Error GoTo ErrorTrap
    '
    Dim Argument1 As String
    Dim CS As Long
    Dim SelectedGroupIDList As String
    Dim GroupCount As Long
    '
    If Not (Main Is Nothing) Then
        '
        GetEventsTab = "events"
    End If
    '
    Exit Function
ErrorTrap:
    Call HandleError("CMngrClass", "GetEventsTab", Err.Number, Err.Source, Err.Description, True, False)
    End Function
'
'=================================================================================
'
'=================================================================================
'
Private Function ProcessSearchForm() As Long
    On Error GoTo ErrorTrap
    '
    Dim TextOption As String
    Dim NumericOption As String
    Dim Argument1 As String
    Dim CS As Long
    Dim SelectedGroupIDList As String
    Dim GroupCount As Long
    Dim GroupPointer As Long
    Dim GroupLabel As String
    Dim ContactGroupCriteria As String
    '
    Dim ContactSearchCriteria As String
    Dim FieldPointer As Long
    Dim GroupChecked As Boolean
    Dim GroupID As Long
    Dim RowEven As Boolean
    Dim Button As String
    Dim SQL As String
    Dim RQS As String
    Dim SubTab As Long
    Dim Criteria As String
    Dim VisitSearchCriteria As String
    Dim CriteriaNames() As String
    Dim CriteriaValues() As String
    Dim CriteriaCount As Long
    Dim CriteriaPointer As Long
    Dim FieldName As String
    Dim FieldValue As String
    Dim FieldCount As Long
    Dim FieldSize As Long
    Dim FormSave As Boolean
    Dim FormClear As Boolean
    Dim NameValues As Variant
    Dim RowPointer As Long
    Dim FieldOperator As String
    Dim fieldType As Long

    '
    Dim ContactContentID As Long
    '
    If Not (Main Is Nothing) Then
        If Main.GetStreamText("SelectionGroupSubTab") <> "" Then
            '
            ' Save the Form
            '
            GroupCount = Main.GetStreamInteger("GroupCount")
            If GroupCount > 0 Then
                For GroupPointer = 0 To GroupCount - 1
                    GroupLabel = "Group" & GroupPointer
                    If Main.GetStreamBoolean(GroupLabel) Then
                        ContactGroupCriteria = ContactGroupCriteria & "," & Main.GetStreamInteger(GroupLabel & ".id")
                    End If
                Next
            End If
            ContactGroupCriteria = ContactGroupCriteria & ","
            Call Main.SetMemberProperty("ContactGroupCriteria", ContactGroupCriteria)
        End If
        If Main.GetStreamText("SelectionContentSubTab") <> "" Then
            '
            ' SelectionContentSubTab
            '
            Call Main.TestPoint("ProcessSearchForm - SelectionContentSubTab")
            ContactContentID = Main.GetStreamInteger("ContentID")
            Call Main.SetMemberProperty("ContactContentID", ContactContentID)
        End If
        If Main.GetStreamText("SelectionSearchSubTab") <> "" Then
            '
            ' SelectionContentSubTab (crlf FieldName tab FieldType tab FieldVAlue tab Operator)
            '
            Call Main.TestPoint("ProcessSearchForm - SelectionSearchSubTab")
            Criteria = "(active<>0)and(ContentID=" & Main.GetContentID("people") & ")and(authorable<>0)"
            CS = Main.OpenCSContent("Content Fields", Criteria, "EditSortPriority")
            Do While Main.IsCSOK(CS)
                FieldName = Main.GetCStext(CS, "name")
                FieldValue = Main.GetStreamText(FieldName)
                fieldType = Main.GetCSInteger(CS, "Type")
                Select Case fieldType
                    Case FieldTypeDate
                        NumericOption = Main.GetStreamText(FieldName & "_D")
                        If NumericOption <> "" Then
                            ContactSearchCriteria = ContactSearchCriteria _
                                & vbCrLf _
                                & FieldName & vbTab _
                                & fieldType & vbTab _
                                & FieldValue & vbTab _
                                & NumericOption
                        End If
                    Case FieldTypeCurrency, FieldTypeFloat, FieldTypeInteger
                        NumericOption = Main.GetStreamText(FieldName & "_N")
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
                        TextOption = Main.GetStreamText(FieldName & "_T")
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
                Call Main.NextCSRecord(CS)
            Loop
            Call Main.CloseCS(CS)
            Call Main.SetMemberProperty("ContactSearchCriteria", ContactSearchCriteria)
        End If
        ProcessSearchForm = FormList
    End If
    '
    Exit Function
ErrorTrap:
    Call HandleError("CMngrClass", "ProcessSearchForm", Err.Number, Err.Source, Err.Description, True, False)
End Function
'
'
'
Private Function GetReportQueryString(ReportPageNumber As Long, ReportPageSize As Long, memberID As Long, ReportID As Long) As String
    On Error GoTo ErrorTrap
    '
    '
    GetReportQueryString = Main.RefreshQueryString
    GetReportQueryString = ModifyQueryString(GetReportQueryString, "aa", 0, False)
    GetReportQueryString = ModifyQueryString(GetReportQueryString, RequestNameAdminForm, AdminFormReports, True)
    GetReportQueryString = ModifyQueryString(GetReportQueryString, RequestNameAdminSourceForm, AdminFormReports, True)
    GetReportQueryString = ModifyQueryString(GetReportQueryString, "rid", CStr(ReportID), True)
    GetReportQueryString = ModifyQueryString(GetReportQueryString, "DateTo", CStr(Int(Now + 1)), True)
    GetReportQueryString = ModifyQueryString(GetReportQueryString, "DateFrom", "1/1/2000", True)
    GetReportQueryString = ModifyQueryString(GetReportQueryString, "MemberID", CStr(memberID), True)
    GetReportQueryString = ModifyQueryString(GetReportQueryString, "ExcludeBrowsers", "0", True)
    GetReportQueryString = ModifyQueryString(GetReportQueryString, "ExcludeIP", "0", True)
    GetReportQueryString = ModifyQueryString(GetReportQueryString, "ExcludeNewVisitors", "0", True)
    GetReportQueryString = ModifyQueryString(GetReportQueryString, "ExcludeOldVisitors", "0", True)
    GetReportQueryString = ModifyQueryString(GetReportQueryString, "PageNumber", CStr(ReportPageNumber), True)
    GetReportQueryString = ModifyQueryString(GetReportQueryString, "PageSize", CStr(ReportPageSize), True)
    '
    Exit Function
ErrorTrap:
    Call HandleError("CMngrClass", "GetReportQueryString", Err.Number, Err.Source, Err.Description, True, False)
    End Function
'
'
'
Private Function GetGroupCaption(GroupID As String) As String
    On Error GoTo ErrorTrap
    '
    Dim CS As Long
    '
    CS = Main.OpenCSContentRecord("groups", kmaEncodeInteger(GroupID), , , "Caption,name")
    If Main.IsCSOK(CS) Then
        GetGroupCaption = Main.GetCStext(CS, "caption")
        If GetGroupCaption = "" Then
            GetGroupCaption = Main.GetCStext(CS, "name")
            If GetGroupCaption = "" Then
                GetGroupCaption = "Group " & GroupID
            End If
        End If
    End If
    Call Main.CloseCS(CS)
'    GetGroupCaption = "Group " & GroupID
    '
    Exit Function
ErrorTrap:
    Call HandleError("CMngrClass", "GetGroupCaption", Err.Number, Err.Source, Err.Description, True, False)
End Function
'
'
'
Private Sub BuildSearch(Criteria As String, SearchCaption As String)
    On Error GoTo ErrorTrap
    '
    Dim LookupName As String
    Dim ContactContentID As Long
    Dim ContactGroupCriteria As String
    Dim ContactSearchCriteria As String
    Dim CriteriaValues() As String
    Dim CriteriaCount As Long
    Dim CriteriaPointer As Long
    Dim FieldParms() As String
    Dim SQLValue As String
    Dim GroupIDs() As String
    Dim GroupDelimiter As String
    Dim GroupPtr As Long
    Dim ContentName As String
    Dim GroupCriteria As String
    Dim SQLNow As String
    Dim SQL As String
    Dim CS As Long
    Dim FieldCaption As String
    Dim FieldContentLookupID As Long
    Dim FieldContentLookupName As String
    Dim hint As String
    '
    hint = "100"
    ContactContentID = kmaEncodeInteger(Main.GetMemberProperty("ContactContentID", Main.GetContentID("people")))
    ContactGroupCriteria = Main.GetMemberProperty("ContactGroupCriteria", "")
    ContactSearchCriteria = Main.GetMemberProperty("ContactSearchCriteria", "")
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
                    SQL = "select caption,LookupContentID from ccfields where name=" & KmaEncodeSQLText(FieldParms(0)) & " and contentid=" & ContactContentID
                    CS = Main.OpenCSSQL("default", SQL)
                    If Main.IsCSOK(CS) Then
                        FieldCaption = Main.GetCStext(CS, "caption")
                        FieldContentLookupID = Main.GetCSInteger(CS, "LookupContentID")
                    End If
                    Call Main.CloseCS(CS)
                    If FieldCaption = "" Then
                        FieldCaption = FieldParms(0)
                        FieldContentLookupID = 0
                    End If
                    Select Case kmaEncodeInteger(FieldParms(1))
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
                                        SQLValue = KmaEncodeSQLText(FieldParms(2))
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
                                FieldContentLookupName = Main.GetRecordName("Content", FieldContentLookupID)
                                If (FieldContentLookupName <> "") And IsNumeric(FieldParms(2)) Then
                                    LookupName = Main.GetRecordName(FieldContentLookupName, FieldParms(2))
                                End If
                            End If
                            If LookupName = "" Then
                                LookupName = FieldParms(2)
                            End If
                            
                            Criteria = Criteria & "AND(ccMembers." & FieldParms(0) & "=" & KmaEncodeSQLNumber(FieldParms(2)) & ")"
                            SearchCaption = SearchCaption & ", " & FieldCaption & " = " & LookupName
                        Case FieldTypeDate
                            '
                            ' date
                            '
                            hint = hint & ",185"
                            Select Case kmaEncodeInteger(FieldParms(3))
                                Case 1
                                    Criteria = Criteria & "AND(ccMembers." & FieldParms(0) & "=" & KmaEncodeSQLDate(FieldParms(2)) & ")"
                                    SearchCaption = SearchCaption & ", " & FieldCaption & " = " & FieldParms(2)
                                Case 2
                                    Criteria = Criteria & "AND(ccMembers." & FieldParms(0) & ">" & KmaEncodeSQLDate(FieldParms(2)) & ")"
                                    SearchCaption = SearchCaption & ", " & FieldCaption & " > " & FieldParms(2)
                                Case 3
                                    Criteria = Criteria & "AND(ccMembers." & FieldParms(0) & "<" & KmaEncodeSQLDate(FieldParms(2)) & ")"
                                    SearchCaption = SearchCaption & ", " & FieldCaption & " < " & FieldParms(2)
                            End Select
                        Case FieldTypeCurrency, FieldTypeFloat, FieldTypeInteger, FieldTypeLookup
                            '
                            ' number
                            '
                            hint = hint & ",190"
                            Select Case kmaEncodeInteger(FieldParms(3))
                                Case 1
                                    Criteria = Criteria & "AND(ccMembers." & FieldParms(0) & "=" & KmaEncodeSQLNumber(FieldParms(2)) & ")"
                                    SearchCaption = SearchCaption & ", " & FieldCaption & " = " & FieldParms(2)
                                Case 2
                                    Criteria = Criteria & "AND(ccMembers." & FieldParms(0) & ">" & KmaEncodeSQLNumber(FieldParms(2)) & ")"
                                    SearchCaption = SearchCaption & ", " & FieldCaption & " > " & FieldParms(2)
                                Case 3
                                    Criteria = Criteria & "AND(ccMembers." & FieldParms(0) & "<" & KmaEncodeSQLNumber(FieldParms(2)) & ")"
                                    SearchCaption = SearchCaption & ", " & FieldCaption & " < " & FieldParms(2)
                            End Select
                        Case FieldTypeBoolean
                            '
                            ' boolean
                            '
                            hint = hint & ",200"
                            If kmaEncodeBoolean(FieldParms(2)) Then
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
                            SearchCaption = SearchCaption & ", " & FieldCaption & " is " & kmaEncodeBoolean(FieldParms(2))
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
        SQLNow = KmaEncodeSQLDate(Now)
        If UBound(GroupIDs) = 0 Then
            hint = hint & ",240"
            If SearchCaption = "" Then
                SearchCaption = " in group " & "'" & GetGroupCaption(GroupIDs(GroupPtr)) & "'"
            Else
                SearchCaption = SearchCaption & ", in group " & "'" & GetGroupCaption(GroupIDs(GroupPtr)) & "'"
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
                    SearchCaption = SearchCaption & " or " & "'" & GetGroupCaption(GroupIDs(GroupPtr)) & "'"
                Else
                    hint = hint & ",280"
                    SearchCaption = SearchCaption & GroupDelimiter & "'" & GetGroupCaption(GroupIDs(GroupPtr)) & "'"
                End If
                GroupDelimiter = ", "
            Next
            Criteria = Criteria & "and(" & Mid(GroupCriteria, 3) & ")"
        End If
    End If
    '
    ' Add Content Criteria
    '
    hint = hint & ",290"
    ContentName = Main.GetContentNameByID(ContactContentID)
    If ContentName = "" Then
        SearchCaption = "Everyone" & SearchCaption
    ElseIf LCase(ContentName) <> "people" Then
        SearchCaption = "Every " & ContentName & SearchCaption
        Criteria = Criteria & "AND" & Main.GetContentControlCriteria(ContentName)
    Else
        SearchCaption = "Everyone " & SearchCaption
    End If
    '
    hint = hint & ",300"
    If Criteria <> "" Then
        Criteria = " WHERE " & Mid(Criteria, 4)
    End If
'    Call BuildSearch(Criteria, SearchCaption)

    '
    Exit Sub
ErrorTrap:
    Call HandleError("CMngrClass", "BuildSearch, hint=[" & hint & "]", Err.Number, Err.Source, Err.Description, True, False)
End Sub
'
'=================================================================================
'
'=================================================================================
'
Private Function GetFormDetail(DetailMemberID As Long, StatusMessage As String, IsGoMethodEnabled As Boolean) As String
    On Error GoTo ErrorTrap
    '
    Dim SubTab As Long
    Dim RQS As String
    Dim Nav As New MenuLiveTabClass
    Dim Header As String
    Dim AdminUI As Object
    Dim Content As String
    Dim ButtonList As String
    Dim Button As String
    Dim CS As Long
    Dim MemberName As String
    Dim hint As String
    '
    If Not (Main Is Nothing) Then
        Set AdminUI = CreateObject("ccWeb3.AdminUIClass")
        CS = Main.OpenCSContent("people", "ID=" & DetailMemberID, , False)
        MemberName = Main.GetCStext(CS, "name")
        If MemberName = "" Then
            MemberName = Trim(Main.GetCStext(CS, "FirstName") & " " & Main.GetCStext(CS, "LastName"))
            If MemberName = "" Then
                MemberName = "Record " & Main.GetCStext(CS, "ID")
            End If
        End If
        '
        ' Determine current Subtab
        '
        SubTab = Main.GetStreamInteger(RequestNameDetailSubtab)
        If SubTab = 0 Then
            SubTab = kmaEncodeInteger(Main.GetMemberProperty(RequestNameDetailSubtab, "1"))
            If SubTab = 0 Then
                SubTab = 1
                Call Main.SetMemberProperty(RequestNameDetailSubtab, CStr(SubTab))
            End If
        Else
            Call Main.SetMemberProperty(RequestNameDetailSubtab, CStr(SubTab))
        End If
        Call Main.AddRefreshQueryString(RequestNameDetailSubtab, SubTab)
        '
        ' SubTab Menu
        '
        Call Main.AddRefreshQueryString("tab", "")
        ButtonList = ButtonCancel & "," & ButtonSave & "," & ButtonOK & "," & ButtonNewSearch
        Header = "<div>" & MemberName & "</div>"
        '
        Call Nav.AddEntry("Contact", GetFormDetail_TabContact(CS), "ccAdminTab")
        Call Nav.AddEntry("Permissions", GetFormDetail_TabPermissions(CS), "ccAdminTab")
        Call Nav.AddEntry("Notes", GetFormDetail_TabNotes(CS), "ccAdminTab")
        Call Nav.AddEntry("Photos", GetFormDetail_TabPhoto(CS), "ccAdminTab")
        Call Nav.AddEntry("Groups", GetFormDetail_TabGroup(CS), "ccAdminTab")
        If IsGoMethodEnabled Then
            Call Nav.AddEntry("Homes&nbsp;Viewed", GetFormDetail_TabHomes(CS), "ccAdminTab")
        End If
        Call Main.CloseCS(CS)
        '
        Content = Nav.GetTabs()
        Content = Content & Main.GetFormInputHidden(RequestNameFormID, FormDetails)
        Content = Content & Main.GetFormInputHidden(RequestNameMemberID, DetailMemberID)
        '
        GetFormDetail = AdminUI.GetBody(Main, "Contact Manager &gt;&gt; Contact Details", ButtonList, "", True, True, Header, "", 0, Content)
    End If
    Exit Function
ErrorTrap:
    Call HandleError("CMngrClass", "GetFormDetail", Err.Number, Err.Source, Err.Description, True, False)
End Function
'
'=================================================================================
'
'=================================================================================
'
Private Function GetFormDetail_TabContact(CS As Long) As String
    On Error GoTo ErrorTrap
    '
    Dim s As String
    Dim Left As String
    Dim Right As String
    Dim Copy As String
    '
    If Not Main.IsCSOK(CS) Then
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
        Left = Left & GetFormDetail_InputTextRow("Full Name", "Name", Main.GetCStext(CS, "name"), False)
        Left = Left & GetFormDetail_InputTextRow("First Name", "FirstName", Main.GetCStext(CS, "FirstName"), False)
        Left = Left & GetFormDetail_InputTextRow("Last Name", "LastName", Main.GetCStext(CS, "LastName"), False)
        Left = Left & GetFormDetail_DividerRow("Contact")
        Left = Left & GetFormDetail_InputTextRow("Email", "EMAIL", Main.GetCStext(CS, "email"), False)
        Left = Left & GetFormDetail_InputTextRow("Phone", "PHONE", Main.GetCStext(CS, "PHONE"), False)
        Left = Left & GetFormDetail_InputTextRow("Fax", "Fax", Main.GetCStext(CS, "Fax"), False)
        Left = Left & GetFormDetail_InputTextRow("Address", "ADDRESS", Main.GetCStext(CS, "ADDRESS"), False)
        Left = Left & GetFormDetail_InputTextRow("Line 2", "ADDRESS2", Main.GetCStext(CS, "ADDRESS2"), False)
        Left = Left & GetFormDetail_InputTextRow("City", "City", Main.GetCStext(CS, "City"), False)
        Left = Left & GetFormDetail_InputTextRow("State", "State", Main.GetCStext(CS, "State"), False)
        Left = Left & GetFormDetail_InputTextRow("Zip", "Zip", Main.GetCStext(CS, "Zip"), False)
        Left = Left & GetFormDetail_DividerRow("Company")
        Left = Left & GetFormDetail_InputTextRow("Name", "Company", Main.GetCStext(CS, "Company"), False)
        Left = Left & GetFormDetail_InputTextRow("Title", "Title", Main.GetCStext(CS, "Title"), False)
        Left = Left & GetFormDetail_DividerRow("Birthday")
        Left = Left & GetFormDetail_InputTextRow("Day", "BirthdayDay", Main.GetCStext(CS, "BirthdayDay"), False)
        Left = Left & GetFormDetail_InputTextRow("Month", "BirthdayMonth", Main.GetCStext(CS, "BirthdayMonth"), False)
        Left = Left & GetFormDetail_InputTextRow("Year", "BirthdayYear", Main.GetCStext(CS, "BirthdayYear"), False)
        Left = Left & "</table>"
        '
        ' Right Side
        '
        Copy = Main.GetFormInputText("AppendNotes", "", 20, 100)
        Copy = Replace(Copy, " cols=""100""", " style=""width:100%;""", , , vbTextCompare)
        '
        Right = Right & "<table border=0 width=100% cellspacing=0 cellpadding=0>" _
                & "<TR>" _
                & "<TD width=200><img src=/cclib/images/spacer.gif width=150 height=1></TD>" _
                & "<TD width=99%><img src=/cclib/images/spacer.gif width=1 height=1></TD>" _
                & "</TR>"
        Right = Right & GetFormDetail_DividerRow("Add to Notes")
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
    GetFormDetail_TabContact = s
    '
    Exit Function
ErrorTrap:
    Call HandleError("CMngrClass", "GetFormDetail_TabContact", Err.Number, Err.Source, Err.Description, True, False)
    End Function
'
'=================================================================================
'
'=================================================================================
'
Private Function GetFormDetail_TabPermissions(CS As Long) As String
    On Error GoTo ErrorTrap
    '
    Dim s As String
    '
    If Not Main.IsCSOK(CS) Then
        '
        s = s & "<div>There was a problem retrieving this person's information.</div>"
    Else
        '
        s = s & "<table border=0 width=100% cellspacing=0 cellpadding=0>" _
                & "<TR>" _
                & "<TD width=200><img src=/cclib/images/spacer.gif width=150 height=1></TD>" _
                & "<TD width=99%><img src=/cclib/images/spacer.gif width=1 height=1></TD>" _
                & "</TR>"
        s = s & GetFormDetail_DividerRow("Login")
        s = s & GetFormDetail_InputTextRow("Username", "Username", Main.GetCStext(CS, "Username"), False)
        s = s & GetFormDetail_InputTextRow("Password", "Password", Main.GetCStext(CS, "Password"), True)
        s = s & GetFormDetail_InputBooleanRow("Allow Auto Login", "AutoLogin", Main.GetCSBoolean(CS, "AutoLogin"))
        s = s & "</table>"
    End If
    '
    s = "<div STYLE=""width:100%;"" class=""cmBody ccPanel3DReverse"">" & s & "</div>"
    GetFormDetail_TabPermissions = s
    '
    Exit Function
ErrorTrap:
    Call HandleError("CMngrClass", "GetFormDetail_TabPermissions", Err.Number, Err.Source, Err.Description, True, False)
    End Function
'
'=================================================================================
'
'=================================================================================
'
Private Function GetFormDetail_TabNotes(CS As Long) As String
    On Error GoTo ErrorTrap
    '
    Dim s As String
    '
    If Not Main.IsCSOK(CS) Then
        '
        s = s & "<div>There was a problem retrieving this person's information.</div>"
    Else
        '
        s = s & "<table border=0 width=100% cellspacing=0 cellpadding=0>" _
                & "<TR>" _
                & "<TD width=200><img src=/cclib/images/spacer.gif width=150 height=1></TD>" _
                & "<TD width=99%><img src=/cclib/images/spacer.gif width=1 height=1></TD>" _
                & "</TR>"
        s = s & GetFormDetail_InputHTMLRow("Notes", "NotesFilename", Main.GetCS(CS, "NotesFilename"))
        s = s & "</table>"
    End If
    '
    s = "<div STYLE=""width:100%;"" class=""cmBody ccPanel3DReverse"">" & s & "</div>"
    GetFormDetail_TabNotes = s
    '
    Exit Function
ErrorTrap:
    Call HandleError("CMngrClass", "GetFormDetail_TabNotes", Err.Number, Err.Source, Err.Description, True, False)
    End Function
'
'=================================================================================
'
'=================================================================================
'
Private Function GetFormDetail_TabPhoto(CS As Long) As String
    On Error GoTo ErrorTrap
    '
    Dim s As String
    '
    If Not Main.IsCSOK(CS) Then
        '
        s = s & "<div>There was a problem retrieving this person's information.</div>"
    Else
        '
        s = s & "<table border=0 width=100% cellspacing=0 cellpadding=0>" _
                & "<TR>" _
                & "<TD width=200><img src=/cclib/images/spacer.gif width=150 height=1></TD>" _
                & "<TD width=99%><img src=/cclib/images/spacer.gif width=1 height=1></TD>" _
                & "</TR>"
        s = s & GetFormDetail_InputImageRow("Thumbnail", "ThumbnailFilename", Main.GetCStext(CS, "ThumbnailFilename"))
        s = s & GetFormDetail_InputImageRow("Image", "ImageFilename", Main.GetCStext(CS, "ImageFilename"))
        s = s & "</table>"
    End If
    '
    s = "<div STYLE=""width:100%;"" class=""cmBody ccPanel3DReverse"">" & s & "</div>"
    GetFormDetail_TabPhoto = s
    '
    Exit Function
ErrorTrap:
    Call HandleError("CMngrClass", "GetFormDetail_TabPhoto", Err.Number, Err.Source, Err.Description, True, False)
    End Function
'
'=================================================================================
'
'=================================================================================
'
Private Function GetFormDetail_InputTextRow(Caption As String, FieldName As String, DefaultValue As String, IsPassword As Boolean) As String
    On Error GoTo ErrorTrap
    '
    Dim s As String
    '
    s = "" _
        & "<TR><TD style=""TEXT-ALIGN:left;PADDING-LEFT:20px;"">" _
        & Caption & ":" _
        & "</TD><TD style=""TEXT-ALIGN:left;"">"
    If IsPassword Then
        s = s & "<input type=password name=""" & FieldName & """ value=""" & kmaEncodeHTML(DefaultValue) & """ style=""width:300px;"">"
    Else
        s = s & "<input type=text name=""" & FieldName & """ value=""" & kmaEncodeHTML(DefaultValue) & """ style=""width:350px;"">"
    End If
    s = s & "</TD></TR>"
    GetFormDetail_InputTextRow = s
    '
    Exit Function
ErrorTrap:
    Call HandleError("CMngrClass", "GetFormDetail_InputTextRow", Err.Number, Err.Source, Err.Description, True, False)
    End Function
'
'=================================================================================
'
'=================================================================================
'
Private Function GetFormDetail_InputBooleanRow(Caption As String, FieldName As String, DefaultValue As String) As String
    On Error GoTo ErrorTrap
    '
    Dim s As String
    '
    s = "" _
        & "<TR><TD style=""TEXT-ALIGN:left;PADDING-LEFT:20px;"">" _
        & Caption & ":" _
        & "</TD><TD style=""TEXT-ALIGN:left;"">"
    s = s & "<input type=checkbox name=""" & FieldName & """ value=""" & DefaultValue & """>"
    s = s & "</TD></TR>"
    GetFormDetail_InputBooleanRow = s
    '
    Exit Function
ErrorTrap:
    Call HandleError("CMngrClass", "GetFormDetail_InputBooleanRow", Err.Number, Err.Source, Err.Description, True, False)
    End Function
'
'=================================================================================
'
'=================================================================================
'
Private Function GetFormDetail_InputHTMLRow(Caption As String, FieldName As String, DefaultValue As String) As String
    On Error GoTo ErrorTrap
    '
    Dim s As String
    '
    s = "" _
        & "<TR><TD style=""TEXT-ALIGN:left;PADDING-LEFT:20px;"">" _
        & Caption & ":" _
        & "</TD><TD style=""TEXT-ALIGN:left;"">"
    s = s & Main.GetFormInputHTML(FieldName, DefaultValue)
    s = s & "</TD></TR>"
    GetFormDetail_InputHTMLRow = s
    '
    Exit Function
ErrorTrap:
    Call HandleError("CMngrClass", "GetFormDetail_InputHTMLRow", Err.Number, Err.Source, Err.Description, True, False)
    End Function
'
'=================================================================================
'
'=================================================================================
'
Private Function GetFormDetail_InputImageRow(Caption As String, FieldName As String, DefaultValue As String) As String
    On Error GoTo ErrorTrap
    '
    Dim s As String
    Dim EncodedLink As String
    Dim Filename As String
    
    '
    s = "" _
        & "<TR><TD style=""TEXT-ALIGN:left;PADDING-LEFT:20px;"">" _
        & Caption & ":" _
        & "</TD><TD style=""TEXT-ALIGN:left;"">"
    If DefaultValue = "" Then
        s = s & Main.GetFormInputFile(FieldName)
        s = s & "</TD></TR>"
    Else
        Filename = kmaEncodeHTML(kmaGetFilename(DefaultValue))
        EncodedLink = kmaEncodeURL("http://" & Main.ServerHost & Main.serverFilePath & DefaultValue)
        s = s _
            & "<table border=0 width=100% cellspacing=0 cellpadding=4><TR>" _
            & "<TD width=200><a href=""" & EncodedLink & """ target=""_blank""><img src=""" & EncodedLink & """ width=200 border=0></a></TD>" _
            & "<TD width=100% valign=top>" _
            & "<div style=""height:20px;"">Filename:&nbsp;" & Filename & "</div>" _
            & "<div style=""height:20px;"">URL:&nbsp;" & EncodedLink & "</div>" _
            & "<div style=""height:20px;""><a href=""" & EncodedLink & """ target=""_blank"">Click for Full Size</A></div>" _
            & "<div style=""height:20px;""><span style=""width:100px;"">Delete:</span>" & Main.GetFormInputCheckBox(FieldName & ".DeleteFlag", False) & "</div>" _
            & "<div style=""height:20px;""><span style=""width:100px;"">Change:</span>" & Main.GetFormInputFile(FieldName) & "</div>" _
            & "</TD></TR>" _
            & "</table>" _
            & "" _
            & "</TD></TR>"
    End If
    GetFormDetail_InputImageRow = s
    '
    Exit Function
ErrorTrap:
    Call HandleError("CMngrClass", "GetFormDetail_InputImageRow", Err.Number, Err.Source, Err.Description, True, False)
    End Function
'
'=================================================================================
'
'=================================================================================
'
Private Function GetFormDetail_DividerRow(Caption As String) As String
    On Error GoTo ErrorTrap
    '
    Dim s As String
    '
    s = Replace(Caption, " ", "&nbsp;")
    s = "<TR><TD colspan=2 style=""Padding-top:10px;"">" _
        & "<TABLE border=0 width=100% cellspacing=0 cellpadding=0>" _
        & "<TR><TD width=1 style=""white-space:nowrap;"">" & s & "&nbsp;&nbsp;</TD><TD width=100%><HR></TD>" _
        & "</TABLE>" _
        & "</tr>"
    GetFormDetail_DividerRow = s
    '
    Exit Function
ErrorTrap:
    Call HandleError("CMngrClass", "GetFormDetail_DividerRow", Err.Number, Err.Source, Err.Description, True, False)
    End Function
'
'=================================================================================
'
'=================================================================================
'
Private Function SaveContactFromStream(DetailMemberID As Long) As String
    On Error GoTo ErrorTrap
    '
    Dim CS As String
    Dim Filename As String
    Dim OriginalFilename As String
    Dim Path As String
    Dim FieldName As String
    Dim Copy As String
    '
    CS = Main.OpenCSContent("people", "id=" & DetailMemberID)
    If Main.IsCSOK(CS) Then
        Call Main.SetCS(CS, "name", Main.GetStreamText("name"))
        Call Main.SetCS(CS, "FirstName", Main.GetStreamText("FirstName"))
        Call Main.SetCS(CS, "LastName", Main.GetStreamText("LastName"))
        '
        ' Contact
        '
        Call Main.SetCS(CS, "email", Main.GetStreamText("email"))
        Call Main.SetCS(CS, "Phone", Main.GetStreamText("Phone"))
        Call Main.SetCS(CS, "Fax", Main.GetStreamText("Fax"))
        Call Main.SetCS(CS, "Address", Main.GetStreamText("Address"))
        Call Main.SetCS(CS, "Address2", Main.GetStreamText("Address2"))
        Call Main.SetCS(CS, "City", Main.GetStreamText("City"))
        Call Main.SetCS(CS, "State", Main.GetStreamText("State"))
        Call Main.SetCS(CS, "Zip", Main.GetStreamText("Zip"))
        '
        ' Company
        '
        Call Main.SetCS(CS, "Company", Main.GetStreamText("Company"))
        Call Main.SetCS(CS, "Title", Main.GetStreamText("Title"))
        '
        ' Birthday
        '
        Call Main.SetCS(CS, "BirthdayDay", Main.GetStreamInteger("BirthdayDay"))
        Call Main.SetCS(CS, "BirthdayMonth", Main.GetStreamInteger("BirthdayMonth"))
        Call Main.SetCS(CS, "BirthdayYear", Main.GetStreamInteger("BirthdayYear"))
        '
        ' Notes
        '
        If Main.SiteProperty_BuildVersion > "3.3.554" Then
            Copy = Main.GetStreamText("AppendNotes")
            If Copy <> "" Then
                Copy = "" _
                    & "<div style=""margin-top:10px;border-top:1px dashed black;"">Added " & Now & " by " & Main.GetRecordName("people", Main.memberID) & "</div>" _
                    & "<div style=""margin-left:20px;margin-top:5px;"">" & Copy & "</div>"
            End If
            Call Main.SetCS(CS, "NotesFilename", Main.GetStreamText("NotesFilename") & Copy)
        End If
        '
        ' Photos
        '
        FieldName = "ThumbnailFilename"
        If Main.GetStreamBoolean(FieldName & ".DeleteFlag") Then
            Call Main.SetCSField(CS, FieldName, Null)
        End If
        OriginalFilename = Main.GetStreamText(FieldName)
        If OriginalFilename <> "" Then
            Filename = Main.GetCSFilename(CS, FieldName, OriginalFilename)
            Path = Filename
            Path = Replace(Path, "/", "\")
            Path = Replace(Path, "\" & OriginalFilename, "")
            Call Main.SetCSField(CS, FieldName, Filename)
            Call Main.SaveStreamFile(FieldName, Path)
        End If
        '
        FieldName = "ImageFilename"
        If Main.GetStreamBoolean(FieldName & ".DeleteFlag") Then
            Call Main.SetCSField(CS, FieldName, Null)
        End If
        OriginalFilename = Main.GetStreamText(FieldName)
        If OriginalFilename <> "" Then
            Filename = Main.GetCSFilename(CS, FieldName, OriginalFilename)
            Path = Filename
            Path = Replace(Path, "/", "\")
            Path = Replace(Path, "\" & OriginalFilename, "")
            Call Main.SetCSField(CS, FieldName, Filename)
            Call Main.SaveStreamFile(FieldName, Path)
        End If
        '
        ' Permissions
        '
        Call Main.SetCS(CS, "Username", Main.GetStreamText("Username"))
        Call Main.SetCS(CS, "Password", Main.GetStreamText("Password"))
        Call Main.SetCS(CS, "AutoLogin", Main.GetStreamBoolean("AutoLogin"))
        '
        ' Groups
        '
        Call SaveMemberRules(CS)
    End If
    Call Main.CloseCS(CS)
    '
    Exit Function
ErrorTrap:
    Call HandleError("CMngrClass", "SaveContactFromStream", Err.Number, Err.Source, Err.Description, True, False)
    End Function
'
'========================================================================
' Print the Member Edit form
'
'   Content must conform to ccMember fields
'========================================================================
'
Private Function GetFormDetail_TabGroup(CSMember As Long) As String
    On Error GoTo ErrorTrap
    '
    Dim f As New fastStringClass
    Dim Copy As String
    Dim SQL As String
    Dim CS As Long
    Dim MembershipCount As Long
    Dim MembershipSize As Long
    Dim MembershipPointer As Long
    Dim SectionName As String
    Dim PrimaryContentID As Long
    Dim SecondaryContentID As Long
    Dim CanSeeHiddenGroups As Boolean
    Dim DateExpireValue As String
    Dim GroupCount As Long
    Dim GroupID As Long
    Dim GroupName As String
    Dim GroupCaption As String
    Dim GroupActive As Boolean
    Dim Membership() As Long
    Dim DateExpires() As Date
    Dim active() As Boolean
    Dim Caption As String
    Dim MethodName As String
    Dim ReportLink As String
    Dim AdminUI As Object
    Dim DetailMemberID As Long
    '
    MethodName = "GetFormDetail_TabGroup"
    '
    Set AdminUI = CreateObject("ccWeb3.AdminUIClass")
    DetailMemberID = Main.GetCSInteger(CSMember, "ID")
    '
    ' ----- Gather all the SecondaryContent that associates to the PrimaryContent
    '
    PrimaryContentID = Main.GetContentID("People")
    SecondaryContentID = Main.GetContentID("Groups")
    '
    MembershipCount = 0
    MembershipSize = 0
    f.Add vbCrLf & "<!-- GroupRule Table --><table border=0 width=100% cellspacing=0 cellpadding=0>" _
        & "<TR>" _
        & "<TD width=150><img src=/cclib/images/spacer.gif width=150 height=1></TD>" _
        & "<TD width=99%><img src=/cclib/images/spacer.gif width=1 height=1></TD>" _
        & "</TR>"
    If DetailMemberID <> 0 Then
        '
        ' ----- read in the groups that this member has subscribed (exclude new member records)
        '
        SQL = "SELECT Active,GroupID,DateExpires" _
            & " FROM " & Main.SQLTableMemberRules & "" _
            & " WHERE MemberID=" & DetailMemberID
        CS = Main.OpenCSSQL("Default", SQL)
        Do While Main.IsCSOK(CS)
            If MembershipCount >= MembershipSize Then
                MembershipSize = MembershipSize + 10
                ReDim Preserve Membership(MembershipSize)
                ReDim Preserve active(MembershipSize)
                ReDim Preserve DateExpires(MembershipSize)
            End If
            Membership(MembershipCount) = Main.GetCSInteger(CS, "GroupID")
            DateExpires(MembershipCount) = Main.GetCSDate(CS, "DateExpires")
            active(MembershipCount) = Main.GetCSBoolean(CS, "Active")
            MembershipCount = MembershipCount + 1
            Main.NextCSRecord (CS)
        Loop
        Call Main.CloseCS(CS)
        '
        ' ----- read in all the groups, sorted by ContentName
        '
        SQL = "SELECT " & Main.SQLTableGroups & ".ID AS ID, ccContent.Name AS SectionName, " & Main.SQLTableGroups & ".Caption AS GroupCaption, " & Main.SQLTableGroups & ".name AS GroupName, " & Main.SQLTableGroups & ".SortOrder" _
            & " FROM " & Main.SQLTableGroups & " LEFT JOIN ccContent ON " & Main.SQLTableGroups & ".ContentControlID = ccContent.ID" _
            & " Where (((" & Main.SQLTableGroups & ".Active) <> " & SQLFalse & ") And ((ccContent.Active) <> " & SQLFalse & "))"
        SQL = SQL _
            & " GROUP BY " & Main.SQLTableGroups & ".ID, ccContent.Name, " & Main.SQLTableGroups & ".Caption, " & Main.SQLTableGroups & ".name, " & Main.SQLTableGroups & ".SortOrder" _
            & " ORDER BY ccContent.Name, " & Main.SQLTableGroups & ".Caption"
        CS = Main.OpenCSSQL("Default", SQL)
        '
        ' Output all the groups, with the active and dateexpires from those joined
        '
        'F.Add AdminUI.EditTableOpen
        SectionName = ""
        GroupCount = 0
        CanSeeHiddenGroups = Main.IsDeveloper
        Do While Main.IsCSOK(CS)
            GroupName = Main.GetCS(CS, "GroupName")
            If (Mid(GroupName, 1, 1) <> "_") Or CanSeeHiddenGroups Then
                GroupCaption = Main.GetCS(CS, "GroupCaption")
                GroupID = Main.GetCSInteger(CS, "ID")
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
                            If DateExpires(MembershipPointer) <> 0 Then
                                DateExpireValue = kmaEncodeText(DateExpires(MembershipPointer))
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
                f.Add Main.GetFormInputHidden("Memberrules." & GroupCount & ".ID", GroupID)
                GroupCaption = Replace(GroupCaption, " ", "&nbsp;")
                f.Add vbCrLf & "<!-- GroupRule Row -->" _
                    & "<TR>" _
                    & "<TD style=""TEXT-ALIGN:left;PADDING-LEFT:20px;border-top:1px solid white;"">" & Main.GetFormInputCheckBox("MemberRules." & GroupCount, GroupActive) & GroupCaption & "</TD>" _
                    & "<TD style=""TEXT-ALIGN:left;PADDING-LEFT:10px;border-top:1px solid white;"">Expires " & Main.GetFormInputText("MemberRules." & GroupCount & ".DateExpires", DateExpireValue, 1, 20) & "</TD>" _
                    & "</TR>"
                GroupCount = GroupCount + 1
            End If
            Main.NextCSRecord (CS)
        Loop
        Main.CloseCS (CS)
    End If
    If DetailMemberID = 0 Then
        f.Add "<TR>" _
            & "<TD valign=middle align=right>" & SpanClassAdminSmall & "Groups</span></TD>" _
            & "<TD>" & SpanClassAdminNormal & "Groups will be available after this record is saved</SPAN></TD>" _
            & "<TR>"
    ElseIf GroupCount = 0 Then
        f.Add "<TR>" _
            & "<TD valign=middle align=right>" & SpanClassAdminSmall & "Groups</span></TD>" _
            & "<TD>" & SpanClassAdminNormal & "There are currently no groups defined</SPAN></TD>" _
            & "<TR>"
    Else
        f.Add "<input type=""hidden"" name=""MemberRules.RowCount"" value=""" & GroupCount & """>"
    End If
    f.Add vbCrLf & "<!-- GroupRule Table End --></table>"
    GetFormDetail_TabGroup = "<div STYLE=""width:100%;"" class=""cmBody ccPanel3DReverse"">" & f.Text & "</div>"
    Exit Function
    '
ErrorTrap:
    Call HandleError("CMngrClass", "GetFormDetail_TabGroup", Err.Number, Err.Source, Err.Description, True, False)
    End Function
'
'========================================================================
' Read and save a GetForm_InputCheckList
'   see GetForm_InputCheckList for an explaination of the input
'========================================================================
'
Private Sub SaveMemberRules(CSMember)
    On Error GoTo ErrorTrap
    '
    Dim GroupCount As Long
    Dim GroupPointer As Long
    Dim CSPointer As Long
    Dim MethodName As String
    Dim GroupID As Long
    Dim RuleNeeded As Boolean
    Dim CSRule As Long
    Dim DateExpires As Date
    Dim DateExpiresVariant As Variant
    Dim RuleActive As Boolean
    Dim RuleDateExpires As Date
    Dim MemberRuleID As Long
    Dim PeopleID As Long
    '
    MethodName = "SaveMemberRules"
    '
    ' --- create MemberRule records for all selected
    '
    PeopleID = Main.GetCSInteger(CSMember, "ID")
    GroupCount = Main.GetStreamNumber("MemberRules" & ".RowCount")
    If GroupCount > 0 Then
        For GroupPointer = 0 To GroupCount - 1
            '
            ' ----- Read Response
            '
            GroupID = Main.GetStreamInteger("MemberRules" & "." & GroupPointer & "." & "ID")
            RuleNeeded = Main.GetStreamBoolean("MemberRules" & "." & GroupPointer)
            DateExpires = Main.GetStreamDate("MemberRules" & "." & GroupPointer & "." & "DateExpires")
            If DateExpires = CDate(0) Then
                DateExpiresVariant = Null
            Else
                DateExpiresVariant = DateExpires
                End If
            '
            ' ----- Update Record
            '
            CSRule = Main.OpenCSContent_Internal("Member Rules", "(MemberID=" & PeopleID & ")and(GroupID=" & GroupID & ")", , False, , , "Active,MemberID,GroupID,DateExpires")
            If Not Main.IsCSOK(CSRule) Then
                '
                ' No record exists
                '
                If RuleNeeded Then
                    '
                    ' No record, Rule needed, add it
                    '
                    Call Main.CloseCS(CSRule)
                    CSRule = Main.InsertCSRecord("Member Rules")
                    If Main.IsCSOK(CSRule) Then
                        Call Main.SetCS(CSRule, "Active", True)
                        Call Main.SetCS(CSRule, "MemberID", PeopleID)
                        Call Main.SetCS(CSRule, "GroupID", GroupID)
                        Call Main.SetCS(CSRule, "DateExpires", DateExpires)
                        End If
                    Call Main.CloseCS(CSRule)
                Else
                    '
                    ' No record, no Rule needed, ignore it
                    '
                    Call Main.CloseCS(CSRule)
                    End If
            Else
                '
                ' Record exists
                '
                If RuleNeeded Then
                    '
                    ' record exists, and it is needed, update the DateExpires if changed
                    '
                    RuleActive = Main.GetCSBoolean(CSRule, "active")
                    RuleDateExpires = Main.GetCSDate(CSRule, "DateExpires")
                    If (Not RuleActive) Or (RuleDateExpires <> DateExpires) Then
                        Call Main.SetCS(CSRule, "Active", True)
                        Call Main.SetCS(CSRule, "DateExpires", DateExpires)
                        End If
                    Call Main.CloseCS(CSRule)
                Else
                    '
                    ' record exists and it is not needed, delete it
                    '
                    MemberRuleID = Main.GetCSInteger(CSRule, "ID")
                    Call Main.CloseCS(CSRule)
                    Call Main.DeleteTableRecord("Default", Main.SQLTableMemberRules, MemberRuleID)
                    End If
                End If
            Next
        End If
    Exit Sub
    '
    ' ----- Error Trap
    '
ErrorTrap:
    Call HandleError("CMngrClass", "SaveMemberRules", Err.Number, Err.Source, Err.Description, True, False)
    End Sub
'
'========================================================================
' Print the Member Edit form
'
'   Content must conform to ccMember fields
'========================================================================
'
Private Function GetFormDetail_TabHomes(CSMember As Long) As String
    On Error GoTo ErrorTrap
    '
    Dim Cell As String
    Dim DetailMemberID As Long
    Dim RowClassOdd As String
    Dim RowClassEven As String
    Dim ColumnCount As Long
    Dim CS As Long
    Dim SQL As String
    Dim SQLOrderBy As String
    Dim DateRangeID As Long
    Dim AllowNonEmail As Boolean
    Dim RQS As String
    Dim PageSize As Long
    Dim PageNumber As Long
    Dim TopCount As Long
    Dim RowCnt As Long
    Dim DataRowCount As Long
    Dim PreTableCopy As String
    Dim PostTableCopy As String
    Dim ColumnPtr As Long
    Dim ColCaption() As String
    Dim ColAlign() As String
    Dim ColWidth() As String
    Dim Cells() As String
    Dim Tab0 As New fastStringClass
    Dim AdminUI As Object
    Dim Stream As String
    Dim SQLWhere As String
    Dim SQLFrom As String
    Dim SortField As String
    Dim SortDirection As Long
    Dim RecordMemberID As Long
    Dim RecordMemberName As String
    Dim Propertyid As Long
    Dim Copy As String
    Dim RowClass As String
    Dim PeopleCID As Long
    Dim RecordID As Long
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
    Set AdminUI = CreateObject("ccWeb3.AdminUIClass")
    If Not IsGrgOK Then
        '
        ' The RealEstate datasource is not available
        '
        Stream = "<P>This site is not configured to display the Homes Viewed tab. This tab requires the grgRealEstate Datasource.</P>"
    Else
        SortField = Main.GetStreamText("SortField")
        SortDirection = Main.GetStreamInteger("SortDirection")
        If SortField = "" Then
            SortField = "DateAdded"
            SortDirection = -1
        End If
        '
        RQS = Main.RefreshQueryString
        PageSize = Main.GetStreamInteger(RequestNamePageSize)
        If PageSize = 0 Then
            PageSize = 50
        End If
        PageNumber = Main.GetStreamInteger(RequestNamePageNumber)
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
            '& " left join Properties P on P.ID=L.PropertyID"
        SQLWhere = " Where" _
            & "(S.ID is not null)" _
            & "and(L.MemberID=" & Main.GetCSInteger(CSMember, "ID") & ")" _
            & ""
        If Not AllowNonEmail Then
            SQLWhere = SQLWhere & "and(L.IsEmailSearch is not null)and(L.IsEmailSearch<>0)"
        End If
        Select Case DateRangeID
            Case DateRangeToday
                SQLWhere = SQLWhere & "and(L.DateAdded>=" & Main.EncodeSQLDate(Int(Now())) & ")"
            Case DateRangeYesterday
                SQLWhere = SQLWhere & "and(L.DateAdded<" & Main.EncodeSQLDate(Int(Now())) & ")and(L.DateAdded>=" & Main.EncodeSQLDate(Int(Now()) - 1) & ")"
            Case DateRangePastWeek
                SQLWhere = SQLWhere & "and(L.DateAdded>" & Main.EncodeSQLDate(Int(Now() - 7)) & ")"
            Case DateRangePastMonth
                SQLWhere = SQLWhere & "and(L.DateAdded>" & Main.EncodeSQLDate(Int(Now() - 30)) & ")"
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
        CS = Main.OpenCSSQL("Default", SQL)
        If Main.IsCSOK(CS) Then
            DataRowCount = Main.GetCSInteger(CS, "cnt")
        End If
        Call Main.CloseCS(CS)
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
        Main.TestPoint ("Get Data, SQL=" & SQL)
        CS = Main.OpenCSSQL("Default", SQL, PageSize, PageNumber)
        If Not Main.IsCSOK(CS) Then
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
            PeopleCID = Main.GetContentID("People")
            If Not Main.IsCSOK(CS) Then
                '
                ' EMPTY
                '
            Else
                Do While Main.IsCSOK(CS) And (RowCnt < PageSize)
                    RecordID = Main.GetCSInteger(CS, "ID")
                    DateCompleted = Main.GetCSDate(CS, "DateCompleted")
                    Cells(RowCnt, 0) = Main.GetCSDate(CS, "DateAdded") & Main.GetFormInputHidden("RowID" & RowCnt, RecordID)
                    Cells(RowCnt, 1) = Main.GetCStext(CS, "name")
                    Cells(RowCnt, 2) = Main.GetYesNo(Main.GetCSBoolean(CS, "ViewingIsEmailSearch"))
                    Propertyid = Main.GetCSInteger(CS, "PropertyID")
                    Copy = Main.GetCStext(CS, "PropertyName")
                    If Copy = "" Then
                        Copy = "&nbsp;"
                    Else
                        Copy = "<a href=/index.asp?grg_pid=" & Propertyid & " target=_blank>" & Copy & "</a>"
                    End If
                    Cells(RowCnt, 4) = Copy
                    RowCnt = RowCnt + 1
                    Call Main.NextCSRecord(CS)
                Loop
            End If
            Stream = Stream & Main.GetFormInputHidden("RowCount", RowCnt)
        End If
        Call Main.CloseCS(CS)
        '
        DetailMemberID = Main.GetCSInteger(CSMember, "ID")
        Stream = Main.GetReport(RowCnt, ColCaption(), ColAlign(), ColWidth(), Cells(), PageSize, PageNumber, PreTableCopy, PostTableCopy, DataRowCount, "")
    End If
    
    GetFormDetail_TabHomes = "<div STYLE=""width:100%;"" class=""cmBody ccPanel3DReverse"">" & Stream & "</div>"
    Exit Function
    '
ErrorTrap:
    Call HandleError("CMngrClass", "GetFormDetail_TabHomes", Err.Number, Err.Source, Err.Description, True, False)
    End Function
'
'
'
Function GetFormInputRadioBox(ElementName, ElementValue, CurrentValue, ElementID)
    GetFormInputRadioBox = Main.GetFormInputRadioBox(ElementName, ElementValue, CurrentValue, ElementID)
End Function
'
'
'
Function GetFormInputText(ElementName, CurrentValue, Height, Width, ElementID, OnFocus)
    GetFormInputText = Main.GetFormInputText(ElementName, CurrentValue, Height, Width, ElementID)
    If OnFocus <> "" Then
        GetFormInputText = Replace(GetFormInputText, ">", "onFocus=""" & OnFocus & """>", 1, 1, vbBinaryCompare)
    End If
End Function
'
'
'
Function GetFormInputDate(ElementName, CurrentValue, Width, ElementID, OnFocus)
    GetFormInputDate = Main.GetFormInputDate(ElementName, CurrentValue, Width, ElementID)
    If OnFocus <> "" Then
        GetFormInputDate = Replace(GetFormInputDate, ">", "onFocus=""" & OnFocus & """>", 1, 1, vbBinaryCompare)
    End If
End Function

