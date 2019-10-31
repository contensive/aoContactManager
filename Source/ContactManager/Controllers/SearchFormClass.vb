
Option Strict On
Option Explicit On

Imports Contensive.BaseClasses
Imports Contensive.Addons.ContactManagerTools.Controllers.GenericController
Imports System.Text
Imports System.Linq

Namespace Views
    Public Class SearchFormClass
        '=================================================================================
        '
        Public Shared Function getResponse(cp As CPBaseClass, IsAdminPath As Boolean) As String
            Dim result As String = ""
            Try
                Dim SubTab As Integer
                Dim Nav As New ContactManagerTools.TabController()
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
                    Call Nav.addEntry("Record&nbsp;Fields", getResponse_TabPeople(cp), "ccAdminTab")
                    Call Nav.addEntry("Groups", getResponse_TabGroup(cp), "ccAdminTab")
                    '
                    Content = "" _
                            & cp.Html.Hidden("SelectionForm", "1") _
                            & Nav.getTabs(cp) _
                            & cp.Html.Hidden(RequestNameFormID, Convert.ToInt32(FormIdEnum.FormSearch).ToString()) _
                            & ""
                    result = ContactManagerTools.AdminUIController.getBody(cp, "Contact Manager &gt;&gt; Selection Criteria", ButtonList, "", True, True, Header, "", 0, Content)
                End If
            Catch ex As Exception
                cp.Site.ErrorReport(ex)
            End Try
            Return result
        End Function
        '
        '=================================================================================
        '
        Public Shared Function getResponse_TabGroup(cp As CPBaseClass) As String
            Dim result As String = ""
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
                    result = result _
                            & "<div>Select groups to narrow your results. If any groups are selected, your search will be limited to people in any of the selected groups.</div>" _
                            & "<div>&nbsp;</div>"
                    '
                    ' Add headers to stream
                    '
                    result = result & "<table border=0 width=100% cellspacing=0 cellpadding=4>"
                    result = result & "<tr>"
                    result = result & "<TD width=30><img src=/cclib/images/spacer.gif width=20 height=1></TD>"
                    result = result & "<TD width=30><img src=/cclib/images/spacer.gif width=20 height=1></TD>"
                    result = result & "<TD width=99%><img src=/cclib/images/spacer.gif width=1 height=1></TD>"
                    result = result & "</tr>"
                    '
                    result = result & "<tr>"
                    result = result & "<TD width=30 align=center class=ccAdminListCaption>Select</TD>"
                    result = result & "<TD width=30 align=center class=ccAdminListCaption>Count</TD>"
                    result = result & "<TD width=99% align=left class=ccAdminListCaption>Group Name</TD>"
                    result = result & "</tr>"
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
                        result = result & "<TR>"
                        result = result & "<TD width=30 align=center class=""" & Style & """>" & cp.Html.CheckBox(GroupLabel, GroupChecked) & cp.Html.Hidden(GroupLabel & ".id", GroupID.ToString()) & "</TD>"
                        result = result & "<TD width=30 align=right class=""" & Style & """>" & CS.GetInteger("CountOfID") & "</TD>"
                        result = result & "<TD width=99% align=left class=""" & Style & """>" & GroupName & "</TD>"
                        result = result & "</TR>"
                        GroupPointer = GroupPointer + 1
                        Call CS.GoNext()
                    Loop
                    Call CS.Close()
                    result = result & "</Table>"
                    '
                    result = result & cp.Html.Hidden("GroupCount", GroupPointer.ToString())
                    result = result & cp.Html.Hidden("SelectionGroupSubTab", "1")
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
        Public Shared Function getResponse_TabPeople(cp As CPBaseClass) As String
            Dim result As String = ""
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
                    Dim fieldList As New List(Of FieldMeta)
                    FieldPtr = 0
                    Do While CS.OK()
                        Dim fieldType As Integer = CS.GetInteger("Type")
                        Dim field = New FieldMeta() With {
                             .currentValue = "",
                             .FieldCaption = CS.GetText("Caption"),
                             .fieldEditTab = CS.GetText("editTab"),
                             .fieldId = CS.GetInteger("ID"),
                             .FieldLookupContentName = If((fieldType <> 7), "", cp.Content.GetRecordName("content", CS.GetInteger("LookupContentID"))),
                             .FieldLookupList = If((fieldType <> 7), "", CS.GetText("LookupList")),
                             .FieldName = CS.GetText("name"),
                             .FieldOperator = 0,
                             .fieldType = fieldType
                            }
                        fieldList.Add(field)
                        '
                        ' set prepoplate value from visit property
                        '
                        If CriteriaCount > 0 Then
                            For CriteriaPointer = 0 To CriteriaCount - 1
                                If InStr(1, CriteriaValues(CriteriaPointer), field.FieldName & "=", vbTextCompare) = 1 Then
                                    NameValues = Split(CriteriaValues(CriteriaPointer), "=")
                                    field.currentValue = NameValues(1)
                                    field.FieldOperator = 1
                                ElseIf InStr(1, CriteriaValues(CriteriaPointer), field.FieldName & ">", vbTextCompare) = 1 Then
                                    NameValues = Split(CriteriaValues(CriteriaPointer), ">")
                                    field.currentValue = NameValues(1)
                                    field.FieldOperator = 2
                                ElseIf InStr(1, CriteriaValues(CriteriaPointer), field.FieldName & "<", vbTextCompare) = 1 Then
                                    NameValues = Split(CriteriaValues(CriteriaPointer), "<")
                                    field.currentValue = NameValues(1)
                                    field.FieldOperator = 3
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
                    result = result _
                            & "<div>Enter criteria for each field to identify and select your results. The results of a search will have to have all of the criteria you enter.</div>" _
                            & "<div>&nbsp;</div>"
                    '
                    ' Add headers to stream
                    '
                    result = result & "<table border=0 width=100% cellspacing=0 cellpadding=4>"
                    result = result & "<tr>"
                    result = result & ContactManagerTools.AdminUIController.kmaStartTableCell("120", 1, RowEven, "right") & "<img src=/cclib/images/spacer.gif width=120 height=1></TD>"
                    result = result & ContactManagerTools.AdminUIController.kmaStartTableCell("99%", 1, RowEven, "left") & "<img src=/cclib/images/spacer.gif width=1 height=1></TD>"
                    result = result & "</tr>"
                    '
                    RowPointer = 0
                    lastEditTab = "-1"
                    currentEditTab = ""
                    groupTab = ""
                    For Each field In fieldList
                        currentEditTab = field.fieldEditTab
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
                        Select Case field.fieldType
                            Case FieldTypeDate
                                '
                                ' Date
                                '
                                result = result & groupTab _
                                & "<TR>" _
                                & "<td class=""cmRowCaption"">" & field.FieldCaption & "</TD>" _
                                & "<TD class=""cmRowField"">" _
                                & "<table border=0 width=100% cellspacing=0 cellpadding=0><TR>" _
                                    & "<TD width=10 align=right>" & getFormInputRadioBox(cp, field.FieldName & "_D", "0", field.FieldOperator.ToString(), "") & "</TD><TD align=left width=100>ignore</TD>" _
                                    & "<TD width=10>&nbsp;&nbsp;</TD><TD width=10 align=right>" & getFormInputRadioBox(cp, field.FieldName & "_D", "1", field.FieldOperator.ToString(), "") & "</TD><TD align=left width=100>=</TD>" _
                                    & "<TD width=10>&nbsp;&nbsp;</TD><TD width=10 align=right>" & getFormInputRadioBox(cp, field.FieldName & "_D", "2", field.FieldOperator.ToString(), "") & "</TD><TD align=left width=100>&gt;</TD>" _
                                    & "<TD width=10>&nbsp;&nbsp;</TD><TD width=10 align=right>" & getFormInputRadioBox(cp, field.FieldName & "_D", "3", field.FieldOperator.ToString(), "") & "</TD><TD align=left width=100>&lt;</TD>" _
                                    & "<TD width=10>&nbsp;&nbsp;</TD><TD align=left width=99%>" & getFormInputDate(cp, field.FieldName, field.currentValue, "10", "", "") & "</TD>" _
                                & "</TR></Table>" _
                                & "</TD>" _
                                & "</TR>"
                                groupTab = ""
                                RowPointer = RowPointer + 1
                            Case FieldTypeCurrency, FieldTypeFloat, FieldTypeInteger
                                '
                                ' Numeric
                                '
                                result = result & groupTab _
                                & "<TR>" _
                                & "<td class=""cmRowCaption"">" & field.FieldCaption & "</TD>" _
                                & "<TD class=""cmRowField"">" _
                                & "<table border=0 width=100% cellspacing=0 cellpadding=0><TR>" _
                                    & "<TD width=10 align=right>" & getFormInputRadioBox(cp, field.FieldName & "_N", "0", field.FieldOperator.ToString(), "") & "</TD><TD align=left width=100>ignore</TD>" _
                                    & "<TD width=10>&nbsp;&nbsp;</TD><TD width=10 align=right>" & getFormInputRadioBox(cp, field.FieldName & "_N", "1", field.FieldOperator.ToString(), "") & "</TD><TD align=left width=100>=</TD>" _
                                    & "<TD width=10>&nbsp;&nbsp;</TD><TD width=10 align=right>" & getFormInputRadioBox(cp, field.FieldName & "_N", "2", field.FieldOperator.ToString(), "") & "</TD><TD align=left width=100>&gt;</TD>" _
                                    & "<TD width=10>&nbsp;&nbsp;</TD><TD width=10 align=right>" & getFormInputRadioBox(cp, field.FieldName & "_N", "3", field.FieldOperator.ToString(), "") & "</TD><TD align=left width=100>&lt;</TD>" _
                                    & "<TD width=10>&nbsp;&nbsp;</TD><TD align=left width=99%>" & getFormInputText(cp, field.FieldName, field.currentValue, "1", "5", "", "") & "</TD>" _
                                & "</TR></Table>" _
                                & "</TD>" _
                                & "</TR>"
                                groupTab = ""
                                RowPointer = RowPointer + 1
                            Case FieldTypeBoolean
                                '
                                ' Boolean
                                '
                                result = result & groupTab _
                                        & "<TR>" _
                                        & "<td class=""cmRowCaption"">" & field.FieldCaption & "</TD>" _
                                        & "<TD class=""cmRowField"">" _
                                        & "<table border=0 width=100% cellspacing=0 cellpadding=0><TR>" _
                                        & "<TD width=10 align=right>" & getFormInputRadioBox(cp, field.FieldName, "", field.currentValue, "") & "</TD><TD align=left width=100>" & "  ignore</TD>" _
                                        & "<TD width=10>&nbsp;&nbsp;</TD><TD width=10 align=right>" & getFormInputRadioBox(cp, field.FieldName, "1", field.currentValue, "") & "</TD><TD align=left width=100>" & "true</TD>" _
                                        & "<TD width=10>&nbsp;&nbsp;</TD><TD width=10 align=right>" & getFormInputRadioBox(cp, field.FieldName, "0", field.currentValue, "") & "</TD><TD align=left width=100>" & "false</TD>" _
                                        & "<TD width=99%>&nbsp;</td>" _
                                        & "</TR></Table>" _
                                        & "</TD>" _
                                        & "</TR>"
                                '& "[" & cp.Html.RadioBox(field.fieldname, "", FieldValue(FieldPtr)) & "ignore]" _
                                '& "[" & cp.Html.RadioBox(field.fieldname, "1", FieldValue(FieldPtr)) & "true]" _
                                '& "[" & cp.Html.RadioBox(field.fieldname, "0", FieldValue(FieldPtr)) & "false]" _
                                RowPointer = RowPointer + 1
                                groupTab = ""
                            Case FieldTypeText, FieldTypeLongText
                                '
                                ' Text
                                '
                                result = result & groupTab _
                                        & "<TR>" _
                                        & "<td class=""cmRowCaption"">" & field.FieldCaption & "</TD>" _
                                        & "<TD class=""cmRowField"" valign=absmiddle>" _
                                        & "<table border=0 width=100% cellspacing=0 cellpadding=0><TR>" _
                                        & "<TD width=10 align=right>" & getFormInputRadioBox(cp, field.FieldName & "_T", "", field.currentValue, "") & "</TD><TD align=left width=100>" & "  ignore</TD>" _
                                        & "<TD width=10>&nbsp;&nbsp;</TD><TD width=10 align=right>" & getFormInputRadioBox(cp, field.FieldName & "_T", "1", field.currentValue, "") & "</TD><TD align=left width=100>" & "empty</TD>" _
                                        & "<TD width=10>&nbsp;&nbsp;</TD><TD width=10 align=right>" & getFormInputRadioBox(cp, field.FieldName & "_T", "2", field.currentValue, "") & "</TD><TD align=left width=100>" & "not&nbsp;empty</TD>" _
                                        & "<TD width=10>&nbsp;&nbsp;</TD><TD width=10 align=right>" & getFormInputRadioBox(cp, field.FieldName & "_T", "3", field.currentValue, field.FieldName & "_r") & "</TD><TD align=center width=100>" & "&nbsp;includes&nbsp;" & "</TD><TD align=left width=99%>" & getFormInputText(cp, field.FieldName, field.currentValue, "1", "20", "", "cmTextInclude") & "</TD>" _
                                        & "</TR></Table>" _
                                        & "</TD>" _
                                        & "</TR>"
                                groupTab = ""
                                RowPointer = RowPointer + 1
                            Case FieldTypeLookup
                                '
                                ' Lookup
                                '
                                result = result & groupTab _
                                        & "<TR><td class=""cmRowCaption"">" & field.FieldCaption & "</TD>" _
                                        & ""
                                If field.FieldLookupContentName <> "" Then
                                    result = result _
                                            & "<TD class=""cmRowField"">" _
                                            & cp.Html.SelectContent(field.FieldName, field.currentValue, field.FieldLookupContentName, "", "Any") & "</TD>"
                                Else
                                    result = result _
                                            & "<TD class=""cmRowField"">" _
                                            & cp.Html.SelectList(field.FieldName, field.currentValue, field.FieldLookupList, "", "Any") & "</TD>"
                                End If
                                result = result & "</TR>"
                                groupTab = ""
                                RowPointer = RowPointer + 1
                        End Select
                    Next
                    result = result & "</Table>"
                    '
                    result = result & cp.Html.Hidden("SelectionSearchSubTab", "1")
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
        Public Shared Function ProcessRequest(cp As CPBaseClass, ae As Controllers.ApplicationController, request As Views.CMngrClass.RequestClass) As FormIdEnum
            Dim result As FormIdEnum = FormIdEnum.FormList
            Try
                Select Case request.Button
                    Case ButtonSearch
                        Dim TextOption As String
                        Dim NumericOption As String
                        Dim CS As CPCSBaseClass = cp.CSNew()
                        Dim GroupCount As Integer
                        Dim GroupPointer As Integer
                        Dim GroupLabel As String
                        Dim ContactGroupCriteria As String = ""
                        Dim ContactSearchCriteria As String = ""
                        Dim Criteria As String
                        Dim FieldName As String
                        Dim FieldValue As String
                        Dim fieldType As Integer
                        '
                        If (Not String.IsNullOrEmpty(request.SelectionGroupSubTab)) Then
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
                            ae.userProperties.contactGroupCriteria = ContactGroupCriteria & ","
                        End If
                        If (Not String.IsNullOrEmpty(request.SelectionSearchSubTab)) Then
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
                End Select
            Catch ex As Exception
                cp.Site.ErrorReport(ex)
            End Try
            Return result
        End Function
        ''
        ''========================================================================
        ''
        'Public Shared Function GetFormInputRadioBox(cp As CPBaseClass, ElementName As String, ElementValue As String, CurrentValue As String, htmlId As String, htmlClass As String) As String
        '    Return cp.Html.RadioBox(ElementName, ElementValue, CurrentValue, htmlClass, htmlId)
        'End Function
        '
        Public Shared Function getFormInputRadioBox(cp As CPBaseClass, ElementName As String, ElementValue As String, CurrentValue As String, htmlId As String) As String
            Return cp.Html.RadioBox(ElementName, ElementValue, CurrentValue, "", htmlId)
        End Function
        '
        '========================================================================
        '
        Public Shared Function getFormInputText(cp As CPBaseClass, htmlName As String, CurrentValue As String, Height As String, Width As String, htmlId As String, htmlClass As String) As String
            Return cp.Html.InputText(htmlName, CurrentValue, Height, Width, False, htmlClass, htmlId)
        End Function
        '
        '========================================================================
        '
        Public Shared Function getFormInputDate(cp As CPBaseClass, ElementName As String, CurrentValue As String, Width As String, ElementID As String, OnFocus As String) As String
            Return cp.Html.InputDate(ElementName, CurrentValue, Width, "", ElementID)
        End Function
        '
    End Class
End Namespace
