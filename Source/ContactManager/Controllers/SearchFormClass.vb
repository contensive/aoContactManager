
Option Strict On
Option Explicit On

Imports Contensive.BaseClasses
Imports Contensive.Addons.ContactManagerTools.Controllers.GenericController
Imports System.Text
Imports System.Linq

Namespace Views
    Public Class SearchFormClass
        '
        '=================================================================================
        '
        Public Shared Function getResponse(cp As CPBaseClass, ae As Controllers.ApplicationController, IsAdminPath As Boolean) As String
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
                        SubTab = ae.userProperties.selectSubTab
                        If SubTab = 0 Then
                            SubTab = 1
                            ae.userProperties.selectSubTab = SubTab
                        End If
                    Else
                        ae.userProperties.selectSubTab = SubTab
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
                    Call Nav.addEntry("Record&nbsp;Fields", getResponse_TabPeople(cp, ae), "ccAdminTab")
                    Call Nav.addEntry("Groups", getResponse_TabGroup(cp, ae), "ccAdminTab")
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
        Public Shared Function getResponse_TabGroup(cp As CPBaseClass, ae As Controllers.ApplicationController) As String
            Dim result As String = ""
            Try
                '
                If True Then
                    Dim Button As String = "GroupSelect"
                    Dim RQS As String = cp.Doc.RefreshQueryString
                    Dim ContactGroupCriteria As String = ae.userProperties.contactGroupCriteria
                    '
                    result = result _
                            & "<div>Select groups to narrow your results. If any groups are selected, your search will be limited to people in any of the selected groups.</div>" _
                            & "<div>&nbsp;</div>"
                    '
                    ' Add headers to stream
                    '
                    result &= "<table border=0 width=100% cellspacing=0 cellpadding=4>"
                    result &= "<tr>"
                    result &= "<td width=30><img src=/cclib/images/spacer.gif width=20 height=1></TD>"
                    result &= "<td width=30><img src=/cclib/images/spacer.gif width=20 height=1></TD>"
                    result &= "<td width=99%><img src=/cclib/images/spacer.gif width=1 height=1></TD>"
                    result &= "</tr>"
                    '
                    result &= "<tr>"
                    result &= "<td width=30 align=center class=ccAdminListCaption>Select</TD>"
                    result &= "<td width=30 align=center class=ccAdminListCaption>Count</TD>"
                    result &= "<td width=99% align=left class=ccAdminListCaption>Group Name</TD>"
                    result &= "</tr>"
                    '
                    Dim SQL As String = "SELECT ccGroups.ID as GroupID, ccGroups.Name as GroupName, ccGroups.Caption as GroupCaption, Count(ccMembers.ID) AS CountOfID" _
                            & " FROM (ccGroups LEFT JOIN ccMemberRules ON ccGroups.ID = ccMemberRules.GroupID) LEFT JOIN ccMembers ON ccMemberRules.MemberID = ccMembers.ID" _
                            & " Where (((ccMemberRules.DateExpires) Is Null Or (ccMemberRules.DateExpires) > " & cp.Db.EncodeSQLDate(Now()) & "))" _
                            & " GROUP BY ccGroups.ID, ccGroups.Name, ccGroups.Caption" _
                            & " ORDER BY ccGroups.Caption;"
                    '
                    Using CS As CPCSBaseClass = cp.CSNew()
                        CS.OpenSQL(SQL)
                        Dim GroupPointer As Integer = 0
                        Do While CS.OK()
                            Dim GroupID As Integer = CS.GetInteger("GroupID")
                            Dim GroupName As String = CS.GetText("GroupCaption")
                            If GroupName = "" Then
                                GroupName = CS.GetText("GroupName")
                                If GroupName = "" Then
                                    GroupName = "Group " & GroupID
                                End If
                            End If
                            Dim GroupLabel As String = "Group" & GroupPointer
                            Dim GroupChecked As Boolean = (InStr(1, ContactGroupCriteria, "," & GroupID & ",") <> 0)

                            Dim Style As String
                            If ((GroupPointer Mod 2) = 0) Then
                                Style = "ccAdminListRowEven"
                            Else
                                Style = "ccAdminListRowOdd"
                            End If
                            result &= "<TR>"
                            result &= "<td class=""p-1"" width=30 align=center class=""" & Style & """>" & cp.Html.CheckBox(GroupLabel, GroupChecked) & cp.Html.Hidden(GroupLabel & ".id", GroupID.ToString()) & "</TD>"
                            result &= "<td class=""p-1"" width=30 align=right class=""" & Style & """>" & CS.GetInteger("CountOfID") & "</TD>"
                            result &= "<td class=""p-1"" width=99% align=left class=""" & Style & """>" & GroupName & "</TD>"
                            result &= "</TR>"
                            GroupPointer += 1
                            Call CS.GoNext()
                        Loop
                        Call CS.Close()
                        result &= "</Table>"
                        '
                        result &= cp.Html.Hidden("GroupCount", GroupPointer.ToString())
                    End Using
                    result &= cp.Html.Hidden("SelectionGroupSubTab", "1")
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
        Public Shared Function getResponse_TabPeople(cp As CPBaseClass, ae As Controllers.ApplicationController) As String
            Dim result As String = ""
            Try
                '
                ' prepare visit property to prepopulate form
                Dim ContactSearchCriteria As String = ae.userProperties.contactSearchCriteria
                Dim CriteriaValues() As String = Array.Empty(Of String)
                Dim CriteriaCount As Integer = 0
                If ContactSearchCriteria <> "" Then
                    CriteriaValues = Split(ContactSearchCriteria, vbCrLf)
                    CriteriaCount = UBound(CriteriaValues) + 1
                End If
                '
                ' Setup fields and capture request changes
                '
                Dim Criteria As String = "(active<>0)and(ContentID=" & cp.Content.GetID("people") & ")and(authorable<>0)"
                '
                Using CS As CPCSBaseClass = cp.CSNew()
                    CS.Open("Content Fields", Criteria, "EditTab,EditSortPriority")
                    Dim fieldList As New List(Of FieldMeta)
                    Dim FieldPtr As Integer = 0
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
                            Dim CriteriaPointer As Integer
                            For CriteriaPointer = 0 To CriteriaCount - 1
                                Dim NameValues() As String
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
                        FieldPtr += 1
                        Call CS.GoNext()
                    Loop
                    Call CS.Close()
                    Dim FieldCount As Integer = FieldPtr
                    '
                    ' header
                    '
                    result = result _
                            & "<div>Enter criteria for each field to identify and select your results. The results of a search will have to have all of the criteria you enter.</div>" _
                            & "<div>&nbsp;</div>"
                    '
                    ' Add headers to stream
                    '
                    result &= "<table border=0 width=100% cellspacing=0 cellpadding=4>"
                    result &= "<tr>"
                    Dim RowEven As Boolean
                    result &= ContactManagerTools.AdminUIController.kmaStartTableCell("120", 1, RowEven, "right") & "<img src=/cclib/images/spacer.gif width=120 height=1></TD>"
                    result &= ContactManagerTools.AdminUIController.kmaStartTableCell("99%", 1, RowEven, "left") & "<img src=/cclib/images/spacer.gif width=1 height=1></TD>"
                    result &= "</tr>"
                    '
                    Dim RowPointer As Integer = 0
                    Dim lastEditTab As String = "-1"
                    Dim currentEditTab As String = ""
                    Dim groupTab As String = ""
                    For Each field In fieldList
                        currentEditTab = field.fieldEditTab
                        If currentEditTab <> lastEditTab Then
                            Dim tabCaption As String
                            If currentEditTab = "" Then
                                tabCaption = "Details"
                            Else
                                tabCaption = currentEditTab
                            End If
                            groupTab = "" _
                                & "<tr>" _
                                & "<td class=""p-1 cmRowTab"" colspan=""2"">" & tabCaption & "</TD>" _
                                & "</TR>"
                            RowPointer += 1
                            lastEditTab = currentEditTab
                        End If
                        RowEven = ((RowPointer Mod 2) = 0)
                        Select Case field.fieldType
                            Case FieldTypeDate
                                '
                                ' Date
                                '
                                result &= groupTab _
                                    & "<TR>" _
                                    & "<td class=""p-1 cmRowCaption"">" & field.FieldCaption & "</TD>" _
                                    & "<td class=""cmRowField"">" _
                                    & "<table border=0 width=100% cellspacing=0 cellpadding=0><TR>" _
                                        & "<td class=""p-1"" width=10 align=right>" & getFormInputRadioBox(cp, field.FieldName & "_D", "0", field.FieldOperator.ToString(), "") & "</TD><td class=""p-1"" align=left width=100>ignore</TD>" _
                                        & "<td class=""p-1"" width=10>&nbsp;&nbsp;</TD><td class=""p-1"" width=10 align=right>" & getFormInputRadioBox(cp, field.FieldName & "_D", "1", field.FieldOperator.ToString(), "") & "</TD><td class=""p-1"" align=left width=100>=</TD>" _
                                        & "<td class=""p-1"" width=10>&nbsp;&nbsp;</TD><td class=""p-1"" width=10 align=right>" & getFormInputRadioBox(cp, field.FieldName & "_D", "2", field.FieldOperator.ToString(), "") & "</TD><td class=""p-1"" align=left width=100>&gt;</TD>" _
                                        & "<td class=""p-1"" width=10>&nbsp;&nbsp;</TD><td class=""p-1"" width=10 align=right>" & getFormInputRadioBox(cp, field.FieldName & "_D", "3", field.FieldOperator.ToString(), "") & "</TD><td class=""p-1"" align=left width=100>&lt;</TD>" _
                                        & "<td class=""p-1"" width=10>&nbsp;&nbsp;</TD><td class=""p-1"" align=left width=99%>" & getFormInputDate(cp, field.FieldName, field.currentValue, "10", "", "") & "</TD>" _
                                    & "</TR></Table>" _
                                    & "</TD>" _
                                    & "</TR>"
                                groupTab = ""
                                RowPointer = RowPointer + 1
                            Case FieldTypeCurrency, FieldTypeFloat, FieldTypeInteger
                                '
                                ' Numeric
                                '
                                result &= groupTab _
                                    & "<TR>" _
                                    & "<td class=""p-1 cmRowCaption"">" & field.FieldCaption & "</TD>" _
                                    & "<td class=""p-1 cmRowField"">" _
                                    & "<table border=0 width=100% cellspacing=0 cellpadding=0><TR>" _
                                        & "<td class=""p-1"" width=10 align=right>" & getFormInputRadioBox(cp, field.FieldName & "_N", "0", field.FieldOperator.ToString(), "") & "</TD><td class=""p-1"" align=left width=100>ignore</TD>" _
                                        & "<td class=""p-1"" width=10>&nbsp;&nbsp;</TD><td class=""p-1"" width=10 align=right>" & getFormInputRadioBox(cp, field.FieldName & "_N", "1", field.FieldOperator.ToString(), "") & "</TD><td class=""p-1"" align=left width=100>=</TD>" _
                                        & "<td class=""p-1"" width=10>&nbsp;&nbsp;</TD><td class=""p-1"" width=10 align=right>" & getFormInputRadioBox(cp, field.FieldName & "_N", "2", field.FieldOperator.ToString(), "") & "</TD><td class=""p-1"" align=left width=100>&gt;</TD>" _
                                        & "<td class=""p-1"" width=10>&nbsp;&nbsp;</TD><td class=""p-1"" width=10 align=right>" & getFormInputRadioBox(cp, field.FieldName & "_N", "3", field.FieldOperator.ToString(), "") & "</TD><td class=""p-1"" align=left width=100>&lt;</TD>" _
                                        & "<td class=""p-1"" width=10>&nbsp;&nbsp;</TD><td class=""p-1"" align=left width=99%>" & getFormInputText(cp, field.FieldName, field.currentValue, "1", "5", "", "") & "</TD>" _
                                    & "</TR></Table>" _
                                    & "</TD>" _
                                    & "</TR>"
                                groupTab = ""
                                RowPointer = RowPointer + 1
                            Case FieldTypeBoolean
                                '
                                ' Boolean
                                '
                                Dim currentValue As String = If(String.IsNullOrWhiteSpace(field.currentValue), "0", field.currentValue)
                                result &= groupTab _
                                    & "<TR>" _
                                    & "<td class=""p-1 cmRowCaption"">" & field.FieldCaption & "</TD>" _
                                    & "<td class=""p-1 cmRowField"">" _
                                    & "<table border=0 width=100% cellspacing=0 cellpadding=0><TR>" _
                                    & "<td class=""p-1"" width=10 align=right>" & getFormInputRadioBox(cp, field.FieldName, "0", currentValue, "") & "</TD><td class=""p-1"" align=left width=100>" & "  ignore</TD>" _
                                    & "<td class=""p-1"" width=10>&nbsp;&nbsp;</TD><td class=""p-1"" width=10 align=right>" & getFormInputRadioBox(cp, field.FieldName, "1", currentValue, "") & "</TD><td class=""p-1"" align=left width=100>" & "true</TD>" _
                                    & "<td class=""p-1"" width=10>&nbsp;&nbsp;</TD><td class=""p-1"" width=10 align=right>" & getFormInputRadioBox(cp, field.FieldName, "2", currentValue, "") & "</TD><td class=""p-1"" align=left width=100>" & "false</TD>" _
                                    & "<td class=""p-1"" width=99%>&nbsp;</td>" _
                                    & "</TR></Table>" _
                                    & "</TD>" _
                                    & "</TR>"
                                RowPointer += 1
                                groupTab = ""
                            Case FieldTypeText, FieldTypeLongText
                                '
                                ' Text
                                '
                                result &= groupTab _
                                    & "<TR>" _
                                    & "<td class=""p-1 cmRowCaption"">" & field.FieldCaption & "</TD>" _
                                    & "<td class=""p-1 cmRowField"" valign=absmiddle>" _
                                    & "<table border=0 width=100% cellspacing=0 cellpadding=0><TR>" _
                                    & "<td class=""p-1"" width=10 align=right>" & getFormInputRadioBox(cp, field.FieldName & "_T", "", field.currentValue, "") & "</TD><td class=""p-1"" align=left width=100>" & "  ignore</TD>" _
                                    & "<td class=""p-1"" width=10>&nbsp;&nbsp;</TD><td class=""p-1"" width=10 align=right>" & getFormInputRadioBox(cp, field.FieldName & "_T", "1", field.currentValue, "") & "</TD><td class=""p-1"" align=left width=100>" & "empty</TD>" _
                                    & "<td class=""p-1"" width=10>&nbsp;&nbsp;</TD><td class=""p-1"" width=10 align=right>" & getFormInputRadioBox(cp, field.FieldName & "_T", "2", field.currentValue, "") & "</TD><td class=""p-1"" align=left width=100>" & "not&nbsp;empty</TD>" _
                                    & "<td class=""p-1"" width=10>&nbsp;&nbsp;</TD><td class=""p-1"" width=10 align=right>" & getFormInputRadioBox(cp, field.FieldName & "_T", "3", field.currentValue, field.FieldName & "_r") & "</TD><td class=""p-1"" align=center width=100>" & "&nbsp;includes&nbsp;" & "</TD><td class=""p-1"" align=left width=99%>" & getFormInputText(cp, field.FieldName, field.currentValue, "1", "20", "", "cmTextInclude") & "</TD>" _
                                    & "</TR></Table>" _
                                    & "</TD>" _
                                    & "</TR>"
                                groupTab = ""
                                RowPointer += 1
                            Case FieldTypeLookup
                                '
                                ' Lookup
                                '
                                result &= groupTab _
                                        & "<TR><td class=""p-1 cmRowCaption"">" & field.FieldCaption & "</TD>" _
                                        & ""
                                If field.FieldLookupContentName <> "" Then
                                    result = result _
                                            & "<td class=""p-1 cmRowField"">" _
                                            & cp.Html.SelectContent(field.FieldName, field.currentValue, field.FieldLookupContentName, "", "Any", "form-control select") & "</TD>"
                                Else
                                    result = result _
                                            & "<td class=""p-1 cmRowField"">" _
                                            & cp.Html.SelectList(field.FieldName, field.currentValue, field.FieldLookupList, "Any", "form-control select") & "</TD>"
                                End If
                                result &= "</TR>"
                                groupTab = ""
                                RowPointer = RowPointer + 1
                        End Select
                    Next
                End Using
                result &= "</Table>"
                '
                result &= cp.Html.Hidden("SelectionSearchSubTab", "1")
                '
            Catch ex As Exception
                cp.Site.ErrorReport(ex)
            End Try
            Return result
        End Function
        '
        '=================================================================================
        '
        Public Shared Function processRequest(cp As CPBaseClass, ae As Controllers.ApplicationController, request As Views.CMngrClass.RequestClass) As FormIdEnum
            Dim result As FormIdEnum = FormIdEnum.FormList
            Try
                Select Case request.Button
                    Case ButtonSearch
                        '
                        If (Not String.IsNullOrEmpty(request.SelectionGroupSubTab)) Then
                            '
                            ' Save the Form
                            '
                            Dim GroupCount As Integer = cp.Doc.GetInteger("GroupCount")
                            Dim ContactGroupCriteria As String = ""
                            If GroupCount > 0 Then
                                Dim GroupPointer As Integer
                                For GroupPointer = 0 To GroupCount - 1
                                    Dim GroupLabel As String = "Group" & GroupPointer
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
                            Using csField As CPCSBaseClass = cp.CSNew()
                                Dim Criteria As String = "(active<>0)and(ContentID=" & cp.Content.GetID("people") & ")and(authorable<>0)"
                                csField.Open("Content Fields", Criteria, "EditSortPriority")
                                Dim ContactSearchCriteria As String = ""
                                Do While csField.OK()
                                    Dim FieldName As String = csField.GetText("name")
                                    Dim FieldValue As String = cp.Doc.GetText(FieldName)
                                    Dim fieldType As Integer = csField.GetInteger("Type")
                                    Dim NumericOption As String
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
                                            Dim TextOption As String = cp.Doc.GetText(FieldName & "_T")
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
                                    Call csField.GoNext()
                                Loop
                                Call csField.Close()
                                ae.userProperties.contactSearchCriteria = ContactSearchCriteria
                            End Using
                        End If
                End Select
            Catch ex As Exception
                cp.Site.ErrorReport(ex)
            End Try
            Return result
        End Function
        '
        '=================================================================================
        '
        Public Shared Function getFormInputRadioBox(cp As CPBaseClass, ElementName As String, ElementValue As String, CurrentValue As String, htmlId As String) As String
            Return cp.Html.RadioBox(ElementName, ElementValue, CurrentValue, "", htmlId)
        End Function
        '
        '=================================================================================
        '
        Public Shared Function getFormInputText(cp As CPBaseClass, htmlName As String, CurrentValue As String, Height As String, Width As String, htmlId As String, htmlClass As String) As String
            Return cp.Html.InputText(htmlName, CurrentValue, Height, Width, False, htmlClass & " form-control", htmlId)
        End Function
        '
        '=================================================================================
        '
        Public Shared Function getFormInputDate(cp As CPBaseClass, ElementName As String, CurrentValue As String, Width As String, ElementID As String, OnFocus As String) As String
            Return cp.Html.InputDate(ElementName, CurrentValue, Width, "", ElementID)
        End Function
        '
    End Class
End Namespace
