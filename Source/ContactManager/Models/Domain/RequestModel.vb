
'
'========================================================================
Imports Contensive.BaseClasses
''' <summary>
''' requests processed by this form
''' </summary>
Public Class RequestModel
    Public TabNumber As Integer
    Public Button As String
    Public DetailMemberID As Integer
    Public FormID As FormIdEnum
    Public RowCount As Integer
    Public GroupID As Integer
    Public GroupToolAction As GroupToolActionEnum
    Public GroupToolSelect As Integer
    Public SelectionGroupSubTab As String
    Public SelectionSearchSubTab As String
    '
    Public Sub New(cp As CPBaseClass)
        '
        ' todo - convert request names to constants
        '
        TabNumber = cp.Doc.GetInteger("tab")
        Button = cp.Doc.GetText("Button")
        DetailMemberID = cp.Doc.GetInteger(RequestNameMemberID)
        RowCount = cp.Doc.GetInteger("M.Count")
        GroupID = cp.Doc.GetInteger("GroupID")
        GroupToolSelect = cp.Doc.GetInteger("GroupToolSelect")
        SelectionGroupSubTab = cp.Doc.GetText("SelectionGroupSubTab")
        SelectionSearchSubTab = cp.Doc.GetText("SelectionSearchSubTab")
        '
        Dim testAction As Integer = cp.Doc.GetInteger("GroupToolAction")
        GroupToolAction = If((testAction > 0 And testAction < 5), CType(testAction, GroupToolActionEnum), GroupToolActionEnum.nop)
        '
        Dim testFormId As Integer = cp.Doc.GetInteger(RequestNameFormID)
        FormID = If((testFormId > 0 And testFormId < 4), CType(testFormId, FormIdEnum), FormIdEnum.FormUnknown)

    End Sub
End Class