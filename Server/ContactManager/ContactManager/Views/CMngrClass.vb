
Option Strict On
Option Explicit On

Imports Contensive.BaseClasses
Imports Contensive.Addons.ContactManagerTools.GenericController
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
        ''' <param name="cp"></param>
        ''' <returns></returns>
        Public Overrides Function Execute(ByVal cp As CPBaseClass) As Object
            Dim result As String = ""
            Dim sw As New Stopwatch : sw.Start()
            Try
                '
                ' -- initialize application. If authentication needed and not login page, pass true
                Using ae As New Controllers.ApplicationController(cp, False)
                    Dim request As RequestClass = New RequestClass(cp)
                    Call cp.Doc.AddRefreshQueryString("tab", request.TabNumber.ToString())
                    '
                    request.FormID = If((request.FormID <> FormIdEnum.FormUnknown), request.FormID, getDefaultFormId(ae, request))
                    If (Not String.IsNullOrEmpty(request.Button)) Then
                        '
                        ' ----- Process Previous Forms
                        '
                        Select Case request.FormID
                            Case FormIdEnum.FormSearch
                                '
                                request.FormID = SearchFormClass.ProcessRequest(cp, ae, request)
                            Case FormIdEnum.FormList
                                '
                                request.FormID = ListFormController.ProcessRequest(cp, ae, request)
                            Case FormIdEnum.FormDetails
                                ' 
                                request.FormID = DetailFormController.ProcessRequest(cp, ae, request)
                        End Select
                    End If
                    Dim IsAdminPath As Boolean = True
                    '
                    ' ----- Output the next form
                    Select Case request.FormID
                        Case FormIdEnum.FormDetails
                            '
                            result = result & DetailFormController.getResponse(cp, request.DetailMemberID)
                        Case FormIdEnum.FormList
                            '
                            result = result & ListFormController.getResponse(cp, ae, IsAdminPath)
                        Case Else
                            '
                            result = result & SearchFormClass.getResponse(cp, IsAdminPath)
                    End Select
                    result = "<div class=ccbodyadmin>" & result & "</div>"
                    '
                    ' wrapper for style strength
                    result = "" _
                    & vbCrLf & vbTab & "<div class=""contactManager"">" _
                    & result _
                    & vbCrLf & vbTab & "</div>" _
                    & ""
                End Using
            Catch ex As Exception
                cp.Site.ErrorReport(ex)
            End Try
            Return result
        End Function
        '
        '========================================================================
        ''' <summary>
        ''' return the default formId
        ''' </summary>
        ''' <param name="ae"></param>
        ''' <param name="request"></param>
        ''' <returns></returns>
        Private Function getDefaultFormId(ae As Controllers.ApplicationController, request As RequestClass) As FormIdEnum
            If (request.DetailMemberID <> 0) Then Return FormIdEnum.FormDetails
            Return FormIdEnum.FormList
            'If ((ae.userProperties.contactSearchCriteria <> "") Or (ae.userProperties.contactGroupCriteria <> "")) Then Return FormIdEnum.FormList
            'Return FormIdEnum.FormSearch
        End Function
        '
        '========================================================================
        ''' <summary>
        ''' requests processed by this form
        ''' </summary>
        Public Class RequestClass
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
    End Class
End Namespace
