
Imports Contensive.BaseClasses

Namespace Views
    Public Class ContactManagerAddon
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
                    Dim request As New RequestModel(cp)
                    Call cp.Doc.AddRefreshQueryString("tab", request.TabNumber.ToString())
                    '
                    request.FormID = If(request.FormID <> FormIdEnum.FormUnknown, request.FormID, getDefaultFormId(request))
                    If (Not String.IsNullOrEmpty(request.Button)) Then
                        '
                        ' ----- Process Previous Forms
                        '
                        Select Case request.FormID
                            Case FormIdEnum.FormSearch
                                '
                                request.FormID = SearchView.processRequest(cp, ae, request)
                            Case FormIdEnum.FormList
                                '
                                request.FormID = ListView.processRequest(cp, ae, request)
                            Case FormIdEnum.FormDetails
                                ' 
                                request.FormID = DetailView.processRequest(cp, request)
                        End Select
                    End If
                    Dim IsAdminPath As Boolean = True
                    '
                    ' ----- Output the next form
                    Select Case request.FormID
                        Case FormIdEnum.FormDetails
                            '
                            result &= DetailView.getResponse(cp, ae, request.DetailMemberID)
                        Case FormIdEnum.FormList
                            '
                            result &= ListView.getResponse(cp, ae)
                        Case Else
                            '
                            result &= SearchView.getResponse(cp, ae, IsAdminPath)
                    End Select
                    'result = "<div class=ccbodyadmin>" & result & "</div>"
                    '
                    ' wrapper for style strength
                    result = "<div class=""contactManager"">" & result & "</div>"
                End Using
            Catch ex As Exception
                cp.Site.ErrorReport(ex)
                Throw
            End Try
            Return result
        End Function
        '
        '========================================================================
        ''' <summary>
        ''' return the default formId
        ''' </summary>
        ''' <param name="request"></param>
        ''' <returns></returns>
        Private Function getDefaultFormId(request As RequestModel) As FormIdEnum
            If (request.DetailMemberID <> 0) Then Return FormIdEnum.FormDetails
            Return FormIdEnum.FormList
            'If ((ae.userProperties.contactSearchCriteria <> "") Or (ae.userProperties.contactGroupCriteria <> "")) Then Return FormIdEnum.FormList
            'Return FormIdEnum.FormSearch
        End Function
    End Class
End Namespace
