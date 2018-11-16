Option Explicit On
Option Strict On

Imports Contensive.BaseClasses

'Imports System.Collections.Generic

Namespace Controllers
    Public Class AdminUIController
        Private Sub New()
        End Sub
        '
        Public Shared Function GetPanel(content As String, stylePanel As String, stylePanelShadow As String, stylePanelHilite As String, widthPercent As String, something As Integer) As String
            Throw New NotImplementedException()
        End Function
        '
        '====================================================================================================
        Public Shared Function GetFormBodyAdminOnly() As String
            Throw New NotImplementedException()
        End Function
        '
        '====================================================================================================
        Public Shared Function GetBody(cp As CPBaseClass, HeaderMaybe As String, buttonMaybe As String, dontKNow As String, dontKnow2 As Boolean, dontKnow3 As Boolean, body As String, dontknow4 As String, dontknow5 As Integer, dontknow6 As String) As String
            Throw New NotImplementedException()
        End Function

        Public Shared Function kmaStartTableCell(widthPercent As String, dontKNow As Integer, isRowEvent As Boolean, align As String) As String
            '            kmaStartTableCell("60", 1, RowEven, "center")
            Throw New NotImplementedException()
        End Function
        '
        Public Shared Function GetReportSortColumnPtr(cp As CPBaseClass, SortColPtr As Integer) As Integer
            Throw New NotImplementedException()
        End Function
        '
        Public Shared Function GetReportSortType(cp As CPBaseClass) As Integer
            Throw New NotImplementedException()
        End Function
        '
        Public Shared Function GetReport2(cp As CPBaseClass, RowPointer As Integer, ColCaption() As String, ColAlign() As String, ColWidth() As String, Cells(,) As String, PageSize As Integer, PageNumber As Integer, PreTableCopy As String, PostTableCopy As String, DataRowCount As Integer, style As String, ColSortable() As Boolean, SortColPtr As Integer) As String
            Throw New NotImplementedException()
        End Function
        '
        Public Shared Function GetReport(cp As CPBaseClass, RowCnt As Integer, ColCaption() As String, ColAlign() As String, ColWidth() As String, Cells(,) As String, PageSize As Integer, PageNumber As Integer, PreTableCopy As String, PostTableCopy As String, DataRowCount As Integer, style As String) As String
            Throw New NotImplementedException()
        End Function

    End Class
End Namespace

