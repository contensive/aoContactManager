
Option Explicit On
Option Strict On

Imports System
Imports System.Collections.Generic
Imports System.Text
Imports Contensive.BaseClasses

Namespace Models     '<------ set namespace
    Public Class VisitSummaryModel        '<------ set set model Name and everywhere that matches this string
        Inherits baseModel
        Implements ICloneable
        '
        '====================================================================================================
        '-- const
        Public Const contentName As String = "Visit Summary"      '<------ set content name
        Public Const contentTableName As String = "ccVisitSummary"   '<------ set to tablename for the primary content (used for cache names)
        Private Shadows Const contentDataSource As String = "default"             '<------ set to datasource if not default
        '
        '====================================================================================================
        ' -- instance properties

        Public Property AuthenticatedVisits As Integer
        Public Property AveTimeOnSite As Integer
        Public Property BotVisits As Integer
        Public Property DateNumber As Integer
        Public Property MobileVisits As Integer
        Public Property NewVisitorVisits As Integer
        Public Property NoCookieVisits As Integer
        Public Property PagesViewed As Integer
        Public Property SinglePageVisits As Integer
        Public Property TimeDuration As Integer
        Public Property TimeNumber As Integer
        Public Property Visits As Integer

        '
        '====================================================================================================
        Public Overloads Shared Function add(cp As CPBaseClass) As VisitSummaryModel
            Return add(Of VisitSummaryModel)(cp)
        End Function
        '
        '====================================================================================================
        Public Overloads Shared Function create(cp As CPBaseClass, recordId As Integer) As VisitSummaryModel
            Return create(Of VisitSummaryModel)(cp, recordId)
        End Function
        '
        '====================================================================================================
        Public Overloads Shared Function create(cp As CPBaseClass, recordGuid As String) As VisitSummaryModel
            Return create(Of VisitSummaryModel)(cp, recordGuid)
        End Function
        '
        '====================================================================================================
        Public Overloads Shared Function createByName(cp As CPBaseClass, recordName As String) As VisitSummaryModel
            Return createByName(Of VisitSummaryModel)(cp, recordName)
        End Function
        '
        '====================================================================================================
        Public Overloads Sub save(cp As CPBaseClass)
            MyBase.save(Of VisitSummaryModel)(cp)
        End Sub
        '
        '====================================================================================================
        Public Overloads Shared Sub delete(cp As CPBaseClass, recordId As Integer)
            delete(Of VisitSummaryModel)(cp, recordId)
        End Sub
        '
        '====================================================================================================
        Public Overloads Shared Sub delete(cp As CPBaseClass, ccGuid As String)
            delete(Of VisitSummaryModel)(cp, ccGuid)
        End Sub
        '
        '====================================================================================================
        Public Overloads Shared Function createList(cp As CPBaseClass, sqlCriteria As String, Optional sqlOrderBy As String = "id") As List(Of VisitSummaryModel)
            Return createList(Of VisitSummaryModel)(cp, sqlCriteria, sqlOrderBy)
        End Function
        '
        '====================================================================================================
        Public Overloads Shared Function getRecordName(cp As CPBaseClass, recordId As Integer) As String
            Return baseModel.getRecordName(Of VisitSummaryModel)(cp, recordId)
        End Function
        '
        '====================================================================================================
        Public Overloads Shared Function getRecordName(cp As CPBaseClass, ccGuid As String) As String
            Return baseModel.getRecordName(Of VisitSummaryModel)(cp, ccGuid)
        End Function
        '
        '====================================================================================================
        Public Overloads Shared Function getRecordId(cp As CPBaseClass, ccGuid As String) As Integer
            Return baseModel.getRecordId(Of VisitSummaryModel)(cp, ccGuid)
        End Function
        '
        '====================================================================================================
        Public Overloads Shared Function getCount(cp As CPBaseClass, sqlCriteria As String) As Integer
            Return baseModel.getCount(Of VisitSummaryModel)(cp, sqlCriteria)
        End Function
        '
        '====================================================================================================
        Public Overloads Function getUploadPath(fieldName As String) As String
            Return MyBase.getUploadPath(Of VisitSummaryModel)(fieldName)
        End Function
        '
        '====================================================================================================
        '
        Public Function Clone(cp As CPBaseClass) As VisitSummaryModel
            Dim result As VisitSummaryModel = DirectCast(Me.Clone(), VisitSummaryModel)
            result.id = cp.Content.AddRecord(contentName)
            result.ccguid = cp.Utils.CreateGuid()
            result.save(cp)
            Return result
        End Function
        '
        '====================================================================================================
        '
        Public Function Clone() As Object Implements ICloneable.Clone
            Return Me.MemberwiseClone()
        End Function

    End Class
End Namespace
