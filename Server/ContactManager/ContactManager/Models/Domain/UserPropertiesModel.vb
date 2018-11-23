
Option Explicit On
Option Strict On

Imports System
Imports System.Collections.Generic
Imports System.Text
Imports Contensive.BaseClasses

Namespace Models.Domain
    ''' <summary>
    ''' Expose public properties for all user properties created for this application.
    ''' </summary>
    Public Class UserPropertiesModel
        '
        Private cp As CPBaseClass
        '
        '
        '====================================================================================================
        ''' <summary>
        ''' constructor
        ''' </summary>
        ''' <param name="cp"></param>
        Public Sub New(cp As CPBaseClass)
            Me.cp = cp
        End Sub
        '
        '====================================================================================================
        ''' <summary>
        ''' 
        ''' </summary>
        ''' <returns></returns>
        Public Property contactSearchCriteria As String
            Get
                Return cp.User.GetText(nameContactSearchCriteria)
            End Get
            Set(value As String)
                cp.User.SetProperty(nameContactSearchCriteria, value)
            End Set
        End Property
        Private Const nameContactSearchCriteria As String = "ContactSearchCriteria"
        '
        '====================================================================================================
        ''' <summary>
        ''' 
        ''' </summary>
        ''' <returns></returns>
        Public Property contactGroupCriteria As String
            Get
                Return cp.User.GetText(nameContactGroupCriteria)
            End Get
            Set(value As String)
                cp.User.SetProperty(nameContactGroupCriteria, value)
            End Set
        End Property
        Private Const nameContactGroupCriteria As String = "ContactGroupCriteria"
        '
    End Class
End Namespace
