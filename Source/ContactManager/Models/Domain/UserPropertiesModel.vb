
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
                Return cp.TempFiles.Read("contactManager\user-contact-search-criteria.txt")
            End Get
            Set(value As String)
                cp.TempFiles.Save("contactManager\user-contact-search-criteria.txt", value)
            End Set
        End Property
        '
        '====================================================================================================
        ''' <summary>
        ''' 
        ''' </summary>
        ''' <returns></returns>
        Public Property contactGroupCriteria As String
            Get
                Return cp.TempFiles.Read("contactManager\user-group-search-criteria.txt")
            End Get
            Set(value As String)
                cp.TempFiles.Save("contactManager\user-group-search-criteria.txt", value)
            End Set
        End Property
        '
        '====================================================================================================
        ''' <summary>
        ''' 
        ''' </summary>
        ''' <returns></returns>
        Public Property selectSubTab As Integer
            Get
                Return cp.User.GetInteger(nameSelectSubTab)
            End Get
            Set(value As Integer)
                cp.User.SetProperty(nameSelectSubTab, value)
            End Set
        End Property
        Private Const nameSelectSubTab As String = "SelectSubTab"
        '
        '====================================================================================================
        ''' <summary>
        ''' 
        ''' </summary>
        ''' <returns></returns>
        Public Property subTab As Integer
            Get
                Return cp.User.GetInteger(nameSubTab, 1)
            End Get
            Set(value As Integer)
                cp.User.SetProperty(nameSubTab, value)
            End Set
        End Property
        Private Const nameSubTab As String = "SubTab"
        '
        '====================================================================================================
        ''' <summary>
        ''' 
        ''' </summary>
        ''' <returns></returns>
        Public Property ContactContentID As Integer
            Get
                Return cp.User.GetInteger(nameContactContentID, cp.Content.GetID("people"))
            End Get
            Set(value As Integer)
                cp.User.SetProperty(nameContactContentID, value)
            End Set
        End Property
        Private Const nameContactContentID As String = "ContactContentID"
        '
        '
        '
    End Class
End Namespace
