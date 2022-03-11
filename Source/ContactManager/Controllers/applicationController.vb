



Imports Contensive.BaseClasses
Imports Contensive.Addons.ContactManager.Models

Namespace Controllers
    '
    '====================================================================================================
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <remarks></remarks>
    Public Class ApplicationController
        Implements IDisposable
        '
        Private ReadOnly cp As CPBaseClass
        '
        '====================================================================================================
        ''' <summary>
        ''' Errors accumulated during rendering.
        ''' </summary>
        ''' <returns></returns>
        Public Property packageErrorList As New List(Of PackageErrorModel)
        '
        '====================================================================================================
        ''' <summary>
        ''' data accumulated during rendering
        ''' </summary>
        ''' <returns></returns>
        Public Property packageNodeList As New List(Of PackageNodeModel)
        '
        '====================================================================================================
        ''' <summary>
        ''' list of name/time used to performance analysis
        ''' </summary>
        ''' <returns></returns>
        Public Property packageProfileList As New List(Of PackageProfileModel)
        '
        '====================================================================================================
        ''' <summary>
        ''' status message displayed on get-form
        ''' </summary>
        ''' <returns></returns>
        Public Property statusMessage As String
            Get
                Return _StatusMessage
            End Get
            Set(value As String)
                _StatusMessage = value
            End Set
        End Property
        Private _StatusMessage As String = ""
        '
        '====================================================================================================
        ''' <summary>
        ''' application's user properties
        ''' </summary>
        ''' <returns></returns>
        Public ReadOnly Property userProperties As Domain.UserPropertiesModel
            Get
                If (_userProperties Is Nothing) Then
                    _userProperties = New Domain.UserPropertiesModel(cp)
                End If
                Return _userProperties
            End Get
        End Property
        Private _userProperties As Domain.UserPropertiesModel = Nothing
        '
        '====================================================================================================
        ''' <summary>
        ''' get the serialized results
        ''' </summary>
        ''' <returns></returns>
        Public Function getSerializedPackage() As String
            Try
                Return cp.JSON.Serialize(New PackageModel With {
                    .success = packageErrorList.Count.Equals(0),
                    .nodeList = packageNodeList,
                    .errorList = packageErrorList,
                    .profileList = packageProfileList
                })
            Catch ex As Exception
                cp.Site.ErrorReport(ex)
                Throw
            End Try
        End Function
        '
        '====================================================================================================
        ''' <summary>
        ''' Constructor
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub New(cp As CPBaseClass, Optional requiresAuthentication As Boolean = True)
            Me.cp = cp
            If (requiresAuthentication And Not cp.User.IsAuthenticated) Then
                packageErrorList.Add(New PackageErrorModel() With {.number = ResultErrorEnum.errAuthentication, .description = "Authorization is required."})
                cp.Response.SetStatus(HttpErrorEnum.forbidden & " Forbidden")
            End If
        End Sub
        '
#Region " IDisposable Support "
        Protected disposed As Boolean = False
        '
        '==========================================================================================
        ''' <summary>
        ''' dispose
        ''' </summary>
        ''' <param name="disposing"></param>
        ''' <remarks></remarks>
        Protected Overridable Overloads Sub Dispose(ByVal disposing As Boolean)
            If Not Me.disposed Then
                If disposing Then
                    '
                    ' ----- call .dispose for managed objects
                    '
                End If
                '
                ' Add code here to release the unmanaged resource.
                '
            End If
            Me.disposed = True
        End Sub
        ' Do not change or add Overridable to these methods.
        ' Put cleanup code in Dispose(ByVal disposing As Boolean).
        Public Overloads Sub Dispose() Implements IDisposable.Dispose
            Dispose(True)
            GC.SuppressFinalize(Me)
        End Sub
        Protected Overrides Sub Finalize()
            Dispose(False)
            MyBase.Finalize()
        End Sub
#End Region
    End Class
End Namespace
