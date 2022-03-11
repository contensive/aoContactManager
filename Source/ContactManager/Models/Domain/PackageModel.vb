'
'====================================================================================================
'
''' <summary>
''' remote method top level data structure
''' </summary>
<Serializable()>
Public Class PackageModel
    Public Property success As Boolean = False
    Public Property errorList As New List(Of PackageErrorModel)
    Public Property nodeList As New List(Of PackageNodeModel)
    Public Property profileList As List(Of PackageProfileModel)
End Class