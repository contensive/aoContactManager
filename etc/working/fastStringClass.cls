VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "fastStringClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit
'
Private iSize As Long
Private Const iChunk = 100
Private iCount As Long
Private Holder() As String
'
'
'
Public Sub Add(NewString As String)
    If iCount >= iSize Then
        iSize = iSize + iChunk
        ReDim Preserve Holder(iSize)
        End If
    Holder(iCount) = NewString
    iCount = iCount + 1
    End Sub
'
'
'
Public Property Get Text() As String
    Text = Join(Holder, "")
    End Property

