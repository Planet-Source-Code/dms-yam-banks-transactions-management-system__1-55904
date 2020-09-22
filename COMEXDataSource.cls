VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "COMEXDataSource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub SetData(ByVal Field As Long, ByVal Record As Long, ByVal newData As Variant)
'
End Sub

Public Function GetFieldName(ByVal Field As Long) As String
'
End Function

Public Function GetRecordCount() As Long
'
End Function

Public Function GetData(ByVal Field As Long, ByVal Record As Long) As Variant
'
End Function

Public Function GetFieldCount() As Long
'
End Function

Public Sub Delete(ByVal Record As Long)
'
End Sub

Public Function Save() As Boolean
'
End Function

Public Property Get IsDeleted(ByVal Record As Long) As Boolean
'
End Property
