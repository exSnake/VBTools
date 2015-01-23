VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IRepository"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub SetConnection(ByRef value As ADODB.Connection)
End Sub

Public Function GetById(ByVal id As Long) As SqlResultRow
End Function

Public Function GetAll() As SqlResult
End Function

Public Function FindId(ByVal naturalKey As SqlResultRow) As Long
End Function

Public Function Search(ByVal terms As SqlResultRow) As SqlResult
End Function

Public Function Count() As Long
End Function

Public Sub Add(ByVal value As SqlResultRow)
End Sub

Public Sub Remove(ByVal id As Long)
End Sub

Public Sub Update(ByVal id As Long, ByVal value As SqlResultRow)
End Sub

Public Function NewItem(ByVal Model As SqlResult, ByVal values As Dictionary) As SqlResultRow
End Function

