VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IUnitOfWork"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub AddRepository(ByVal key As String, ByRef repo As IRepository)
End Sub

Public Property Get Repository(ByVal key As String) As IRepository
End Property

Public Sub Commit()
End Sub

Public Sub Rollback()
End Sub
