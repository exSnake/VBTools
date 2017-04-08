VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IPresenter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Property Get UnitOfWork() As IUnitOfWork
End Property

Public Property Set UnitOfWork(ByVal value As IUnitOfWork)
End Property



Public Sub Show()
End Sub

Public Function ExecuteCommand(ByVal commandId As CommandType) As Variant
End Function

Public Function CanExecuteCommand(ByVal commandId As CommandType) As Boolean
End Function

Public Property Get DetailsPresenter() As IPresenter
End Property

Public Property Set DetailsPresenter(ByVal value As IPresenter)
End Property

Public Property Get MasterId() As Long
End Property

Public Property Let MasterId(ByVal value As Long)
End Property

