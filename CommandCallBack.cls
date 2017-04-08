VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CommandCallBack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private owner As IPresenter
Implements ICommandCallback

Public Property Get CallbackOwner() As IPresenter
    Set CallbackOwner = owner
End Property

Public Property Set CallbackOwner(ByVal value As IPresenter)
    Set owner = value
End Property

Private Property Set ICommandCallback_CallbackOwner(ByVal value As IPresenter)
    Set owner = value
End Property

Private Property Get ICommandCallback_CallbackOwner() As IPresenter
    Set ICommandCallback_CallbackOwner = owner
End Property

Private Function ICommandCallback_CanExecute(ByVal cmd As CommandType) As Boolean
    If owner Is Nothing Then Exit Function
    ICommandCallback_CanExecute = CallByName(owner, "CanExecuteCommand", VbMethod, cmd)
End Function

Private Sub ICommandCallback_Execute(ByVal cmd As CommandType)
    If owner Is Nothing Then Exit Sub
    If Not ICommandCallback_CanExecute(cmd) Then Exit Sub
    CallByName owner, "ExecuteCommand", VbMethod, cmd
End Sub
