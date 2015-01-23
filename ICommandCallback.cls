VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ICommandCallback"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum CommandType
    CloseCommand
    AcceptCommand
    AddCommand
    EditCommand
    DeleteCommand
    RefreshCommand
    ShowDetailsCommand
    EditKeyCommand
    EditDateCommand
    EditDescriptionCommand
    FindCommand
End Enum

Public Function CanExecute(ByVal cb As CommandType) As Boolean
End Function

Public Sub Execute(ByVal cb As CommandType)
End Sub

Public Property Get CallbackOwner() As IPresenter
End Property

Public Property Set CallbackOwner(ByVal value As IPresenter)
End Property
