VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UnitOfWork"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const CONNECTION_STRING As String = ""

Private repositories As New Dictionary
Private adoConnection As New ADODB.Connection

Private disposed As Boolean

Implements IUnitOfWork
Implements IDisposable

Private Sub Class_Initialize()
    adoConnection.ConnectionString = CONNECTION_STRING
    adoConnection.Open
    adoConnection.BeginTrans
End Sub

Private Sub Class_Terminate()
    If Not disposed Then Dispose
End Sub

Private Sub Dispose()
    Set repositories = Nothing
    If Not adoConnection Is Nothing Then
        If adoConnection.State = adStateOpen Then
            adoConnection.RollbackTrans 'rollback any uncommitted changes
            adoConnection.Close
        End If
        Set adoConnection = Nothing
    End If
    disposed = True
End Sub

Private Sub IDisposable_Dispose()
    If Not disposed Then Dispose
End Sub

Public Sub AddRepository(ByVal key As String, ByRef repo As IRepository)
    repo.SetConnection adoConnection
    repositories.Add key, repo
End Sub

Public Property Get Repository(ByVal key As String) As IRepository
    Set Repository = repositories(key)
End Property

Public Sub Commit()
    adoConnection.CommitTrans
    adoConnection.BeginTrans
End Sub

Public Sub Rollback()
    adoConnection.RollbackTrans
    adoConnection.BeginTrans
End Sub

Private Sub IUnitOfWork_AddRepository(ByVal key As String, ByRef repo As IRepository)
    AddRepository key, repo
End Sub

Private Sub IUnitOfWork_Commit()
    Commit
End Sub

Private Property Get IUnitOfWork_Repository(ByVal key As String) As IRepository
    Set IUnitOfWork_Repository = Repository(key)
End Property

Private Sub IUnitOfWork_Rollback()
    Rollback
End Sub
