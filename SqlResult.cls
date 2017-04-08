VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SqlResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private Type tSqlResult
    FieldNames As List
    Values As List
    ToStringValueSeparator As String
End Type

Private nameIndices As New Dictionary

Private this As tSqlResult
Option Explicit

Private Sub Class_Initialize()
    Set this.FieldNames = List.Create
    Set this.Values = List.Create
    this.ToStringValueSeparator = ","
End Sub

Public Property Get ValueSeparator() As String
    ValueSeparator = this.ToStringValueSeparator
End Property

Public Property Let ValueSeparator(ByVal value As String)
    this.ToStringValueSeparator = value
End Property

Public Sub AddFieldName(name As String)
    this.FieldNames.Add name
    nameIndices.Add name, this.FieldNames.Count - 1
End Sub

Public Function FieldNameIndex(ByVal name As String) As Long
    
    Dim i As Long
    If nameIndices.TryGetValue(name, i) Then
        FieldNameIndex = i
    Else
        FieldNameIndex = -1
    End If
    
End Function

Public Sub AddValue(value As SqlResultRow)
    this.Values.Add value
End Sub

Public Property Get Count() As Long
    Count = this.Values.Count
End Property

Public Property Get Item(ByVal index As Long) As SqlResultRow
Attribute Item.VB_UserMemId = 0
    Set Item = this.Values(index + 1)
End Property

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
'Gets an enumerator that iterates through the List.
    
    Set NewEnum = this.Values.NewEnum

End Property

Public Function First() As SqlResultRow
    Set First = Item(0)
End Function

Public Function Create(adoRecordset As ADODB.Recordset) As SqlResult
    
    Dim result As New SqlResult
    
    Dim grabFieldNames As Boolean
    grabFieldNames = True
    
    Dim rowFactory As New SqlResultRow
    Dim field As ADODB.field
    
    While Not adoRecordset.BOF And Not adoRecordset.EOF
        
        For Each field In adoRecordset.fields
            If grabFieldNames Then result.AddFieldName LCase$(StringType.Coalesce(field.name, vbNullString))
        Next
        
        result.AddValue rowFactory.Create(result, adoRecordset.fields)
        
        grabFieldNames = False
        adoRecordset.MoveNext
    Wend
    
    Set Create = result
    
End Function


