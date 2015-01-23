VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SqlResultRow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private Type tRow
    ParentResult As SqlResult
    Values() As Variant
    IsEmpty As Boolean
End Type

Private this As tRow
Option Explicit

Private Sub Class_Initialize()
    ReDim this.Values(0 To 0)
    this.IsEmpty = True
End Sub

Public Property Set ParentResult(value As SqlResult)
    Set this.ParentResult = value
End Property

Friend Sub AddValue(ByVal value As Variant)
    
    If Not this.IsEmpty Then ReDim Preserve this.Values(0 To UBound(this.Values) + 1)
    this.Values(UBound(this.Values)) = value
    
    this.IsEmpty = False
    
End Sub

Public Property Get Item(nameOrIndex As Variant) As Variant
Attribute Item.VB_UserMemId = 0
    
    If TypeName(nameOrIndex) = "String" Then
    
        Item = GetFieldValueByName(nameOrIndex)
    
    ElseIf IsNumeric(nameOrIndex) Then
    
        Item = GetFieldValueByIndex(nameOrIndex)
    
    Else
        
        'return empty variant
        
    End If
    
End Property

Private Function GetFieldValueByName(ByVal name As String) As Variant
    If Not this.IsEmpty Then GetFieldValueByName = this.Values(this.ParentResult.FieldNameIndex(name))
End Function

Private Function GetFieldValueByIndex(ByVal index As Integer) As Variant
    If Not this.IsEmpty Then GetFieldValueByIndex = this.Values(index)
End Function

Public Function Create(parent As SqlResult, fields As ADODB.fields) As SqlResultRow
    
    Dim result As New SqlResultRow
    Set result.ParentResult = parent
    
    Dim field As ADODB.field
    Dim value As Variant
    
    For Each field In fields
        If TypeName(field.value) = "String" Then
            value = LTrim(RTrim(Coalesce(field.value, vbNullString)))
        Else
            value = Coalesce(field.value, vbEmpty)
        End If
        result.AddValue value
    Next
    
    Set Create = result
    
End Function

Public Function ToString() As String
        
    If this.IsEmpty Then
        ToString = TypeName(Me)
        Exit Function
    End If
    
    Dim result As String
    result = Join(this.Values, this.ParentResult.ValueSeparator)
    
    ToString = result
    
End Function
