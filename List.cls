VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "List"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Type tList
    Encapsulated As Collection
    ItemTypeName As String
End Type

Private this As tList
Option Explicit

Public Function Create(ParamArray Values() As Variant) As List
                
    Dim valuesArray() As Variant
    valuesArray = Values

    Dim result As New List
    result.AddArray valuesArray
    
    Set Create = result

End Function

'' ***************************************************************************************
'' Private functions
'' ***************************************************************************************

Private Function IsReferenceType() As Boolean
    If this.ItemTypeName = vbNullString Then Exit Function
    IsReferenceType = IsObject(this.Encapsulated(1))
End Function

Private Function IsComparable() As Boolean
    If IsReferenceType Then
        IsComparable = TypeOf First Is IComparable
    End If
End Function

Private Function CompareReferenceTypes(value As Variant, other As Variant) As Integer

    Dim comparable As IComparable

    If IsComparable Then

        Set comparable = value
        CompareReferenceTypes = comparable.CompareTo(other)

    Else

        RaiseErrorMustImplementIComparable "CompareReferenceTypes()"

    End If

End Function

Private Function CompareValueTypes(value As Variant, other As Variant) As Integer

    If value < other Then

        CompareValueTypes = -1

    ElseIf value > other Then

        CompareValueTypes = 1

    End If

End Function

Private Function IsEquatable() As Boolean
    If IsReferenceType Then
        IsEquatable = TypeOf First Is IEquatable
    End If
End Function

Private Function EquateReferenceTypes(value As Variant, other As Variant) As Boolean

    Dim equatable As IEquatable
    If IsEquatable Then

        Set equatable = value
        EquateReferenceTypes = equatable.Equals(other)

    Else

        Debug.Print "WARNING: Reference type doesn't implement IEquatable, using reference equality."
        EquateReferenceTypes = (ObjPtr(value) = ObjPtr(other))

    End If

End Function

Private Function EquateValueTypes(value As Variant, other As Variant) As Boolean

    EquateValueTypes = (value = other)

End Function

Private Function ValidateItemType(value As Variant)

    If this.ItemTypeName = vbNullString Then this.ItemTypeName = TypeName(value)
    ValidateItemType = IsTypeSafe(value)

End Function

Private Sub RaiseErrorUnsafeType(member As String, suppliedType As String)
    Err.Raise 13, Strings.Format("{0}.{1}", ToString, member), _
                  Strings.Format("Type Mismatch. Expected: '{0}', '{1}' was supplied.", this.ItemTypeName, suppliedType)
End Sub

Private Sub RaiseErrorMustImplementIComparable(member As String)
    Err.Raise 5, Strings.Format("{0}.{1}", ToString, member), "Invalid operation: method requires a list of numeric, date or string values, or a list of objects implementing the IComparable interface."
End Sub

Private Sub Class_Initialize()
    Set this.Encapsulated = New Collection
End Sub

Private Sub Class_Terminate()
    Set this.Encapsulated = Nothing
End Sub

'' ***************************************************************************************
'' Properties
'' ***************************************************************************************

Public Property Get Item(ByVal index As Long) As Variant
Attribute Item.VB_Description = "Gets or sets the element at the specified index."
Attribute Item.VB_UserMemId = 0
'Gets the element at the specified index.

    If IsReferenceType Then
        Set Item = this.Encapsulated(index)
    Else
        Item = this.Encapsulated(index)
    End If

End Property

Public Property Let Item(ByVal index As Long, ByVal value As Variant)
'Sets the element at the specified index.

    If Not IsTypeSafe(value) Then RaiseErrorUnsafeType "Item(Let)", TypeName(value)

    RemoveAt index
    If index = Count Then
        Add value
    Else
        Insert index, value
    End If

End Property

Public Property Set Item(ByVal index As Long, ByVal value As Variant)
'Sets the element at the specified index.

    If Not IsTypeSafe(value) Then RaiseErrorUnsafeType "Item(Set)", TypeName(value)

    RemoveAt index
    If index = Count Then
        Add value
    Else
        Insert index, value
    End If

End Property

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_Description = "Gets an enumerator that iterates through the List."
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
'Gets an enumerator that iterates through the List.

    Set NewEnum = this.Encapsulated.[_NewEnum]

End Property

Public Property Get Count() As Long
Attribute Count.VB_Description = "Gets the number of elements contained in the List."
'Gets the number of elements contained in the List.

    Count = this.Encapsulated.Count

End Property

'' ***************************************************************************************
'' Public functions
'' ***************************************************************************************

Public Sub Add(ParamArray Values())
Attribute Add.VB_Description = "Adds the specified element(s) to the end of the List."
'Adds the specified element(s) to the end of the List.

    Dim valuesArray() As Variant
    valuesArray = Values

    AddArray valuesArray

End Sub

Public Sub AddRange(Values As List)
Attribute AddRange.VB_Description = "Adds the specified elements to the end of the List."
'Adds the specified elements to the end of the List.

    AddArray Values.ToArray

End Sub

Public Sub AddArray(Values() As Variant)
Attribute AddArray.VB_Description = "Adds the specified elements to the end of the List."
'Adds the specified elements to the end of the List.

    Dim value As Variant, i As Long
    For i = LBound(Values) To UBound(Values)
        If ValidateItemType(value) Then
            this.Encapsulated.Add Values(i)
        Else
            RaiseErrorUnsafeType "AddArray()", TypeName(value)
        End If
    Next

End Sub

Public Sub Clear()
Attribute Clear.VB_Description = "Removes all elements from the List."
'Removes all elements from the List.

    Do Until Count = 0
        this.Encapsulated.Remove 1
    Loop

End Sub

Public Function Contains(value As Variant) As Boolean
Attribute Contains.VB_Description = "Determines whether an element is in the List."
'Determines whether an element is in the List.

    Contains = (IndexOf(value) <> -1)

End Function

Public Function First() As Variant
Attribute First.VB_Description = "Returns the first element of the List."
'Returns the first element of the List.

    If Count = 0 Then Exit Function
    If IsReferenceType Then
        Set First = Item(1)
    Else
        First = Item(1)
    End If

End Function

Public Function GetRange(ByVal index As Long, ByVal valuesCount As Long) As List
Attribute GetRange.VB_Description = "Creates a copy of a range of elements in the source List."
'Creates a copy of a range of elements in the source List.

    Dim result As List
    If index > Count Then Err.Raise 9 'index out of range

    Dim lastIndex As Long
    lastIndex = IIf(index + valuesCount > Count, Count, index + valuesCount)

    Set result = New List

    Dim i As Long
    For i = index To lastIndex
        result.Add Item(i)
    Next

    Set GetRange = result

End Function

Public Function IndexOf(value As Variant) As Long
Attribute IndexOf.VB_Description = "Searches for the specified object and returns the 1-based index of the first occurrence within the entire List."
'Searches for the specified object and returns the 1-based index of the first occurrence within the entire List.

    Dim found As Boolean
    Dim isRef As Boolean
    isRef = IsReferenceType

    Dim i As Long

    If Count = 0 Then IndexOf = -1: Exit Function
    For i = 1 To Count

        If isRef Then

            found = EquateReferenceTypes(value, Item(i))

        Else

            found = EquateValueTypes(value, Item(i))

        End If

        If found Then IndexOf = i: Exit Function

    Next

    IndexOf = -1

End Function

Public Sub InsertArray(ByVal index As Long, Values() As Variant)
Attribute InsertArray.VB_Description = "Inserts the specified elements into the List at the specified index."
'Inserts the specified elements into the List at the specified index.

    Dim tmp As List
    Set tmp = GetRange(index, Count)

    RemoveRange index, Count

    AddArray Values
    AddRange tmp

End Sub

Public Sub InsertRange(ByVal index As Long, Values As List)
Attribute InsertRange.VB_Description = "Inserts the specified elements into the List at the specified index."
'Inserts the specified elements into the List at the specified index.

    Dim tmp As List
    Set tmp = GetRange(index, Count)

    RemoveRange index, Count

    AddRange Values
    AddRange tmp

End Sub

Public Sub Insert(ByVal index As Long, ParamArray Values())
Attribute Insert.VB_Description = "Inserts the specified element(s) into the List at the specified index."
'Inserts the specified element(s) into the List at the specified index.

    Dim valuesArray() As Variant
    valuesArray = Values

    InsertArray index, valuesArray

End Sub

Public Function IsSortable() As Boolean
Attribute IsSortable.VB_Description = "Determines whether the List can be sorted."
'Determines whether the List can be sorted.

    If Count = 0 Then Exit Function

    Dim firstItem As Variant
    If IsReferenceType Then
        Set firstItem = First
    Else
        firstItem = First
    End If

    IsSortable = IsNumeric(firstItem) _
                Or IsDate(firstItem) _
                Or this.ItemTypeName = "String" _
                Or IsComparable

End Function

Public Function IsTypeSafe(value As Variant) As Boolean
Attribute IsTypeSafe.VB_Description = "Determines whether a value can be safely added to the List."
'Determines whether a value can be safely added to the List.

'Returns true if the type of specified value matches the type of items already in the list,
'or it the type of specified value is a numeric type smaller than the type of items already in the list.
'This means a List<Long> can contain Integer values, but a List<Integer> cannot contain Long values.

    Dim result As Boolean

    'most common cases: this.ItemTypeName isn't yet defined, or matches TypeName(value):
    result = this.ItemTypeName = vbNullString Or this.ItemTypeName = TypeName(value)
    If result Then IsTypeSafe = result: Exit Function

    'all other cases demand more processing:
    IsTypeSafe = result _
        Or this.ItemTypeName = "Integer" And Strings.MatchesAny(TypeName(value), "Byte") _
        Or this.ItemTypeName = "Long" And Strings.MatchesAny(TypeName(value), "Integer", "Byte") _
        Or this.ItemTypeName = "Single" And Strings.MatchesAny(TypeName(value), "Long", "Integer", "Byte") _
        Or this.ItemTypeName = "Double" And Strings.MatchesAny(TypeName(value), "Long", "Integer", "Byte", "Single") _
        Or this.ItemTypeName = "Currency" And Strings.MatchesAny(TypeName(value), "Long", "Integer", "Byte", "Single", "Double")

End Function

Public Function Last() As Variant
Attribute Last.VB_Description = "Returns the last element of the List."
'Returns the last element of the List.

    If Count = 0 Then Exit Function
    If IsReferenceType Then
        Set Last = Item(Count)
    Else
        Last = Item(Count)
    End If

End Function

Public Function LastIndexOf(value As Variant) As Long
Attribute LastIndexOf.VB_Description = "Searches for the specified object and returns the 1-based index of the last occurrence within the entire List."
'Searches for the specified object and returns the 1-based index of the last occurrence within the entire List.

    Dim found As Boolean
    Dim isRef As Boolean
    isRef = IsReferenceType

    LastIndexOf = -1
    If Count = 0 Then Exit Function

    Dim i As Long
    For i = 1 To Count

        If isRef Then

            found = EquateReferenceTypes(value, Item(i))

        Else

            found = EquateValueTypes(value, Item(i))

        End If

        If found Then LastIndexOf = i

    Next

End Function

Public Function Max() As Variant
Attribute Max.VB_Description = "Returns the maximum value in the List."
'Returns the maximum value in the List.

    Dim isRef As Boolean
    isRef = IsReferenceType

    Dim largest As Variant
    Dim isLarger As Boolean

    Dim i As Long
    For i = 1 To Count

        If isRef Then

            If IsEmpty(largest) Then Set largest = Item(i)
            isLarger = CompareReferenceTypes(Item(i), largest) > 0

            If isLarger Or IsEmpty(Max) Then
                Set largest = Item(i)
                Set Max = largest
            End If

        Else

            If IsEmpty(largest) Then largest = Item(i)
            isLarger = CompareValueTypes(Item(i), largest) > 0

            If isLarger Or IsEmpty(Max) Then
                largest = Item(i)
                Max = largest
            End If

        End If


    Next

End Function

Public Function Min() As Variant
Attribute Min.VB_Description = "Returns the minimum value in the List."
'Returns the minimum value in the List.

    Dim isRef As Boolean
    isRef = IsReferenceType

    Dim smallest As Variant
    Dim isSmaller As Boolean

    Dim i As Long
    For i = 1 To Count

        If isRef Then

            If IsEmpty(smallest) Then Set smallest = Item(i)
            isSmaller = CompareReferenceTypes(Item(i), smallest) < 0

            If isSmaller Or IsEmpty(Min) Then
                Set smallest = Item(i)
                Set Min = smallest
            End If

        Else

            If IsEmpty(smallest) Then smallest = Item(i)
            isSmaller = CompareValueTypes(Item(i), smallest) < 0

            If isSmaller Or IsEmpty(Min) Then
                smallest = Item(i)
                Min = smallest
            End If

        End If


    Next

End Function

Public Sub Reverse()
Attribute Reverse.VB_Description = "Reverses the order of the elements in the entire List."
'Reverses the order of the elements in the entire List.

    Dim tmp As New List
    Do Until Count = 0

        tmp.Add Item(Count)
        RemoveAt Count

    Loop

    AddRange tmp

End Sub

Public Sub Remove(ParamArray Values())
Attribute Remove.VB_Description = "Removes the first occurrence of specified object(s) from the List."
'Removes the first occurrence of specified object(s) from the List.

    Dim i As Long
    Dim index As Long

    For i = LBound(Values) To UBound(Values)

        index = IndexOf(Values(i))
        If index <> -1 Then RemoveAt index

    Next

End Sub

Public Sub RemoveAt(ByVal index As Long)
Attribute RemoveAt.VB_Description = "Removes the element at the specified index of the List."
'Removes the element at the specified index of the List.

    this.Encapsulated.Remove index

End Sub

Public Sub RemoveRange(ByVal index As Long, ByVal valuesCount As Long)
Attribute RemoveRange.VB_Description = "Removes a range of elements from the List."
'Removes a range of elements from the List.

    Dim i As Long
    For i = index To index + valuesCount - 1

        RemoveAt index

    Next

End Sub

Public Sub Sort()
Attribute Sort.VB_Description = "Sorts the elements in the entire List."
'Sorts the elements in the entire List.

    Dim tmp As List
    Dim minValue As Variant

    If Not IsSortable Then RaiseErrorMustImplementIComparable "Sort()"

    Dim isRef As Boolean
    isRef = IsReferenceType

    Set tmp = New List
    Do Until Count = 0

        If isRef Then

            Set minValue = Min

        Else

            minValue = Min

        End If

        tmp.Add minValue
        Remove minValue

    Loop

    AddRange tmp

End Sub

Public Sub SortDescending()
Attribute SortDescending.VB_Description = "Sorts the elements in the entire List, in descending order."
'Sorts the elements in the entire List, in descending order.

    Dim tmp As List
    Dim maxValue As Variant

    If Not IsSortable Then RaiseErrorMustImplementIComparable "SortDescending()"

    Dim isRef As Boolean
    isRef = IsReferenceType

    Set tmp = New List
    Do Until Count = 0

        If isRef Then
            Set maxValue = Max
        Else
            maxValue = Max
        End If

        tmp.Add maxValue
        Remove maxValue

    Loop

    AddRange tmp

End Sub

Public Function ToArray() As Variant()
Attribute ToArray.VB_Description = "Copies the elements of the List to a new array."
'Copies the elements of the List to a new array.

    Dim result() As Variant
    ReDim result(1 To Count)

    Dim i As Long
    If Count = 0 Then Exit Function

    If IsReferenceType Then
        For i = 1 To Count
            Set result(i) = Item(i)
        Next
    Else
        For i = 1 To Count
            result(i) = Item(i)
        Next
    End If

    ToArray = result

End Function

Public Function ToString() As String
Attribute ToString.VB_Description = "Returns a string that represents the current List object."
'Returns a string that represents the current List object.

    ToString = Strings.Format("{0}<{1}>", TypeName(Me), Coalesce(this.ItemTypeName, "Variant"))

End Function
