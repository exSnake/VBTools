VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SqlCommand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private converter As New AdoValueConverter
Private connString As String
Private resultFactory As New SqlResult
Option Explicit

Public Property Get ConnectionString() As String
    ConnectionString = connString
End Property

Public Property Let ConnectionString(ByVal value As String)
    connString = value
End Property

Public Property Get ParameterFactory() As AdoValueConverter
Attribute ParameterFactory.VB_Description = "Gets an object that can create ADODB Parameters and configure how ADODB Parameters are created."
    Set ParameterFactory = converter
End Property

Public Function Execute(connection As ADODB.connection, ByVal sql As String, ParamArray parameterValues()) As ADODB.Recordset
Attribute Execute.VB_Description = "Returns a connected ADODB.Recordset that contains the results of the specified parameterized query."
'Returns a connected ADODB.Recordset that contains the results of the specified parameterized query.

    Dim parameters() As Variant
    parameters = parameterValues
    
    Set Execute = ExecuteInternal(connection, sql, parameters)
    
End Function

Public Function ExecuteNonQuery(connection As ADODB.connection, ByVal sql As String, ParamArray parameterValues()) As Boolean
Attribute ExecuteNonQuery.VB_Description = "Returns a Boolean that indicates whether the specified parameterized SQL command (update, delete, etc.) executed without throwing an error."
'Returns a Boolean that indicates whether the specified parameterized SQL command (update, delete, etc.) executed without throwing an error.

    Dim parameters() As Variant
    parameters = parameterValues
    
    ExecuteNonQuery = ExecuteNonQueryInternal(connection, sql, parameters)
    
End Function

Public Function ExecuteStoredProc(connection As ADODB.connection, ByVal spName As String, ParamArray parameterValues()) As ADODB.Recordset
Attribute ExecuteStoredProc.VB_Description = "Executes the specified parameterized stored procedure, passing specified parameter values."
'Executes the specified parameterized stored procedure, passing specified parameter values.

    Dim parameters() As Variant
    parameters = parameterValues
    
    Set ExecuteStoredProc = ExecuteStoredProcInternal(connection, spName, parameters)
    
End Function

Public Function ExecuteWithParameters(connection As ADODB.connection, ByVal spName As String, parameters As List) As ADODB.Recordset
Attribute ExecuteWithParameters.VB_Description = "Executes specified parameterized stored procedure by passing the specified preconfigured parameters. Use this method over ExecuteStoredProcedure for more complex parameter scenarios, such as running a stored procedure with output parameters."
'Executes specified parameterized stored procedure by passing the specified preconfigured parameters.
'Use this method over ExecuteStoredProcedure for more complex parameter scenarios, such as running a stored procedure with output parameters.

    'parameters is expected to be a List containing ADODB.Parameter objects: parameters.ToString = "List<Parameter>"
    If Not parameters.IsTypeSafe(New ADODB.parameter) Or parameters.ToString = "List<Variant>" Then Err.Raise vbObjectError + 2, "SqlCommand.ExecuteWithParameters", "Wrong list type. Argument 'parameters' must be a List<Parameter>."
    
    Dim emptyArray() As Variant
    Dim cmd As New ADODB.Command
    Set cmd = CreateCommand(connection, adCmdStoredProc, spName, emptyArray)
    
    Dim param As ADODB.parameter
    For Each param In parameters
        cmd.parameters.Append param
    Next
    
    Set ExecuteWithParameters = cmd.Execute

End Function

Public Function SelectSingleValue(connection As ADODB.connection, ByVal sql As String, ParamArray parameterValues()) As Variant
Attribute SelectSingleValue.VB_Description = "Returns the value of the first field of the first record of the results of the specified parameterized SQL query."
'Returns the value of the first field of the first record of the results of the specified parameterized SQL query.

    Dim parameters() As Variant
    parameters = parameterValues
    
    SelectSingleValue = SelectSingleValueInternal(connection, sql, parameters)
    
End Function

Public Function QuickExecute(ByVal sql As String, ParamArray parameterValues()) As SqlResult
Attribute QuickExecute.VB_Description = "Returns an object representing the result of the specified parameterized SQL statement. Field values for each row can then be accessed by field name or by field index."
'Returns an object representing the result of the specified parameterized SQL statement. Field values for each row can then be accessed by field name or by field index.
    
    Dim parameters() As Variant
    parameters = parameterValues

    Set QuickExecute = QuickExecuteInternal(sql, parameters)
        
End Function

Public Function QuickExecuteNonQuery(ByVal sql As String, ParamArray parameterValues()) As Boolean
Attribute QuickExecuteNonQuery.VB_Description = "Opens a connection and returns a Boolean that indicates whether the specified parameterized SQL command (update, delete, etc.) executed without throwing an error, then closes the connection."
'Opens a connection and returns a Boolean that indicates whether the specified parameterized SQL command (update, delete, etc.) executed without throwing an error,
'then closes the connection.

    Dim parameters() As Variant
    parameters = parameterValues

    Dim connection As New ADODB.connection
    connection.ConnectionString = connString
    
    connection.Open
    
    QuickExecuteNonQuery = ExecuteNonQueryInternal(connection, sql, parameters)
    
    connection.Close
    Set connection = Nothing
    
End Function

Public Function QuickExecuteStoredProc(ByVal spName As String, ParamArray parameterValues()) As SqlResult
Attribute QuickExecuteStoredProc.VB_Description = "Executes the specified parameterized stored procedure, passing specified parameter values."
'Executes the specified parameterized stored procedure, passing specified parameter values.
    
    Dim parameters() As Variant
    parameters = parameterValues

    Dim connection As New ADODB.connection
    connection.ConnectionString = connString
    
    connection.Open
    
    Dim rs As ADODB.Recordset
    Set rs = ExecuteStoredProcInternal(connection, spName, parameters)
    
    Set QuickExecuteStoredProc = resultFactory.Create(rs)
    
    rs.Close
    Set rs = Nothing
    
    connection.Close
    Set connection = Nothing
        
End Function

Public Function QuickExecuteWithParameters(ByVal spName As String, parameters As List) As SqlResult
'Executes specified parameterized stored procedure by passing the specified preconfigured parameters.
'Use this method over ExecuteStoredProcedure for more complex parameter scenarios, such as running a stored procedure with output parameters.
    
    Dim connection As New ADODB.connection
    connection.ConnectionString = connString
    
    connection.Open
    
    Dim rs As ADODB.Recordset
    Set rs = ExecuteWithParameters(connection, spName, parameters)
    
    Set QuickExecuteWithParameters = resultFactory.Create(rs)
    
    rs.Close
    Set rs = Nothing
    
    connection.Close
    Set connection = Nothing
    
End Function

Public Function QuickSelectFirstRow(ByVal sql As String, ParamArray parameterValues()) As SqlResultRow
Attribute QuickSelectFirstRow.VB_Description = "Opens a connection and returns an object representing the first record of the results of the specified parameterized SQL query, then closes the connection."
'Opens a connection and returns an object representing the first record of the results of the specified parameterized SQL query, then closes the connection.
    
    Dim parameters() As Variant
    parameters = parameterValues
    
    Dim result As SqlResult
    Set result = QuickExecuteInternal(sql, parameters)
    
    Set QuickSelectFirstRow = result.First
    
End Function

Public Function QuickSelectSingleValue(ByVal sql As String, ParamArray parameterValues()) As Variant
Attribute QuickSelectSingleValue.VB_Description = "Opens a connection and returns the value of the first field of the first record of the results of the specified parameterized SQL query, then closes the connection."
'Opens a connection and returns the value of the first field of the first record of the results of the specified parameterized SQL query, then closes the connection.

    Dim parameters() As Variant
    parameters = parameterValues

    Dim connection As New ADODB.connection
    connection.ConnectionString = connString
    
    connection.Open
    
    Set QuickSelectSingleValue = SelectSingleValueInternal(connection, sql, parameters)
    
    connection.Close
    Set connection = Nothing
    
End Function

Private Function CreateCommand(connection As ADODB.connection, ByVal cmdType As ADODB.CommandTypeEnum, ByVal sql As String, parameterValues() As Variant) As ADODB.Command
    
    Dim cmd As New ADODB.Command
    cmd.ActiveConnection = connection
    cmd.CommandType = cmdType
    cmd.CommandText = sql
        
    Dim i As Integer
    Dim value As Variant
    
    For i = LBound(parameterValues) To UBound(parameterValues)
        value = parameterValues(i)
        If TypeName(value) <> "Variant()" Then cmd.parameters.Append ToSqlInputParameter(value)
    Next
    
    Set CreateCommand = cmd

End Function

Private Function ToSqlInputParameter(ByVal value As Variant) As ADODB.parameter
    
    If IsObject(value) Then Err.Raise vbObjectError + 911, "SqlCommand.ToSqlInputParameter", "Invalid argument, parameter value cannot be an object."
    
    Dim result As ADODB.parameter
    Set result = CallByName(converter, "To" & TypeName(value) & "Parameter", VbMethod, value, ADODB.ParameterDirectionEnum.adParamInput)
    
    Set ToSqlInputParameter = result
    
End Function


Private Function ExecuteInternal(connection As ADODB.connection, ByVal sql As String, parameterValues()) As ADODB.Recordset
    
    Dim cmd As ADODB.Command
    Set cmd = CreateCommand(connection, adCmdText, sql, parameterValues)
    
    Set ExecuteInternal = cmd.Execute
    
End Function

Private Function ExecuteNonQueryInternal(connection As ADODB.connection, ByVal sql As String, parameterValues()) As Boolean

    Dim cmd As ADODB.Command
    Set cmd = CreateCommand(connection, adCmdText, sql, parameterValues)
    
    Dim result As Boolean
    On Error Resume Next
        cmd.Execute
        result = (Err.Number = 0)
    On Error GoTo 0
    
    ExecuteNonQueryInternal = result
    
End Function

Private Function ExecuteStoredProcInternal(connection As ADODB.connection, ByVal spName As String, parameterValues()) As ADODB.Recordset
    
    Dim cmd As ADODB.Command
    Set cmd = CreateCommand(connection, adCmdStoredProc, spName, parameterValues)
    
    Set ExecuteStoredProcInternal = cmd.Execute

End Function

Private Function SelectSingleValueInternal(connection As ADODB.connection, ByVal sql As String, parameterValues()) As Variant
    
    Dim parameters() As Variant
    parameters = parameterValues
    
    Dim cmd As ADODB.Command
    Set cmd = CreateCommand(connection, adCmdText, sql, parameters)
    
    Dim rs As ADODB.Recordset
    Set rs = cmd.Execute
    
    Dim result As Variant
    If Not rs.BOF And Not rs.EOF Then result = rs.fields(0).value
    
    rs.Close
    Set rs = Nothing
    
    SelectSingleValueInternal = result

End Function

Private Function QuickExecuteInternal(ByVal sql As String, parameterValues()) As SqlResult
    
    Dim parameters() As Variant
    parameters = parameterValues
    
    Dim connection As New ADODB.connection
    connection.ConnectionString = connString
    
    connection.Open
    
    Dim rs As ADODB.Recordset
    Set rs = ExecuteInternal(connection, sql, parameters)
    
    Set QuickExecuteInternal = resultFactory.Create(rs)
    
    rs.Close
    Set rs = Nothing
    
    connection.Close
    Set connection = Nothing
        
End Function
