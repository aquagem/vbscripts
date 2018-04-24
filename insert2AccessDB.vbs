Public Function accessDb_operations(accessDB_Path,DB_NAME,sql)
	'This function is used to insert records into the Access DB'
	On Error Resume Next
	Err.Number = 0
	
	'Variable Declaration Start'
		Dim connStr
		Dim objConn
		Dim results
		Dim rs
	'Variable Declaration End'
	
'	connStr = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=D:\DATA\TIC_FILE_DB_V1.0.accdb"
	connStr = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source="&accessDB_Path&"\"&DB_NAME
	
	'Define object type
	Set objConn = CreateObject("ADODB.Connection")
	SET rs = CreateObject("adodb.recordset")
	
	'Open Connection
	objConn.open connStr
	objConn.Execute(sql) 'Execute the sql'
	
	'Check if any error exists'
	If Err.Number <> 0 Then
		MsgBox "Error occured while executing sql: "&Err.Description
	Else
		MsgBox "SQL executed successfully"
	End If
	
	'objConn.Close 
	'Set objConn = Nothing
End Function
