Attribute VB_Name = "SQL_to_Excel"
Sub export_sql()

'Fast export data from SQL to Excel, by ADODB.Command

'just run script with settings connection base (change: User ID, Password, Initial Catalog, Data Source, Workstation)
'change SQL string

'script export data to active sheet

Set objConn = CreateObject("ADODB.Connection")
Set objComm = CreateObject("ADODB.Command")

ConnStr = "Provider=SQLOLEDB.1;" & _
        "Persist Security Info=True;" & _
        "User ID=user_id;" & _
        "Password = password;" & _
        "Initial Catalog=need_base;" & _
        "Data Source=ip_address;" & _
        "Use Procedure for Prepare=1;" & _
        "Auto Translate=True;" & _
        "Packet Size=4096;" & _
        "Workstation ID=user_name;" & _
        "Use Encryption for Data=False;" & _
        "Tag with column collation when possible=False"
                
objConn.ConnectionString = ConnStr
objConn.Open
objComm.ActiveConnection = objConn

'string SQL query
objComm.CommandText = "select * from test_table"
objComm.CommandTimeout = 0

Set objRecordset = objComm.Execute

'number row begin paste data
row_var = 4

While Not objRecordset.EOF
    strRes = vbNullString
    For i = 0 To objRecordset.Fields.Count - 1
        Cells(row_var, i + 1) = objRecordset.Fields(i).Value
    Next
    row_var = row_var + 1
    objRecordset.MoveNext
Wend

objConn.Close

Set objConn = Nothing
Set objComm = Nothing
Set objRecordset = Nothing

End Sub






