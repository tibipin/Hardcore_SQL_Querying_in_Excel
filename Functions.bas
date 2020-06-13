Attribute VB_Name = "Functions"
Global connection As ADODB.connection
Global rs As ADODB.Recordset
Global strSQL As String


Public Function connect(file)

    Dim conn_str As String
    
    Set connection = New ADODB.connection
    Set rs = New ADODB.Recordset
    
    
    conn_str = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & file & ";Extended Properties=""Excel 12.0 Macro;HDR=YES"";"
    
    
    connection.Open conn_str


End Function

Public Function reset_connection()

    connection.Close
    
    Set connection = Nothing
    Set rs = Nothing


End Function

Public Function allocation_sheet()

Dim sheeturi
Dim sheet_name
Dim a() As String

        Set sheeturi = connection.OpenSchema(adSchemaTables)
        Do While Not sheeturi.EOF
            sheet_name = sheeturi.Fields("table_name").Value
            If InStr(1, sheet_name, "To be processed") Then
                a = Split(sheet_name, "_")
                allocation_sheet = Mid(a(0), 2, 27)
            End If
            sheeturi.MoveNext
        Loop
        
End Function

Public Function findlastrow(ws As Worksheet)

findlastrow = ws.Range("A1").CurrentRegion.Rows.Count

End Function

