Attribute VB_Name = "Main"
Sub Import()

Dim selected_folder As String
Dim fsoObject As FileSystemObject
Dim fsoFolder As Folder
Dim Subfolder As Folder
Dim fisier As file
Dim ur As Double
Dim ur2 As Double


'Identify Root Allocations based on a user selected month folder and extract data

arrFileNameFilter = Array("nd.xlsm", "rd.xlsm", "th.xlsm")

country_filter = Sheets("Frontsheet").[E3].Value
If [E3].Value = "" Then
    MsgBox "Select the country first."
    Exit Sub
End If



With Application.FileDialog(msoFileDialogFolderPicker)
    If .Show = -1 Then
        selected_folder = .SelectedItems(1)
    Else
        MsgBox "You haven't selected any folder"
        Exit Sub
    End If
End With

i = 1

Set fsoObject = New FileSystemObject
Set fsoFolder = fsoObject.GetFolder(selected_folder)
For Each Subfolder In fsoFolder.SubFolders
    For Each fisier In Subfolder.Files
        If Len(fisier.Name) - 5 = 32 And InStr(1, fisier.Name, country_filter) Then
                connect (fisier)
                Sheets("activity splits").Range("A" & i).Value = fisier
                alocare = allocation_sheet
                strSQL = "SELECT * FROM [" & alocare & "] WHERE [SGBS] IS NOT NULL AND [SGBS] <> ""Yes"""
                Set rs = connection.Execute(strSQL)
                On Error Resume Next
                '---Add allocation file name in raw data
                ur = findlastrow(Sheets("Raw Data"))
                Sheets("Raw Data").Range("B" & ur + 1).CopyFromRecordset rs
                ur2 = findlastrow(Sheets("Raw Data"))
                Sheets("Raw Data").Range("R" & ur + 1 & ":R" & ur2).Value = fisier.Name
                '---
                Call reset_connection
                i = i + 1
        Else
            For s = LBound(arrFileNameFilter) To UBound(arrFileNameFilter)
                    If InStr(1, fisier.Name, arrFileNameFilter(s)) Then
                        connect (fisier)
                        Sheets("activity splits").Range("A" & i).Value = fisier
                        alocare = allocation_sheet
                        strSQL = "SELECT * FROM [" & alocare & "] WHERE [SGBS] IS NOT NULL AND [SGBS] <> ""Yes"""
                        Set rs = connection.Execute(strSQL)
                        On Error Resume Next
                        '---Add allocation file name in raw data
                        ur = findlastrow(Sheets("Raw Data"))
                        Sheets("Raw Data").Range("B" & ur + 1).CopyFromRecordset rs
                        ur2 = findlastrow(Sheets("Raw Data"))
                        Sheets("Raw Data").Range("R" & ur + 1 & ":R" & ur2).Value = fisier.Name
                        '---
                        Call reset_connection
                        i = i + 1
                    End If
            Next
        End If
    Next
Next

     MsgBox "Data Extraction Successful"
        
End Sub


Sub CalculateReworks()

ur = findlastrow(Sheets("Raw Data"))
If ur < 2 Then
    MsgBox "You haven't imported any data, so what am I supposed to calculate?"
    Exit Sub
End If

'Sort based on columns: [HE_Transaction Number], [HE_Last Change Workflow Status] and [Allocation Name]

With Sheets("Raw Data").Sort
    .SortFields.Clear
    .SortFields.Add Key:=Sheets("Raw Data").Range("B1"), Order:=xlAscending
    .SortFields.Add Key:=Sheets("Raw Data").Range("J1"), Order:=xlAscending
    .SortFields.Add Key:=Sheets("Raw Data").Range("R1"), Order:=xlAscending
    .SetRange Sheets("Raw Data").Range("A1").CurrentRegion
    .Header = xlYes
    .Apply
End With


'Add index

ur = findlastrow(Sheets("Raw Data"))
For i = 2 To ur
    Sheets("Raw Data").Cells(i, 1).Value = i - 1
Next

'Create a copy of Raw Data

connect (ThisWorkbook.Path & "\" & ThisWorkbook.Name)
strSQL = "SELECT * FROM [Raw Data$]"
Set rs = connection.Execute(strSQL)
For i = 0 To rs.Fields.Count - 1
    Sheets("Temp").Cells(1, i + 1).Value = rs.Fields(i).Name
Next
Sheets("Temp").Range("A2").CopyFromRecordset rs
Call reset_connection

'Extract Results - part 1 - processed and reworks

strSQL = "SELECT [Raw Data$].[Index], [Raw Data$].[HE_Transaction Number], [Raw Data$].[Rework Status], [Raw Data$].[Processing Status] FROM [Raw Data$] LEFT JOIN [Temp$] " & _
        "ON  [Raw Data$].[Index] = [Temp$].[Index] + 1 WHERE [Raw Data$].[HE_Transaction Number] = [Temp$].[HE_Transaction Number] " & _
        "AND [Raw Data$].[HE_Last Change Workflow Status] <> [Temp$].[HE_Last Change Workflow Status]"

connect (ThisWorkbook.Path & "\" & ThisWorkbook.Name)
Set rs = connection.Execute(strSQL)
Sheets("Results").Range("A2").CopyFromRecordset rs
For i = 0 To rs.Fields.Count - 1
    Sheets("Results").Cells(1, i + 1).Value = rs.Fields(i).Name
Next
Call reset_connection

ur = findlastrow(Sheets("Results"))

For Each cell In Sheets("Results").Range("C2:C" & ur)
    cell.Value = "Rework"
    cell.Offset(0, 1).Value = "Processed"
Next cell

connect (ThisWorkbook.Path & "\" & ThisWorkbook.Name)
strSQL = "SELECT [Results$].[Rework Status], [Results$].[Processing status] FROM [Raw Data$] LEFT JOIN [Results$] ON [Results$].[Index] = [Raw Data$].[Index]"
Set rs = connection.Execute(strSQL)
Sheets("Raw Data").Range("S2").CopyFromRecordset rs
Call reset_connection

Sheets("Results").Cells.Delete

'Extract Results - part 2 - processed:
'-single appearances of values in column [HE_Transaction Number] in the reporting period
'-changes in workflow status for values in column [HE_Transaction Number] that appear more than once
'-changes in last change workflow status for values in column [HE_Transaction Number] that appear more than once

strSQL = "SELECT [Raw Data$].[Index], [Raw Data$].[HE_Transaction Number], [Raw Data$].[Processing Status] FROM [Raw Data$] LEFT JOIN [Temp$] " & _
        "ON  [Raw Data$].[Index] = [Temp$].[Index] - 1 WHERE [Raw Data$].[HE_Transaction Number] <> [Temp$].[HE_Transaction Number] OR " & _
        "([Raw Data$].[HE_Transaction Number] = [Temp$].[HE_Transaction Number] AND [Raw Data$].[HE_Last Change Workflow Status] <> [Temp$].[HE_Last Change Workflow Status]) OR " & _
        "([Raw Data$].[HE_Transaction Number] = [Temp$].[HE_Transaction Number] AND [Raw Data$].[HE_Last Change Workflow Status] = [Temp$].[HE_Last Change Workflow Status] AND " & _
        "[Raw Data$].[HE_Workflow Status] <> [Temp$].[HE_Workflow Status])"
        
connect (ThisWorkbook.Path & "\" & ThisWorkbook.Name)
Set rs = connection.Execute(strSQL)
Sheets("Results").Range("A2").CopyFromRecordset rs
For i = 0 To rs.Fields.Count - 1
    Sheets("Results").Cells(1, i + 1).Value = rs.Fields(i).Name
Next
Call reset_connection
ur = findlastrow(Sheets("Results"))
    
For Each cell In Sheets("Results").Range("C2:C" & ur)
    cell.Value = "Processed"
Next cell

connect (ThisWorkbook.Path & "\" & ThisWorkbook.Name)
strSQL = "SELECT [Results$].[Processing status] FROM [Raw Data$] LEFT JOIN [Results$] ON [Results$].[Index] = [Raw Data$].[Index]"
Set rs = connection.Execute(strSQL)
Sheets("Raw Data").Range("T2").CopyFromRecordset rs
Call reset_connection

Sheets("Results").Cells.Delete

'Extract results - part 3 - not processed & not reworks (remaining blanks in [Processing status] column)

ur = findlastrow(Sheets("Raw Data"))
For Each cell In Sheets("Raw Data").Range("T2:T" & ur)
    If cell.Value <> "Processed" Then
        cell.Value = "Not Processed"
    End If
    If cell.Offset(0, -1).Value <> "Rework" Then
        cell.Offset(0, -1).Value = "Not Rework"
    End If
Next

'Extract results - part 4 - vendor changes

strSQL = "SELECT [Raw Data$].[Index], [Raw Data$].[HE_Transaction Number], [Raw Data$].[Vendor change] FROM [Raw Data$] LEFT JOIN [Temp$] " & _
        "ON  [Raw Data$].[Index] = [Temp$].[Index] + 1 WHERE [Raw Data$].[HE_Transaction Number] = [Temp$].[HE_Transaction Number] " & _
        "AND [Raw Data$].[HE_Creditor Number] <> [Temp$].[HE_Creditor Number]"
connect (ThisWorkbook.Path & "\" & ThisWorkbook.Name)
Set rs = connection.Execute(strSQL)
If rs.EOF Then
    GoTo pas1
End If
Sheets("Results").Range("A2").CopyFromRecordset rs
For i = 0 To rs.Fields.Count - 1
    Sheets("Results").Cells(1, i + 1).Value = rs.Fields(i).Name
Next
Call reset_connection

ur = findlastrow(Sheets("Results"))
    
For Each cell In Sheets("Results").Range("C2:C" & ur)
    cell.Value = "Vendor changed"
Next cell

connect (ThisWorkbook.Path & "\" & ThisWorkbook.Name)
strSQL = "SELECT [Results$].[Vendor change] FROM [Raw Data$] LEFT JOIN [Results$] ON [Results$].[Index] = [Raw Data$].[Index]"
Set rs = connection.Execute(strSQL)
Sheets("Raw Data").Range("U2").CopyFromRecordset rs
pas1:
Call reset_connection

Sheets("Results").Cells.Delete

'Extract results - part 5 - invoice type changes

strSQL = "SELECT [Raw Data$].[Index], [Raw Data$].[HE_Transaction Number], [Raw Data$].[Invoice type change] FROM [Raw Data$] LEFT JOIN [Temp$] " & _
        "ON  [Raw Data$].[Index] = [Temp$].[Index] + 1 WHERE [Raw Data$].[HE_Transaction Number] = [Temp$].[HE_Transaction Number] " & _
        "AND [Raw Data$].[HE_Invoice Type] <> [Temp$].[HE_Invoice Type]"
connect (ThisWorkbook.Path & "\" & ThisWorkbook.Name)
Set rs = connection.Execute(strSQL)
If rs.EOF Then
    GoTo pas2
End If

Sheets("Results").Range("A2").CopyFromRecordset rs
For i = 0 To rs.Fields.Count - 1
    Sheets("Results").Cells(1, i + 1).Value = rs.Fields(i).Name
Next
Call reset_connection

ur = findlastrow(Sheets("Results"))
    
For Each cell In Sheets("Results").Range("C2:C" & ur)
    cell.Value = "Invoice type changed"
Next cell

connect (ThisWorkbook.Path & "\" & ThisWorkbook.Name)
strSQL = "SELECT [Results$].[Invoice type change] FROM [Raw Data$] LEFT JOIN [Results$] ON [Results$].[Index] = [Raw Data$].[Index]"
Set rs = connection.Execute(strSQL)
Sheets("Raw Data").Range("V2").CopyFromRecordset rs
pas2:
Call reset_connection

Sheets("Results").Cells.Delete

'Extract results - part 6 - company code changes

strSQL = "SELECT [Raw Data$].[Index], [Raw Data$].[HE_Transaction Number], [Raw Data$].[Company code change] FROM [Raw Data$] LEFT JOIN [Temp$] " & _
        "ON  [Raw Data$].[Index] = [Temp$].[Index] + 1 WHERE [Raw Data$].[HE_Transaction Number] = [Temp$].[HE_Transaction Number] " & _
        "AND [Raw Data$].[HE_Company Code] <> [Temp$].[HE_Company Code]"
connect (ThisWorkbook.Path & "\" & ThisWorkbook.Name)
Set rs = connection.Execute(strSQL)
If rs.EOF Then
    GoTo pas3
End If
Sheets("Results").Range("A2").CopyFromRecordset rs
For i = 0 To rs.Fields.Count - 1
    Sheets("Results").Cells(1, i + 1).Value = rs.Fields(i).Name
Next
Call reset_connection

ur = findlastrow(Sheets("Results"))
    
For Each cell In Sheets("Results").Range("C2:C" & ur)
    cell.Value = "Company code changed"
Next cell

connect (ThisWorkbook.Path & "\" & ThisWorkbook.Name)
strSQL = "SELECT [Results$].[Company code change] FROM [Raw Data$] LEFT JOIN [Results$] ON [Results$].[Index] = [Raw Data$].[Index]"
Set rs = connection.Execute(strSQL)
Sheets("Raw Data").Range("W2").CopyFromRecordset rs
pas3:
Call reset_connection

Sheets("Results").Cells.Delete

Sheets("Temp").Cells.Delete


Call comments

MsgBox "The following calculations have been performed:" & vbNewLine & vbNewLine & _
        "- Invoice processing performance per user" & vbNewLine & _
        "- Number of reworks from total number of invoices allocated" & vbNewLine & _
        "- Number of vendor changes, invoice type changes and company code changes" & vbNewLine & _
        "- Comments for the invoices not processed"


End Sub

Private Sub comments()
Dim cell, cell2 As Range
Dim nume_alocare As String

ur = findlastrow(Sheets("Raw Data"))
ur2 = findlastrow(Sheets("activity splits"))

For Each cell In Sheets("Raw Data").Range("T2:T" & ur)
    If cell.Value = "Not Processed" Then
        nume_alocare = cell.Offset(0, -2).Value
        For Each cell2 In Sheets("activity splits").Range("A1:A" & ur2)
            If InStr(1, cell2, nume_alocare) Then
                user_file = Left(cell2, Len(cell2) - Len(nume_alocare)) & _
                            Left(nume_alocare, Len(nume_alocare) - 5) & " " & _
                            cell.Offset(0, -6).Value & ".xlsx"
                connect (user_file)
                strSQL = "SELECT [Comments] from [Sheet1$] WHERE [HE_Transaction Number] = " & """" & cell.Offset(0, -18) & """"
                Set rs = connection.Execute(strSQL)
                On Error Resume Next
                cell.Offset(0, -3).CopyFromRecordset rs
                Call reset_connection
            End If
        Next
    End If
Next


End Sub

Sub delete_data()


ur = findlastrow(Sheets("Raw Data"))
If ur = 1 Then
    GoTo Pas
Else
    Sheets("Raw Data").Range("A2:W" & ur).Cells.Delete
End If
Pas:
Sheets("Temp").Cells.Delete
Sheets("Results").Cells.Delete
Sheets("activity splits").Cells.Delete

MsgBox "Raw data cleared."

End Sub

