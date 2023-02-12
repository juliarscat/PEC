'Imports data from a CSV, creates a sheet for every CSV, then add the name of the file to an index with the date and time of the data added, then creates an hyperlink for every file, it moves the file to a subfolder'

Sub ImportCSVFiles()
Dim folderPath As String
Dim fileName As String
Dim ws As Worksheet
Dim indexSheet As Worksheet
Dim lastRow As Long
Dim wb As Workbook


'Set the folder path where the CSV files are located
folderPath = "C:\CSVFiles\"

'Check if the index sheet already exists, if not, create it
On Error Resume Next
Set indexSheet = ThisWorkbook.Sheets("Index")
On Error GoTo 0
If indexSheet Is Nothing Then
    Set indexSheet = ThisWorkbook.Sheets.Add
    indexSheet.Name = "Index"
End If

'Loop through all the files in the folder
fileName = Dir(folderPath & "*.csv")
Do While fileName <> ""

    'Check if the sheet already exists, if not, create it
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(Replace(fileName, ".csv", ""))
    On Error GoTo 0
    If ws Is Nothing Then
    
        'Import the data from the CSV file into a new sheet
        On Error Resume Next
        Set wb = Workbooks.Open(folderPath & fileName, ReadOnly:=True)
        If Err.Number = 0 Then
            wb.Sheets(1).Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
            wb.Close False
            FileSystem.FileCopy folderPath & fileName, "C:\CSVFiles\read\" & fileName
FileSystem.Kill folderPath & fileName
            Set ws = ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
            ws.Name = Replace(fileName, ".csv", "")
            ws.Range("A1").CurrentRegion.TextToColumns Destination:=ws.Range("A1"), DataType:=xlDelimited, _
    TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
    Semicolon:=True, Comma:=True, Space:=False, Other:=False, FieldInfo:=Array(Array(1, 1), Array(2, 1))
        Else
            MsgBox "Error opening file: " & fileName & ". Error code: " & Err.Number
        End If
        On Error GoTo 0
    End If
    
    'Add the date and time the file was imported to the index sheet
    If Not IsError(Application.Match(Replace(fileName, ".csv", ""), indexSheet.Range("A:A"), 0)) Then
    
        'File name already exists in the index sheet, skip adding a new entry
        GoTo NextFile
    End If

    lastRow = indexSheet.Cells(indexSheet.Rows.Count, "A").End(xlUp).Row + 1
    indexSheet.Cells(lastRow, "A").value = Replace(fileName, ".csv", "")
    indexSheet.Cells(lastRow, "B").value = Format(Now, "yyyy-MM-dd hh:mm:ss")
    indexSheet.Hyperlinks.Add Anchor:=indexSheet.Cells(lastRow, "A"), Address:="", SubAddress:= _
        ws.Name & "!A1", TextToDisplay:=Replace(fileName, ".csv", "")

NextFile:
    'Move to the next file
    fileName = Dir()



Loop

End Sub





    'code for the add button'
    Sub AddImportButton()
    Dim btn As Button
    Set btn = ThisWorkbook.Sheets(1).Buttons.Add(10, 10, 70, 20)
    btn.Caption = "Import Data"
    btn.OnAction = "ImportCSVFiles"
    btn.Enabled = True
End Sub
