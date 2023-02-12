'Macro que crea X nº de códigos con nº de lote e identificador único, crea una pestaña por cada creación, y crea un índice donde agrupa con hipervínculos las hojas generadas con fecha y hora para trazabilidad'

Sub CreateCodes()
Dim numCodes As Integer
Dim lotNumber As Long
Dim ws As Worksheet
Dim wsIndex As Worksheet
Dim code As Integer
Dim sheetName As String
Dim creationDate As String
Dim creationTime As String
Dim uniqueSheetName As String
Dim lastRow As Long
numCodes = InputBox("Introduzca el número de códigos a crear:")
lotNumber = Int((9999999 - 1000000 + 1) * Rnd + 1000000)

'Comprobar si la hoja "Índice" existe
On Error Resume Next
Set wsIndex = ThisWorkbook.Sheets("Índice")
On Error GoTo 0
If wsIndex Is Nothing Then
    Set wsIndex = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    sheetName = "Índice"
    wsIndex.Name = sheetName
    wsIndex.Range("A1").Value = "Nombre de Hoja"
    wsIndex.Range("B1").Value = "Fecha de creación"
    wsIndex.Range("C1").Value = "Tiempo de creación"
End If

creationDate = Format(Now, "dd-MM-yyyy")
creationTime = Format(Now, "HH:mm:ss")
uniqueSheetName = "Codes " & Int((9999999 - 1000000 + 1) * Rnd + 1000000)
Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
ws.Name = uniqueSheetName
ws.Range("A1").Value = "Lote"
ws.Range("B1").Value = "Nº"

For code = 1 To numCodes
    ws.Range("A" & (code + 1)).Value = lotNumber
    ws.Range("B" & (code + 1)).Value = code
Next code

lastRow = wsIndex.Range("A" & wsIndex.Rows.Count).End(xlUp).Row + 1
wsIndex.Hyperlinks.Add Anchor:=wsIndex.Range("A" & lastRow), Address:="", SubAddress:= _
"'" & ws.Name & "'!A1", TextToDisplay:=ws.Name
wsIndex.Range("A" & lastRow).Value = ws.Name
wsIndex.Range("B" & lastRow).Value = creationDate
wsIndex.Range("C" & lastRow).Value = creationTime

MsgBox "¡Códigos creados con éxito"

End Sub
