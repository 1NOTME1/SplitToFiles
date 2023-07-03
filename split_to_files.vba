Sub Podziel()

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Arkusz1")

    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    For i = 2 To lastRow
        If Not dict.Exists(ws.Cells(i, "A").Value) Then
            dict.Add ws.Cells(i, "A").Value, ws.Cells(i, "A").Value
        End If
    Next

    For Each Key In dict.keys
        Dim newWB As Workbook
        Set newWB = Workbooks.Add
        With newWB
            .Title = Key
            .Subject = "Data for " & Key
        End With
        
        ' Kopiowanie nagłówka
        ws.Range("A1:F1").Copy
        newWB.Sheets(1).Range("A1").PasteSpecial xlPasteAllUsingSourceTheme
       
        ' Filtrowanie i kopiowanie widocznych komórek
        ws.Range("A1:F" & lastRow).AutoFilter field:=1, Criteria1:=Key
        ws.Range("A2:F" & lastRow).SpecialCells(xlCellTypeVisible).Copy
        newWB.Sheets(1).Range("A2").PasteSpecial xlPasteAllUsingSourceTheme

        ' Zapis i zamknięcie nowego pliku
        newWB.SaveAs Filename:="C:\Users\USER_NAME\Desktop\test\" & Key & ".xlsx"
        newWB.Close
        ws.AutoFilterMode = False
    Next

End Sub

