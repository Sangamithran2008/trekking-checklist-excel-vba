Attribute VB_Name = "Module2"
Sub extract()

'extract selected cells + header in to a new worksheet

    Dim newWorksheet As Worksheet
    Dim currentSheet As Worksheet
    Dim pickedData As Range

    Set pickedData = Selection

    Set currentSheet = ActiveSheet
    Set newWorksheet = Sheets.Add

    pickedData.Copy
    newWorksheet.Range("B3").PasteSpecial (xlPasteAll)
    
    currentSheet.Range("A4:C15").Copy
    newWorksheet.Range("B2").PasteSpecial (xlPasteAll)

    newWorksheet.Range("B2").PasteSpecial (xlPasteColumnWidths)

End Sub

