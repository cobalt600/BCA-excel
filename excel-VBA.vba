Sub Sample3()

With Workbooks.Open("H:\BCA\*.xlsx")
   
    .Sheets("End point_1").Range("A4:A8").Copy
    ThisWorkbook.Sheets("Result").Cells(1, 1).PasteSpecial _
    Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
   
    .Sheets("End point_1").Range("B49:M56").Copy
    ThisWorkbook.Sheets("Result").Cells(23, 2).PasteSpecial _
    Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
   
    .Sheets("End point_1").Shapes("Picture 1").Copy
    ThisWorkbook.Worksheets("Result").Activate
    ThisWorkbook.Sheets("Result").Range("B7").Select
    ActiveSheet.Paste
    Selection.ShapeRange.LockAspectRatio = msoTrue
    Selection.ShapeRange.Height = 160
   
    .Sheets("Linear regression fit").Range("A15:C18").Copy
    ThisWorkbook.Sheets("Result").Cells(2, 9).PasteSpecial _
    Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=True, Transpose:=False
    .Close SaveChanges:=False
   
   
End With
   
    Dim wb1 As Workbook
    Dim wb2 As Workbook
   
    ThisWorkbook.Activate
    Set wb1 = ActiveWorkbook
   
    Workbooks.Open ("H:\BCA\*.xlsx")
    Set wb2 = ActiveWorkbook
   
    wb1.Sheets("Result").Copy Before:=wb2.Sheets(1)
    wb1.Close SaveChanges:=False
   
   
End Sub
