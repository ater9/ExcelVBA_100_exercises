'The 1st ExcelVBA exercise.
'To copy A1:C5 on Sheet1 and paste them on A1:C5 on Sheet2.
'(Copy all values, formats, and fomulas.)
'(Do not use "select" method.)

Sub copy_paste()
    Worksheets("Sheet1").Range("A1:C5").Copy
    Worksheets("Sheet2").Range("A1").PasteSpecial
    Paste:=xlPasteAll
    Application.CutCopyMode = False
End Sub