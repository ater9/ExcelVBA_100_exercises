'The 1st ExcelVBA exercise.
'To copy A1:C5 on Sheet1 and paste them on A1:C5 on Sheet2.
'(Copy all values, formats, but fomulas need to be changed into values.)
'(Comments aren't fromats. Ruby(phonetic guide) is optional.)

Sub copy_paste()
    Worksheets("Sheet1").Range("A1:C5").Copy
    Worksheets("Sheet2").Range("A1").PasteSpecial
    Paste:=xlPasteFormats 'Formats

    Worksheets("Sheet2").Range("A1").PasteSpecial
    Paste:=xlPasteValues 'Values

    Application.CutCopyMode = False
End sub