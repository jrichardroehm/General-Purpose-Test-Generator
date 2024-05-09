Attribute VB_Name = "Module2"
Sub Macro5()
Attribute Macro5.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro5 Macro
'

'
    Range("F2,F4,F8,F10").Select
    Range("F10").Activate
    Selection.Copy
    Range("K4").Select
    Application.CutCopyMode = False
    Range("J4").Select
End Sub
Sub NewSheet()
Attribute NewSheet.VB_ProcData.VB_Invoke_Func = " \n14"
'
' NewSheet Macro
'

'
    Sheets.Add After:=ActiveSheet
    Sheets("Sheet1").Select
    Sheets("Sheet1").Name = "TempSheet"
    Range("A1").Select
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Columns("A:A").ColumnWidth = 35.86
    Columns("B:B").ColumnWidth = 35.86
End Sub
