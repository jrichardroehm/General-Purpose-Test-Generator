Attribute VB_Name = "MakeTest3"
Function RandString(n As Long) As String
    'Assumes that Randomize has been invoked by caller
    Dim i As Long, j As Long, m As Long, s As String, pool As String
    pool = "0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ"
    m = Len(pool)
    For i = 1 To n
        j = 1 + Int(m * Rnd())
        s = s & Mid(pool, j, 1)
    Next i
    RandString = s
End Function
Function ErrorHandler()
    MsgBox ("You've entered something funny... Please, only enter numbers, no words or weird characters.")
    MsgBox ("Restarting...")
    GoTo Line1
End Function
Function MakeNewSheet()
    Randomize
    strname = RandString(5)
    Sheets.Add(Before:=Sheets(1)).Name = strname
    Sheets(1).Select
    Application.Union(Range("A:A"), Range("E:E"), Range("G:G")).Select
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlVAlignTop
    End With

    Application.Union(Range("B:B"), Range("F:F"), Range("C:C")).Select
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection.Font
        .Name = "Times New Roman"
    End With
    Range("A:A", "I:I").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Columns("A:A").ColumnWidth = 6
    Columns("B:B").ColumnWidth = 47.86
    Columns("C:C").ColumnWidth = 14.29
    Columns("D:D").ColumnWidth = 12.86
    Columns("E:E").ColumnWidth = 6
    Columns("F:F").ColumnWidth = 47.86
    MakeNewSheet = strname
End Function

Function SimplePrintToPDF(Name)
    SvAs = Application.ThisWorkbook.Path & "\" & Name
    'MsgBox (SvAs)
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:=SvAs, Quality:=xlQualityStandard, _
        IncludeDocProperties:=False, IgnorePrintAreas:=False, OpenAfterPublish:=False
 
End Function

Function FirstCol(sheet)
    Dim first(1) As Integer
    Worksheets(sheet).Activate
    'Gets length of question bank
    r = 1
    c = 1
    While (Cells(r, c).Value = "")
        c = c + 1
    Wend
    While (Cells(r, c).Value = "")
        r = r + 1
    Wend
    first(0) = r
    first(1) = c
    FirstCol = first

End Function


Function FindCol(sheet, Name)
    Worksheets(sheet).Activate
    'Gets length of question bank
    r = 1
    c = 1
    While (Cells(r, c).Value <> Name)
        'MsgBox (Cells(r, c).Value)
        c = c + 1
    Wend
    FindCol = c
End Function
Function IsInArray(numArray, Test, size)
    For i = 0 To size - 1 Step 1
        If numArray(i) = Test Then
            IsInArray = True
            Exit Function
        End If
    Next
    IsInArray = False
End Function

Function GetPriQuestionList(length, QBSize)
    Dim PriList() As Variant
    ReDim PriList(length - 1)
    Col = FindCol(QuestionBank.Name, "Priority")
    Counter = -1
    For i = 2 To QBSize Step 1
        If Cells(i, Col).Value = 1 Then
            Counter = Counter + 1
                If Counter > length - 1 Then
                    MsgBox ("The number of priority questions exceeds the length of the test. Some questions will not be included in the final test.")
                    Exit For
                End If
            Debug.Print (Counter)
            PriList(Counter) = i - 1
        End If
    Next
    GetPriQuestionList = PriList
    Debug.Print (PriList(1))
End Function

Function GetIndexList(length, QBSize, Pri)
    Dim INDEX() As Variant
    ReDim INDEX(length - 1)
    'MsgBox (length & " " & QBSize)
    For Each num In Pri
        If num = "" Then
            Exit For
        End If
        Test = WorksheetFunction.RandBetween(0, length - 1)
        While INDEX(Test) <> ""
            Test = WorksheetFunction.RandBetween(0, length - 1)
        Wend
        INDEX(Test) = num
    Next
    
    
    For i = 0 To length - 1 Step 1
        Test = WorksheetFunction.RandBetween(2, QBSize)
        If IsInArray(INDEX, Test, i) = True Then
            i = i - 1
        ElseIf INDEX(i) = "" Then
            INDEX(i) = Test
        End If
    Next i
    GetIndexList = INDEX
End Function
Function GetSelectionList(NumList, length, Col)
    Dim RetList() As String
    ReDim RetList(length)
    'Dim UDum As Variant
    'ReDim UDum(Length)
    'UDum = Cells(NumList(0), Col).Address
    For i = 0 To length - 1 Step 1
        'MsgBox (NumList(i) + 1 & " " & Col & Length)
        RetList(i) = Cells(NumList(i) + 1, Col).Value
    Next i

    GetSelectionList = RetList
End Function
Function GetQuestionObList(questions, answers, length) As cQuestion
    'MsgBox ("In the function")
    Dim quests() As New cQuestion
    ReDim quests(length - 1)
    'MsgBox ("Created question list")
    For i = 0 To length - 1 Step 1
        quests(i).Set_Answer = answers(i)
        quests(i).Set_Question = questions(i)
    Next i
    For i = 0 To length - 1 Step 1
        MsgBox (quests(i).Answer)
    Next
    GetQuestionsObList = quests
End Function
Function AddQuestions(cQuest, leng)
    Cells(1, 1).Value = "#"
    Cells(1, 2).Value = "Questions"
    Cells(1, 3).Value = "Ref."
    'Next i
    'For i = TEST_LEN To (2 * TEST_LEN) - 1 Step 1
    Cells(1, 5).Value = "#"
    Cells(1, 6).Value = "Answer"
    Cells(1, 7).Value = "Question Bank Number"
    For i = 0 To leng - 1 Step 1
        Cells(i + 2, 1).Value = i + 1
        Cells(i + 2, 2).Value = cQuest(i).Question
        Cells(i + 2, 3).Value = cQuest(i).Refer
    'Next i
    'For i = TEST_LEN To (2 * TEST_LEN) - 1 Step 1
        Cells(i + 2, 5).Value = i + 1
        Cells(i + 2, 6).Value = cQuest(i).Answer
        Cells(i + 2, 7).Value = cQuest(i).Loc
    Next i
End Function

'Created by LTJG Joseph Roehm, feel free to use whatever code you would like!
'Unapologetically leaving my name here

Sub MakeTest3()
'
' makeTest Macro
' Makes the test
'

'
    Application.ScreenUpdating = False
    QBName = QuestionBank.Name
    Worksheets(QBName).Activate
    'Gets length of question bank
    rowCol = FirstCol(QBName)
    
    'MsgBox ("Row and Col are: " & rowCol(0) & " " & rowCol(1))
    r = rowCol(0)
    c = rowCol(1)
    
    TITLE_ROW = r

    Cells(r, c).Select
    
    TOP_ROW = r + 1

    BOTTOM_ROW = Cells(Rows.Count, c).End(xlUp).Row
    QUESTION_TOTAL = BOTTOM_ROW - 1
    
    'MsgBox ("BOTTOM_ROW = " & BOTTOM_ROW)
    QuestionCol = FindCol(QBName, "Question")
    AnswerCol = FindCol(QBName, "Answer")
    RefCol = FindCol(QBName, "Ref")
    'MsgBox ("Question Col is: " & QuestionCol)
    
    'Get number of questions per test from user
    Dim TEST_LEN As Integer
    'On Error GoTo ErrorHandler()
'Line1:
    TEST_LEN = InputBox("Number of Questions")
    While (TEST_LEN > QUESTION_TOTAL) Or (TEST_LEN < 0)
        TEST_LEN = InputBox("Only " & QUESTION_TOTAL & " Available. Enter Again")
    Wend
    
    
    'Create list of Question Objects
    Dim cQuest() As New cQuestion
    ReDim cQuest(TEST_LEN - 1)
    'Get number of unique test versions from user
    Dim NUM_OF_TESTS As Integer
    NUM_OF_TESTS = InputBox("Number of Test Versions")
    While (NUM_OF_TESTS > 15) Or (NUM_OF_TESTS < 0)
        NUM_OF_TESTS = InputBox("Number must be 15 or less, enter again")
    Wend
    
    'Get priority questions from user
    Dim PRIORITY_QUESTIONS() As Variant
    ReDim PRIORITY_QUESTIONS(TEST_LEN - 1)
    PRIORITY_QUESTIONS = GetPriQuestionList(TEST_LEN, QUESTION_TOTAL)
    
    For j = 0 To NUM_OF_TESTS - 1 Step 1
    
        'MsgBox ("In Loop " & j)
        
        'Activate the Question Data Base Sheet
        Worksheets(QBName).Activate
        
        'Get random question numbers
        Dim QUESTION_INDEX() As Variant
        ReDim QUESTION_INDEX(TEST_LEN - 1)
        QUESTION_INDEX = GetIndexList(TEST_LEN, QUESTION_TOTAL, PRIORITY_QUESTIONS)
        'For i = 0 To TEST_LEN - 1 Step 1
        '    MsgBox (QUESTION_INDEX(i))
        'Next
        
        'Get list of references
        Dim SELECTION_R_LIST() As String
        ReDim SELECTION_R_LIST(TEST_LEN)
        SELECTION_R_LIST = GetSelectionList(QUESTION_INDEX, TEST_LEN, RefCol)
        
        'Get list of question strings
        Dim SELECTION_Q_LIST() As String
        ReDim SELECTION_Q_LIST(TEST_LEN)
        SELECTION_Q_LIST = GetSelectionList(QUESTION_INDEX, TEST_LEN, QuestionCol)
        'For i = 0 To TEST_LEN - 1 Step 1
        '    MsgBox (SELECTION_Q_LIST(i))
        'Next
        
        'Get list of answer strings
        Dim SELECTION_A_LIST() As String
        ReDim SELECTION_A_LIST(TEST_LEN)
        SELECTION_A_LIST = GetSelectionList(QUESTION_INDEX, TEST_LEN, AnswerCol)
        'For i = 0 To TEST_LEN - 1 Step 1
        '    MsgBox (SELECTION_A_LIST(i))
        'Next

        
        'Insert info into the Question Objects
        For i = 0 To TEST_LEN - 1 Step 1
            cQuest(i).Set_Answer = SELECTION_A_LIST(i)
            cQuest(i).Set_Question = SELECTION_Q_LIST(i)
            cQuest(i).Set_Loc = QUESTION_INDEX(i)
            cQuest(i).Set_Refer = SELECTION_R_LIST(i)
        Next i
        'For i = 0 To TEST_LEN - 1 Step 1
        '    cQuest(i).Print_Questions
        'Next
        
        'Make a new sheet with proper formatting for questions
        SheetName = MakeNewSheet()
        
        'Add questions to sheet
        none = AddQuestions(cQuest, TEST_LEN)
        
        'Print to PDF
        SimplePrintToPDF (SheetName)

    'Iterate to next unique version
    Next j
    
    Application.ScreenUpdating = True
    
End Sub


