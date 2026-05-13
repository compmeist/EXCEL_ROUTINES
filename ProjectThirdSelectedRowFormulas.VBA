'
'   Select Three Rows, where you want the formulas to copy linearly and porportionally to the third row
'
'     =A3
'     =A7     
' ( will set to =A11)
'
'

Option Explicit

Public Sub ProjectThirdSelectedRowFormulas()
    Dim sel As Range
    Dim r1 As Long, r2 As Long, r3 As Long
    Dim firstCol As Long, lastCol As Long
    Dim c As Long
    Dim f1 As String, f2 As String
    Dim projected As String
    Dim ws As Worksheet
    Dim usedRng As Range

    If TypeName(Selection) <> "Range" Then Exit Sub
    If Selection.Areas.Count > 1 Then
        MsgBox "Please select one contiguous 3-row range.", vbExclamation
        Exit Sub
    End If

    Set sel = Selection
    Set ws = sel.Worksheet

    If sel.Rows.Count <> 3 Then
        MsgBox "Please select exactly three rows, or a range that is exactly three rows high.", vbExclamation
        Exit Sub
    End If

    r1 = sel.Rows(1).Row
    r2 = sel.Rows(2).Row
    r3 = sel.Rows(3).Row

    'If whole rows are selected, limit work to UsedRange columns.
    Set usedRng = ws.UsedRange

    If sel.Columns.Count = ws.Columns.Count Then
        firstCol = usedRng.Column
        lastCol = usedRng.Column + usedRng.Columns.Count - 1
    Else
        firstCol = sel.Column
        lastCol = sel.Column + sel.Columns.Count - 1
    End If

    Application.ScreenUpdating = False

    For c = firstCol To lastCol
        If ws.Cells(r1, c).HasFormula And ws.Cells(r2, c).HasFormula Then
            f1 = ws.Cells(r1, c).Formula
            f2 = ws.Cells(r2, c).Formula

            projected = ProjectFormulaFromTwoExamples(f1, f2)

            If Len(projected) > 0 Then
                ws.Cells(r3, c).Formula = projected
            End If
        End If
    Next c

    Application.ScreenUpdating = True
End Sub

Private Function ProjectFormulaFromTwoExamples(ByVal formula1 As String, ByVal formula2 As String) As String
    Dim re As Object
    Dim m1 As Object, m2 As Object
    Dim i As Long
    Dim result As String
    Dim ref1 As String, ref2 As String, ref3 As String
    Dim rA As Long, cA As Long
    Dim rB As Long, cB As Long
    Dim dRow As Long, dCol As Long

    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.IgnoreCase = False

    'Matches normal A1-style references such as N2, $N2, N$2, $N$2.
    re.Pattern = "\$?[A-Z]{1,3}\$?\d+"

    Set m1 = re.Execute(UCase$(formula1))
    Set m2 = re.Execute(UCase$(formula2))

    If m1.Count = 0 Or m1.Count <> m2.Count Then
        ProjectFormulaFromTwoExamples = vbNullString
        Exit Function
    End If

    result = formula2

    'Work backwards so character positions remain valid.
    For i = m2.Count - 1 To 0 Step -1
        ref1 = m1(i).Value
        ref2 = m2(i).Value

        ParseCellRef ref1, rA, cA
        ParseCellRef ref2, rB, cB

        dRow = rB - rA
        dCol = cB - cA

        ref3 = BuildCellRefLike(ref2, rB + dRow, cB + dCol)

        result = Left$(result, m2(i).FirstIndex) & _
                 ref3 & _
                 Mid$(result, m2(i).FirstIndex + m2(i).Length + 1)
    Next i

    ProjectFormulaFromTwoExamples = result
End Function

Private Sub ParseCellRef(ByVal refText As String, ByRef rowNum As Long, ByRef colNum As Long)
    Dim s As String
    Dim letters As String
    Dim numbers As String
    Dim i As Long
    Dim ch As String

    s = Replace$(UCase$(refText), "$", "")

    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)

        If ch Like "[A-Z]" Then
            letters = letters & ch
        ElseIf ch Like "[0-9]" Then
            numbers = numbers & ch
        End If
    Next i

    rowNum = CLng(numbers)
    colNum = ColumnLetterToNumber(letters)
End Sub

Private Function BuildCellRefLike(ByVal templateRef As String, ByVal rowNum As Long, ByVal colNum As Long) As String
    Dim hasAbsCol As Boolean
    Dim hasAbsRow As Boolean
    Dim i As Long
    Dim dollarPos As Long

    'Preserve $ style from the second formula's reference.
    hasAbsCol = Left$(templateRef, 1) = "$"

    dollarPos = InStrRev(templateRef, "$")
    hasAbsRow = dollarPos > 1 And Mid$(templateRef, dollarPos + 1, 1) Like "[0-9]"

    BuildCellRefLike = IIf(hasAbsCol, "$", "") & _
                       ColumnNumberToLetter(colNum) & _
                       IIf(hasAbsRow, "$", "") & _
                       CStr(rowNum)
End Function

Private Function ColumnLetterToNumber(ByVal colLetters As String) As Long
    Dim i As Long
    Dim n As Long

    For i = 1 To Len(colLetters)
        n = n * 26 + Asc(Mid$(colLetters, i, 1)) - Asc("A") + 1
    Next i

    ColumnLetterToNumber = n
End Function

Private Function ColumnNumberToLetter(ByVal colNum As Long) As String
    Dim s As String
    Dim n As Long

    n = colNum

    Do While n > 0
        s = Chr$(((n - 1) Mod 26) + Asc("A")) & s
        n = (n - 1) \ 26
    Loop

    ColumnNumberToLetter = s
End Function

