' Final macro to be used: Run_Shading_Blocks


Sub ShadeThreeLowestUnderOne_Dynamic()
    Dim rng As Range, area As Range
    Set rng = Selection
    If rng Is Nothing Then Exit Sub
    If rng.Areas.Count = 0 Then Exit Sub

    Dim sep As String
    sep = Application.International(xlListSeparator)

    ' Build VSTACK(arg1,arg2,...) over each selected Area
    Dim vs As String
    vs = "VSTACK(" & BuildAreaListForVStack(rng, sep) & ")"

    ' b = numeric values < 1 from stacked array
    ' threshold = 3rd smallest if at least 3 values, otherwise MAX(b)
    ' return FALSE when there are no qualifying values
    Dim thresholdExpr As String
    thresholdExpr = "LET(a," & vs & _
                    sep & "b,FILTER(a,ISNUMBER(a)*(a<1))" & _
                    sep & "IF(ROWS(b)=0" & sep & "NA()" & _
                    sep & "IF(ROWS(b)>=3" & sep & "SMALL(b,3)" & sep & "MAX(b)))" & ")"

    ' Apply CF per area (needed for noncontiguous selections)
    For Each area In rng.Areas
        area.FormatConditions.Delete

        Dim cellRef As String
        cellRef = area.Cells(1, 1).Address(RowAbsolute:=False, ColumnAbsolute:=False)

        Dim fmla As String
        fmla = "=IFERROR(AND(" & _
               "ISNUMBER(" & cellRef & ")" & sep & _
               cellRef & "<1" & sep & _
               cellRef & "<=" & thresholdExpr & _
               ")" & sep & "FALSE)"

        Dim fc As FormatCondition
        Set fc = area.FormatConditions.Add(Type:=xlExpression, Formula1:=fmla)
        fc.Interior.color = RGB(217, 217, 217)
        fc.ModifyAppliesToRange area
    Next area
End Sub

Sub ShadeThreeLowestUnderZero_Dynamic()
    Dim rng As Range, area As Range
    Set rng = Selection
    If rng Is Nothing Then Exit Sub
    If rng.Areas.Count = 0 Then Exit Sub

    Dim sep As String
    sep = Application.International(xlListSeparator)

    ' Build VSTACK(arg1,arg2,...) over each selected Area
    Dim vs As String
    vs = "VSTACK(" & BuildAreaListForVStack(rng, sep) & ")"

    ' b = numeric values < 0 from stacked array
    ' threshold = 3rd smallest if at least 3 values, otherwise MAX(b)
    ' return FALSE when there are no qualifying values
    Dim thresholdExpr As String
    thresholdExpr = "LET(a," & vs & _
                    sep & "b,FILTER(a,ISNUMBER(a)*(a<0))" & _
                    sep & "IF(ROWS(b)=0" & sep & "NA()" & _
                    sep & "IF(ROWS(b)>=3" & sep & "SMALL(b,3)" & sep & "MAX(b)))" & ")"

    ' Apply CF per area (needed for noncontiguous selections)
    For Each area In rng.Areas
        area.FormatConditions.Delete

        Dim cellRef As String
        cellRef = area.Cells(1, 1).Address(RowAbsolute:=False, ColumnAbsolute:=False)

        Dim fmla As String
        fmla = "=IFERROR(AND(" & _
               "ISNUMBER(" & cellRef & ")" & sep & _
               cellRef & "<1" & sep & _
               cellRef & "<=" & thresholdExpr & _
               ")" & sep & "FALSE)"

        Dim fc As FormatCondition
        Set fc = area.FormatConditions.Add(Type:=xlExpression, Formula1:=fmla)
        fc.Interior.color = RGB(217, 217, 217)
        fc.ModifyAppliesToRange area
    Next area
End Sub

Private Function BuildAreaListForVStack(ByVal rng As Range, ByVal sep As String) As String
    ' Returns: Sheet1!$A$1:$A$3,Sheet1!$C$1:$C$3 (with correct list separator)
    Dim area As Range
    Dim s As String: s = ""

    For Each area In rng.Areas
        If Len(s) > 0 Then s = s & sep
        s = s & area.Parent.Name & "!" & area.Address(True, True)
    Next area

    BuildAreaListForVStack = s
End Function


Public Sub Run_Shading_Blocks()
    Dim ws As Worksheet
    Set ws = ActiveSheet   ' change if you want a specific sheet, e.g. ThisWorkbook.Worksheets("Sheet1")

    Dim r As Long
    Dim rowsToProcess As Variant
    rowsToProcess = Array(Array(6, 32), Array(40, 66))

    Dim block As Variant
    For Each block In rowsToProcess
        For r = CLng(block(0)) To CLng(block(1))
            ' Under zero group: D,G,L,P,T,Y,AC,AG,AK,AO (same row)
            ws.Range("D" & r & ",H" & r & ",L" & r & ",P" & r & ",T" & r & _
                     ",Y" & r & ",AC" & r & ",AG" & r & ",AK" & r & ",AO" & r).Select
            Application.Run "ShadeThreeLowestUnderZero_Dynamic"

            ' Under one group: E,I,M,Q,U,Z,AD,AH,AL,AP (same row)
            ws.Range("E" & r & ",I" & r & ",M" & r & ",Q" & r & ",U" & r & _
                     ",Z" & r & ",AD" & r & ",AH" & r & ",AL" & r & ",AP" & r).Select
            Application.Run "ShadeThreeLowestUnderOne_Dynamic"
        Next r
    Next block

End Sub
