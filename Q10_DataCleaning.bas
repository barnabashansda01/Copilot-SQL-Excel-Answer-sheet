' ================================================================
' Question 10
' Prompt given to Excel Copilot:
' "Generate VBA code for cleaning data (removing duplicates,
'  filling null values, format table)."
' ================================================================
' Usage: Open the workbook -> Alt+F11 -> Insert Module -> paste
' this code -> run CleanData()

Sub CleanData()

    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim lastRow As Long, lastCol As Long
    Dim rng As Range
    Dim c As Range
    Dim colIdx As Long

    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    Set rng = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))

    Application.ScreenUpdating = False

    ' --------------------------------------------------------
    ' Step 1: Fill blank/null cells
    '   - Text columns  -> "Unknown"
    '   - Numeric columns -> column median
    ' --------------------------------------------------------
    Dim r As Long, colValues As Object
    For colIdx = 1 To lastCol
        Dim isNumeric As Boolean
        isNumeric = True
        Dim vals() As Double
        Dim n As Long
        n = 0
        ReDim vals(1 To lastRow)

        For r = 2 To lastRow
            Dim v As Variant
            v = ws.Cells(r, colIdx).Value
            If Not IsEmpty(v) And v <> "" Then
                If Not IsNumeric(v) Then
                    isNumeric = False
                Else
                    n = n + 1
                    vals(n) = CDbl(v)
                End If
            End If
        Next r

        Dim fillValue As Variant
        If isNumeric And n > 0 Then
            fillValue = Median(vals, n)
        Else
            fillValue = "Unknown"
        End If

        For r = 2 To lastRow
            If Trim(ws.Cells(r, colIdx).Value & "") = "" Then
                ws.Cells(r, colIdx).Value = fillValue
            End If
        Next r
    Next colIdx

    ' --------------------------------------------------------
    ' Step 2: Remove duplicate rows (based on all columns)
    ' --------------------------------------------------------
    Dim colArr() As Long
    ReDim colArr(1 To lastCol)
    For colIdx = 1 To lastCol
        colArr(colIdx) = colIdx
    Next colIdx
    rng.RemoveDuplicates Columns:=colArr, Header:=xlYes

    ' Refresh lastRow after de-duplication
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Set rng = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))

    ' --------------------------------------------------------
    ' Step 3: Format as a proper Excel Table
    ' --------------------------------------------------------
    On Error Resume Next
    ws.ListObjects("CleanedTable").Unlist
    On Error GoTo 0

    Set tbl = ws.ListObjects.Add(xlSrcRange, rng, , xlYes)
    tbl.Name = "CleanedTable"
    tbl.TableStyle = "TableStyleMedium9"

    ' Autofit columns for readability
    rng.Columns.AutoFit

    Application.ScreenUpdating = True

    MsgBox "Data cleaning complete: duplicates removed, blanks filled," & _
           " and range formatted as table 'CleanedTable'.", vbInformation

End Sub

' Helper function: median of an array
Function Median(arr() As Double, n As Long) As Double
    Dim i As Long, j As Long, temp As Double
    ' simple bubble sort (fine for typical column sizes)
    For i = 1 To n - 1
        For j = 1 To n - i
            If arr(j) > arr(j + 1) Then
                temp = arr(j)
                arr(j) = arr(j + 1)
                arr(j + 1) = temp
            End If
        Next j
    Next i

    If n Mod 2 = 0 Then
        Median = (arr(n \ 2) + arr(n \ 2 + 1)) / 2
    Else
        Median = arr((n + 1) \ 2)
    End If
End Function
