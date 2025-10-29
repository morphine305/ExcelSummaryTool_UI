Attribute VB_Name = "Module1"
Sub BatchTranspose_AB_to_OneColumn_Fix()
    Dim fDialog As FileDialog
    Dim filePath As Variant
    Dim wb As Workbook, ws As Worksheet
    Dim arrA As Variant, arrB As Variant
    Dim tmpA As Variant, tmpB As Variant
    Dim outArr() As Variant
    Dim i As Long, n As Long, lastA As Long, lastB As Long, lastRow As Long

    Set fDialog = Application.FileDialog(msoFileDialogFilePicker)
    fDialog.AllowMultiSelect = True
    fDialog.Filters.Add "Excel Files", "*.xlsx;*.xlsm", 1
    If fDialog.Show <> -1 Then Exit Sub

    For Each filePath In fDialog.SelectedItems
        Set wb = Workbooks.Open(filePath)
        Set ws = wb.Sheets(2) ' 第二分頁

        lastA = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        lastB = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
        lastRow = Application.WorksheetFunction.Max(lastA, lastB)

        ' 如果 A、B 都是空白，跳過
        If lastRow = 1 And IsEmpty(ws.Range("A1")) And IsEmpty(ws.Range("B1")) Then
            wb.Close SaveChanges:=False
            GoTo NextFile
        End If

        ' 讀入範圍（即使只有一格也要轉成二維陣列）
        arrA = ws.Range("A1:A" & lastRow).Value
        arrB = ws.Range("B1:B" & lastRow).Value

        If Not IsArray(arrA) Then
            ReDim tmpA(1 To 1, 1 To 1)
            tmpA(1, 1) = arrA
            arrA = tmpA
        End If
        If Not IsArray(arrB) Then
            ReDim tmpB(1 To 1, 1 To 1)
            tmpB(1, 1) = arrB
            arrB = tmpB
        End If

        ReDim outArr(1 To lastRow * 2, 1 To 1)
        n = 1
        For i = 1 To lastRow
            outArr(n, 1) = arrA(i, 1)
            n = n + 1
            outArr(n, 1) = arrB(i, 1)
            n = n + 1
        Next i

        ws.Cells.ClearContents
        ws.Range("A1").Resize(UBound(outArr, 1), 1).Value = outArr

        wb.Close SaveChanges:=True
NextFile:
    Next filePath
End Sub

