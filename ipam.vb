Option Explicit

Private Sub Label1_Click()

End Sub

Private Sub txtIP_Change()

End Sub

Private Sub txtLoc_Change()

End Sub

Private Sub UserForm_Initialize()

    With lstRecords
        .ColumnCount = 7
        .ColumnHeads = False
        .ColumnWidths = "40;150;400;180;100;165;60"
    End With

    LoadRecords

End Sub

Private Sub LoadRecords()

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long

    Set ws = ThisWorkbook.Worksheets("IPData")

    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    lstRecords.Clear

    If lastRow < 2 Then Exit Sub

    For r = 2 To lastRow

        lstRecords.AddItem r

        lstRecords.List(lstRecords.ListCount - 1, 1) = ws.Cells(r, 1).Value
        lstRecords.List(lstRecords.ListCount - 1, 2) = ws.Cells(r, 2).Value
        lstRecords.List(lstRecords.ListCount - 1, 3) = ws.Cells(r, 3).Value
        lstRecords.List(lstRecords.ListCount - 1, 4) = ws.Cells(r, 4).Value
        lstRecords.List(lstRecords.ListCount - 1, 5) = ws.Cells(r, 5).Value
        lstRecords.List(lstRecords.ListCount - 1, 6) = ws.Cells(r, 6).Value

    Next r

End Sub

Private Sub SearchRecords()

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim searchText As String
    Dim r As Long
    Dim c As Long
    Dim matchFound As Boolean

    Set ws = ThisWorkbook.Worksheets("IPData")

    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    lstRecords.Clear

    searchText = LCase(Trim(txtSearch.Text))

    If searchText = "" Then

        LoadRecords
        Exit Sub

    End If

    For r = 1 To lastRow

        matchFound = False

        For c = 1 To 6

            If InStr(1, LCase(CStr(ws.Cells(r, c).Value)), searchText) > 0 Then

                matchFound = True
                Exit For

            End If

        Next c

        If matchFound Then

            lstRecords.AddItem r

            lstRecords.List(lstRecords.ListCount - 1, 1) = ws.Cells(r, 1).Value
            lstRecords.List(lstRecords.ListCount - 1, 2) = ws.Cells(r, 2).Value
            lstRecords.List(lstRecords.ListCount - 1, 3) = ws.Cells(r, 3).Value
            lstRecords.List(lstRecords.ListCount - 1, 4) = ws.Cells(r, 4).Value
            lstRecords.List(lstRecords.ListCount - 1, 5) = ws.Cells(r, 5).Value
            lstRecords.List(lstRecords.ListCount - 1, 6) = ws.Cells(r, 6).Value

        End If

    Next r

End Sub

Private Sub txtSearch_Change()

    SearchRecords

End Sub

Private Sub ClearInputs()

    txtIP.Text = ""
    txtBIN.Text = ""
    txtHost.Text = ""
    txtDesc.Text = ""
    txtLoc.Text = ""
    txtDup.Text = ""

    lstRecords.ListIndex = -1

End Sub

Private Sub cmdAdd_Click()

    Dim ws As Worksheet
    Dim nextRow As Long

    Set ws = ThisWorkbook.Worksheets("IPData")

    nextRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row + 1

    ws.Cells(nextRow, 1).Value = txtIP.Text
    ws.Cells(nextRow, 2).Value = txtBIN.Text
    ws.Cells(nextRow, 3).Value = txtHost.Text
    ws.Cells(nextRow, 4).Value = txtDesc.Text
    ws.Cells(nextRow, 5).Value = txtLoc.Text
    ws.Cells(nextRow, 6).Value = txtDup.Text

    LoadRecords
    ClearInputs

End Sub

Private Sub cmdUpdate_Click()

    Dim ws As Worksheet
    Dim rowToUpdate As Long

    If lstRecords.ListIndex = -1 Then

        MsgBox "Select a record to update.", vbInformation
        Exit Sub

    End If

    Set ws = ThisWorkbook.Worksheets("IPData")

    rowToUpdate = CLng(lstRecords.List(lstRecords.ListIndex, 0))

    ws.Cells(rowToUpdate, 1).Value = txtIP.Text
    ws.Cells(rowToUpdate, 2).Value = txtBIN.Text
    ws.Cells(rowToUpdate, 3).Value = txtHost.Text
    ws.Cells(rowToUpdate, 4).Value = txtDesc.Text
    ws.Cells(rowToUpdate, 5).Value = txtLoc.Text
    ws.Cells(rowToUpdate, 6).Value = txtDup.Text

    LoadRecords

End Sub

Private Sub cmdDelete_Click()

    Dim ws As Worksheet
    Dim rowToDelete As Long

    If lstRecords.ListIndex = -1 Then

        MsgBox "Select a record to delete.", vbInformation
        Exit Sub

    End If

    If MsgBox("Delete this record?", vbYesNo + vbQuestion) = vbNo Then Exit Sub

    Set ws = ThisWorkbook.Worksheets("IPData")

    rowToDelete = CLng(lstRecords.List(lstRecords.ListIndex, 0))

    ws.Rows(rowToDelete).Delete

    LoadRecords
    ClearInputs

End Sub

Private Sub cmdClear_Click()

    ClearInputs

End Sub

Private Sub cmdClose_Click()

    Unload Me

End Sub

Private Sub lstRecords_Click()

    If lstRecords.ListIndex = -1 Then Exit Sub

    txtIP.Text = lstRecords.List(lstRecords.ListIndex, 1)
    txtBIN.Text = lstRecords.List(lstRecords.ListIndex, 2)
    txtHost.Text = lstRecords.List(lstRecords.ListIndex, 3)
    txtDesc.Text = lstRecords.List(lstRecords.ListIndex, 4)
    txtLoc.Text = lstRecords.List(lstRecords.ListIndex, 5)
    txtDup.Text = lstRecords.List(lstRecords.ListIndex, 6)

End Sub
