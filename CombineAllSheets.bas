
Sub CombineAllSheets()

    Dim ws As Worksheet
    Dim dest As Worksheet
    Dim lastRow As Long
    Dim lastCol As Long
    Dim destRow As Long
    Dim headerCopied As Boolean
    Dim lastCell As Range

    Application.ScreenUpdating = False

    'Create or find Combined sheet
    On Error Resume Next
    Set dest = ThisWorkbook.Worksheets("Combined")
    On Error GoTo 0

    If dest Is Nothing Then
        Set dest = ThisWorkbook.Worksheets.Add( _
            After:=ThisWorkbook.Worksheets( _
            ThisWorkbook.Worksheets.Count))
        dest.Name = "Combined"
    End If

    'Clear previous combined data
    dest.Cells.ClearContents

    destRow = 1
    headerCopied = False

    'Loop through all worksheets
    For Each ws In ThisWorkbook.Worksheets

        If ws.Name <> dest.Name Then

            Set lastCell = ws.Cells.Find( _
                What:="*", _
                After:=ws.Cells(1, 1), _
                LookIn:=xlFormulas, _
                LookAt:=xlPart, _
                SearchOrder:=xlByRows, _
                SearchDirection:=xlPrevious)

            If Not lastCell Is Nothing Then

                lastRow = lastCell.Row

                Set lastCell = ws.Cells.Find( _
                    What:="*", _
                    After:=ws.Cells(1, 1), _
                    LookIn:=xlFormulas, _
                    LookAt:=xlPart, _
                    SearchOrder:=xlByColumns, _
                    SearchDirection:=xlPrevious)

                lastCol = lastCell.Column

                'Copy header only once
                If headerCopied = False Then

                    dest.Cells(1, 1).Resize(1, lastCol).Value = _
                        ws.Cells(1, 1).Resize(1, lastCol).Value

                    destRow = 2
                    headerCopied = True

                End If

                'Copy data excluding header
                If lastRow >= 2 Then

                    dest.Cells(destRow, 1).Resize( _
                        lastRow - 1, lastCol).Value = _
                        ws.Cells(2, 1).Resize( _
                        lastRow - 1, lastCol).Value

                    destRow = destRow + lastRow - 1

                End If

            End If

        End If

    Next ws

    Application.ScreenUpdating = True

    MsgBox "All sheets combined successfully!", vbInformation

End Sub
