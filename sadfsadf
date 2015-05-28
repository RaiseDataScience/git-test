Sub CustomHighlightRows()

'created by Brian Dobray
'Description:  Highlights rows based on 1 column's alternating values

'Aqua = 42
'Black = 1
'Blue = 5
'Blue -Gray = 47
'Bright Green = 4
'Brown = 53
'Dark Blue = 11
'Dark Green = 51
'Dark Red = 9
'Dark Teal = 49
'Dark Yellow = 12
'Gold = 44
'Gray -25 = 15
'Gray -40 = 48
'Gray -50 = 16
'Gray -80 = 56
'Green = 10
'Indigo = 55
'Lavendar = 39
'Light Blue = 41
'Light Green = 35
'Light Orange = 45
'Light Turqoise = 34
'Light Yellow = 36
'Lime = 43
'Olive Green = 52
'Orange = 46
'Pale Blue = 37
'Pink = 7
'Plum = 54
'Red = 3
'Rose = 38
'Sea Green = 50
'Sky Blue = 33
'Tan = 40
'Teal = 14
'Turqoise = 8
'Violet = 13
'White = 2
'Yellow = 6

  'change colors here
  Const iColor1 As Integer = 34
  Const iColor2 As Integer = 36

  Dim sCurrValue As String
  Dim iMinRow As Long
  Dim iMaxRow As Long
  Dim iLastRow As Long
  Dim bColor As Boolean
  Dim sChangeCol As String
  Dim iStartRow As Long
  Dim sLastCol As String
  Dim sFirstCol As String
  
  Application.ScreenUpdating = False
  'Assign last row
  'Range("A1").Select
  'iLastRow = ActiveCell.SpecialCells(xlLastCell).Row
  iLastRow = Selection.Row + Selection.Rows.Count - 1
  sFirstCol = Split(Cells(1, Selection.Column).Address, "$")(1)
  sLastCol = Split(Cells(1, (Selection.Column + Selection.Columns.Count) - 1).Address, "$")(1)
  
  'prompt for column
  'sFirstCol = InputBox("Enter the first column in the data set.", "Enter Column Info", "A")
  sChangeCol = InputBox("Enter the column with the changing values", "Enter Column Info", "A")
  'sLastCol = InputBox("Enter the last column in the data set.", "Enter last column", "A")
  iStartRow = InputBox("Enter row # where data begins.", "Enter Start Row", 1)
  
  Debug.Print "Change col: " & sChangeCol
  Debug.Print "Last col: " & sLastCol
  Debug.Print "Start row: " & CStr(iStartRow)
    ' check if ascii decimal number corresponds to a lower case alpha character or space
  On Error GoTo ErrHandler
  Range(sChangeCol & CStr(iStartRow)).Select

  
  sCurrValue = ActiveCell.Value
  iMinRow = iStartRow
  iMaxRow = iStartRow
  
  Do While ActiveCell.Row <= iLastRow
    If UCase(ActiveCell.Offset(1, 0).Value) <> UCase(sCurrValue) Then
      iMaxRow = ActiveCell.Row
      'highlight rows
      If bColor Then
        Range(sFirstCol & iMinRow & ":" & sLastCol & iMaxRow).Interior.ColorIndex = iColor1
      Else
        Range(sFirstCol & iMinRow & ":" & sLastCol & iMaxRow).Interior.ColorIndex = iColor2
      End If
      iMinRow = ActiveCell.Row + 1
      iMaxRow = ActiveCell.Row + 1
      bColor = Not bColor
      sCurrValue = ActiveCell.Offset(1, 0).Value
    End If
    ActiveCell.Offset(1, 0).Select

  Loop
  
  If ActiveCell.Row > iLastRow Then
  Exit Sub
  End If
      
  iMaxRow = ActiveCell.Row - 1
  'highlight rows
  If bColor Then
    Range(sFirstCol & iMinRow & ":" & sLastCol & iMaxRow).Interior.ColorIndex = iColor1
  Else
    Range(sFirstCol & iMinRow & ":" & sLastCol & iMaxRow).Interior.ColorIndex = iColor2
  End If
  
  Range("A1").Select
  Application.ScreenUpdating = True

Exit Sub
'error handler here
ErrHandler:
  Error = MsgBox("Invalid entry. Procedure terminating.", vbOKOnly, "Error!")
  Exit Sub

End Sub



