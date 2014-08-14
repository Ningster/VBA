VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub clean()
Dim n As Integer
For Each ws In ActiveWorkbook.Worksheets
  ws.AutoFilterMode = False
  ws.Range("A:AD").FormatConditions.Delete
Next ws
'remove all rules in the worksheet'

For n = 2 To Range("A1").End(xlDown).Row
  If Range("B" & n).Value = "Paint/Process" Or Range("B" & n).Font.Strikethrough = True Then
     Rows(n).Delete
     n = n - 1
  End If
Next n
'remove "Paint/Process" & "strikethrough" DPN'

Range("E:K").Replace ".*", "", xlPart
'remove string after "." '

Range("E:K").NumberFormat = "@"
'make sure the format is Text'

Dim xR As Range, TT
For Each xR In [E:K].SpecialCells(2)
   TT = Application.Trim(xR)
   xR = Replace(TT, " " & Chr(10), Chr(10))
 Next
'remove all blank after P/N'

Columns("E:K").EntireColumn.AutoFit

End Sub

Sub colorOrphanParts()

Dim i, j, p, q As Integer
Dim myArray(6) As String
myArray(0) = "F"
myArray(1) = "G"
myArray(2) = "H"
myArray(3) = "I"
myArray(4) = "J"
myArray(5) = "K"


For j = 0 To 5
  For i = 2 To Range("A1").End(xlDown).Row
      If Len(Range(myArray(j) & i).Value) = 5 Then
         p = -1
         Do Until Range(myArray(j) & i).Offset(p, -1).Value <> ""
           p = p - 1
         Loop
         If Range(myArray(j) & i).Offset(p, -1).Row <> 1 And (Len(Range(myArray(j) & i).Offset(p, -1).Value) > 5) Then
           Range(myArray(j) & i).Interior.Color = RGB(200, 160, 35)
         End If
      End If
  Next i
Next j
'If the Range report to an non-DPN, fill out with color(Orphan parts)'


For q = 0 To 5
  For i = 2 To Range("A1").End(xlDown).Row
      If Len(Range(myArray(q) & i).Value) = 5 Then
         p = -1
         Do Until Range(myArray(q) & i).Offset(p, -1).Value <> ""
           p = p - 1
         Loop
         If Range(myArray(q) & i).Offset(p, -1).Interior.Color = RGB(200, 160, 35) And (Len(Range(myArray(q) & i).Offset(p, -1).Value) = 5) Then
           Range(myArray(q) & i).Interior.Color = RGB(200, 160, 35)
         End If
      End If
  Next i
Next q
'If the Range report to an colored DPN(Orphan parts), fill out with color'

Dim A As Boolean
A = False

For q = 0 To 5
  For i = 2 To Range("A1").End(xlDown).Row
         If Range(myArray(q) & i).Interior.Color = RGB(200, 160, 35) Then
             A = True
         End If
  Next i
Next q

If A = Flase Then
MsgBox ("No orphan parts have found.")
End If







End Sub

Sub addOrphanSheet()

Dim lastrowOrphans As Long
Dim ws As Worksheet
Dim i, j, k As Integer
Dim sheetName As String
Dim myArray(6) As String
myArray(0) = "F"
myArray(1) = "G"
myArray(2) = "H"
myArray(3) = "I"
myArray(4) = "J"
myArray(5) = "K"


Set bk = Application.ActiveWorkbook
'if you don't "set", you will get the error(object not found) if your add-in is the only workbook loaded. In that case there is no active workbook and your code is failing on the line.'

sheetName = bk.ActiveSheet.Name
'Add a new worksheet named Orphans'

Dim B As Boolean
B = True
For Each ws In ActiveWorkbook.Worksheets
  If ws.Name = sheetName + "-Orphan" Then
    B = False
  End If
Next ws
'Check if there's duplicate worksheets.'

If B = False Then
MsgBox ("The orphan sheet already existed.")   'Error message if there's an existed one'
ElseIf Len(sheetName) > 24 Then
MsgBox ("Sheet name is too long")   'Error message if sheet name is too long'
Else
bk.Sheets.Add.Name = sheetName + "-Orphan"

For j = 0 To 5
   For i = 2 To bk.Worksheets(sheetName).Range("A1").End(xlDown).Row
      If bk.Worksheets(sheetName).Range(myArray(j) & i).Interior.Color = RGB(200, 160, 35) Then
        With bk.Sheets(sheetName + "-Orphan")
          k = .Cells(.Rows.Count, "A").End(xlUp).Row + 1
        End With
        
        bk.Worksheets(sheetName).Range(myArray(j) & i).Copy
        bk.Sheets(sheetName + "-Orphan").Range("A" & k).PasteSpecial xlPasteValues
        
        bk.Worksheets(sheetName).Range("L" & (Range(myArray(j) & i).Row)).Copy
        bk.Sheets(sheetName + "-Orphan").Range("B" & k).PasteSpecial xlPasteValues
        
        bk.Worksheets(sheetName).Range("O" & (Range(myArray(j) & i).Row)).Copy
        bk.Sheets(sheetName + "-Orphan").Range("C" & k).PasteSpecial xlPasteValues
      End If
   Next i
Next j

bk.Sheets(sheetName + "-Orphan").Range("A1").Value = "Orphan DPN"
bk.Sheets(sheetName + "-Orphan").Range("B1").Value = "Rev"
bk.Sheets(sheetName + "-Orphan").Range("C1").Value = "Agile Description"
'Copy & Paste all colored parts to the worksheet'

Columns("A:C").EntireColumn.AutoFit

Range("A:C").NumberFormat = "@"
'make sure the format is Text'

bk.Sheets(sheetName + "-Orphan").Range("G2").Formula = "=BOMChain(A2:A100)"
End If
'Create an orphan sheet if there's no duplicate one'

End Sub


Sub eliminateOrpanAndNonDPN()
Dim i, j As Integer

For i = 3 To Range("A2").End(xlDown).Row

  If Range("E" & i).Value <> "" Then        'if E has value, check is DPN or non-DPN'
     If Len(Range("E" & i)) > 5 Or Range("E" & i).Interior.Color = RGB(200, 160, 35) Then
     Rows(i).Delete
     i = i - 1
     End If
  Else   'if E doesn't have value, check to right'
     j = Range("E" & i).End(xlToRight).Column
     If Len(Cells(i, j)) > 5 Or Cells(i, j).Interior.Color = RGB(200, 160, 35) Then
     Rows(i).Delete
     i = i - 1
     End If
  End If

Next i

Columns("E:K").EntireColumn.AutoFit

End Sub



