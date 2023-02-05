Attribute VB_Name = "Module1"
Option Explicit


'
Sub Hello()
  MsgBox "Hello World"
End Sub


Sub Goodbye()
  MsgBox "Good Bye!"
End Sub

sub copy()
  Set atai = 1
  Worksheets("Sheet2").Range("A1:D5").Activate
  Worksheets("Sheet2").Range("A1:D5").Copy
  Worksheets("Sheet2").Range("A1:D5").PasteSpecial Paste:=xlPasteAll
End sub








