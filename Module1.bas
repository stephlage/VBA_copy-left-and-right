Attribute VB_Name = "Module1"

Sub KeyPressLCopyLeft()
' USE KEY PRESS L
' Macro1 Macro
'

'
    
    'Application.CutCopyMode = False
    Selection.Copy
    ActiveCell.Offset(0, 1).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
     ActiveCell.Offset(1, 0).Select
Do Until ActiveCell.Height <> 0
 ActiveCell.Offset(1, 0).Select
Loop
     ActiveCell.Offset(0, -1).Select
End Sub


Sub KeyPressRCopyRight()
' USE KEY PRESS R
' Macro1 Macro
'

'
    
    'Application.CutCopyMode = False
    Selection.Copy
    ActiveCell.Offset(0, -1).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
     ActiveCell.Offset(1, 0).Select
Do Until ActiveCell.Height <> 0
 ActiveCell.Offset(1, 0).Select
Loop
     ActiveCell.Offset(0, 1).Select
End Sub

