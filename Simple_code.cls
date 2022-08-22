VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)

If Target.Row >= 7 And Target.Row <= 20 And Target.Column = 2 Then
Cancel = True

Sheet3.Range("D20") = Target.Value

End If

End Sub


