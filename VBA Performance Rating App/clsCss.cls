VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCss"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public WithEvents btnLabel          As MSForms.Label
Attribute btnLabel.VB_VarHelpID = -1
Public WithEvents btnImage          As MSForms.Image
Attribute btnImage.VB_VarHelpID = -1
Public WithEvents btnCheckBox       As MSForms.CheckBox
Attribute btnCheckBox.VB_VarHelpID = -1
Public WithEvents btnButton         As MSForms.CommandButton
Attribute btnButton.VB_VarHelpID = -1

Private Sub btnButton_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call Maozinha
End Sub

Private Sub btnCheckBox_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call Maozinha
End Sub

Private Sub btnImage_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call Maozinha
End Sub

Private Sub btnLabel_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call Maozinha
End Sub
