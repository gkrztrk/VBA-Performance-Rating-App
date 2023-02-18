VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public WithEvents Cards As MSForms.frame
Attribute Cards.VB_VarHelpID = -1
Public WithEvents CardBody As MSForms.frame
Attribute CardBody.VB_VarHelpID = -1
Public WithEvents Menus As MSForms.frame
Attribute Menus.VB_VarHelpID = -1
Public WithEvents ContainerMenus As MSForms.frame
Attribute ContainerMenus.VB_VarHelpID = -1

Private Sub ContainerMenus_Click()
    MsgBox "Menu"
End Sub

Private Sub ContainerMenus_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call moverForm(frmPerformance, ContainerMenus, Button)
End Sub

Private Sub ContainerMenus_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim ctrl As Control
    
    For Each ctrl In ContainerMenus.Controls
        If TypeName(ctrl) = "Frame" Then
            With ctrl
                .BorderColor = &H492B27
                .BackColor = &H492B27
            End With
        End If
    Next
    
End Sub

Private Sub Menus_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call Maozinha
    With Menus
        .BackColor = &H3A1F1A
        .BorderColor = &H3A1F1A
    End With
    
End Sub
