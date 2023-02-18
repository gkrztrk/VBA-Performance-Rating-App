VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Calendary 
   Caption         =   "Filter"
   ClientHeight    =   870
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5325
   OleObjectBlob   =   "Calendary.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Calendary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmbDateStart_Change()

End Sub

Private Sub cmdDateEnd_Change()

End Sub

Private Sub Label56_Click()

End Sub

Private Sub UserForm_Initialize()
'    Call removeTudo(Me)
    Set dpFrom = New DateTimePicker
    With dpFrom
        .Add Me.cmbDateStart
        .Add Me.cmdDateEnd
        .Create Me, "DD/MM/YYYY" ', _
'            BackColor:=&H492B27, _
'            TitleBack:=RGB(39, 56, 151), _
'            Trailing:=&H80000010, _
'            TitleFore:=&HFFFFFF
    End With
End Sub

Private Sub UserForm_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call moverForm(Me, Me, Button)
End Sub
