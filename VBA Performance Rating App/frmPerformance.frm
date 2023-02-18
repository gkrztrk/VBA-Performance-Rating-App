VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmPerformance 
   Caption         =   " "
   ClientHeight    =   10710
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   18465
   OleObjectBlob   =   "frmPerformance.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmPerformance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub btnFilter_Click()
    Calendary.Show
End Sub

Private Sub CardBody_Click()

End Sub

Private Sub CommandButton1_Click()



frmPerformance.Hide

Grading.Show
End Sub

Private Sub Container_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call moverForm(Me, Container, Button)
End Sub

Private Sub btnMenu_Click()
    Dim i As Long
'    On Error Resume Next
    If Me.sidebar.Width = 186 Then
        DoEvents
        i = 186
        Do Until i = 60
            Sleep 0.000000000001
            Me.sidebar.Width = i
            Me.Container.Left = Me.sidebar.Width
            Me.Container.Width = Me.Width
            i = i - 1
        Loop
        'Camisa.Visible = False
    Else
        DoEvents
        For i = 60 To 186
            Sleep 0.000000000001
            Me.sidebar.Width = i
            Me.Container.Left = Me.sidebar.Width
            Me.Container.Width = Me.Width
        Next
        'Camisa.Visible = True
    End If
    Call UserForm_Resize
End Sub

Private Sub Frame10_Click()

End Sub

Private Sub Frame9_Click()

End Sub

Private Sub Image13_BeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal DragState As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)

End Sub

Private Sub Label12_Click()

End Sub

Private Sub Image15_BeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal DragState As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)

End Sub

Private Sub Image15_Click()
On Error GoTo Err

ThisWorkbook.RefreshAll
DefineRange
Exit Sub

Err:
MsgBox ("Connection of data source is not available for now!" & vbNewLine & "Please try again later")
End Sub

Private Sub imgChart_BeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal DragState As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)

End Sub

Private Sub Label4_Click()

DataNick
    
    

TextBox1 = CDbl(Nick.Range("I2"))
TextBox2 = CDbl(Nick.Range("J2"))
TextBox3 = CDbl(Nick.Range("K2"))
TextBox1 = Format(TextBox1, "#,##0.0")
    TextBox2 = Format(TextBox2, "#,##0.0")
    TextBox3 = Format(TextBox3, "#,##0.0")

End Sub

Private Sub Label5_Click()
DataAlanJackpot
    
TextBox1 = AlanJackpot.Range("I2")
TextBox2 = AlanJackpot.Range("J2")
TextBox3 = AlanJackpot.Range("K2")

TextBox1 = Format(TextBox1, "#,##0.0")
    TextBox2 = Format(TextBox2, "#,##0.0")
    TextBox3 = Format(TextBox3, "#,##0.0")
    

End Sub

Private Sub Label57_Click()

End Sub

Private Sub Label59_Click()
DataNick
    

TextBox1 = CDbl(Nick.Range("I2"))
TextBox2 = CDbl(Nick.Range("J2"))
TextBox3 = CDbl(Nick.Range("K2"))

TextBox1 = Format(TextBox1, "#,##0.0")
    TextBox2 = Format(TextBox2, "#,##0.0")
    TextBox3 = Format(TextBox3, "#,##0.0")
    

End Sub

Private Sub Label6_Click()
DataIsac
    
    
TextBox1 = Isac.Range("I2")
TextBox2 = Isac.Range("J2")
TextBox3 = Isac.Range("K2")

TextBox1 = Format(TextBox1, "#,##0.0")
    TextBox2 = Format(TextBox2, "#,##0.0")
    TextBox3 = Format(TextBox3, "#,##0.0")

End Sub

Private Sub Label60_Click()
DataAlanJackpot
    
    
TextBox1 = AlanJackpot.Range("I2")
TextBox2 = AlanJackpot.Range("J2")
TextBox3 = AlanJackpot.Range("K2")

TextBox1 = Format(TextBox1, "#,##0.0")
    TextBox2 = Format(TextBox2, "#,##0.0")
    TextBox3 = Format(TextBox3, "#,##0.0")
End Sub

Private Sub Label61_Click()
DataIsac
    
    
TextBox1 = Isac.Range("I2")
TextBox2 = Isac.Range("J2")
TextBox3 = Isac.Range("K2")

TextBox1 = Format(TextBox1, "#,##0.0")
    TextBox2 = Format(TextBox2, "#,##0.0")
    TextBox3 = Format(TextBox3, "#,##0.0")

End Sub

Private Sub Label9_Click()

End Sub

Private Sub MenuCustomers_Click()

End Sub

Private Sub MenuDashboard_Click()

End Sub

Private Sub MenuLogout_Click()
    Unload Me
    Application.Visible = False
    ThisWorkbook.Close True
End Sub

Private Sub sidebar_Click()

End Sub

Private Sub TextBox1_Change()

End Sub

Private Sub TextBox2_Change()

End Sub

Private Sub UserForm_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Unload app
End Sub

Private Sub UserForm_Initialize()
    

    Application.ScreenUpdating = True
Application.EnableEvents = True
Application.Calculation = xlCalculationAutomatic
    Call removeTudo(Me)
    Call Maocursor(Me)
    Call UserForm_Resize

End Sub

Private Sub UserForm_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call moverForm(Me, Me, Button)
End Sub

Private Sub UserForm_Resize()
    Dim ctrl As Control
    Dim img  As Control
    
    'CONTAINER
    With Container
        .Width = Me.Width - (sidebar.Width + 15)
    End With
    
    'USER
    With CardUser
        .Left = Container.Width - .Width + 5
    End With
    
    'CARDS
    For Each ctrl In Container.Controls
        If TypeName(ctrl) = "Frame" And ctrl.Tag = "Cards" Then
            ctrl.Width = (Container.Width / 45) * 38
            'Image15.Left = 100
            imgChart.Width = ctrl.Width - 10
            
        End If
    Next
    'Frame9.Left = 18
    'Frame10.Left = Frame9.Left + Frame9.Width + 12
    'Frame11.Left = Frame10.Left + Frame10.Width + 12
    'Frame12.Left = Frame11.Left + Frame11.Width + 12
    
    'IMG CARD
    For Each ctrl In Container.Controls
        If TypeName(ctrl) = "Frame" And ctrl.Tag = "Cards" Then
           For Each img In ctrl.Controls
              If TypeName(img) = "Image" Then
                img.Left = ctrl.Width - (img.Width + 6)
                
                
                
              End If
           Next
        End If
    Next

    'CARD-BODY
    'With CardBody
        '.Width = Calc(Container.Width) - 24
    'End With
    
    'CARD CHARTS
    'Me.CardCharts.Left = CardBody.Width + CardBody.Left + 18
    'Me.btnFilter.Left = CardBody.Width - (Me.btnFilter.Width + 12)
    
    
    'SIDEBAR
    Me.sidebar.Height = Me.Height
    
    'MENU LOGOUT
    Me.MenuLogout.Top = Me.sidebar.Height - (Me.MenuLogout.Height + 18)
    
    
    
End Sub

Function Calc(value As Double)
    Calc = value * 0.67
End Function


