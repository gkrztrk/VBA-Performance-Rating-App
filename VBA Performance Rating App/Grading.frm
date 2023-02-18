VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Grading 
   Caption         =   " "
   ClientHeight    =   10710
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   18465
   OleObjectBlob   =   "Grading.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Grading"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnFilter_Click()
    Calendary.Show
End Sub

Private Sub btnMenu_BeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal DragState As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)

End Sub

Private Sub CommandButton1_Click()

Grading.Hide

Unload frmResults
Load frmResults
frmResults.Show



End Sub

Private Sub CommandButton2_Click()
Dim elcevab As String

elcevab = MsgBox("All scores will be deleted permanently!" & vbNewLine & "Are you still want to continue?", vbYesNo, "Reset Scores")

If elcevab = vbYes Then



Data.Range("F4:I9").Clear
Data.Range("L4:S9").Clear
Unload Me
Load Grading
Grading.Show

End If


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
            Sleep 0.00000000000001
            
            Me.sidebar.Width = i
            Me.Container.Left = Me.sidebar.Width
            Me.Container.Width = Me.Width
            i = i - 1
        Loop
        'Camisa.Visible = False
    Else
        DoEvents
        For i = 60 To 186
            Sleep 0.00000000000001
            
            Me.sidebar.Width = i
            Me.Container.Left = Me.sidebar.Width
            Me.Container.Width = Me.Width
        Next
        'Camisa.Visible = True
    End If
    Call UserForm_Resize
End Sub

Private Sub Image13_BeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal DragState As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)

End Sub

Private Sub Frame10_Click()

End Sub

Private Sub Frame13_Click()

End Sub

Private Sub Frame14_Click()

End Sub

Private Sub Frame9_Click()

End Sub



Private Sub Image15_BeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal DragState As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)

End Sub

Private Sub Image15_Click()
    
    If Now - Data.Range("L4") >= 1 Then
    
        Data.Range("F4") = Data.Range("F4") + 100
        Data.Range("L4") = Now
        Image15.Visible = False
        Image69.Visible = True
        
    End If

End Sub

Private Sub Image18_Click()
    
    If Now - Data.Range("L5") >= 1 Then
    
        Data.Range("F5") = Data.Range("F5") + 100
        Data.Range("L5") = Now
        Image18.Visible = False
        Image70.Visible = True
    End If

End Sub

Private Sub Image20_Click()
    
    If Now - Data.Range("L6") >= 1 Then
    
        Data.Range("F6") = Data.Range("F6") + 100
        Data.Range("L6") = Now
        Image20.Visible = False
        Image71.Visible = True
    End If

End Sub

Private Sub Image22_Click()
    
    If Now - Data.Range("L7") >= 1 Then
    
        Data.Range("F7") = Data.Range("F7") + 100
        Data.Range("L7") = Now
        Image22.Visible = False
        Image72.Visible = True
    End If

End Sub

Private Sub Image24_Click()
    
    If Now - Data.Range("L8") >= 1 Then
    
        Data.Range("F8") = Data.Range("F8") + 100
        Data.Range("L8") = Now
        Image24.Visible = False
        Image73.Visible = True
    End If

End Sub

Private Sub Image25_Click()
    
    If Now - Data.Range("L9") >= 1 Then
    
        Data.Range("F9") = Data.Range("F9") + 100
        Data.Range("L9") = Now
        Image25.Visible = False
        Image74.Visible = True
    End If

End Sub

Private Sub Image30_Click()
    
    If Now - Data.Range("N4") >= 1 Then
    
        Data.Range("G4") = Data.Range("G4") + 100
        Data.Range("N4") = Now
        Image30.Visible = False
        Image75.Visible = True
    End If

End Sub

Private Sub Image27_Click()
    
    If Now - Data.Range("N5") >= 1 Then
    
        Data.Range("G5") = Data.Range("G5") + 100
        Data.Range("N5") = Now
        Image27.Visible = False
        Image76.Visible = True
    End If

End Sub


Private Sub Image32_Click()
    
    If Now - Data.Range("N6") >= 1 Then
    
        Data.Range("G6") = Data.Range("G6") + 100
        Data.Range("N6") = Now
        Image32.Visible = False
        Image77.Visible = True
    End If

End Sub


Private Sub Image34_Click()
    
    If Now - Data.Range("N7") >= 1 Then
    
        Data.Range("G7") = Data.Range("G7") + 100
        Data.Range("N7") = Now
        Image34.Visible = False
        Image78.Visible = True
    End If

End Sub


Private Sub Image36_Click()
    
    If Now - Data.Range("N8") >= 1 Then
    
        Data.Range("G8") = Data.Range("G8") + 100
        Data.Range("N8") = Now
        Image36.Visible = False
        Image79.Visible = True
    End If

End Sub


Private Sub Image37_Click()
    
    If Now - Data.Range("N9") >= 1 Then
    
        Data.Range("G9") = Data.Range("G9") + 100
        Data.Range("N9") = Now
        Image37.Visible = False
        Image80.Visible = True
    End If

End Sub


Private Sub Image42_Click()
    
    If Now - Data.Range("P4") >= 1 Then
    
        Data.Range("H4") = Data.Range("H4") + 100
        Data.Range("P4") = Now
        Image42.Visible = False
        Image81.Visible = True
    End If

End Sub

Private Sub Image39_Click()
    
    If Now - Data.Range("P5") >= 1 Then
    
        Data.Range("H5") = Data.Range("H5") + 100
        Data.Range("P5") = Now
        Image39.Visible = False
        Image82.Visible = True
    End If

End Sub

Private Sub Image44_Click()
    
    If Now - Data.Range("P6") >= 1 Then
    
        Data.Range("H6") = Data.Range("H6") + 100
        Data.Range("P6") = Now
        Image44.Visible = False
        Image83.Visible = True
    End If

End Sub


Private Sub Image46_Click()
    
    If Now - Data.Range("P7") >= 1 Then
    
        Data.Range("H7") = Data.Range("H7") + 100
        Data.Range("P7") = Now
        Image46.Visible = False
        Image84.Visible = True
    End If

End Sub


Private Sub Image48_Click()
    
    If Now - Data.Range("P8") >= 1 Then
    
        Data.Range("H8") = Data.Range("H8") + 100
        Data.Range("P8") = Now
        Image48.Visible = False
        Image85.Visible = True
    End If

End Sub


Private Sub Image49_Click()
    
    If Now - Data.Range("P9") >= 1 Then
    
        Data.Range("H9") = Data.Range("H9") + 100
        Data.Range("P9") = Now
        Image49.Visible = False
        Image86.Visible = True
    End If

End Sub


Private Sub Image54_Click()
    
    If Now - Data.Range("R4") >= 1 Then
    
        Data.Range("I4") = Data.Range("I4") + 100
        Data.Range("R4") = Now
        Image54.Visible = False
        Image87.Visible = True
    End If

End Sub


Private Sub Image51_Click()
    
    If Now - Data.Range("R5") >= 1 Then
    
        Data.Range("I5") = Data.Range("I5") + 100
        Data.Range("R5") = Now
        Image51.Visible = False
        Image88.Visible = True
    End If

End Sub


Private Sub Image56_Click()
    
    If Now - Data.Range("R6") >= 1 Then
    
        Data.Range("I6") = Data.Range("I6") + 100
        Data.Range("R6") = Now
        Image56.Visible = False
        Image89.Visible = True
    End If

End Sub


Private Sub Image58_Click()
    
    If Now - Data.Range("R7") >= 1 Then
    
        Data.Range("I7") = Data.Range("I7") + 100
        Data.Range("R7") = Now
        Image58.Visible = False
        Image90.Visible = True
    End If

End Sub


Private Sub Image60_Click()
    
    If Now - Data.Range("R8") >= 1 Then
    
        Data.Range("I8") = Data.Range("I8") + 100
        Data.Range("R8") = Now
        Image60.Visible = False
        Image91.Visible = True
    End If

End Sub


Private Sub Image61_Click()
    
    If Now - Data.Range("R9") >= 1 Then
    
        Data.Range("I9") = Data.Range("I9") + 100
        Data.Range("R9") = Now
        Image61.Visible = False
        Image92.Visible = True
    End If

End Sub


Private Sub Image16_Click()
    
    If Now - Data.Range("M4") >= 1 Then
    
        Data.Range("F4") = Data.Range("F4") - 100
        Data.Range("M4") = Now
        Image16.Visible = False
        Image93.Visible = True
    End If

End Sub


Private Sub Image17_Click()
    
    If Now - Data.Range("M5") >= 1 Then
    
        Data.Range("F5") = Data.Range("F5") - 100
        Data.Range("M5") = Now
        Image17.Visible = False
        Image94.Visible = True
    End If

End Sub


Private Sub Image19_Click()
    
    If Now - Data.Range("M6") >= 1 Then
    
        Data.Range("F6") = Data.Range("F6") - 100
        Data.Range("M6") = Now
        Image19.Visible = False
        Image95.Visible = True
    End If

End Sub


Private Sub Image21_Click()
    
    If Now - Data.Range("M7") >= 1 Then
    
        Data.Range("F7") = Data.Range("F7") - 100
        Data.Range("M7") = Now
        Image21.Visible = False
        Image96.Visible = True
    End If

End Sub


Private Sub Image23_Click()
    
    If Now - Data.Range("M8") >= 1 Then
    
        Data.Range("F8") = Data.Range("F8") - 100
        Data.Range("M8") = Now
        Image23.Visible = False
        Image97.Visible = True
    End If

End Sub


Private Sub Image26_Click()
    
    If Now - Data.Range("M9") >= 1 Then
    
        Data.Range("F9") = Data.Range("F9") - 100
        Data.Range("M9") = Now
        Image26.Visible = False
        Image98.Visible = True
    End If

End Sub



Private Sub Image29_Click()
    
    If Now - Data.Range("O4") >= 1 Then
    
        Data.Range("G4") = Data.Range("G4") - 100
        Data.Range("O4") = Now
        Image29.Visible = False
        Image99.Visible = True
    End If

End Sub


Private Sub Image28_Click()
    
    If Now - Data.Range("O5") >= 1 Then
    
        Data.Range("G5") = Data.Range("G5") - 100
        Data.Range("O5") = Now
        Image28.Visible = False
        Image100.Visible = True
    End If

End Sub


Private Sub Image31_Click()
    
    If Now - Data.Range("O6") >= 1 Then
    
        Data.Range("G6") = Data.Range("G6") - 100
        Data.Range("O6") = Now
        Image31.Visible = False
        Image101.Visible = True
    End If

End Sub


Private Sub Image33_Click()
    
    If Now - Data.Range("O7") >= 1 Then
    
        Data.Range("G7") = Data.Range("G7") - 100
        Data.Range("O7") = Now
        Image33.Visible = False
        Image102.Visible = True
    End If

End Sub


Private Sub Image35_Click()
    
    If Now - Data.Range("O8") >= 1 Then
    
        Data.Range("G8") = Data.Range("G8") - 100
        Data.Range("O8") = Now
        Image35.Visible = False
        Image103.Visible = True
    End If

End Sub


Private Sub Image38_Click()
    
    If Now - Data.Range("O9") >= 1 Then
    
        Data.Range("G9") = Data.Range("G9") - 100
        Data.Range("O9") = Now
        Image38.Visible = False
        Image104.Visible = True
    End If

End Sub


Private Sub Image41_Click()
    
    If Now - Data.Range("Q4") >= 1 Then
    
        Data.Range("H4") = Data.Range("H4") - 100
        Data.Range("Q4") = Now
        Image41.Visible = False
        Image105.Visible = True
    End If

End Sub


Private Sub Image40_Click()
    
    If Now - Data.Range("Q5") >= 1 Then
    
        Data.Range("H5") = Data.Range("H5") - 100
        Data.Range("Q5") = Now
        Image40.Visible = False
        Image106.Visible = True
    End If

End Sub


Private Sub Image43_Click()
    
    If Now - Data.Range("Q6") >= 1 Then
    
        Data.Range("H6") = Data.Range("H6") - 100
        Data.Range("Q6") = Now
        Image43.Visible = False
        Image107.Visible = True
    End If

End Sub


Private Sub Image45_Click()
    
    If Now - Data.Range("Q7") >= 1 Then
    
        Data.Range("H7") = Data.Range("H7") - 100
        Data.Range("Q7") = Now
        Image45.Visible = False
        Image108.Visible = True
    End If

End Sub


Private Sub Image47_Click()
    
    If Now - Data.Range("Q8") >= 1 Then
    
        Data.Range("H8") = Data.Range("H8") - 100
        Data.Range("Q8") = Now
        Image47.Visible = False
        Image109.Visible = True
    End If

End Sub


Private Sub Image50_Click()
    
    If Now - Data.Range("Q9") >= 1 Then
    
        Data.Range("H9") = Data.Range("H9") - 100
        Data.Range("Q9") = Now
        Image50.Visible = False
        Image110.Visible = True
    End If

End Sub


Private Sub Image53_Click()
    
    If Now - Data.Range("S4") >= 1 Then
    
        Data.Range("I4") = Data.Range("I4") - 100
        Data.Range("S4") = Now
        Image53.Visible = False
        Image111.Visible = True
    End If

End Sub


Private Sub Image52_Click()
    
    If Now - Data.Range("S5") >= 1 Then
    
        Data.Range("I5") = Data.Range("I5") - 100
        Data.Range("S5") = Now
        Image52.Visible = False
        Image112.Visible = True
    End If

End Sub


Private Sub Image55_Click()
    
    If Now - Data.Range("S6") >= 1 Then
    
        Data.Range("I6") = Data.Range("I6") - 100
        Data.Range("S6") = Now
        Image55.Visible = False
        Image113.Visible = True
    End If

End Sub


Private Sub Image57_Click()
    
    If Now - Data.Range("S7") >= 1 Then
    
        Data.Range("I7") = Data.Range("I7") - 100
        Data.Range("S7") = Now
        Image57.Visible = False
        Image114.Visible = True
    End If

End Sub


Private Sub Image59_Click()
    
    If Now - Data.Range("S8") >= 1 Then
    
        Data.Range("I8") = Data.Range("I8") - 100
        Data.Range("S8") = Now
        Image59.Visible = False
        Image115.Visible = True
    End If

End Sub


Private Sub Image62_Click()
    
    If Now - Data.Range("S9") >= 1 Then
    
        Data.Range("I9") = Data.Range("I9") - 100
        Data.Range("S9") = Now
        Image62.Visible = False
        Image116.Visible = True
    End If

End Sub


Private Sub Image69_BeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal DragState As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)

End Sub

'////////////////// Grey Buttons \\\\\\\\\\\\\\\\\\\\\\\\\\\\\



Private Sub Image69_Click()
    
    If Now - Data.Range("L4") <= 1 Then
    
        MsgBox "You can rate once in a day"
        
    End If

End Sub

Private Sub Image70_Click()
    
    If Now - Data.Range("L5") <= 1 Then
    
       MsgBox "You can rate once in a day"
       
    End If

End Sub

Private Sub Image71_Click()
    
    If Now - Data.Range("L6") <= 1 Then
    
        MsgBox "You can rate once in a day"
       
    End If

End Sub

Private Sub Image72_Click()
    
    If Now - Data.Range("L7") <= 1 Then
    
        MsgBox "You can rate once in a day"
       
    End If

End Sub

Private Sub Image73_Click()
    
    If Now - Data.Range("L8") <= 1 Then
    
        MsgBox "You can rate once in a day"
       
    End If

End Sub

Private Sub Image74_Click()
    
    If Now - Data.Range("L9") <= 1 Then
    
        MsgBox "You can rate once in a day"
       
    End If

End Sub

Private Sub Image75_Click()
    
    If Now - Data.Range("N4") <= 1 Then
    
        MsgBox "You can rate once in a day"
       
    End If

End Sub

Private Sub Image76_Click()
    
    If Now - Data.Range("N5") <= 1 Then
    
        MsgBox "You can rate once in a day"
       
    End If

End Sub


Private Sub Image77_Click()
    
    If Now - Data.Range("N6") <= 1 Then
    
        MsgBox "You can rate once in a day"
       
    End If

End Sub


Private Sub Image78_Click()
    
    If Now - Data.Range("N7") <= 1 Then
    
        MsgBox "You can rate once in a day"
       
    End If

End Sub


Private Sub Image79_Click()
    
    If Now - Data.Range("N8") <= 1 Then
    
        MsgBox "You can rate once in a day"
       
    End If

End Sub


Private Sub Image80_Click()
    
    If Now - Data.Range("N9") <= 1 Then
    
        MsgBox "You can rate once in a day"
       
    End If

End Sub


Private Sub Image81_Click()
    
    If Now - Data.Range("P4") <= 1 Then
    
        MsgBox "You can rate once in a day"
       
    End If

End Sub

Private Sub Image82_Click()
    
    If Now - Data.Range("P5") <= 1 Then
    
        MsgBox "You can rate once in a day"
       
    End If

End Sub

Private Sub Image83_Click()
    
    If Now - Data.Range("P6") <= 1 Then
    
        MsgBox "You can rate once in a day"
       
    End If

End Sub


Private Sub Image84_Click()
    
    If Now - Data.Range("P7") <= 1 Then
    
        MsgBox "You can rate once in a day"
       
    End If

End Sub


Private Sub Image85_Click()
    
    If Now - Data.Range("P8") <= 1 Then
    
        MsgBox "You can rate once in a day"
       
    End If

End Sub


Private Sub Image86_Click()
    
    If Now - Data.Range("P9") <= 1 Then
    
        MsgBox "You can rate once in a day"
       
    End If

End Sub


Private Sub Image87_Click()
    
    If Now - Data.Range("R4") <= 1 Then
    
       MsgBox "You can rate once in a day"
       
    End If

End Sub


Private Sub Image88_Click()
    
    If Now - Data.Range("R5") <= 1 Then
    
        MsgBox "You can rate once in a day"
       
    End If

End Sub


Private Sub Image89_Click()
    
    If Now - Data.Range("R6") <= 1 Then
    
        MsgBox "You can rate once in a day"
       
    End If

End Sub


Private Sub Image90_Click()
    
    If Now - Data.Range("R7") <= 1 Then
    
      MsgBox "You can rate once in a day"
       
    End If

End Sub


Private Sub Image91_Click()
    
    If Now - Data.Range("R8") <= 1 Then
    
        MsgBox "You can rate once in a day"
       
    End If

End Sub


Private Sub Image92_Click()
    
    If Now - Data.Range("R9") <= 1 Then
    
       MsgBox "You can rate once in a day"
       
    End If

End Sub


Private Sub Image93_Click()
    
    If Now - Data.Range("M4") <= 1 Then
    
        MsgBox "You can rate once in a day"
       
    End If

End Sub


Private Sub Image94_Click()
    
    If Now - Data.Range("M5") <= 1 Then
    
        MsgBox "You can rate once in a day"
       
    End If

End Sub


Private Sub Image95_Click()
    
    If Now - Data.Range("M6") <= 1 Then
    
        MsgBox "You can rate once in a day"
       
    End If

End Sub


Private Sub Image96_Click()
    
    If Now - Data.Range("M7") <= 1 Then
    
        MsgBox "You can rate once in a day"
       
    End If

End Sub


Private Sub Image97_Click()
    
    If Now - Data.Range("M8") <= 1 Then
    
        MsgBox "You can rate once in a day"
       
    End If

End Sub


Private Sub Image98_Click()
    
    If Now - Data.Range("M9") <= 1 Then
    
        MsgBox "You can rate once in a day"
       
    End If

End Sub



Private Sub Image99_Click()
    
    If Now - Data.Range("O4") <= 1 Then
    
        MsgBox "You can rate once in a day"
       
    End If

End Sub


Private Sub Image100_Click()
    
    If Now - Data.Range("O5") <= 1 Then
    
        MsgBox "You can rate once in a day"
       
    End If

End Sub


Private Sub Image101_Click()
    
    If Now - Data.Range("O6") <= 1 Then
    
        MsgBox "You can rate once in a day"
       
    End If

End Sub


Private Sub Image102_Click()
    
    If Now - Data.Range("O7") <= 1 Then
    
        MsgBox "You can rate once in a day"
       
    End If

End Sub


Private Sub Image103_Click()
    
    If Now - Data.Range("O8") <= 1 Then
    
        MsgBox "You can rate once in a day"
       
    End If

End Sub


Private Sub Image104_Click()
    
    If Now - Data.Range("O9") <= 1 Then
    
        MsgBox "You can rate once in a day"
       
    End If

End Sub


Private Sub Image105_Click()
    
    If Now - Data.Range("Q4") <= 1 Then
    
    MsgBox "You can rate once in a day"
       
    End If

End Sub


Private Sub Image106_Click()
    
    If Now - Data.Range("Q5") <= 1 Then
    
    MsgBox "You can rate once in a day"
       
    End If

End Sub


Private Sub Image107_Click()
    
    If Now - Data.Range("Q6") <= 1 Then
    
        MsgBox "You can rate once in a day"
       
    End If

End Sub


Private Sub Image108_Click()
    
    If Now - Data.Range("Q7") <= 1 Then
    
       MsgBox "You can rate once in a day"
       
    End If

End Sub


Private Sub Image109_Click()
    
    If Now - Data.Range("Q8") <= 1 Then
    
      MsgBox "You can rate once in a day"
       
    End If

End Sub


Private Sub Image110_Click()
    
    If Now - Data.Range("Q9") <= 1 Then
    
       MsgBox "You can rate once in a day"
       
    End If

End Sub


Private Sub Image111_Click()
    
    If Now - Data.Range("S4") <= 1 Then
    
        MsgBox "You can rate once in a day"
       
    End If

End Sub


Private Sub Image112_Click()
    
    If Now - Data.Range("S5") <= 1 Then
    
  MsgBox "You can rate once in a day"
       
    End If

End Sub


Private Sub Image113_Click()
    
    If Now - Data.Range("S6") <= 1 Then
    
        MsgBox "You can rate once in a day"
       
    End If

End Sub


Private Sub Image114_Click()
    
    If Now - Data.Range("S7") <= 1 Then
    
        MsgBox "You can rate once in a day"
       
    End If

End Sub


Private Sub Image115_Click()
    
    If Now - Data.Range("S8") <= 1 Then
    
        MsgBox "You can rate once in a day"
       
    End If

End Sub


Private Sub Image116_Click()
    
    If Now - Data.Range("S9") <= 1 Then
    
        MsgBox "You can rate once in a day"
       
    End If

End Sub





'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\



Private Sub Label12_Click()

End Sub

Private Sub Label16_Click()

End Sub

Private Sub Label4_Click()


End Sub

Private Sub Label5_Click()

End Sub

Private Sub Label55_Click()
DataAlanJackpot
Grading.Hide

    frmPerformance.TextBox1 = Format(frmPerformance.TextBox1, "#,##0.0")
    frmPerformance.TextBox2 = Format(frmPerformance.TextBox2, "#,##0.0")
    frmPerformance.TextBox3 = Format(frmPerformance.TextBox3, "#,##0.0")
    

frmPerformance.TextBox1 = CDbl(AlanJackpot.Range("I2"))
frmPerformance.TextBox2 = CDbl(AlanJackpot.Range("J2"))
frmPerformance.TextBox3 = CDbl(AlanJackpot.Range("K2"))


frmPerformance.Show

End Sub

Private Sub Label56_Click()
DataNick
Grading.Hide

   frmPerformance.TextBox1 = Format(frmPerformance.TextBox1, "#,##0.0")
    frmPerformance.TextBox2 = Format(frmPerformance.TextBox2, "#,##0.0")
    frmPerformance.TextBox3 = Format(frmPerformance.TextBox3, "#,##0.0")
    

frmPerformance.TextBox1 = CDbl(Nick.Range("I2"))
frmPerformance.TextBox2 = CDbl(Nick.Range("J2"))
frmPerformance.TextBox3 = CDbl(Nick.Range("K2"))

frmPerformance.Show
End Sub

Private Sub Label57_Click()
DataIsac
Grading.Hide

   frmPerformance.TextBox1 = Format(frmPerformance.TextBox1, "#,##0.0")
    frmPerformance.TextBox2 = Format(frmPerformance.TextBox2, "#,##0.0")
    frmPerformance.TextBox3 = Format(frmPerformance.TextBox3, "#,##0.0")
    

frmPerformance.TextBox1 = CDbl(Isac.Range("I2"))
frmPerformance.TextBox2 = CDbl(Isac.Range("J2"))
frmPerformance.TextBox3 = CDbl(Isac.Range("K2"))

frmPerformance.Show

End Sub

Private Sub Label6_Click()

End Sub

Private Sub Label9_Click()

End Sub

Private Sub MenuCustomers_Click()
DataAlanJackpot
Grading.Hide

    frmPerformance.TextBox1 = Format(frmPerformance.TextBox1, "#,##0.0")
    frmPerformance.TextBox2 = Format(frmPerformance.TextBox2, "#,##0.0")
    frmPerformance.TextBox3 = Format(frmPerformance.TextBox3, "#,##0.0")
    

frmPerformance.TextBox1 = CDbl(AlanJackpot.Range("I2"))
frmPerformance.TextBox2 = CDbl(AlanJackpot.Range("J2"))
frmPerformance.TextBox3 = CDbl(AlanJackpot.Range("K2"))


frmPerformance.Show




End Sub

Private Sub MenuDashboard_Click()
DataNick
Grading.Hide

   frmPerformance.TextBox1 = Format(frmPerformance.TextBox1, "#,##0.0")
    frmPerformance.TextBox2 = Format(frmPerformance.TextBox2, "#,##0.0")
    frmPerformance.TextBox3 = Format(frmPerformance.TextBox3, "#,##0.0")
    

frmPerformance.TextBox1 = CDbl(Nick.Range("I2"))
frmPerformance.TextBox2 = CDbl(Nick.Range("J2"))
frmPerformance.TextBox3 = CDbl(Nick.Range("K2"))

frmPerformance.Show

 

End Sub

Private Sub MenuLogout_Click()
    Unload Me
    Application.Visible = False
    ThisWorkbook.Close True
End Sub

Private Sub MenuProjects_Click()
DataIsac
Grading.Hide

   frmPerformance.TextBox1 = Format(frmPerformance.TextBox1, "#,##0.0")
    frmPerformance.TextBox2 = Format(frmPerformance.TextBox2, "#,##0.0")
    frmPerformance.TextBox3 = Format(frmPerformance.TextBox3, "#,##0.0")
    

frmPerformance.TextBox1 = CDbl(Isac.Range("I2"))
frmPerformance.TextBox2 = CDbl(Isac.Range("J2"))
frmPerformance.TextBox3 = CDbl(Isac.Range("K2"))

frmPerformance.Show


 

End Sub

Private Sub sidebar_Click()

End Sub

Private Sub UserForm_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Unload app
End Sub

Private Sub UserForm_Initialize()
    Call visibleButtons
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
    Dim a As Integer
    
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
            a = ctrl.Width
            ctrl.Width = (Container.Width / 168) * 38
        End If
    Next
    Frame9.Left = 18
    Frame10.Left = Frame9.Left + Frame9.Width + 12
    Frame11.Left = Frame10.Left + Frame10.Width + 12
    Frame12.Left = Frame11.Left + Frame11.Width + 12
    
    
    Frame13.Left = 18
    Frame14.Left = Frame13.Left + Frame13.Width + 12
    Frame15.Left = Frame14.Left + Frame14.Width + 12
    Frame16.Left = Frame15.Left + Frame15.Width + 12
    
    'IMG CARD
    For Each ctrl In Container.Controls
        If TypeName(ctrl) = "Frame" And ctrl.Tag = "Cards" Then
           For Each img In ctrl.Controls
              If TypeName(img) = "Image" Then
                img.Left = (img.Left + (ctrl.Width - a) / 2)
              End If
           Next
        End If
    Next

    'CARD -BODY
'    With CardBody
'        .Width = Calc(Container.Width) - 24
'    End With

    'CARD Charts
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



Sub visibleButtons()


' Image15
    
    If Now - Data.Range("L4") <= 1 Then

        Image15.Visible = False
        Image69.Visible = True
    Else
    
        Image15.Visible = True
        Image69.Visible = False
        
    End If



' Image18
    
    If Now - Data.Range("L5") <= 1 Then
    
        Image18.Visible = False
        Image70.Visible = True
    Else
    
        Image18.Visible = True
        Image70.Visible = False
        
    End If



' Image20
    
    If Now - Data.Range("L6") <= 1 Then
    
        Image20.Visible = False
        Image71.Visible = True
    Else
    
        Image20.Visible = True
        Image71.Visible = False
        
    End If



' Image22
    
    If Now - Data.Range("L7") <= 1 Then
    
        Image22.Visible = False
        Image72.Visible = True
    Else
    
        Image22.Visible = True
        Image72.Visible = False
        
    End If



' Image24
    
    If Now - Data.Range("L8") <= 1 Then
    
        Image24.Visible = False
        Image73.Visible = True
    Else
    
        Image24.Visible = True
        Image73.Visible = False
        
    End If



' Image25
    
    If Now - Data.Range("L9") <= 1 Then
    
        Image25.Visible = False
        Image74.Visible = True
    Else
    
        Image25.Visible = True
        Image74.Visible = False
        
    End If



' Image30
    
    If Now - Data.Range("N4") <= 1 Then
    
        Image30.Visible = False
        Image75.Visible = True
    Else
    
        Image30.Visible = True
        Image75.Visible = False
        
    End If



' Image27
    
    If Now - Data.Range("N5") <= 1 Then
    
        Image27.Visible = False
        Image76.Visible = True
    Else
    
        Image27.Visible = True
        Image76.Visible = False
        
    End If




' Image32
    
    If Now - Data.Range("N6") <= 1 Then
    
        Image32.Visible = False
        Image77.Visible = True
    Else
    
        Image32.Visible = True
        Image77.Visible = False
        
    End If




' Image34
    
    If Now - Data.Range("N7") <= 1 Then
    
        Image34.Visible = False
        Image78.Visible = True
    Else
    
        Image34.Visible = True
        Image78.Visible = False
        
    End If




' Image36
    
    If Now - Data.Range("N8") <= 1 Then
    
        Image36.Visible = False
        Image79.Visible = True
    Else
    
        Image36.Visible = True
        Image79.Visible = False
        
    End If




' Image37
    
    If Now - Data.Range("N9") <= 1 Then
    
        Image37.Visible = False
        Image80.Visible = True
    Else
    
        Image37.Visible = True
        Image80.Visible = False
        
    End If




' Image42
    
    If Now - Data.Range("P4") <= 1 Then
    
        Image42.Visible = False
        Image81.Visible = True
    Else
    
        Image42.Visible = True
        Image81.Visible = False
        
    End If



' Image39
    
    If Now - Data.Range("P5") <= 1 Then
    
        Image39.Visible = False
        Image82.Visible = True
    Else
    
        Image39.Visible = True
        Image82.Visible = False
        
    End If



' Image44
    
    If Now - Data.Range("P6") <= 1 Then
    
        Image44.Visible = False
        Image83.Visible = True
    Else
    
        Image44.Visible = True
        Image83.Visible = False
        
    End If




' Image46
    
    If Now - Data.Range("P7") <= 1 Then
    
        Image46.Visible = False
        Image84.Visible = True
    Else
    
        Image46.Visible = True
        Image84.Visible = False
        
    End If




' Image48
    
    If Now - Data.Range("P8") <= 1 Then
    
        Image48.Visible = False
        Image85.Visible = True
    Else
    
        Image48.Visible = True
        Image85.Visible = False
        
    End If




' Image49
    
    If Now - Data.Range("P9") <= 1 Then
    
        Image49.Visible = False
        Image86.Visible = True
    Else
    
        Image49.Visible = True
        Image86.Visible = False
        
    End If




' Image54
    
    If Now - Data.Range("R4") <= 1 Then
    
        Image54.Visible = False
        Image87.Visible = True
    Else
    
        Image54.Visible = True
        Image87.Visible = False
        
    End If




' Image51
    
    If Now - Data.Range("R5") <= 1 Then
    
        Image51.Visible = False
        Image88.Visible = True
    Else
    
        Image51.Visible = True
        Image88.Visible = False
        
    End If




' Image56
    
    If Now - Data.Range("R6") <= 1 Then
    
        Image56.Visible = False
        Image89.Visible = True
    Else
    
        Image56.Visible = True
        Image89.Visible = False
        
    End If




' Image58
    
    If Now - Data.Range("R7") <= 1 Then
    
        Image58.Visible = False
        Image90.Visible = True
    Else
    
        Image58.Visible = True
        Image90.Visible = False
        
    End If




' Image60
    
    If Now - Data.Range("R8") <= 1 Then
    
        Image60.Visible = False
        Image91.Visible = True
    Else
    
        Image60.Visible = True
        Image91.Visible = False
        
    End If




' Image61
    
    If Now - Data.Range("R9") <= 1 Then
    
        Image61.Visible = False
        Image92.Visible = True
    Else
    
        Image61.Visible = True
        Image92.Visible = False
        
    End If




' Image16
    
    If Now - Data.Range("M4") <= 1 Then
    
        Image16.Visible = False
        Image93.Visible = True
    Else
    
        Image16.Visible = True
        Image93.Visible = False
        
    End If




' Image17
    
    If Now - Data.Range("M5") <= 1 Then
    
        Image17.Visible = False
        Image94.Visible = True
    Else
    
        Image17.Visible = True
        Image94.Visible = False
        
    End If




' Image19
    
    If Now - Data.Range("M6") <= 1 Then
    
        Image19.Visible = False
        Image95.Visible = True
    Else
    
        Image19.Visible = True
        Image95.Visible = False
        
    End If




' Image21
    
    If Now - Data.Range("M7") <= 1 Then
    
        Image21.Visible = False
        Image96.Visible = True
    Else
    
        Image21.Visible = True
        Image96.Visible = False
        
    End If




' Image23
    
    If Now - Data.Range("M8") <= 1 Then
    
        Image23.Visible = False
        Image97.Visible = True
    Else
    
        Image23.Visible = True
        Image97.Visible = False
        
    End If




' Image26
    
    If Now - Data.Range("M9") <= 1 Then
    
        Image26.Visible = False
        Image98.Visible = True
    Else
    
        Image26.Visible = True
        Image98.Visible = False
        
    End If





' Image29
    
    If Now - Data.Range("O4") <= 1 Then
    
        Image29.Visible = False
        Image99.Visible = True
    Else
    
        Image29.Visible = True
        Image99.Visible = False
        
    End If




' Image28
    
    If Now - Data.Range("O5") <= 1 Then
    
        Image28.Visible = False
        Image100.Visible = True
    Else
    
        Image28.Visible = True
        Image100.Visible = False
        
    End If




' Image31
    
    If Now - Data.Range("O6") <= 1 Then
    
        Image31.Visible = False
        Image101.Visible = True
    Else
    
        Image31.Visible = True
        Image101.Visible = False
        
    End If




' Image33
    
    If Now - Data.Range("O7") <= 1 Then
    
        Image33.Visible = False
        Image102.Visible = True
    Else
    
        Image33.Visible = True
        Image102.Visible = False
        
    End If




' Image35
    
    If Now - Data.Range("O8") <= 1 Then
    
        Image35.Visible = False
        Image103.Visible = True
    Else
    
        Image35.Visible = True
        Image103.Visible = False
        
    End If




' Image38
    
    If Now - Data.Range("O9") <= 1 Then
    
        Image38.Visible = False
        Image104.Visible = True
    Else
    
        Image38.Visible = True
        Image104.Visible = False
        
    End If




' Image41
    
    If Now - Data.Range("Q4") <= 1 Then
    
       Image41.Visible = False
        Image105.Visible = True
    Else
    
        Image41.Visible = True
        Image105.Visible = False
        
    End If




' Image40
    
    If Now - Data.Range("Q5") <= 1 Then
    
        Image40.Visible = False
        Image106.Visible = True
    Else
    
        Image40.Visible = True
        Image106.Visible = False
        
    End If




' Image43
    
    If Now - Data.Range("Q6") <= 1 Then
    
        Image43.Visible = False
        Image107.Visible = True
    Else
    
        Image43.Visible = True
        Image107.Visible = False
        
    End If




' Image45
    
    If Now - Data.Range("Q7") <= 1 Then
    
       Image45.Visible = False
        Image108.Visible = True
    Else
    
        Image45.Visible = True
        Image108.Visible = False
        
    End If




' Image47
    
    If Now - Data.Range("Q8") <= 1 Then
    
        Image47.Visible = False
        Image109.Visible = True
    Else
    
        Image47.Visible = True
        Image109.Visible = False
        
    End If




' Image50
    
    If Now - Data.Range("Q9") <= 1 Then
    
        Image50.Visible = False
        Image110.Visible = True
    Else
    
        Image50.Visible = True
        Image110.Visible = False
        
    End If




' Image53
    
    If Now - Data.Range("S4") <= 1 Then
    
        Image53.Visible = False
        Image111.Visible = True
    Else
    
        Image53.Visible = True
        Image111.Visible = False
        
    End If




' Image52
    
    If Now - Data.Range("S5") <= 1 Then
    
        Image52.Visible = False
        Image112.Visible = True
    Else
    
        Image52.Visible = True
        Image112.Visible = False
        
    End If




' Image55
    
    If Now - Data.Range("S6") <= 1 Then
    
        Image55.Visible = False
        Image113.Visible = True
    Else
    
        Image55.Visible = True
        Image113.Visible = False
        
    End If




' Image57
    
    If Now - Data.Range("S7") <= 1 Then
    
        Image57.Visible = False
        Image114.Visible = True
    Else
    
        Image57.Visible = True
        Image114.Visible = False
        
    End If




' Image59
    
    If Now - Data.Range("S8") <= 1 Then
    
       Image59.Visible = False
        Image115.Visible = True
    Else
    
        Image59.Visible = True
        Image115.Visible = False
        
    End If




' Image62
    
    If Now - Data.Range("S9") <= 1 Then
    
        Image62.Visible = False
        Image116.Visible = True
    Else
    
        Image62.Visible = True
        Image116.Visible = False
        
    End If




End Sub






