Attribute VB_Name = "PerformanceData"
Sub DataNick()


Dim CurrentChart As Chart
Dim FName As String

FName = ThisWorkbook.Path & "\Temp.jpg"

Set CurrentChart = Nick.ChartObjects("ChartNick").Chart

CurrentChart.Export Filename:=FName, filtername:="JPG"

frmPerformance.imgChart.Picture = LoadPicture(FName)



End Sub

Sub DataIsac()


Dim CurrentChart As Chart
Dim FName As String

FName = ThisWorkbook.Path & "\Temp.gif"

Set CurrentChart = Isac.ChartObjects("ChartIsac").Chart

CurrentChart.Export Filename:=FName, filtername:="GIF"

frmPerformance.imgChart.Picture = LoadPicture(FName)



End Sub


Sub DataAlanJackpot()


Dim CurrentChart As Chart
Dim FName As String

FName = ThisWorkbook.Path & "\Temp.gif"

Set CurrentChart = AlanJackpot.ChartObjects("ChartAlanJackpot").Chart

CurrentChart.Export Filename:=FName, filtername:="GIF"

frmPerformance.imgChart.Picture = LoadPicture(FName)



End Sub

