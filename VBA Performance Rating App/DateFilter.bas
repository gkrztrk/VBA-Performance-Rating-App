Attribute VB_Name = "DateFilter"
Sub DefineRange()

Dim dt1, dt2 As Date


dt1 = Data.Range("A3")
dt2 = Data.Range("B3")
' ws.PivotTables("RMN BCKFLL LST WEEK M3").PivotFields("DATE").PivotFilters.Add Type:=xlDateBetween, Value1:=dt2 - 2, Value2:=dt1 - 2
Nick.PivotTables("PvtNick").PivotFields("Date").ClearAllFilters
Nick.PivotTables("PvtNick").PivotFields("Date").PivotFilters.Add Type:=xlDateBetween, Value1:=dt1, Value2:=dt2

Isac.PivotTables("PvtIsac").PivotFields("Date").ClearAllFilters
Isac.PivotTables("PvtIsac").PivotFields("Date").PivotFilters.Add Type:=xlDateBetween, Value1:=dt1, Value2:=dt2

AlanJackpot.PivotTables("PvtAJ").PivotFields("Date").ClearAllFilters
AlanJackpot.PivotTables("PvtAJ").PivotFields("Date").PivotFilters.Add Type:=xlDateBetween, Value1:=dt1, Value2:=dt2



End Sub
