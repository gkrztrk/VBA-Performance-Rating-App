Attribute VB_Name = "SQLite"
Option Explicit

Dim cn As New ADODB.Connection, rs As New ADODB.Recordset
Dim sql As String
Dim Linha As Long, Coluna As Long, i As Long
Dim Pagina()



Public Sub ConectaSQLite()
    Set cn = New ADODB.Connection
    Set rs = New ADODB.Recordset
    
    ' OPEN CONNECTION
    cn.Open "DRIVER=SQLite3 ODBC Driver;Database=" & ThisWorkbook.Path & "\database.db"

End Sub

Public Sub Insert(username As String, password As String)
    Call ConectaSQLite
    
    sql = "SELECT * FROM tbUser"
    rs.Open sql, cn, 3, 3

    With rs
        .AddNew
        .Fields("username") = username
        .Fields("password") = password
        .Update
    End With

    rs.Close
    Set rs = Nothing: Set cn = Nothing
End Sub

Public Sub Update(id As Long, username As String, password As String)
    Call ConectaSQLite
    
    sql = "SELECT * FROM tbUser WHERE id = " + id
    rs.Open sql, cn, 3, 3
    

    With rs
        .Fields("username") = username
        .Fields("password") = password
        .Update
    End With

    rs.Close
    Set rs = Nothing: Set cn = Nothing
End Sub

Public Sub Delete(id As Long)
    Call ConectaSQLite
    
    sql = "DELETE FROM tbUser WHERE id = " + id
    
    rs.Close
    Set rs = Nothing: Set cn = Nothing
End Sub

Sub SelectAll()

    Call ConectaSQLite
    sql = "SELECT * FROM tbUser"
    rs.Source = sql
    Set rs.ActiveConnection = cn
    rs.CursorLocation = adUseClient
    rs.Open

'    Worksheets("Folha1").Range("A2").CopyFromRecordset rs   '// Este c�digo fica desabilitado no momento
    
    'Codigo Novo
    
    Linha = rs.RecordCount: Coluna = rs.Fields.count
    ReDim ArraysListaUsuarios(1 To Linha, 1 To Coluna)
    
    For i = 1 To rs.RecordCount
        ArraysListaUsuarios(1, 1) = rs(0)
        ArraysListaUsuarios(1, 2) = rs(1)
        ArraysListaUsuarios(1, 3) = rs(2)
        rs.MoveNext
    Next
    
    Call StorageList(ArraysListaUsuarios, UBound(ArraysListaUsuarios), Pagina)
    
    'Fim C�digo Novo
    
    
    rs.Close
    Set rs = Nothing: Set cn = Nothing
End Sub

Public Sub RefleshAll()
    Dim i As Long
    Dim ultimalinha As Long
    
    ultimalinha = Worksheets("Folha1").Cells(Worksheets("Folha1").Rows.count, 2).End(xlUp).Row

    For i = 2 To ultimalinha
        Select Case Cells(i, 4)
            Case "Inserir"
                Insert Cells(i, 2), Cells(i, 3)
            Case "Alterar"
                Update Cells(i, 1), Cells(i, 2), Cells(i, 3)
            Case "Excluir"
                Delete Cells(i, 1)
        End Select
    Next i
    
    Call SelectAll
    
End Sub

