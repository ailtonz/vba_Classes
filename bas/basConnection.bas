Attribute VB_Name = "basConnection"
Public Bnc As New clsBancos

Public Function OpenConnection(banco As clsBancos) As ADODB.Connection
'' Build the connection string depending on the source
Dim connectionString As String
    
Select Case banco.Source
    Case "Access"
        connectionString = "Provider=" & banco.Driver & ";Data Source=" & banco.Database
    Case "Access2003"
        connectionString = "Driver={" & banco.Driver & "};Dbq=" & banco.Location & banco.Database & ";Uid=" & banco.User & ";PWD=" & banco.Password & ""
    Case "SQLite"
        connectionString = "Driver={" & banco.Driver & "};Database=" & banco.Database
    Case "MySQL"
        connectionString = "Driver={" & banco.Driver & "};Server=" & banco.Location & ";Database=" & banco.Database & ";PORT=" & banco.Port & ";UID=" & banco.User & ";PWD=" & banco.Password
    Case "PostgreSQL"
        connectionString = "Driver={" & banco.Driver & "};Server=" & banco.Location & ";Database=" & banco.Database & ";UID=" & banco.User & ";PWD=" & banco.Password
End Select

'' Create and open a new connection to the selected source
Set OpenConnection = New ADODB.Connection
Call OpenConnection.Open(connectionString)
   
End Function

Public Sub carregarBanco()
Dim wsBnc As Worksheet
Set wsBnc = Worksheets("BANCOS")

    With Bnc
        .Source = wsBnc.Range("F2").Value
        .Driver = wsBnc.Range("F3").Value
        .Location = wsBnc.Range("F4").Value
        .Database = wsBnc.Range("F5").Value
        .User = wsBnc.Range("F6").Value
        .Password = wsBnc.Range("F7").Value
        .Port = wsBnc.Range("F8").Value
        .add Bnc
    End With

Set wsBnc = Nothing

End Sub
