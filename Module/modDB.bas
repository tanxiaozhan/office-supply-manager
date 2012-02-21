Attribute VB_Name = "modDB"
Public Conn  As New ADODB.Connection

'连接ACCESS数据库
Sub DBConnect()
    
    strconn = "Provider=Microsoft.Jet.OLEDB.4.0;jet oledb:database Password=oasis;Data Source=" & GetApp & "data.gds"
    
    If Conn.state <> 0 Then Conn.Close
    Conn.Open strconn
    
End Sub
