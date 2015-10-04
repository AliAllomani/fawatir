Attribute VB_Name = "db_mod"
Public DB As Connection
Public rs As Recordset


Public Sub OpenDB()
    Set DB = New Connection
    DB.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;jet OLEDB:database password=master;Data Source=" & MyDBF & ";"
End Sub


