Attribute VB_Name = "Module1"
 Dim cmd As String
  Dim cn As ADODB.connection
Public Sub con()
 
  cmd = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=dsn1;Initial Catalog=vb"
  Set cn = New ADODB.connection
  With cn
  .ConnectionString = cmd
  .Open
  End With
End Sub
