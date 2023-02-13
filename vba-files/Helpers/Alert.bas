Attribute VB_Name = "Alert"


'namespace=vba-files\Helpers


'/*
'
'This comment block is used by XVBA to
' show the sub info
'
'@return void
'*/
Public Sub Show()

 linhas = worksheetfunction.countA(Planilha2.columns("a"))
 
 MsgBox "test" & linhas
 
End Sub
