Attribute VB_Name = "Sql"

Public cnn As New ADODB.Connection


Public Sub ConectaDB()

if cnn.State = 0 then

cnn.Open "Provider=SQLNCLI11;Server=DESKTOP-84NKS8P\SQLEXPRESS;Database=UserDB;Trusted_Connection=yes;"

end if


End Sub

Sub NovoCadastro()

ConectaDB



End Sub