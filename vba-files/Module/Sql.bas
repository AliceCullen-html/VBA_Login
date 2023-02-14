Attribute VB_Name = "Sql"

Public cnn As New ADODB.Connection




Public Sub ConectaDB()


cnn.Open "Provider=SQLNCLI11;Server=DESKTOP-84NKS8P\SQLEXPRESS;Database=UserDB;Trusted_Connection=yes;"

Debug.Print cnn.State


End Sub
