Attribute VB_Name = "Sql"
Option Explicit

Public cnn As New ADODB.Connection
Public SQL As String

Public Sub ConectaDB()

If cnn.State = 0 Then

cnn.Open "Provider=SQLNCLI11;Server=DESKTOP-84NKS8P\SQLEXPRESS;Database=UserDB;Trusted_Connection=yes;"

End If

Debug.Print cnn.State



End Sub

Sub NovoCadastro()
Dim Pl As Worksheet

Set Pl = Planilha2


ConectaDB

SQL = "INSERT INTO USERSDB ("
SQL = SQL & "      Usuario,"
SQL = SQL & "       senha)"
SQL = SQL & "VALUES "
SQL = SQL & " ('" & Pl.Range("A2").Value & "',"
SQL = SQL & " '" & Pl.Range("B2").Value & "');"



cnn.Execute SQL

cnn.Close


End Sub
