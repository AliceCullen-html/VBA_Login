Attribute VB_Name = "Sql"
Option Explicit

Public Pl As Worksheet
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
Dim Lin As Integer
Set Pl = Planilha2

Lin = 2

ConectaDB

With Pl
    Do While .Cells(Lin, 1).Value <> ""
        SQL = "INSERT INTO USERSDB ("
        SQL = SQL & "      Usuario,"
        SQL = SQL & "       senha)"
        SQL = SQL & "VALUES "
        SQL = SQL & " ('" & .Cells(Lin, 1).Value & "',"
        SQL = SQL & " '" & .Cells(Lin, 2).Value & "');"



cnn.Execute SQL
    Lin = Lin + 1
    Loop
    

End With

cnn.Close


End Sub
