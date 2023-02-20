VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2 
   Caption         =   "UserForm2"
   ClientHeight    =   10845
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15465
   OleObjectBlob   =   "UserForm2.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
Private Sub btncadastro_Click()
 
 
    If senha_add <> senha_add2 Then
 
        ConectaDB
 

    
        SQL = "INSERT INTO USERSDB ("
        SQL = SQL & "      Usuario,"
        SQL = SQL & "       senha)"
        SQL = SQL & "VALUES "
        SQL = SQL & " ('" & .user_add.Value & "',"
        SQL = SQL & " '" & .senha_add.Value & "');"



        cnn.Execute SQL
    
    Else

    MsgBox "Senhas não coincidem", vbInformation, ""
    
End If
End With

cnn.Close

MsgBox "USER CADASTRADO !", vbInformation, ""


End Sub
Sub comp()

If senha_add <> senha_add Then

Call btncadastro_Click

Else

MsgBox "Senhas não coincidem", vbInformation, ""

End If



End Sub

Private Sub edit_Click()

End Sub

Private Sub Label3_Click()

End Sub

Private Sub sair_Click()

Unload UserForm2



End Sub

Private Sub senha_add2_Change()

End Sub

Private Sub UserForm_Initialize()

MakeUserformTransparent Me
HideTitleBarAndBordar Me


End Sub

