VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   12585
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   18240
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub btnlogin_Click()

Dim usuario As String
Dim senha As String
Dim combinacao As Boolean
Dim combi As String
Dim combi2 As String


usuario = Me.usuario.Value
senha = Me.senha.Value

combinacao = False

combi = usuario & senha

linhas = worksheetfunction.countA(Planilha2.columns("a"))

For cont = 2 To linhas

combi2 = Planilha2.Cells(cont, 1).Value & Planilha2.Cells(cont, 2).Value

combinacao = combi = combi2

If combinacao = True Then

MsgBox "USER AUTORIZADO !", vbInformation, ""

Unload Me

Application.Visible = True


Exit Sub
End If

Next

MsgBox "USER OU SENHA INCORRETOS !", vbCritical, ""


End Sub

Private Sub Label2_Click()

End Sub

Private Sub Label3_Click()

End Sub

Private Sub Label4_Click()

End Sub

Private Sub Label7_Click()


ThisWorkbook.Close




End Sub

Private Sub UserForm_Activate()
MakeUserformTransparent Me
HideTitleBarAndBordar Me

End Sub
