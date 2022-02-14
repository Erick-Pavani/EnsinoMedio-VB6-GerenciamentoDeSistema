VERSION 5.00
Begin VB.MDIForm MDIMenu 
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   7620
   ClientLeft      =   165
   ClientTop       =   810
   ClientWidth     =   13260
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu mnuCadastros 
      Caption         =   "Cadastros"
      Begin VB.Menu mnuClientes 
         Caption         =   "Clientes"
      End
      Begin VB.Menu mnuFornecedores 
         Caption         =   "Fornecedores"
      End
      Begin VB.Menu mnuSetores 
         Caption         =   "Setores "
      End
      Begin VB.Menu mnuCadastroDeUsuario 
         Caption         =   "Cadastro De Usuário"
      End
      Begin VB.Menu mnuCadastroDeProdutos 
         Caption         =   "Cadastro De Produtos"
      End
      Begin VB.Menu mnuProdutos 
         Caption         =   "Produtos"
      End
      Begin VB.Menu mnuTerminal 
         Caption         =   "Terminal"
      End
   End
   Begin VB.Menu mnuSair 
      Caption         =   "Sair"
   End
End
Attribute VB_Name = "MDIMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MDIForm_Load()
Call MsgBox("Bem vindo " & frmAutenticacao.txtLogin.Text & "!")
End Sub
Private Sub mnuCadastroDeUsuario_Click()
frmCadastroDeUsuario.Show
End Sub
Private Sub mnuClientes_Click()
frmClientes.Show
End Sub
Private Sub mnuFornecedores_Click()
frmCadastrosFornecedores.Show
End Sub
Private Sub mnuCadastroDeProdutos_Click()
frmProdutos.Show
End Sub
Private Sub mnuProdutos_Click()
frmProdutos.Show
End Sub
Private Sub mnuSair_Click()
If MsgBox("Deseja realmente Sair?", vbYesNo + vbQuestion, "Sair") = vbYes Then
    End
Else
    Cancel = True
End If
End Sub
Private Sub mnuSetores_Click()
frmCadastroDeSetores.Show
End Sub
Private Sub mnuTerminal_Click()
frmTerminal.Show
End Sub
