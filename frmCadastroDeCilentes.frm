VERSION 5.00
Begin VB.Form frmCadastroDeUsuario 
   BackColor       =   &H00FF0000&
   Caption         =   "Cadastro de Usuário"
   ClientHeight    =   7200
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   10485
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7200
   ScaleWidth      =   10485
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdLimpar 
      Caption         =   "Limpar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3330
      TabIndex        =   12
      Top             =   5460
      Width           =   1785
   End
   Begin VB.CommandButton cmdExcluir 
      Caption         =   "Excluir"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   5280
      TabIndex        =   11
      Top             =   5460
      Width           =   1635
   End
   Begin VB.CommandButton cmdAlterar 
      Caption         =   "Alterar"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   5280
      TabIndex        =   10
      Top             =   4440
      Width           =   1605
   End
   Begin VB.ListBox lstUsuarios 
      Height          =   2205
      Left            =   120
      TabIndex        =   9
      Top             =   4740
      Width           =   2175
   End
   Begin VB.CommandButton cmdSair 
      Caption         =   "Sair"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7770
      TabIndex        =   8
      Top             =   4920
      Width           =   1815
   End
   Begin VB.CommandButton cmdCadastrar 
      Caption         =   "Cadastrar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3330
      TabIndex        =   7
      Top             =   4410
      Width           =   1815
   End
   Begin VB.TextBox txtConfirmarSenha 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      IMEMode         =   3  'DISABLE
      Left            =   7470
      MaxLength       =   15
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   2250
      Width           =   2325
   End
   Begin VB.TextBox txtSenha 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      IMEMode         =   3  'DISABLE
      Left            =   4140
      MaxLength       =   15
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   2280
      Width           =   2445
   End
   Begin VB.TextBox txtLogin 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   750
      MaxLength       =   15
      TabIndex        =   1
      Top             =   2250
      Width           =   2475
   End
   Begin VB.Label lblMensagem 
      BackColor       =   &H00FF0000&
      Caption         =   "* A senha deve conter uma letra maiúscula, uma minúscula, um número e um caractere especial"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1005
      Left            =   4080
      TabIndex        =   13
      Top             =   3150
      Visible         =   0   'False
      Width           =   2565
   End
   Begin VB.Label lblConfirmarSenha 
      BackColor       =   &H00FF0000&
      Caption         =   "Confirmar Senha"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   7830
      TabIndex        =   6
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Label lblSenha 
      BackColor       =   &H00FF0000&
      Caption         =   "Senha"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   4650
      TabIndex        =   5
      Top             =   1530
      Width           =   1215
   End
   Begin VB.Label lblLogin 
      BackColor       =   &H00FF0000&
      Caption         =   "Login"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   1440
      TabIndex        =   4
      Top             =   1560
      Width           =   1005
   End
   Begin VB.Label lblCadastroDeUsuario 
      BackColor       =   &H00FF0000&
      Caption         =   "Cadastro de Usuario"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   2880
      TabIndex        =   0
      Top             =   540
      Width           =   4575
   End
End
Attribute VB_Name = "frmCadastroDeUsuario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SQL As String
Dim CaracterE As Boolean
Dim CaracterL As Boolean
Dim CaracterN As Boolean
Private Sub cmdAlterar_Click()
SQL = "Update Autenticacao Set Senha = '" & txtSenha.Text & "' Where Login = '" & txtlogin.Text & "' "
Banco.Execute (SQL)
Unload Me
frmCadastroDeUsuario.Show
End Sub
Private Sub cmdCadastrar_Click()
If txtSenha.Text <> txtConfirmarSenha.Text Or txtSenha.Text = "" Or txtlogin.Text = "" Or txtConfirmarSenha.Text = "" Then
    Call MsgBox("Senhas diferentes ou campos em branco!")
End If
SQL = "Select * From Autenticacao Where Senha = '" & txtSenha.Text & "'"
Set Tabela_Dinamica = Banco.Execute(SQL)
If Not Tabela_Dinamica.EOF Then
Call MsgBox("Essa senha já existe!")
Else
CaracterE = False
CaracterL = False
CaracterN = False
For X = 1 To Len(txtSenha.Text)
A = Mid(txtSenha.Text, X, 1)
    If A = "@" Or A = "!" Or A = "?" Or A = "*" Then
        CaracterE = True
    ElseIf A = "a" Or A = "b" Or A = "c" Or A = "d" Or A = "e" Or A = "f" Or A = "g" Or A = "h" Or A = "i" Or A = "j" Then
        CaracterL = True
    ElseIf A = "1" Or A = "2" Or A = "3" Then
        CaracterN = True
    End If
Next
If Not CaracterE = True Or Not CaracterL = True Or Not CaracterN = True Then
    C = 1
    txtSenha.SetFocus
    txtSenha.BackColor = &HFF&
    txtConfirmarSenha.BackColor = &HFF&
Else
    txtSenha.BackColor = &H80000005
    txtConfirmarSenha.BackColor = &H80000005
    C = 0
    SQL = " Insert Into Autenticacao(Login,Senha) Values ('" & txtlogin.Text & "', '" & txtSenha.Text & "')"
    Banco.Execute (SQL)
    Call MsgBox("Cadastrado!")
    Unload Me
    frmCadastroDeUsuario.Show
    cmdCadastrar.Enabled = True
End If
End If
End Sub
Private Sub cmdExcluir_Click()
SQL = "Delete * From Autenticacao Where Login = '" & txtlogin.Text & "'"
Banco.Execute (SQL)
Unload Me
frmCadastroDeUsuario.Show
cmdCadastrar.Enabled = True
End Sub
Private Sub cmdLimpar_Click()
txtlogin.Text = ""
txtConfirmarSenha.Text = ""
txtSenha.Text = ""
End Sub
Private Sub cmdSair_Click()
If MsgBox("Deseja realmente Sair?", vbYesNo + vbQuestion, "Sair") = vbYes Then
    End
Else
    Cancel = True
End If
End Sub
Private Sub Form_Load()
SQL = "Select * From Autenticacao"
Set TabelaDinamica = Banco.Execute(SQL)
If Not TabelaDinamica.EOF Then
    Do While Not TabelaDinamica.EOF
        lstUsuarios.AddItem TabelaDinamica("Login") & "-" & TabelaDinamica("Senha")
        TabelaDinamica.MoveNext
    Loop
End If
End Sub
Private Sub lstUsuarios_Click()
A = InStr(1, lstUsuarios, "-")
B = Left(lstUsuarios, A - 1)
 SQL = "Select * From Autenticacao Where Login = '" & B & "'"
Set TabelaDinamica = Banco.Execute(SQL)
If TabelaDinamica.EOF Then
    Call MsgBox("Usuário não encontrado!")
Else
    txtlogin.Text = TabelaDinamica("Login")
    txtSenha.Text = TabelaDinamica("Senha")
    cmdAlterar.Enabled = True
    cmdExcluir.Enabled = True
    cmdCadastrar.Enabled = False
End If
End Sub
Private Sub txtSenha_GotFocus()
If C <> 1 Then
    lblMensagem.Visible = True
    lblMensagem.ForeColor = &H80000012
Else
    lblMensagem.Visible = True
    lblMensagem.ForeColor = &HFF&
End If
End Sub
Private Sub txtSenha_LostFocus()
If C <> 1 Then
    lblMensagem.Visible = False
End If
End Sub
