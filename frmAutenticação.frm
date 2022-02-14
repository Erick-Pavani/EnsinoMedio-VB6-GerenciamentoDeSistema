VERSION 5.00
Begin VB.Form frmAutenticacao 
   BackColor       =   &H0000FF00&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Autenticação"
   ClientHeight    =   7080
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7020
   ControlBox      =   0   'False
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7080
   ScaleWidth      =   7020
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmr1 
      Interval        =   100
      Left            =   3060
      Top             =   6420
   End
   Begin VB.CommandButton cmdSair 
      Caption         =   "Sair"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3960
      TabIndex        =   6
      Top             =   5160
      Width           =   2175
   End
   Begin VB.CommandButton cmdEntrar 
      Caption         =   "Entrar"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   360
      TabIndex        =   5
      Top             =   5160
      Width           =   2415
   End
   Begin VB.TextBox txtSenha 
      BackColor       =   &H0000FF00&
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      IMEMode         =   3  'DISABLE
      Left            =   3120
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   3480
      Width           =   3375
   End
   Begin VB.TextBox txtlogin 
      BackColor       =   &H0000FF00&
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3120
      MaxLength       =   10
      TabIndex        =   3
      Top             =   2280
      Width           =   3375
   End
   Begin VB.Label lblSenha 
      BackColor       =   &H0000FF00&
      Caption         =   "Senha"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   420
      TabIndex        =   2
      Top             =   3600
      Width           =   1575
   End
   Begin VB.Label lblLogin 
      BackColor       =   &H0000FF00&
      Caption         =   "Login"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   480
      TabIndex        =   1
      Top             =   2430
      Width           =   1425
   End
   Begin VB.Label lblAutenticação 
      BackColor       =   &H0000FF00&
      Caption         =   "Autenticação"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   525
      Left            =   1680
      TabIndex        =   0
      Top             =   540
      Width           =   3585
   End
End
Attribute VB_Name = "frmAutenticacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SQL As String
Dim Incorreto As Integer
Private Sub cmdEntrar_Click()
'Verifica se a senha ou o Usuário estão corretos'
For X = 1 To Len(txtlogin.Text)
A = Mid(txtlogin.Text, X, 1)
If A = "'" Then
    Call MsgBox("Tire o apóstrofo! (')")
    Cancel = True
    txtlogin.Text = ""
    txtSenha.Text = ""
End If
Next
SQL = " Select * From Autenticacao Where Login = '" & txtlogin.Text & "' And Senha = '" & txtSenha.Text & "'"
Set TabelaDinamica = Banco.Execute(SQL)
    If TabelaDinamica.EOF Then
        Call MsgBox("Senha ou Login incorretos!")
        'Fecha o programa se o usuário errar a senha ou usuário 3 vezes'
        Incorreto = Incorreto + 1
        txtSenha.Text = ""
        txtlogin.Text = ""
            If Incorreto = 3 Then
            Call MsgBox("Você errou a senha 3 vezes, seu programa será fechado por questões de segurança!")
            End
            End If
Else
    MDIMenu.Show
    Unload Me
End If
End Sub
Private Sub cmdSair_Click()
'Sai do programa'
If MsgBox("Deseja realmente Sair?", vbYesNo + vbQuestion, "Sair") = vbYes Then
    End
Else
    Cancel = True
End If
End Sub
Private Sub Form_Load()
'Conexão com o Banco'
Incorreto = 0
Banco.Open "Provider = microsoft.jet.oleDB.4.0;data source = " & App.Path & "\Banco.mdb"
End Sub
Private Sub tmr1_Timer()
'Faz a label Piscar'
If lblAutenticação.ForeColor = &H0& Then
    lblAutenticação.ForeColor = &HFF&
Else
    lblAutenticação.ForeColor = &H0&
End If
End Sub
Private Sub txtLogin_Change()
If txtlogin.Text = "" Or txtSenha.Text = "" Then
    cmdEntrar.Enabled = False
Else
    cmdEntrar.Enabled = True
End If
If Len(txtSenha.Text) > 10 Or Len(txtlogin.Text) > 10 Then
    Call MsgBox("Sua Senha ou Login podem ter no máximo 10 caracteres!")
    txtSenha.Text = ""
    txtlogin.Text = ""
End If
End Sub
Private Sub txtlogin_GotFocus()
txtlogin.BackColor = &H808080
End Sub
Private Sub txtlogin_LostFocus()
txtlogin.BackColor = &HFF00&
End Sub
Private Sub txtSenha_Change()
If txtSenha.Text = "" Or txtlogin.Text = "" Then
    cmdEntrar.Enabled = False
Else
    cmdEntrar.Enabled = True
End If
If Len(txtSenha.Text) > 10 Or Len(txtlogin.Text) > 10 Then
    Call MsgBox("Sua Senha e Login podem ter no máximo 10 caracteres!")
    txtSenha.Text = ""
    txtlogin.Text = ""
End If
End Sub
Private Sub txtSenha_GotFocus()
txtSenha.BackColor = &H808080
End Sub
Private Sub txtSenha_LostFocus()
txtSenha.BackColor = &HFF00&
End Sub
