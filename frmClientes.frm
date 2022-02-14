VERSION 5.00
Begin VB.Form frmClientes 
   BackColor       =   &H00FF0000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Clientes"
   ClientHeight    =   8385
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7620
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8385
   ScaleWidth      =   7620
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   30
      Left            =   1050
      TabIndex        =   20
      Top             =   7140
      Width           =   30
   End
   Begin VB.CommandButton cmdLimpar 
      Caption         =   "Limpar"
      Height          =   525
      Left            =   3210
      TabIndex        =   18
      Top             =   7650
      Width           =   1545
   End
   Begin VB.CommandButton cmdExcluir 
      Caption         =   "Excluir"
      Enabled         =   0   'False
      Height          =   495
      Left            =   3210
      TabIndex        =   17
      Top             =   7050
      Width           =   1545
   End
   Begin VB.CommandButton cmdAlterar 
      Caption         =   "Alterar"
      Enabled         =   0   'False
      Height          =   495
      Left            =   3210
      TabIndex        =   16
      Top             =   6450
      Width           =   1545
   End
   Begin VB.CommandButton cmdIncluir 
      Caption         =   "Incluir"
      Enabled         =   0   'False
      Height          =   495
      Left            =   3210
      TabIndex        =   15
      Top             =   5880
      Width           =   1545
   End
   Begin VB.ListBox lstClientes 
      Height          =   1230
      Left            =   540
      TabIndex        =   14
      Top             =   6780
      Width           =   1905
   End
   Begin VB.Frame fraDados 
      BackColor       =   &H00FF0000&
      Caption         =   "Dados"
      Enabled         =   0   'False
      Height          =   2445
      Left            =   1740
      TabIndex        =   2
      Top             =   2820
      Width           =   4335
      Begin VB.TextBox txtEmail 
         Height          =   435
         Left            =   3120
         TabIndex        =   12
         Top             =   1800
         Width           =   1095
      End
      Begin VB.ComboBox cmbEstado 
         Height          =   315
         ItemData        =   "frmClientes.frx":0000
         Left            =   210
         List            =   "frmClientes.frx":0055
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   1830
         Width           =   2565
      End
      Begin VB.TextBox txtCidade 
         Height          =   555
         Left            =   3000
         MaxLength       =   30
         TabIndex        =   8
         Top             =   660
         Width           =   1125
      End
      Begin VB.TextBox txtEndereco 
         Height          =   525
         Left            =   1620
         MaxLength       =   50
         TabIndex        =   6
         Top             =   660
         Width           =   1185
      End
      Begin VB.TextBox txtNome 
         Height          =   525
         Left            =   150
         MaxLength       =   50
         TabIndex        =   3
         Top             =   660
         Width           =   1185
      End
      Begin VB.Label lblEmail 
         BackColor       =   &H00FF0000&
         Caption         =   "Email"
         Height          =   315
         Left            =   3450
         TabIndex        =   11
         Top             =   1410
         Width           =   585
      End
      Begin VB.Label lblEstado 
         BackColor       =   &H00FF0000&
         Caption         =   "Estado"
         Height          =   345
         Left            =   1170
         TabIndex        =   10
         Top             =   1350
         Width           =   675
      End
      Begin VB.Label lblCidade 
         BackColor       =   &H00FF0000&
         Caption         =   "Cidade"
         Height          =   315
         Left            =   3240
         TabIndex        =   7
         Top             =   330
         Width           =   615
      End
      Begin VB.Label lblEndereco 
         BackColor       =   &H00FF0000&
         Caption         =   "Endereço"
         Height          =   315
         Left            =   1800
         TabIndex        =   5
         Top             =   300
         Width           =   795
      End
      Begin VB.Label lblNome 
         BackColor       =   &H00FF0000&
         Caption         =   "Nome"
         Height          =   315
         Left            =   420
         TabIndex        =   4
         Top             =   270
         Width           =   615
      End
   End
   Begin VB.Frame fraCodigo 
      BackColor       =   &H00FF0000&
      Caption         =   "Código"
      Height          =   1245
      Left            =   2340
      TabIndex        =   0
      Top             =   1320
      Width           =   3135
      Begin VB.TextBox txtCodigo 
         Height          =   765
         Left            =   360
         Locked          =   -1  'True
         MaxLength       =   14
         TabIndex        =   1
         Top             =   360
         Width           =   2355
      End
   End
   Begin VB.Label lblModuloClientes 
      BackColor       =   &H00FF0000&
      Caption         =   "Módulo Clientes"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1740
      TabIndex        =   19
      Top             =   300
      Width           =   4245
   End
   Begin VB.Label lblClientesCadastrados 
      BackColor       =   &H00FF0000&
      Caption         =   "Clientes Cadastrados"
      Height          =   285
      Left            =   660
      TabIndex        =   13
      Top             =   6180
      Width           =   1665
   End
End
Attribute VB_Name = "frmClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SQL As String
Private Sub cmdAlterar_Click()
'Altera o Cadastro!'
SQL = "Update Clientes Set Nome = '" & txtNome.Text & "',Endereco = '" & txtEndereco.Text & "', Cidade = '" & txtCidade.Text & "', Estado = '" & cmbEstado.Text & "', Email = '" & txtEmail.Text & "' Where Codigo = '" & txtCodigo.Text & "' "
Banco.Execute (SQL)
Unload Me
frmClientes.Show
End Sub
Private Sub cmdExcluir_Click()
'Deleta o Cadastro'
SQL = "Delete *From Clientes Where Codigo = '" & txtCodigo.Text & "'"
Banco.Execute (SQL)
Unload Me
frmClientes.Show
End Sub
Private Sub cmdIncluir_Click()
If txtNome.Text = "" Or txtEndereco.Text = "" Or txtCidade.Text = "" Or cmbEstado.Text = "" Or txtEmail.Text = "" Then
 Call MsgBox("Existem Campos em Branco!")
    Else
        SQL = " Insert Into Clientes (Codigo, Nome, Endereco, Cidade, Estado, Email) Values ('" & txtCodigo.Text & "', '" & txtNome.Text & "', '" & txtEndereco.Text & "', '" & txtCidade.Text & "', '" & cmbEstado.Text & "', '" & txtEmail.Text & "')"
        'Insere os registros no SQL'
        Banco.Execute (SQL)
        'Fecha o Form e abre de novo para economizar linhas'
        Unload Me
        frmClientes.Show
End If
End Sub
Private Sub cmdLimpar_Click()
'Limpa Tudo'
If MsgBox("Deseja realmente Limpar?", vbYesNo + vbQuestion, "Limpar") = vbYes Then
    txtNome.Text = ""
    txtEndereco.Text = ""
    txtCidade.Text = ""
    txtEmail.Text = ""
    txtCodigo.Text = ""
    fraDados.Enabled = False
    fraCodigo.Enabled = True
Else
    Cancel = True
End If
End Sub
Private Sub Form_Load()
'Garante que o código sempre tenha 14 digitos'
If Month(Now) < 10 Then
    txtCodigo.Text = Year(Now) & "0" & Month(Now) & Day(Now) & Hour(Now) & Minute(Now) & Second(Now)
Else
    txtCodigo.Text = Year(Now) & Month(Now) & Day(Now) & Hour(Now) & Minute(Now) & Second(Now)
End If
'Rotina para Carregar a List'
SQL = "Select * From Clientes"
Set TabelaDinamica = Banco.Execute(SQL)
If Not TabelaDinamica.EOF Then
    Do While Not TabelaDinamica.EOF
        lstClientes.AddItem TabelaDinamica("Codigo") & "-" & TabelaDinamica("Nome")
        TabelaDinamica.MoveNext
    Loop
End If
End Sub
Private Sub lstClientes_Click()
'Rotina para carregar a List'
A = InStr(1, lstClientes, "-")
B = Left(lstClientes, A - 1)
SQL = "Select * From Clientes Where Codigo = '" & B & "'"
Set TabelaDinamica = Banco.Execute(SQL)
If TabelaDinamica.EOF Then
    Call MsgBox("Código não encontrado")
Else
txtCodigo.Text = TabelaDinamica("Codigo")
txtNome.Text = TabelaDinamica("Nome")
txtEndereco.Text = TabelaDinamica("Endereco")
txtCidade.Text = TabelaDinamica("Cidade")
cmbEstado.Text = TabelaDinamica("Estado")
txtEmail.Text = TabelaDinamica("Email")
fraDados.Enabled = True
cmdIncluir.Enabled = True
cmdExcluir.Enabled = True
cmdAlterar.Enabled = True
End If
End Sub
Private Sub txtCodigo_KeyPress(KeyAscii As Integer)
'Quando o Usuário Pressionar Enter Desbloqueia o fraDados'
If KeyAscii = 13 Then
SQL = "Select * From Clientes Where Codigo = '" & txtCodigo.Text & "'"
Set Tabela_Dinamica = Banco.Execute(SQL)
    If Tabela_Dinamica.EOF Then
        fraDados.Enabled = True
        cmdIncluir.Enabled = True
        fraCodigo.Enabled = False
    Else
        Call MsgBox("Esse Código Já Existe!")
    End If
End If
End Sub
