VERSION 5.00
Begin VB.Form frmCadastrosFornecedores 
   BackColor       =   &H00FF0000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastros de Fornecedores"
   ClientHeight    =   8220
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9705
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8220
   ScaleWidth      =   9705
   Begin VB.Frame fraCNPJ 
      BackColor       =   &H00FF0000&
      Caption         =   "CNPJ"
      Height          =   1245
      Left            =   3060
      TabIndex        =   17
      Top             =   1110
      Width           =   3105
      Begin VB.TextBox txtCNPJ 
         Height          =   765
         Left            =   360
         Locked          =   -1  'True
         MaxLength       =   18
         TabIndex        =   18
         Top             =   360
         Width           =   2355
      End
   End
   Begin VB.Frame fraDados 
      BackColor       =   &H00FF0000&
      Caption         =   "Dados"
      Enabled         =   0   'False
      Height          =   2445
      Left            =   960
      TabIndex        =   6
      Top             =   2580
      Width           =   7005
      Begin VB.TextBox txtTelefone 
         Height          =   375
         Left            =   2940
         MaxLength       =   15
         TabIndex        =   21
         Top             =   1800
         Width           =   1455
      End
      Begin VB.TextBox txtEmpresa 
         Height          =   525
         Left            =   210
         MaxLength       =   40
         TabIndex        =   11
         Top             =   600
         Width           =   1815
      End
      Begin VB.TextBox txtEndereco 
         Height          =   525
         Left            =   2520
         MaxLength       =   50
         TabIndex        =   10
         Top             =   600
         Width           =   1665
      End
      Begin VB.TextBox txtCidade 
         Height          =   525
         Left            =   4800
         MaxLength       =   30
         TabIndex        =   9
         Top             =   600
         Width           =   1725
      End
      Begin VB.ComboBox cmbEstado 
         Height          =   315
         ItemData        =   "frmCadastrosProdutos.frx":0000
         Left            =   210
         List            =   "frmCadastrosProdutos.frx":0055
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1830
         Width           =   2565
      End
      Begin VB.TextBox txtEmail 
         Height          =   435
         Left            =   4770
         MaxLength       =   50
         TabIndex        =   7
         Top             =   1770
         Width           =   1965
      End
      Begin VB.Label lblTelefone 
         BackColor       =   &H00FF0000&
         Caption         =   "Telefone"
         Height          =   345
         Left            =   3300
         TabIndex        =   22
         Top             =   1470
         Width           =   765
      End
      Begin VB.Label lblEmpresa 
         BackColor       =   &H00FF0000&
         Caption         =   "Empresa"
         Height          =   315
         Left            =   750
         TabIndex        =   16
         Top             =   270
         Width           =   615
      End
      Begin VB.Label lblEndereco 
         BackColor       =   &H00FF0000&
         Caption         =   "Endereço"
         Height          =   315
         Left            =   3000
         TabIndex        =   15
         Top             =   300
         Width           =   795
      End
      Begin VB.Label lblCidade 
         BackColor       =   &H00FF0000&
         Caption         =   "Cidade"
         Height          =   315
         Left            =   5400
         TabIndex        =   14
         Top             =   300
         Width           =   615
      End
      Begin VB.Label lblEstado 
         BackColor       =   &H00FF0000&
         Caption         =   "Estado"
         Height          =   345
         Left            =   1140
         TabIndex        =   13
         Top             =   1500
         Width           =   675
      End
      Begin VB.Label lblEmail 
         BackColor       =   &H00FF0000&
         Caption         =   "Email"
         Height          =   315
         Left            =   5520
         TabIndex        =   12
         Top             =   1470
         Width           =   585
      End
   End
   Begin VB.ListBox lstFornecedores 
      Height          =   1230
      Left            =   960
      TabIndex        =   5
      Top             =   6420
      Width           =   1875
   End
   Begin VB.CommandButton cmdIncluir 
      Caption         =   "Incluir"
      Enabled         =   0   'False
      Height          =   495
      Left            =   3870
      TabIndex        =   4
      Top             =   5430
      Width           =   1515
   End
   Begin VB.CommandButton cmdAlterar 
      Caption         =   "Alterar"
      Enabled         =   0   'False
      Height          =   495
      Left            =   3870
      TabIndex        =   3
      Top             =   6030
      Width           =   1515
   End
   Begin VB.CommandButton cmdExcluir 
      Caption         =   "Excluir"
      Enabled         =   0   'False
      Height          =   495
      Left            =   3870
      TabIndex        =   2
      Top             =   6600
      Width           =   1515
   End
   Begin VB.CommandButton cmdLimpar 
      Caption         =   "Limpar"
      Height          =   525
      Left            =   3870
      TabIndex        =   1
      Top             =   7170
      Width           =   1515
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   30
      Left            =   510
      TabIndex        =   0
      Top             =   6840
      Width           =   30
   End
   Begin VB.Label lblFornecedoresCadastrados 
      BackColor       =   &H00FF0000&
      Caption         =   "Fornecedores Cadastrados"
      Height          =   285
      Left            =   930
      TabIndex        =   20
      Top             =   5940
      Width           =   1935
   End
   Begin VB.Label lblCadastroProdutos 
      BackColor       =   &H00FF0000&
      Caption         =   "Cadastro de Fornecedores"
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
      Left            =   1200
      TabIndex        =   19
      Top             =   180
      Width           =   7155
   End
End
Attribute VB_Name = "frmCadastrosFornecedores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAlterar_Click()
'Altera o Cadastro!'
SQL = "Update Fornecedores Set Empresa = '" & txtEmpresa.Text & "',Endereco = '" & txtEndereco.Text & "', Cidade = '" & txtCidade.Text & "', Estado = '" & cmbEstado.Text & "', Telefone = '" & txtTelefone.Text & "' , Email = '" & txtEmail.Text & "' Where Codigo = '" & txtCodigo.Text & "' "
Banco.Execute (SQL)
Unload Me
frmCadastrosFornecedores.Show
End Sub

Private Sub cmdExcluir_Click()
'Deleta o Cadastro'
SQL = "Delete *From Fornecedores Where CNPJ = '" & txtCNPJ.Text & "'"
Banco.Execute (SQL)
Unload Me
frmCadastrosFornecedores.Show
End Sub
Private Sub cmdIncluir_Click()
If txtEmpresa.Text = "" Or txtEndereco.Text = "" Or txtCidade.Text = "" Or cmbEstado.Text = "" Or txtTelefone.Text = "" Or txtEmail.Text = "" Then
 Call MsgBox("Existem Campos em Branco!")
    Else
        SQL = " Insert Into Fornecedores (CNPJ, Empresa, Endereco, Cidade, Estado, Telefone, Email) Values ('" & txtCNPJ.Text & "', '" & txtEmpresa.Text & "', '" & txtEndereco.Text & "', '" & txtCidade.Text & "', '" & cmbEstado.Text & "', '" & txtTelefone.Text & "' , '" & txtEmail.Text & "')"
        'Insere os registros no SQL'
        Banco.Execute (SQL)
        'Fecha o Form e abre de novo para economizar linhas'
        Unload Me
        frmCadastrosFornecedores.Show
End If
End Sub
Private Sub cmdLimpar_Click()
'Limpa Tudo'
If MsgBox("Deseja realmente Limpar?", vbYesNo + vbQuestion, "Limpar") = vbYes Then
    txtCNPJ.Text = ""
    txtEmpresa.Text = ""
    txtEndereco.Text = ""
    txtCidade.Text = ""
    txtTelefone.Text = ""
    txtEmail.Text = ""
    fraCNPJ.Enabled = True
    fraDados.Enabled = False
Else
    Cancel = True
End If
End Sub
Private Sub Form_Load()
txtCNPJ.Text = Year(Now) & Month(Now) & Day(Now) & Hour(Now) & Minute(Now) & Second(Now)
'Rotina para Carregar a List'
SQL = "Select * From Fornecedores"
Set TabelaDinamica = Banco.Execute(SQL)
If Not TabelaDinamica.EOF Then
    Do While Not TabelaDinamica.EOF
        lstFornecedores.AddItem TabelaDinamica("CNPJ") & "-" & TabelaDinamica("Empresa")
        TabelaDinamica.MoveNext
    Loop
End If
End Sub
Private Sub lstFornecedores_Click()
'Rotina para carregar a List'
A = InStr(1, lstFornecedores, "-")
B = Left(lstFornecedores, A - 1)
 SQL = "Select * From Fornecedores Where CNPJ = '" & B & "'"
Set TabelaDinamica = Banco.Execute(SQL)
If TabelaDinamica.EOF Then
    Call MsgBox("CNPJ não encontrado")
Else
txtCNPJ.Text = TabelaDinamica("CNPJ")
txtEmpresa.Text = TabelaDinamica("Empresa")
txtEndereco.Text = TabelaDinamica("Endereco")
txtCidade.Text = TabelaDinamica("Cidade")
cmbEstado.Text = TabelaDinamica("Estado")
txtTelefone.Text = TabelaDinamica("Telefone")
txtEmail.Text = TabelaDinamica("Email")
fraDados.Enabled = True
cmdIncluir.Enabled = True
cmdExcluir.Enabled = True
cmdAlterar.Enabled = True
End If
End Sub

Private Sub txtCNPJ_KeyPress(KeyAscii As Integer)
'Quando o Usuário Pressionar Enter Desbloqueia o fraDados'
If KeyAscii = 13 Then
SQL = "Select * From Fornecedores Where CNPJ = '" & txtCNPJ.Text & "'"
Set Tabela_Dinamica = Banco.Execute(SQL)
    If Tabela_Dinamica.EOF Then
        fraDados.Enabled = True
        cmdIncluir.Enabled = True
        fraCNPJ.Enabled = False
    Else
        Call MsgBox("Esse CNPJ Já Existe")
    End If
End If
End Sub
