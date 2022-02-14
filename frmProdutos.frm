VERSION 5.00
Begin VB.Form frmProdutos 
   BackColor       =   &H00FF0000&
   Caption         =   "Produtos"
   ClientHeight    =   8670
   ClientLeft      =   60
   ClientTop       =   705
   ClientWidth     =   11685
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8670
   ScaleWidth      =   11685
   WindowState     =   2  'Maximized
   Begin VB.ListBox lstProdutosCadastrados 
      Height          =   1620
      Left            =   510
      TabIndex        =   23
      Top             =   6870
      Width           =   3075
   End
   Begin VB.CommandButton cmdLimpar 
      Caption         =   "Limpar"
      Height          =   705
      Left            =   7230
      TabIndex        =   22
      Top             =   6750
      Width           =   1455
   End
   Begin VB.CommandButton cmdAlterar 
      Caption         =   "Alterar"
      Enabled         =   0   'False
      Height          =   705
      Left            =   5490
      TabIndex        =   21
      Top             =   7620
      Width           =   1455
   End
   Begin VB.CommandButton cmdExcluir 
      Caption         =   "Excluir"
      Enabled         =   0   'False
      Height          =   705
      Left            =   7230
      TabIndex        =   20
      Top             =   7620
      Width           =   1455
   End
   Begin VB.CommandButton cmdIncluir 
      Caption         =   "Incluir"
      Enabled         =   0   'False
      Height          =   705
      Left            =   5490
      TabIndex        =   19
      Top             =   6750
      Width           =   1455
   End
   Begin VB.Frame fraDadosDoProduto 
      BackColor       =   &H00FF0000&
      Caption         =   "Dados Do Produto"
      Enabled         =   0   'False
      Height          =   3345
      Left            =   450
      TabIndex        =   3
      Top             =   2520
      Width           =   9465
      Begin VB.TextBox txtDataValidade 
         Height          =   555
         Left            =   300
         TabIndex        =   14
         Top             =   2520
         Width           =   1725
      End
      Begin VB.TextBox txtValorUnitario 
         Height          =   615
         Left            =   2460
         TabIndex        =   13
         Top             =   2490
         Width           =   1725
      End
      Begin VB.TextBox txtPorcentagemLucro 
         Height          =   615
         Left            =   4650
         TabIndex        =   12
         Top             =   2490
         Width           =   2145
      End
      Begin VB.TextBox txtValorVenda 
         Height          =   615
         Left            =   7320
         TabIndex        =   11
         Top             =   2520
         Width           =   1725
      End
      Begin VB.ComboBox cmbSetor 
         Height          =   315
         Left            =   6360
         TabIndex        =   10
         Top             =   810
         Width           =   2055
      End
      Begin VB.ComboBox cmbFornecedor 
         Height          =   315
         ItemData        =   "frmProdutos.frx":0000
         Left            =   3390
         List            =   "frmProdutos.frx":0002
         TabIndex        =   6
         Top             =   810
         Width           =   2415
      End
      Begin VB.TextBox txtDescricao 
         Height          =   615
         Left            =   360
         TabIndex        =   4
         Top             =   720
         Width           =   2505
      End
      Begin VB.Label lblDataValidade 
         BackColor       =   &H00FF0000&
         Caption         =   "Data Validade"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   180
         TabIndex        =   18
         Top             =   1890
         Width           =   2025
      End
      Begin VB.Label lblValorUnitario 
         BackColor       =   &H00FF0000&
         Caption         =   "Valor Unitário"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2400
         TabIndex        =   17
         Top             =   1890
         Width           =   1845
      End
      Begin VB.Label lblPorcentagemLucro 
         BackColor       =   &H00FF0000&
         Caption         =   "Porcentagem Lucro"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4470
         TabIndex        =   16
         Top             =   1920
         Width           =   2475
      End
      Begin VB.Label lblValorVenda 
         BackColor       =   &H00FF0000&
         Caption         =   "Valor Venda"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7320
         TabIndex        =   15
         Top             =   1950
         Width           =   1755
      End
      Begin VB.Label lblSetor 
         BackColor       =   &H00FF0000&
         Caption         =   "Setor"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6930
         TabIndex        =   9
         Top             =   390
         Width           =   945
      End
      Begin VB.Label lblFornecedor 
         BackColor       =   &H00FF0000&
         Caption         =   "Fornecedor"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3570
         TabIndex        =   8
         Top             =   360
         Width           =   2025
      End
      Begin VB.Label lblDescricao 
         BackColor       =   &H00FF0000&
         Caption         =   "Descrição"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   660
         TabIndex        =   5
         Top             =   270
         Width           =   1725
      End
   End
   Begin VB.Frame fraCodigo 
      BackColor       =   &H00FF0000&
      Caption         =   "Código"
      Height          =   1215
      Left            =   510
      TabIndex        =   1
      Top             =   1020
      Width           =   3525
      Begin VB.TextBox txtCodigo 
         Height          =   675
         Left            =   240
         MaxLength       =   13
         TabIndex        =   2
         Top             =   300
         Width           =   2955
      End
   End
   Begin VB.Label lblProdutosCadastrados 
      BackColor       =   &H00FF0000&
      Caption         =   "Produtos Cadastrados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   210
      TabIndex        =   24
      Top             =   6210
      Width           =   3915
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF0000&
      Caption         =   "Descrição"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3150
      TabIndex        =   7
      Top             =   3150
      Width           =   1635
   End
   Begin VB.Label lblCadastro 
      BackColor       =   &H00FF0000&
      Caption         =   "Cadastro De Produtos"
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
      Left            =   3720
      TabIndex        =   0
      Top             =   150
      Width           =   3615
   End
   Begin VB.Menu mnuVoltar 
      Caption         =   "Voltar"
   End
End
Attribute VB_Name = "frmProdutos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAlterar_Click()
If txtDescricao.Text = "" Or txtDataValidade.Text = "" Or txtValorUnitario.Text = "" Or txtPorcentagemLucro.Text = "" Or txtValorVenda.Text = "" Or cmbFornecedor.Text = "" Or cmbSetor.Text = "" Then
    Call MsgBox("Existem campos em branco!")
Else
    SQL = "Update Produtos Set Descricao = '" & txtDescricao.Text & "',Fornecedor = '" & cmbFornecedor.Text & "', Setor = '" & cmbSetor.Text & "', DataValidade = '" & txtDataValidade.Text & "', ValorUnitario = '" & txtValorUnitario.Text & "' , PorcentagemLucro = '" & txtPorcentagemLucro & "',ValorVenda = '" & txtValorVenda & "' Where Codigo = '" & txtCodigo.Text & "' "
    Set TabelaDinamica = Banco.Execute(SQL)
    Unload Me
    frmProdutos.Show
    End If
End Sub
Private Sub cmdExcluir_Click()
SQL = "Delete * From Produtos Where Codigo = '" & txtCodigo.Text & "'"
Set TabelaDinamica = Banco.Execute(SQL)
Unload Me
frmProdutos.Show
End Sub
Private Sub cmdIncluir_Click()
If txtDescricao.Text = "" Or txtDataValidade.Text = "" Or txtValorUnitario.Text = "" Or txtPorcentagemLucro.Text = "" Or txtValorVenda.Text = "" Or cmbFornecedor.Text = "" Or cmbSetor.Text = "" Then
    Call MsgBox("Existem campos em branco!")
Else
    SQL = "Select Codigo From Fornecedores Where Empresa = cmbFornecedor.Text "
    Set TabelaDinamica = Banco.Execute(SQL)
    SQL = "Select Codigo From Setores Where Nome = cmbSetor.Text "
    Set TabelaDinamica = Banco.Execute(SQL)
    SQL = "Insert Into Produtos (Codigo,Descricao,Fornecedor,Setor,DataValidade,ValorUnitario,PorcentagemLucro,ValorVenda) Values ('" & txtCodigo.Text & "','" & txtDescricao.Text & "','" & (SQL1) & "','" & (SQL2) & "','" & txtDataValidade.Text & "', '" & txtValorUnitario.Text & "', '" & txtPorcentagemLucro.Text & "', '" & txtValorVenda.Text & "')"
    Set TabelaDinamica = Banco.Execute(SQL)
    Unload Me
    frmProdutos.Show
End If
End Sub
Private Sub cmdLimpar_Click()
If fraCodigo.Enabled = True Then
    txtCodigo.Text = ""
Else
    txtDescricao.Text = ""
    txtDataValidade.Text = ""
    txtValorUnitario.Text = ""
    txtPorcentagemLucro.Text = ""
    txtValorVenda.Text = ""
End If
End Sub
Private Sub Form_Load()
SQL = "Select * From Fornecedores"
Set TabelaDinamica = Banco.Execute(SQL)
If Not TabelaDinamica.EOF Then
    Do While Not TabelaDinamica.EOF
        cmbFornecedor.AddItem TabelaDinamica("Empresa")
        TabelaDinamica.MoveNext
    Loop
End If
SQL = "Select * From Setores "
Set TabelaDinamica = Banco.Execute(SQL)
If Not TabelaDinamica.EOF Then
    Do While Not TabelaDinamica.EOF
        cmbSetor.AddItem TabelaDinamica("Nome")
        TabelaDinamica.MoveNext
    Loop
End If
SQL = "Select * From Produtos"
Set TabelaDinamica = Banco.Execute(SQL)
If Not TabelaDinamica.EOF Then
    Do While Not TabelaDinamica.EOF
        lstProdutosCadastrados.AddItem TabelaDinamica("Codigo") & "-" & TabelaDinamica("Descricao")
        TabelaDinamica.MoveNext
    Loop
End If
End Sub
Private Sub lstProdutosCadastrados_Click()
A = InStr(1, lstProdutosCadastrados, "-")
B = Left(lstProdutosCadastrados, A - 1)
SQL = "Select * From Produtos Where Codigo = '" & B & "'"
Set TabelaDinamica = Banco.Execute(SQL)
If TabelaDinamica.EOF Then
    Call MsgBox("Código não encontrado")
Else
    txtCodigo.Text = TabelaDinamica("Codigo")
    txtDescricao.Text = TabelaDinamica("Descricao")
    cmbFornecedor.Text = TabelaDinamica("Fornecedor")
    cmbSetor.Text = TabelaDinamica("Setor")
    txtDataValidade.Text = TabelaDinamica("DataValidade")
    txtValorUnitario.Text = TabelaDinamica("ValorUnitario")
    txtPorcentagemLucro.Text = TabelaDinamica("PorcentagemLucro")
    txtValorVenda.Text = TabelaDinamica("ValorVenda")
    cmdExcluir.Enabled = True
    cmdAlterar.Enabled = True
    fraCodigo.Enabled = False
    fraDadosDoProduto.Enabled = True
End If
End Sub
Private Sub mnuVoltar_Click()
Unload Me
End Sub
Private Sub txtCodigo_Change()
If Len(txtCodigo.Text) = 13 Then
    fraCodigo.Enabled = False
    fraDadosDoProduto.Enabled = True
    cmdIncluir.Enabled = True
    SQL = "Select * From Produtos Where Codigo = '" & txtCodigo.Text & "'"
    Set Tabela_Dinamica = Banco.Execute(SQL)
    If Tabela_Dinamica.EOF Then
        fraDadosDoProduto.Enabled = True
        cmdIncluir.Enabled = True
        fraCodigo.Enabled = False
    Else
        Call MsgBox("Esse Codigo Já Existe")
    End If
End If
End Sub
