VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmTerminal 
   BackColor       =   &H00FF0000&
   Caption         =   "Terminal"
   ClientHeight    =   8925
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   14175
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8925
   ScaleWidth      =   14175
   Begin MSFlexGridLib.MSFlexGrid flxItens 
      Height          =   8445
      Left            =   4830
      TabIndex        =   15
      Top             =   420
      Width           =   9285
      _ExtentX        =   16378
      _ExtentY        =   14896
      _Version        =   393216
   End
   Begin VB.Frame fraCodigo 
      BackColor       =   &H00FF0000&
      Caption         =   "Codigo"
      Height          =   1605
      Left            =   300
      TabIndex        =   10
      Top             =   540
      Width           =   4185
      Begin VB.TextBox txtCodigo 
         Height          =   825
         Left            =   750
         TabIndex        =   11
         Top             =   480
         Width           =   2685
      End
   End
   Begin VB.Frame fraQuantidade 
      BackColor       =   &H00FF0000&
      Caption         =   "Quantidade"
      Height          =   1665
      Left            =   3150
      TabIndex        =   8
      Top             =   6840
      Width           =   1545
      Begin VB.Label lblQuantidade 
         BackColor       =   &H00FF0000&
         Caption         =   "1"
         Height          =   555
         Left            =   330
         TabIndex        =   9
         Top             =   600
         Width           =   885
      End
   End
   Begin VB.Frame fraProdutos 
      BackColor       =   &H00FF0000&
      Caption         =   "Produtos"
      Height          =   1815
      Left            =   270
      TabIndex        =   3
      Top             =   2730
      Width           =   4395
      Begin VB.Label lblDescricao 
         BackColor       =   &H00FF0000&
         Caption         =   "Descrição"
         Height          =   315
         Left            =   390
         TabIndex        =   7
         Top             =   390
         Width           =   825
      End
      Begin VB.Label lblValorUnitario 
         BackColor       =   &H00FF0000&
         Caption         =   "Valor Unitario"
         Height          =   285
         Left            =   2700
         TabIndex        =   6
         Top             =   420
         Width           =   1065
      End
      Begin VB.Label lblDescricao1 
         BackColor       =   &H00FF0000&
         Height          =   555
         Left            =   300
         TabIndex        =   5
         Top             =   1020
         Width           =   1065
      End
      Begin VB.Label lblValorUnitario1 
         BackColor       =   &H00FF0000&
         Height          =   615
         Left            =   2580
         TabIndex        =   4
         Top             =   840
         Width           =   1275
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF0000&
      Caption         =   "Cliente"
      Height          =   1425
      Left            =   270
      TabIndex        =   0
      Top             =   5100
      Width           =   4425
      Begin VB.Label lblNome 
         BackColor       =   &H00FF0000&
         Caption         =   "Nome"
         Height          =   345
         Left            =   180
         TabIndex        =   2
         Top             =   390
         Width           =   945
      End
      Begin VB.Label lblClientes 
         BackColor       =   &H00FF0000&
         Height          =   555
         Left            =   1200
         TabIndex        =   1
         Top             =   600
         Width           =   2505
      End
   End
   Begin VB.Label lblItensCadastrados 
      BackColor       =   &H00FF0000&
      Caption         =   "Itens Cadastrados"
      Height          =   225
      Left            =   8430
      TabIndex        =   14
      Top             =   60
      Width           =   1395
   End
   Begin VB.Label lblValorTotal 
      BackColor       =   &H00FF0000&
      Caption         =   "Valor Total"
      Height          =   435
      Left            =   150
      TabIndex        =   13
      Top             =   6900
      Width           =   945
   End
   Begin VB.Label lblValorTotal1 
      BackColor       =   &H00FF0000&
      Height          =   495
      Left            =   1320
      TabIndex        =   12
      Top             =   6900
      Width           =   1695
   End
End
Attribute VB_Name = "frmTerminal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
flxItens.TextMatrix(0, 0) = "Codigo"
flxItens.TextMatrix(0, 1) = "Descricao"
flxItens.TextMatrix(0, 2) = "VL Unit"
flxItens.TextMatrix(0, 3) = "Quantidade"
flxItens.TextMatrix(0, 4) = "Valor Total"
flxItens.Coldwidth(0) = 2000
flxItens.Coldwidth(1) = 4000
End Sub
Private Sub mnuVoltar_Click()
Unload Me
End Sub
Private Sub txtCodigo_Change()
lblDescricao.Caption = Produtos("Descricao")
lblValorUnitario.Caption = Produtos("Valor Unitario")
flxItens.TextMatrix(flxItens.Rows - 1, 0) = TabelaDinamica("Codigo")
flxItens.TextMatrix(flxItens.Rows - 1, 1) = TabelaDinamica("Descricao")
flxItens.TextMatrix(flxItens.Rows - 1, 2) = TabelaDinamica("ValorUnitario")
flxItens.TextMatrix(flxItens.Rows - 1, 3) = lblQuantidade.Text
flxItens.TextMatrix(flxItens.Rows - 1, 4) = TabelaDinamica("Valor Venda")
End Sub
Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then
    lblQuantidade.Caption = InputBox("Digite a quantidade a ser comprada!", "Caixa Registradora")
ElseIf KeyCode = 114 Then
    X = InputBox("Digite o código do cliente!")
    SQL = "Select * From Clientes Where Codigo = '" & X & "' "
    Set TabelaDinamica = Banco.Execute(SQL)
        If TabelaDinamica.EOF Then
            Call MsgBox("Cliente não encontrado!")
        Else
        lblClientes.Caption = Tabela("Nome")
        End If
End If
End Sub

