VERSION 5.00
Begin VB.Form frmCadastroDeSetores 
   BackColor       =   &H00FF0000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Setores"
   ClientHeight    =   7920
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6885
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7920
   ScaleWidth      =   6885
   Begin VB.Frame fraCodigo 
      BackColor       =   &H00FF0000&
      Caption         =   "Código Setor"
      Height          =   1245
      Left            =   420
      TabIndex        =   9
      Top             =   1890
      Width           =   2955
      Begin VB.TextBox txtCodigo 
         Height          =   585
         Left            =   420
         MaxLength       =   3
         TabIndex        =   0
         Top             =   420
         Width           =   1965
      End
   End
   Begin VB.Frame fraDados 
      BackColor       =   &H00FF0000&
      Caption         =   "Dados do Setor"
      Enabled         =   0   'False
      Height          =   1515
      Left            =   630
      TabIndex        =   7
      Top             =   3660
      Width           =   1725
      Begin VB.TextBox txtNome 
         Height          =   525
         Left            =   240
         MaxLength       =   50
         TabIndex        =   1
         Top             =   780
         Width           =   1185
      End
      Begin VB.Label lblNome 
         BackColor       =   &H00FF0000&
         Caption         =   "Nome Setor"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   210
         TabIndex        =   8
         Top             =   360
         Width           =   1275
      End
   End
   Begin VB.ListBox lstSetores 
      Height          =   1230
      Left            =   690
      TabIndex        =   6
      Top             =   6090
      Width           =   1905
   End
   Begin VB.CommandButton cmdIncluir 
      Caption         =   "Incluir"
      Enabled         =   0   'False
      Height          =   495
      Left            =   3870
      TabIndex        =   2
      Top             =   4530
      Width           =   1545
   End
   Begin VB.CommandButton cmdAlterar 
      Caption         =   "Alterar"
      Enabled         =   0   'False
      Height          =   495
      Left            =   3870
      TabIndex        =   3
      Top             =   5190
      Width           =   1545
   End
   Begin VB.CommandButton cmdExcluir 
      Caption         =   "Excluir"
      Enabled         =   0   'False
      Height          =   495
      Left            =   3870
      TabIndex        =   4
      Top             =   5790
      Width           =   1545
   End
   Begin VB.CommandButton cmdLimpar 
      Caption         =   "Limpar"
      Height          =   525
      Left            =   3870
      TabIndex        =   5
      Top             =   6420
      Width           =   1545
   End
   Begin VB.Label lblCadastroSetores 
      BackColor       =   &H00FF0000&
      Caption         =   "Cadastro Setores"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   1440
      TabIndex        =   11
      Top             =   510
      Width           =   4095
   End
   Begin VB.Label lblClientesCadastrados 
      BackColor       =   &H00FF0000&
      Caption         =   "Setores Cadastrados"
      Height          =   285
      Left            =   960
      TabIndex        =   10
      Top             =   5520
      Width           =   1665
   End
End
Attribute VB_Name = "frmCadastroDeSetores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SQL As String
Private Sub cmdAlterar_Click()
If txtNome.Text = "" Then
    Call MsgBox("Digite alguma coisa na caixa texto!")
Else
SQL = "Update Setores Set Nome = '" & txtNome.Text & "' Where Codigo = '" & txtCodigo.Text & "' "
Banco.Execute (SQL)
Unload Me
frmCadastroDeSetores.Show
End If
End Sub
Private Sub cmdExcluir_Click()
SQL = "Delete * From Setores Where Codigo = '" & txtCodigo.Text & "'"
Banco.Execute (SQL)
Unload Me
frmCadastroDeSetores.Show
End Sub
Private Sub cmdIncluir_Click()
If txtNome.Text = "" Then
    Call MsgBox("Digite alguma coisa na caixa texto!")
Else
SQL = "Insert Into Setores (Codigo,Nome) Values ('" & txtCodigo.Text & "','" & txtNome.Text & "')"
Banco.Execute (SQL)
Unload Me
frmCadastroDeSetores.Show
End If
End Sub
Private Sub cmdLimpar_Click()
If MsgBox("Deseja realmente Limpar?", vbYesNo + vbQuestion, "Limpar") = vbYes Then
    txtCodigo.Text = ""
    txtNome.Text = ""
    fraCodigo.Enabled = True
    fraDados.Enabled = False
Else
    Cancel = True
End If
End Sub
Private Sub Form_Load()
SQL = "Select * From Setores"
Set TabelaDinamica = Banco.Execute(SQL)
If Not TabelaDinamica.EOF Then
    Do While Not TabelaDinamica.EOF
        lstSetores.AddItem TabelaDinamica("Codigo") & "-" & TabelaDinamica("Nome")
        TabelaDinamica.MoveNext
    Loop
End If
End Sub
Private Sub lstSetores_Click()
A = InStr(1, lstSetores, "-")
B = Left(lstSetores, A - 1)
 SQL = "Select * From Setores Where Codigo = '" & B & "'"
Set TabelaDinamica = Banco.Execute(SQL)
If TabelaDinamica.EOF Then
    Call MsgBox("Codigo não encontrado")
Else
    txtCodigo.Text = TabelaDinamica("Codigo")
    txtNome.Text = TabelaDinamica("Nome")
    cmdExcluir.Enabled = True
    cmdAlterar.Enabled = True
End If
End Sub
Private Sub txtCodigo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If txtCodigo.Text = "" Then
        Call MsgBox("Por Favor Digite Um Número!")
            ElseIf Not Len(txtCodigo.Text) = 3 Then
                If Len(txtCodigo.Text) = 2 Then
                    A = txtCodigo.Text
                    txtCodigo.Text = "0" & A
                Else
                    A = txtCodigo.Text
                    txtCodigo.Text = "00" & A
                    End If
    End If
fraDados.Enabled = True
fraCodigo.Enabled = False
cmdIncluir.Enabled = True
End If
End Sub
