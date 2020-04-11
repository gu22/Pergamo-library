VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Cadastro"
   ClientHeight    =   6840
   ClientLeft      =   2520
   ClientTop       =   3300
   ClientWidth     =   9975
   ControlBox      =   0   'False
   Icon            =   "Cadastro.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   6840
   ScaleWidth      =   9975
   Begin VB.TextBox Text14 
      DataField       =   "Rg"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   5040
      TabIndex        =   33
      Top             =   4440
      Width           =   1455
   End
   Begin VB.Data Data1 
      Caption         =   "Navegação"
      Connect         =   "Access"
      DatabaseName    =   "C:\Users\Gustavo\Documents\Programa VB\Pergamo.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   6600
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Cadastro"
      ToolTipText     =   "Navega pelos registros"
      Top             =   5880
      Width           =   2655
   End
   Begin VB.CommandButton Command6 
      Caption         =   "&Voltar"
      Height          =   735
      Left            =   8280
      Picture         =   "Cadastro.frx":0CCA
      Style           =   1  'Graphical
      TabIndex        =   31
      ToolTipText     =   "Voltar para tela de login"
      Top             =   5040
      Width           =   855
   End
   Begin VB.CommandButton Command5 
      Caption         =   "&Excluir"
      Height          =   735
      Left            =   7440
      Picture         =   "Cadastro.frx":110C
      Style           =   1  'Graphical
      TabIndex        =   30
      ToolTipText     =   "Apaguar registro"
      Top             =   5040
      Width           =   855
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Al&terar"
      Height          =   735
      Left            =   6600
      Picture         =   "Cadastro.frx":154E
      Style           =   1  'Graphical
      TabIndex        =   29
      ToolTipText     =   "Alterar Registro"
      Top             =   5040
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Pesquisar"
      Height          =   735
      Left            =   8280
      Picture         =   "Cadastro.frx":1990
      Style           =   1  'Graphical
      TabIndex        =   28
      ToolTipText     =   "pesquisar um registro"
      Top             =   4320
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Novo"
      Height          =   735
      Left            =   7440
      Picture         =   "Cadastro.frx":1DD2
      Style           =   1  'Graphical
      TabIndex        =   27
      ToolTipText     =   "Novo registro"
      Top             =   4320
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Salvar"
      Height          =   735
      Left            =   6600
      Picture         =   "Cadastro.frx":2214
      Style           =   1  'Graphical
      TabIndex        =   26
      ToolTipText     =   "Salvar registro"
      Top             =   4320
      Width           =   855
   End
   Begin VB.TextBox Text13 
      DataField       =   "Observaçoes"
      DataSource      =   "Data1"
      Height          =   1335
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   25
      Top             =   5400
      Width           =   3495
   End
   Begin VB.TextBox Text12 
      DataField       =   "Cpf/cnpj"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   2880
      TabIndex        =   23
      Top             =   4440
      Width           =   1815
   End
   Begin VB.TextBox Text11 
      DataField       =   "E-mail"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   120
      TabIndex        =   22
      Top             =   4440
      Width           =   2535
   End
   Begin VB.TextBox Text10 
      DataField       =   "Celular"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   5760
      TabIndex        =   19
      Top             =   3480
      Width           =   2295
   End
   Begin VB.TextBox Text9 
      DataField       =   "Telefone Comercial"
      DataSource      =   "Data1"
      Height          =   315
      Left            =   2880
      TabIndex        =   18
      Top             =   3480
      Width           =   2295
   End
   Begin VB.TextBox Text8 
      DataField       =   "Telefone Residencial"
      DataSource      =   "Data1"
      Height          =   315
      Left            =   120
      TabIndex        =   17
      Top             =   3480
      Width           =   1815
   End
   Begin VB.TextBox Text7 
      DataField       =   "Estado"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   5640
      TabIndex        =   13
      Top             =   2520
      Width           =   2775
   End
   Begin VB.TextBox Text6 
      DataField       =   "Cidade"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   2760
      TabIndex        =   12
      Top             =   2520
      Width           =   2415
   End
   Begin VB.TextBox Text5 
      DataField       =   "Cep"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   120
      TabIndex        =   9
      Top             =   2520
      Width           =   2055
   End
   Begin VB.TextBox Text4 
      DataField       =   "Complemento"
      DataSource      =   "Data1"
      Height          =   525
      Left            =   2760
      MultiLine       =   -1  'True
      TabIndex        =   8
      Top             =   1440
      Width           =   2415
   End
   Begin VB.TextBox Text3 
      DataField       =   "Endereço"
      DataSource      =   "Data1"
      Height          =   495
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   7
      Top             =   1440
      Width           =   2295
   End
   Begin VB.TextBox Text2 
      DataField       =   "Nome"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   2760
      TabIndex        =   3
      Top             =   600
      Width           =   3015
   End
   Begin VB.TextBox Text1 
      DataField       =   "Codigo"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   2055
   End
   Begin VB.Label Label14 
      Caption         =   "Rg"
      Height          =   255
      Left            =   5040
      TabIndex        =   32
      Top             =   3960
      Width           =   495
   End
   Begin VB.Label Label13 
      Caption         =   "Observações"
      Height          =   255
      Left            =   120
      TabIndex        =   24
      ToolTipText     =   "Observaçoes sobre a pessoa"
      Top             =   4920
      Width           =   1815
   End
   Begin VB.Label Label12 
      Caption         =   "Cpf/Cnpj"
      Height          =   255
      Left            =   2880
      TabIndex        =   21
      Top             =   3960
      Width           =   1935
   End
   Begin VB.Label Label11 
      Caption         =   "E-Mail"
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   3960
      Width           =   1695
   End
   Begin VB.Label Label10 
      Caption         =   "Celular"
      Height          =   255
      Left            =   5640
      TabIndex        =   16
      Top             =   3000
      Width           =   2415
   End
   Begin VB.Label Label9 
      Caption         =   "Telefone Comercial"
      Height          =   255
      Left            =   2880
      TabIndex        =   15
      Top             =   3000
      Width           =   2175
   End
   Begin VB.Label Label8 
      Caption         =   "Telefone Residencial"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   3000
      Width           =   1935
   End
   Begin VB.Label Label7 
      Caption         =   "Estado"
      Height          =   255
      Left            =   5640
      TabIndex        =   11
      Top             =   2040
      Width           =   2535
   End
   Begin VB.Label Label6 
      Caption         =   "Cidade"
      Height          =   255
      Left            =   2880
      TabIndex        =   10
      Top             =   2040
      Width           =   2175
   End
   Begin VB.Label Label5 
      Caption         =   "Cep"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   2040
      Width           =   2175
   End
   Begin VB.Label Label4 
      Caption         =   "Complemento"
      Height          =   255
      Left            =   2760
      TabIndex        =   5
      Top             =   1080
      Width           =   2535
   End
   Begin VB.Label Label3 
      Caption         =   "Endereço"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "Nome"
      Height          =   255
      Left            =   2760
      TabIndex        =   1
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "Codigo"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
   Begin VB.Menu menu_arquivo 
      Caption         =   "&Arquivo"
      Begin VB.Menu menu_salvar 
         Caption         =   "Salvar"
      End
      Begin VB.Menu menu_alterar 
         Caption         =   "Alterar"
      End
      Begin VB.Menu menu_novo 
         Caption         =   "Novo"
      End
      Begin VB.Menu menu_pesquisar 
         Caption         =   "Pesquisar"
      End
      Begin VB.Menu menu_linha 
         Caption         =   "-"
      End
      Begin VB.Menu menu_voltar 
         Caption         =   "Voltar"
      End
      Begin VB.Menu menu_linha2 
         Caption         =   "-"
      End
      Begin VB.Menu menu_sair 
         Caption         =   "Sair"
      End
   End
   Begin VB.Menu menu_abrir 
      Caption         =   "A&brir"
      Begin VB.Menu menu_Acervo 
         Caption         =   "Acervo"
      End
      Begin VB.Menu menu_despezas 
         Caption         =   "Despesas"
      End
   End
   Begin VB.Menu menu_informçoes 
      Caption         =   "&Informações"
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub menu_estoque_Click()
Form3.Show
End Sub


Private Sub Command1_Click()
If Len(Text2.Text) = Empty Then
MsgBox "é necessario o nome"
Text2.SetFocus
Else
Data1.UpdateRecord
End If
End Sub

Private Sub Command2_Click()
Data1.Recordset.AddNew
End Sub

Private Sub Command3_Click()
Dim pesquisa As String
pesquisa = InputBox("Qual o nome usuario?")
If pesquisa = "" Then
Exit Sub
Else
Data1.Recordset.FindFirst "nome= '" & pesquisa & "'"
If Data1.Recordset.NoMatch Then
MsgBox "Registro nao localizado"
Exit Sub
End If
End If
End Sub

Private Sub Command4_Click()
MsgBox "Alteraçao realizada com sucesso"
Data1.UpdateRecord
End Sub

Private Sub Command5_Click()
Dim excluir As String
excluir = MsgBox("Deseja excluir este registro?", vbQuestion + vbYesNo, "cadastro")
If excluir <> vbYes Then
Exit Sub
Else
Data1.Recordset.Delete
Data1.Refresh
End If

End Sub

Private Sub Command6_Click()
Form1.Show
Unload Form2
Unload Form3
Unload Form4
Unload Form5
Unload Form6
Unload Form7
Unload Form8
End Sub

Private Sub menu_Acervo_Click()
Form3.Show
End Sub

Private Sub menu_alterar_Click()
MsgBox "Alteraçao realizada com sucesso"
Data1.UpdateRecord
End Sub

Private Sub menu_despezas_Click()
Form4.Show
End Sub

Private Sub menu_informçoes_Click()
Form6.Show
End Sub





Private Sub menu_novo_Click()
Data1.Recordset.AddNew
End Sub

Private Sub menu_pesquisar_Click()
Dim pesquisa As String
pesquisa = InputBox("Qual o nome usuario?")
If pesquisa = "" Then
Exit Sub
Else
Data1.Recordset.FindFirst "nome= '" & pesquisa & "'"
If Data1.Recordset.NoMatch Then
MsgBox "Registro nao localizado"
Exit Sub
End If
End If
End Sub

Private Sub menu_sair_Click()
If MsgBox("Deseja sair?", vbYesNo, "Aviso") = vbYes Then
End
End If

End Sub

Private Sub menu_salvar_Click()
If Len(Text2.Text) = Empty Then
MsgBox "é necessario o nome"
Text2.SetFocus
Else
Data1.UpdateRecord
End If
End Sub

Private Sub menu_voltar_Click()
Form1.Show
Unload Form2
Unload Form3
Unload Form4
Unload Form5
Unload Form6
Unload Form7
Unload Form8
End Sub
