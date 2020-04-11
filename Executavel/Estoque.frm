VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Acervo"
   ClientHeight    =   4845
   ClientLeft      =   3705
   ClientTop       =   3495
   ClientWidth     =   8190
   ControlBox      =   0   'False
   Icon            =   "Estoque.frx":0000
   LinkTopic       =   "Form3"
   ScaleHeight     =   4845
   ScaleWidth      =   8190
   Begin VB.TextBox Text9 
      DataField       =   "data devoloçao"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   6360
      TabIndex        =   25
      Top             =   2400
      Width           =   1575
   End
   Begin VB.TextBox Text8 
      DataField       =   "usuario"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   4440
      TabIndex        =   23
      Top             =   2400
      Width           =   1575
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Sim"
      DataField       =   "Empretimo"
      DataSource      =   "Data1"
      Height          =   255
      Left            =   3120
      TabIndex        =   21
      Top             =   2400
      Width           =   855
   End
   Begin VB.TextBox Text7 
      DataField       =   "observaçao"
      DataSource      =   "Data1"
      Height          =   1365
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   19
      Top             =   2880
      Width           =   2415
   End
   Begin VB.TextBox Text6 
      DataField       =   "ISBN"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   3600
      TabIndex        =   17
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Data Data1 
      Caption         =   "Navegação"
      Connect         =   "Access"
      DatabaseName    =   "C:\Users\Gustavo\Documents\Programa VB\Pergamo.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Acervo"
      ToolTipText     =   "Navega pelos registros"
      Top             =   4440
      Width           =   2655
   End
   Begin VB.CommandButton Command6 
      Caption         =   "&Voltar"
      Height          =   735
      Left            =   6960
      Picture         =   "Estoque.frx":0CCA
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Voltar para a tela de login"
      Top             =   3840
      Width           =   855
   End
   Begin VB.CommandButton Command5 
      Caption         =   "&Excluir"
      Height          =   735
      Left            =   6120
      Picture         =   "Estoque.frx":110C
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Excluir registro"
      Top             =   3840
      Width           =   855
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Al&terar"
      Height          =   735
      Left            =   5400
      Picture         =   "Estoque.frx":154E
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Alterar registro"
      Top             =   3840
      Width           =   735
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Pesquisar"
      Height          =   735
      Left            =   6960
      Picture         =   "Estoque.frx":1990
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Pesquisar registro"
      Top             =   3120
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Novo"
      Height          =   735
      Left            =   6120
      Picture         =   "Estoque.frx":1DD2
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Novo registro"
      Top             =   3120
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Salvar"
      Height          =   735
      Left            =   5400
      Picture         =   "Estoque.frx":2214
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Salvar registro"
      Top             =   3120
      Width           =   735
   End
   Begin VB.TextBox Text5 
      DataField       =   "Ano"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   2160
      TabIndex        =   9
      Top             =   1560
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      DataField       =   "Ediçao"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   120
      TabIndex        =   7
      Top             =   1560
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      DataField       =   "Editora"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   6120
      TabIndex        =   6
      Top             =   480
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      DataField       =   "Autor"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   3600
      TabIndex        =   5
      Top             =   480
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      DataField       =   "Livro"
      DataSource      =   "Data1"
      Height          =   525
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   480
      Width           =   3015
   End
   Begin VB.Label Label10 
      Caption         =   "Data de Devoluçao"
      Height          =   255
      Left            =   6360
      TabIndex        =   24
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Label Label9 
      Caption         =   "Para "
      Height          =   255
      Left            =   4440
      TabIndex        =   22
      Top             =   2040
      Width           =   1575
   End
   Begin VB.Label Label8 
      Caption         =   "Empestado"
      Height          =   255
      Left            =   3120
      TabIndex        =   20
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label Label7 
      Caption         =   "Editora"
      Height          =   255
      Left            =   6120
      TabIndex        =   18
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label6 
      Caption         =   "ISBN"
      Height          =   255
      Left            =   3600
      TabIndex        =   16
      ToolTipText     =   "Numero Universal do livro"
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label Label5 
      Caption         =   "Observações"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      ToolTipText     =   "Observaçoes sobre o livro"
      Top             =   2520
      Width           =   1575
   End
   Begin VB.Label Label4 
      Caption         =   "Ano"
      Height          =   255
      Left            =   2160
      TabIndex        =   3
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Ediçao "
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Autor"
      Height          =   255
      Left            =   3600
      TabIndex        =   1
      Top             =   120
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "Livro"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1575
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
      Begin VB.Menu linha 
         Caption         =   "-"
      End
      Begin VB.Menu menu_voltar 
         Caption         =   "Voltar"
      End
      Begin VB.Menu linha2 
         Caption         =   "-"
      End
      Begin VB.Menu menu_sair 
         Caption         =   "Sair"
      End
   End
   Begin VB.Menu menu_abrir 
      Caption         =   "A&brir"
      Begin VB.Menu menu_cadastro 
         Caption         =   "Cadastro"
      End
      Begin VB.Menu menu_despeza 
         Caption         =   "Despesa"
      End
   End
   Begin VB.Menu menu_informaçoes 
      Caption         =   "&Informações"
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Len(Text2.Text) = Empty Then
MsgBox "É necessario o nome"
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
pesquisa = InputBox("Qual o autor?")
If pesquisa = "" Then
Exit Sub
Else
Data1.Recordset.FindFirst "autor= '" & pesquisa & "'"
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

Private Sub menu_alterar_Click()
MsgBox "Alteraçao realizada com sucesso"
Data1.UpdateRecord
End Sub



Private Sub menu_cadastro_Click()
Form2.Show

End Sub

Private Sub menu_despeza_Click()
Form4.Show
End Sub

Private Sub menu_informaçoes_Click()
Form7.Show
End Sub

Private Sub menu_novo_Click()
Data1.Recordset.AddNew
End Sub

Private Sub menu_pesquisar_Click()
Dim pesquisa As String
pesquisa = InputBox("Qual o nome autor?")
If pesquisa = "" Then
Exit Sub
Else
Data1.Recordset.FindFirst "autor= '" & pesquisa & "'"
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
MsgBox "É necessario o nome"
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
