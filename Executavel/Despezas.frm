VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "Despesas"
   ClientHeight    =   5865
   ClientLeft      =   3510
   ClientTop       =   4095
   ClientWidth     =   8010
   ControlBox      =   0   'False
   Icon            =   "Despezas.frx":0000
   LinkTopic       =   "Form4"
   ScaleHeight     =   5865
   ScaleWidth      =   8010
   Begin VB.CommandButton Command3 
      Caption         =   "&Pesquisar"
      Height          =   735
      Index           =   1
      Left            =   5880
      Picture         =   "Despezas.frx":0CCA
      Style           =   1  'Graphical
      TabIndex        =   33
      ToolTipText     =   "Pesquisar registro"
      Top             =   3600
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Al&terar"
      Height          =   735
      Index           =   0
      Left            =   5160
      Picture         =   "Despezas.frx":110C
      Style           =   1  'Graphical
      TabIndex        =   32
      ToolTipText     =   "Alterar registro"
      Top             =   4320
      Width           =   735
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&Voltar"
      Height          =   735
      Left            =   5880
      Picture         =   "Despezas.frx":154E
      Style           =   1  'Graphical
      TabIndex        =   31
      ToolTipText     =   "Voltar para tela de login"
      Top             =   4320
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Salvar"
      Height          =   735
      Index           =   1
      Left            =   5160
      Picture         =   "Despezas.frx":1990
      Style           =   1  'Graphical
      TabIndex        =   30
      ToolTipText     =   "Salvar registro"
      Top             =   3600
      Width           =   735
   End
   Begin VB.Data Data1 
      Caption         =   "Navegação"
      Connect         =   "Access"
      DatabaseName    =   "C:\Users\Gustavo\Documents\Programa VB\Pergamo.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   5040
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Despezas"
      ToolTipText     =   "Navega pelos registro"
      Top             =   5160
      Width           =   2535
   End
   Begin VB.TextBox Text14 
      DataField       =   "Total"
      DataSource      =   "Data1"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1680
      TabIndex        =   29
      Top             =   5160
      Width           =   1335
   End
   Begin VB.TextBox Text13 
      DataField       =   "Diversos"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   3360
      TabIndex        =   28
      Top             =   4440
      Width           =   1335
   End
   Begin VB.TextBox Text12 
      DataField       =   "Funcionarios"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   1920
      TabIndex        =   27
      Top             =   4440
      Width           =   1095
   End
   Begin VB.TextBox Text11 
      DataField       =   "Proaganda"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   240
      TabIndex        =   26
      Top             =   4440
      Width           =   1215
   End
   Begin VB.TextBox Text10 
      DataField       =   "Imposto"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   1800
      TabIndex        =   25
      Top             =   3360
      Width           =   855
   End
   Begin VB.TextBox Text9 
      DataField       =   "Contador"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   240
      TabIndex        =   24
      Top             =   3360
      Width           =   855
   End
   Begin VB.TextBox Text8 
      DataField       =   "Internet"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   3720
      TabIndex        =   23
      Top             =   2400
      Width           =   975
   End
   Begin VB.TextBox Text7 
      DataField       =   "Celular"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   1800
      TabIndex        =   22
      Top             =   2400
      Width           =   1095
   End
   Begin VB.TextBox Text6 
      DataField       =   "Telefone"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   240
      TabIndex        =   21
      Top             =   2400
      Width           =   1095
   End
   Begin VB.TextBox Text5 
      DataField       =   "Gas"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   2880
      TabIndex        =   17
      Top             =   1440
      Width           =   1095
   End
   Begin VB.TextBox Text4 
      DataField       =   "Agua"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   1680
      TabIndex        =   16
      Top             =   1440
      Width           =   975
   End
   Begin VB.TextBox Text3 
      DataField       =   "Luz"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   120
      TabIndex        =   15
      Top             =   1440
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      DataField       =   "Ano"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   4680
      TabIndex        =   4
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox Text1 
      DataField       =   "Mes"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   720
      TabIndex        =   3
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label16 
      Caption         =   "Gas"
      Height          =   255
      Left            =   3000
      TabIndex        =   20
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label15 
      Caption         =   "Agua"
      Height          =   255
      Left            =   1800
      TabIndex        =   19
      Top             =   960
      Width           =   735
   End
   Begin VB.Label Label14 
      Caption         =   "Luz"
      Height          =   255
      Left            =   240
      TabIndex        =   18
      Top             =   960
      Width           =   855
   End
   Begin VB.Label Label13 
      Caption         =   "Total do mês"
      Height          =   255
      Left            =   240
      TabIndex        =   14
      Top             =   5160
      Width           =   1095
   End
   Begin VB.Label Label12 
      Caption         =   "Diversos"
      Height          =   255
      Left            =   3480
      TabIndex        =   13
      Top             =   3960
      Width           =   1455
   End
   Begin VB.Label Label11 
      Caption         =   "Funcionarios"
      Height          =   255
      Left            =   1920
      TabIndex        =   12
      Top             =   3960
      Width           =   1095
   End
   Begin VB.Label Label10 
      Caption         =   "Proaganda"
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   3960
      Width           =   1095
   End
   Begin VB.Label Label9 
      Caption         =   "Impostos"
      Height          =   255
      Left            =   1800
      TabIndex        =   10
      Top             =   2880
      Width           =   735
   End
   Begin VB.Label Label8 
      Caption         =   "Contador"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   2880
      Width           =   735
   End
   Begin VB.Label Label7 
      Caption         =   "Internet"
      Height          =   255
      Left            =   3720
      TabIndex        =   8
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label Label6 
      Caption         =   "Celular"
      Height          =   255
      Left            =   1800
      TabIndex        =   7
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label Label5 
      Caption         =   "Telefone"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "Contas"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Ano"
      Height          =   255
      Left            =   3960
      TabIndex        =   2
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label2 
      Height          =   255
      Left            =   1800
      TabIndex        =   1
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Mês"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   375
   End
   Begin VB.Menu menu_arquivo 
      Caption         =   "&Arquivo"
      Begin VB.Menu menu_salvar 
         Caption         =   "Salvar"
      End
      Begin VB.Menu menu_alterar 
         Caption         =   "Alterar"
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
      Begin VB.Menu menu_estoque 
         Caption         =   "Acervo"
      End
   End
   Begin VB.Menu menu_informaçoes 
      Caption         =   "&Informações"
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False













Private Sub Command1_Click(Index As Integer)
If Len(Text2.Text) = Empty Then
MsgBox "É necessario o nome"
Text2.SetFocus
Else
Data1.UpdateRecord
End If
End Sub





Private Sub Command2_Click(Index As Integer)
MsgBox "Alteraçao realizada com sucesso"
Data1.UpdateRecord
End Sub

Private Sub Command3_Click(Index As Integer)
Dim pesquisa As String
pesquisa = InputBox("Qual o nome Codigo?")
If pesquisa = "" Then
Exit Sub
Else
Data1.Recordset.FindFirst "Total= '" & pesquisa & "'"
If Data1.Recordset.NoMatch Then
MsgBox "registro nao localizado"
Exit Sub
End If
End If
End Sub

Private Sub Command4_Click()
Form1.Show
Unload Form2
Unload Form3
Unload Form4
Unload Form5
Unload Form6
Unload Form7
Unload Form8
End Sub

Private Sub Command5_Click()
Data1.Recordset.AddNew
End Sub

Private Sub menu_alterar_Click()
MsgBox "Alteraçao realizada com sucesso"
Data1.UpdateRecord
End Sub

Private Sub menu_cadastro_Click()
Form2.Show
End Sub

Private Sub menu_estoque_Click()
Form3.Show

End Sub

Private Sub menu_informaçoes_Click()
Form8.Show

End Sub

Private Sub menu_pesquisar_Click()
Dim pesquisa As String
pesquisa = InputBox("Qual o total que procura?")
If pesquisa = "" Then
Exit Sub
Else
Data1.Recordset.FindFirst "total= '" & pesquisa & "'"
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

Private Sub Text1_Change()
Dim n As Integer
n = Val(Text1.Text)
Select Case n
Case 1
Label2 = " Janeiro"
Case 2
Label2 = " Fevereiro"
Case 3
Label2 = " Março"
Case 4
Label2 = " Abril"
Case 5
Label2 = " Maio"
Case 6
Label2 = " Junho"
Case 7
Label2 = " Julho"
Case 8
Label2 = " Agosto"
Case 9
Label2 = " Setembro"
Case 10
Label2 = " Outubro"
Case 11
Label2 = " Novembro"
Case 12
Label2 = " Dezembro"
Case Else
Label2 = ""
End Select

End Sub






'PARTE DA CONTA AUTOMATICA

Private Sub Text10_Change()
Dim a, b, c, d, e, f, g, h, i, j, k, r As String
'variaveis
a = Val(Text3.Text)
b = Val(Text4.Text)
c = Val(Text5.Text)
d = Val(Text6.Text)
e = Val(Text7.Text)
f = Val(Text8.Text)
g = Val(Text9.Text)
h = Val(Text10.Text)
i = Val(Text11.Text)
j = Val(Text12.Text)
k = Val(Text13.Text)
'resoluçao
r = Val(a + b + c + d + e + f + g + h + i + j + k)
Text14.Text = r
End Sub

Private Sub Text11_Change()
Dim a, b, c, d, e, f, g, h, i, j, k, r As String
'variaveis
a = Val(Text3.Text)
b = Val(Text4.Text)
c = Val(Text5.Text)
d = Val(Text6.Text)
e = Val(Text7.Text)
f = Val(Text8.Text)
g = Val(Text9.Text)
h = Val(Text10.Text)
i = Val(Text11.Text)
j = Val(Text12.Text)
k = Val(Text13.Text)
'resoluçao
r = Val(a + b + c + d + e + f + g + h + i + j + k)
Text14.Text = r
End Sub

Private Sub Text12_Change()
Dim a, b, c, d, e, f, g, h, i, j, k, r As String
'variaveis
a = Val(Text3.Text)
b = Val(Text4.Text)
c = Val(Text5.Text)
d = Val(Text6.Text)
e = Val(Text7.Text)
f = Val(Text8.Text)
g = Val(Text9.Text)
h = Val(Text10.Text)
i = Val(Text11.Text)
j = Val(Text12.Text)
k = Val(Text13.Text)
'resoluçao
r = Val(a + b + c + d + e + f + g + h + i + j + k)
Text14.Text = r
End Sub








Private Sub Text13_Change()
Dim a, b, c, d, e, f, g, h, i, j, k, r As String
'variaveis
a = Val(Text3.Text)
b = Val(Text4.Text)
c = Val(Text5.Text)
d = Val(Text6.Text)
e = Val(Text7.Text)
f = Val(Text8.Text)
g = Val(Text9.Text)
h = Val(Text10.Text)
i = Val(Text11.Text)
j = Val(Text12.Text)
k = Val(Text13.Text)
'resoluçao
r = Val(a + b + c + d + e + f + g + h + i + j + k)
Text14.Text = r
End Sub

Private Sub Text14_Change()
Dim a, b, c, d, e, f, g, h, i, j, k, r As String
'variaveis
a = Val(Text3.Text)
b = Val(Text4.Text)
c = Val(Text5.Text)
d = Val(Text6.Text)
e = Val(Text7.Text)
f = Val(Text8.Text)
g = Val(Text9.Text)
h = Val(Text10.Text)
i = Val(Text11.Text)
j = Val(Text12.Text)
k = Val(Text13.Text)
'resoluçao
r = Val(a + b + c + d + e + f + g + h + i + j + k)
Text14.Text = r

End Sub

Private Sub Text3_Change()
Dim a, b, c, d, e, f, g, h, i, j, k, r As String
'variaveis
a = Val(Text3.Text)
b = Val(Text4.Text)
c = Val(Text5.Text)
d = Val(Text6.Text)
e = Val(Text7.Text)
f = Val(Text8.Text)
g = Val(Text9.Text)
h = Val(Text10.Text)
i = Val(Text11.Text)
j = Val(Text12.Text)
k = Val(Text13.Text)
'resoluçao
r = Val(a + b + c + d + e + f + g + h + i + j + k)
Text14.Text = r
End Sub

Private Sub Text4_Change()
Dim a, b, c, d, e, f, g, h, i, j, k, r As String
'variaveis
a = Val(Text3.Text)
b = Val(Text4.Text)
c = Val(Text5.Text)
d = Val(Text6.Text)
e = Val(Text7.Text)
f = Val(Text8.Text)
g = Val(Text9.Text)
h = Val(Text10.Text)
i = Val(Text11.Text)
j = Val(Text12.Text)
k = Val(Text13.Text)
'resoluçao
r = Val(a + b + c + d + e + f + g + h + i + j + k)
Text14.Text = r
End Sub

Private Sub Text5_Change()
Dim a, b, c, d, e, f, g, h, i, j, k, r As String
'variaveis
a = Val(Text3.Text)
b = Val(Text4.Text)
c = Val(Text5.Text)
d = Val(Text6.Text)
e = Val(Text7.Text)
f = Val(Text8.Text)
g = Val(Text9.Text)
h = Val(Text10.Text)
i = Val(Text11.Text)
j = Val(Text12.Text)
k = Val(Text13.Text)
'resoluçao
r = Val(a + b + c + d + e + f + g + h + i + j + k)
Text14.Text = r
End Sub

Private Sub Text6_Change()
Dim a, b, c, d, e, f, g, h, i, j, k, r As String
'variaveis
a = Val(Text3.Text)
b = Val(Text4.Text)
c = Val(Text5.Text)
d = Val(Text6.Text)
e = Val(Text7.Text)
f = Val(Text8.Text)
g = Val(Text9.Text)
h = Val(Text10.Text)
i = Val(Text11.Text)
j = Val(Text12.Text)
k = Val(Text13.Text)
'resoluçao
r = Val(a + b + c + d + e + f + g + h + i + j + k)
Text14.Text = r
End Sub

Private Sub Text7_Change()
Dim a, b, c, d, e, f, g, h, i, j, k, r As String
'variaveis
a = Val(Text3.Text)
b = Val(Text4.Text)
c = Val(Text5.Text)
d = Val(Text6.Text)
e = Val(Text7.Text)
f = Val(Text8.Text)
g = Val(Text9.Text)
h = Val(Text10.Text)
i = Val(Text11.Text)
j = Val(Text12.Text)
k = Val(Text13.Text)
'resoluçao
r = Val(a + b + c + d + e + f + g + h + i + j + k)
Text14.Text = r
End Sub

Private Sub Text8_Change()
Dim a, b, c, d, e, f, g, h, i, j, k, r As String
'variaveis
a = Val(Text3.Text)
b = Val(Text4.Text)
c = Val(Text5.Text)
d = Val(Text6.Text)
e = Val(Text7.Text)
f = Val(Text8.Text)
g = Val(Text9.Text)
h = Val(Text10.Text)
i = Val(Text11.Text)
j = Val(Text12.Text)
k = Val(Text13.Text)
'resoluçao
r = Val(a + b + c + d + e + f + g + h + i + j + k)
Text14.Text = r
End Sub

Private Sub Text9_Change()
Dim a, b, c, d, e, f, g, h, i, j, k, r As String
'variaveis
a = Val(Text3.Text)
b = Val(Text4.Text)
c = Val(Text5.Text)
d = Val(Text6.Text)
e = Val(Text7.Text)
f = Val(Text8.Text)
g = Val(Text9.Text)
h = Val(Text10.Text)
i = Val(Text11.Text)
j = Val(Text12.Text)
k = Val(Text13.Text)
'resoluçao
r = Val(a + b + c + d + e + f + g + h + i + j + k)
Text14.Text = r
End Sub
