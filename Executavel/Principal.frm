VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H8000000D&
   Caption         =   "Pérgamo Library"
   ClientHeight    =   7230
   ClientLeft      =   2520
   ClientTop       =   3000
   ClientWidth     =   10080
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "Principal.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "Principal.frx":0CCA
   ScaleHeight     =   7230
   ScaleWidth      =   10080
   Begin VB.CommandButton Command7 
      Caption         =   "Sair do Login"
      Enabled         =   0   'False
      Height          =   495
      Left            =   9120
      TabIndex        =   21
      ToolTipText     =   "Sair do Login"
      Top             =   5400
      Width           =   855
   End
   Begin VB.CommandButton Command6 
      Caption         =   "OK"
      Height          =   495
      Left            =   8640
      TabIndex        =   18
      ToolTipText     =   "Efetuar login"
      Top             =   4680
      Width           =   735
   End
   Begin VB.TextBox Text2 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   6120
      PasswordChar    =   "*"
      TabIndex        =   17
      Top             =   5160
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   6120
      TabIndex        =   16
      Top             =   4200
      Width           =   2295
   End
   Begin VB.Timer Timer1 
      Interval        =   60
      Left            =   9600
      Top             =   120
   End
   Begin VB.Frame Frame4 
      BorderStyle     =   0  'None
      Caption         =   "Frame4"
      Height          =   495
      Left            =   480
      TabIndex        =   3
      Top             =   6240
      Width           =   9255
      Begin VB.Label Label7 
         Caption         =   "Label7"
         Height          =   255
         Left            =   5400
         TabIndex        =   10
         ToolTipText     =   "Horas"
         Top             =   120
         Width           =   1575
      End
      Begin VB.Label Label6 
         Caption         =   "Ver. 1.0"
         Height          =   255
         Left            =   8160
         TabIndex        =   9
         ToolTipText     =   "Versão do Programa"
         Top             =   120
         Width           =   735
      End
      Begin VB.Label Label5 
         Height          =   255
         Left            =   4920
         TabIndex        =   8
         Top             =   120
         Width           =   1695
      End
      Begin VB.Label Label4 
         Caption         =   "Label4"
         Height          =   375
         Left            =   3000
         TabIndex        =   7
         Top             =   120
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Data:"
         Height          =   495
         Left            =   2520
         TabIndex        =   6
         Top             =   120
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "Gustavo"
         Height          =   255
         Left            =   360
         TabIndex        =   5
         ToolTipText     =   "Criador do Programa"
         Top             =   120
         Width           =   855
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   495
      Left            =   360
      TabIndex        =   2
      Top             =   6360
      Width           =   9255
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   4215
      Left            =   480
      TabIndex        =   1
      Top             =   1680
      Width           =   1215
      Begin VB.CommandButton Command5 
         Caption         =   "&Sair"
         Height          =   735
         Left            =   120
         Picture         =   "Principal.frx":15249
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Sair do programa"
         Top             =   3360
         Width           =   975
      End
      Begin VB.CommandButton Command4 
         Caption         =   "&Informações"
         Height          =   855
         Left            =   120
         Picture         =   "Principal.frx":1568B
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Ver informações"
         Top             =   2520
         Width           =   975
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Despesas"
         Enabled         =   0   'False
         Height          =   855
         Left            =   120
         Picture         =   "Principal.frx":15ACD
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Ver despesas"
         Top             =   1680
         Width           =   975
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Acervo"
         Enabled         =   0   'False
         Height          =   735
         Left            =   120
         Picture         =   "Principal.frx":15F0F
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Ir para acervo de livros"
         Top             =   960
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Cadastro"
         Enabled         =   0   'False
         Height          =   855
         Left            =   120
         Picture         =   "Principal.frx":16351
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Ir para cadastro dos clietes"
         Top             =   120
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      ForeColor       =   &H8000000E&
      Height          =   4215
      Left            =   360
      TabIndex        =   0
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Software Freeware"
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   480
      TabIndex        =   22
      Top             =   6960
      Width           =   1695
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Senha"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6240
      TabIndex        =   20
      Top             =   4800
      Width           =   975
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Usuario"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6240
      TabIndex        =   19
      Top             =   3840
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Pérgamo-Library"
      BeginProperty Font 
         Name            =   "Tw Cen MT"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   1695
      Left            =   1920
      TabIndex        =   4
      Top             =   360
      Width           =   8175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Form2.Show
End Sub

Private Sub Command2_Click()
Form3.Show
End Sub

Private Sub Command3_Click()
Form4.Show
End Sub

Private Sub Command4_Click()
MsgBox "Software Freeware feito por Gustavo nº15  2ºAno de Informatica", 64, "Pérgamos Library"
Form5.Show
End Sub

Private Sub Command5_Click()
If MsgBox("Deseja sair?", vbYesNo, "Aviso") = vbYes Then
End
End If
End Sub

Private Sub Command6_Click()


'começo:

'usuarios
stradm = "admin"
strsuper = "super"
Strgerente = "gerente"

'senhas
strsenha1 = "master"
strsenha2 = "123"
Strsenha3 = "mestre"



'montagem
If Text1 = Empty And Text2 = Empty Then
MsgBox "informe usuario e senha", vbInformation, "aviso"
Text1 = ""
Text2 = ""
Text1.SetFocus
Else


If Text1 = Empty Then
MsgBox "informe o nome do usuário", vbInformation, "aviso"
Text1 = ""
Text1.SetFocus
Else


If Text2 = Empty Then
MsgBox "informe a senha", vbInformation, "aviso"
Text2 = ""
Text2.SetFocus



Else

'campo usuario



If Text1 <> stradm And Text1 <> strsuper And Text1 <> Strgerente Then
MsgBox "usuario invalido", vbInformation, "aviso"
Text1 = ""
Text1.SetFocus
Else


'campo senha

If Text2 <> strsenha1 And Text2 <> strsenha2 And Text2 <> Strsenha3 Then
MsgBox "senha invalida", vbInformation, "aviso"
Text2 = ""
Text2.SetFocus

End If
End If
End If
End If
End If


' abilitaçao dos botoes
If Text1 = stradm And Text2 = strsenha1 Then
Command1.Enabled = True
Command2.Enabled = True
Command3.Enabled = True
Command7.Enabled = True

Form1.Enabled = True
Form3.Enabled = True
Form4.Enabled = True

Else
If Text1 = strsuper And Text2 = strsenha2 Then
Command1.Enabled = True
Command3.Enabled = True
Command7.Enabled = True

Form2.Enabled = True
Form3.Enabled = True
Form4.Enabled = True

Else
If Text1 = Strgerente And Text2 = Strsenha3 Then
Command1.Enabled = True
Command2.Enabled = True
Command3.Enabled = True
Command7.Enabled = True





End If
End If
End If


End Sub


Private Sub Command7_Click()
Command1.Enabled = False
Command2.Enabled = False
Command3.Enabled = False
Command7.Enabled = False
Text1 = ""
Text2 = ""
Text1.SetFocus



Unload Form2
Unload Form3
Unload Form4
Unload Form5
Unload Form6
Unload Form7
Unload Form8
End Sub

Private Sub Form_Load()
Label4 = Date
End Sub

Private Sub Timer1_Timer()
Label7 = Time
End Sub
