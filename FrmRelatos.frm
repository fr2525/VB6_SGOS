VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form FrmRelatos 
   Caption         =   "Relatórios"
   ClientHeight    =   5310
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8550
   LinkTopic       =   "Form1"
   ScaleHeight     =   5310
   ScaleWidth      =   8550
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdimprimir 
      Height          =   510
      Left            =   7740
      Picture         =   "FrmRelatos.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Imprime o orçamento"
      Top             =   2580
      Width           =   510
   End
   Begin VB.CommandButton CmdSair 
      Height          =   510
      Left            =   7740
      Picture         =   "FrmRelatos.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Sai "
      Top             =   3120
      Width           =   510
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   3705
      Left            =   240
      TabIndex        =   0
      Top             =   990
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   6535
      _Version        =   393216
      Rows            =   1
      FixedCols       =   0
      BackColor       =   16777215
      BackColorSel    =   8454143
      ForeColorSel    =   0
      ScrollBars      =   2
      FormatString    =   "Nome do Relatório                                                                                 |"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "FrmRelatos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdimprimir_Click()
    If MSFlexGrid1.Row > 0 Then
       MSFlexGrid1.Col = 1
       relato = MSFlexGrid1.Text
    End If
End Sub

Private Sub CmdSair_Click()
    Unload Me
End Sub

Private Sub Form_Load()
 Dim F As Long, Linha As String
 
F = FreeFile
Open App.Path & "\relatos.txt" For Input As F   'abre o arquivo texto

'db.Execute "CREATE TABLE Clientes (ID LONG, [Nome] TEXT (50), " _
'& "[Endereco] TEXT (50), [telefone] TEXT (15), [Nascimento] TEXT (10))" 'cria a tabela c/a estrutura

'Set rs = db.OpenRecordset("Clientes", dbOpenTable)   'abre a tabela para receber os dados

Do While Not EOF(F)
   Line Input #F, Linha 'lê uma linha do arquivo texto

   'extrai a informação do arquivo texto usando a função MID
   descricao = Mid(Linha, 1, 40)
   relato = Mid(Linha, 41, 12)
   Me.MSFlexGrid1.AddItem descricao & vbTab & relato
   
Loop
End Sub

Private Sub MSFlexGrid1_Click()
   PintaGrid MSFlexGrid1
End Sub
