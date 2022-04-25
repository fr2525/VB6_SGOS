VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form FrmPedAten 
   Caption         =   "Pedidos de Venda Atendidos"
   ClientHeight    =   4440
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9300
   LinkTopic       =   "Form1"
   ScaleHeight     =   4440
   ScaleWidth      =   9300
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid MsflxOrca 
      Height          =   3375
      Left            =   240
      TabIndex        =   2
      Top             =   450
      Width           =   7665
      _ExtentX        =   13520
      _ExtentY        =   5953
      _Version        =   393216
      Cols            =   3
      FixedCols       =   0
      BackColorSel    =   65535
      ForeColorSel    =   0
      FocusRect       =   0
      HighLight       =   2
      ScrollBars      =   2
      SelectionMode   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton CmdSair 
      Caption         =   "&Sair"
      Height          =   540
      Left            =   8160
      Picture         =   "FrmPedAten.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Tag             =   "&Update"
      Top             =   2145
      Width           =   660
   End
   Begin VB.CommandButton CmdAlterar 
      Caption         =   "Detalhe"
      Height          =   615
      Left            =   8145
      Picture         =   "FrmPedAten.frx":00FA
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Alterar Item "
      Top             =   1260
      Width           =   675
   End
End
Attribute VB_Name = "FrmPedAten"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdAdd_Click()
    Unload Me
    gnSequencia = Format(0, "000000000")
    FrmVendas.Show vbModal
End Sub

Private Sub CmdAlterar_Click()
    MsflxOrca.Col = 0
    gnSequencia = MsflxOrca.Text
    Unload Me
    FrmPeDet.Show vbModal
    
End Sub

Private Sub CmdSair_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    gSql = "select nsu,tab_clientes.nome,dta_venda from tab_vendas,tab_clientes "
    gSql = gSql & "where tab_vendas.tipovenda > 0 and tab_clientes.codcli = tab_vendas.codcli"
    gRs.Open gSql, ConDb, adOpenKeyset
    Carrega_Grid_orca
    gRs.Close
End Sub

Private Sub Form_Load()
   'Centraliza a tela no video
   Me.Move (Screen.Width - Me.Width) / 2, _
           (Screen.Height - Me.Height) / 2
    
End Sub

Private Sub Carrega_Grid_orca()
'Teste do MsFlexgrid1
  
  MsflxOrca.Row = 0
  
  MsflxOrca.Col = 0
  MsflxOrca.Text = "Numero"
  MsflxOrca.ColWidth(0) = 900
  MsflxOrca.Col = 1
  MsflxOrca.Text = "Nome do Cliente"
  MsflxOrca.ColWidth(1) = 4400
  MsflxOrca.Col = 2
  MsflxOrca.Text = "Data"
  MsflxOrca.ColWidth(2) = 2000
  
  MsflxOrca.Row = 0
    
  With gRs
      '.MoveLast
      'nItem = .RecordCount
      '.MoveFirst
      MsflxOrca.Rows = 1
      Do While Not .EOF
         MsflxOrca.Rows = MsflxOrca.Rows + 1
         MsflxOrca.Row = MsflxOrca.Rows - 1
         MsflxOrca.Col = 0: MsflxOrca.Text = Format(f_nulo(!nsu, ""), "000000")
         MsflxOrca.Col = 1: MsflxOrca.Text = f_nulo(!nome, "")
         MsflxOrca.Col = 2: MsflxOrca.Text = f_nulo(!dta_venda, "")
         .MoveNext
       Loop
       MsflxOrca.FixedRows = 1
          
  End With

  MsflxOrca.Row = 1
  MsflxOrca.Col = 0
  
  End Sub

