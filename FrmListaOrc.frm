VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form FrmListaOrc 
   Caption         =   "Orçamentos Pendentes"
   ClientHeight    =   4440
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8580
   LinkTopic       =   "Form1"
   ScaleHeight     =   4440
   ScaleWidth      =   8580
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   2565
      Left            =   7500
      TabIndex        =   1
      Top             =   705
      Width           =   780
      Begin VB.CommandButton CmdSair 
         Caption         =   "&Sair"
         Height          =   540
         Left            =   75
         Picture         =   "FrmListaOrc.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   5
         Tag             =   "&Update"
         Top             =   1905
         Width           =   615
      End
      Begin VB.CommandButton CmdExcluir 
         Caption         =   "Excluir"
         Height          =   540
         Left            =   75
         Picture         =   "FrmListaOrc.frx":00FA
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Alterar Item "
         Top             =   1320
         Width           =   615
      End
      Begin VB.CommandButton CmdAlterar 
         Caption         =   "Alterar"
         Height          =   540
         Left            =   75
         Picture         =   "FrmListaOrc.frx":026C
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Alterar Item "
         Top             =   735
         Width           =   615
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Novo"
         Height          =   540
         Left            =   75
         Picture         =   "FrmListaOrc.frx":03DE
         Style           =   1  'Graphical
         TabIndex        =   2
         Tag             =   "&Add"
         Top             =   165
         Width           =   615
      End
   End
   Begin MSFlexGridLib.MSFlexGrid MsflxOrca 
      Height          =   3375
      Left            =   360
      TabIndex        =   0
      Top             =   435
      Width           =   6915
      _ExtentX        =   12197
      _ExtentY        =   5953
      _Version        =   393216
      Cols            =   3
      FixedCols       =   0
      BackColorSel    =   16777215
      ForeColorSel    =   0
      FocusRect       =   0
      HighLight       =   2
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
End
Attribute VB_Name = "FrmListaOrc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdAdd_Click()
    Unload Me
    'gnSequencia = Format(0, "000000000")
    gnSequencia = 0
    FrmVendas.Show vbModal
End Sub

Private Sub CmdAlterar_Click()
    MsflxOrca.Col = 0
    gnSequencia = MsflxOrca.Text
    Unload Me
    FrmVendas.Show vbModal
    
End Sub

Private Sub CmdSair_Click()
    Unload Me
End Sub

Private Sub Command1_Click()
   MsflxOrca.Col = 0
   gnSequencia = Val(MsflxOrca.Text)
   If MsgBox("Deseja Realmente excluir o orçamento de no. " & Format(gnSequencia, "000000") & " ??? ", vbYesNo, "Atenção " & gOperador) = vbYes Then
      gSql = "DELETE FROM tab_vendas WHERE nsu = '" & Format(gnSequencia, "000000000") & "'"
      ConDb.Execute gSql
      gSql = "DELETE FROM tab_itemvenda WHERE nsu = '" & Format(gnSequencia, "000000000") & "'"
      ConDb.Execute gSql
      MsgBox "Orçamento de no. " & Format(gnSequencia, "000000") & " Foi Excluido ", vbOKOnly, " Olá " & gOperador
      suCarregaDados
      
   End If
   
End Sub

Private Sub Form_Activate()
    suCarregaDados
End Sub

Private Sub Form_Load()
   'Centraliza a tela no video
   Me.Move (Screen.Width - Me.Width) / 2, _
           (Screen.Height - Me.Height) / 2
    
End Sub

Private Sub suCarregaDados()

    gSql = "select nsu,tab_clientes.nome,dta_venda from tab_vendas,tab_clientes "
    gSql = gSql & "where tab_vendas.tipovenda = 0 and tab_clientes.codcli = tab_vendas.codcli"
    gRs.Open gSql, ConDb, adOpenKeyset
    If gRs.BOF And gRs.EOF Then
       MsgBox "Arquivo de Orçamentos está vazio. Entre com um novo ", vbOKOnly, "Atenção " & gOperador
       Me.cmdAdd.SetFocus
       Me.CmdAlterar.Enabled = False
       Me.CmdExcluir.Enabled = False
    Else
       Carrega_Grid_orca
    End If
    gRs.Close
    
End Sub
Private Sub Carrega_Grid_orca()
'Teste do MsFlexgrid1
  
  MsflxOrca.Row = 0
  
  MsflxOrca.Col = 0
  MsflxOrca.Text = "Numero"
  MsflxOrca.ColWidth(0) = 1100
  MsflxOrca.Col = 1
  MsflxOrca.Text = "Nome do Cliente"
  MsflxOrca.ColWidth(1) = 4400
  MsflxOrca.Col = 2
  MsflxOrca.Text = "Data"
  MsflxOrca.ColWidth(2) = 1300
  
  MsflxOrca.Row = 0
    
  With gRs
      '.MoveLast
      'nItem = .RecordCount
      '.MoveFirst
      MsflxOrca.Rows = 1
      Do While Not .EOF
         MsflxOrca.Rows = MsflxOrca.Rows + 1
         MsflxOrca.Row = MsflxOrca.Rows - 1
         MsflxOrca.Col = 0: MsflxOrca.Text = f_nulo(!nsu, "")
         MsflxOrca.Col = 1: MsflxOrca.Text = f_nulo(!nome, "")
         MsflxOrca.Col = 2: MsflxOrca.Text = f_nulo(!dta_venda, "")
         .MoveNext
       Loop
       MsflxOrca.FixedRows = 1
          
  End With

  MsflxOrca.Row = 1
  MsflxOrca.Col = 0
  
  End Sub

Private Sub MsflxOrca_Click()
    
    PintaGrid MsflxOrca
    
End Sub
