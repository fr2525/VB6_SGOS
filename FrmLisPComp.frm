VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form FrmLisPComp 
   Caption         =   "Pedidos de Compra"
   ClientHeight    =   4350
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8640
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   4350
   ScaleWidth      =   8640
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   2940
      Left            =   7440
      TabIndex        =   1
      Top             =   630
      Width           =   930
      Begin VB.CommandButton CmdSair 
         Caption         =   "&Sair"
         Height          =   615
         Left            =   75
         Picture         =   "FrmLisPComp.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   5
         Tag             =   "&Update"
         Top             =   2235
         Width           =   750
      End
      Begin VB.CommandButton CmdExcluir 
         Caption         =   "Excluir"
         Height          =   615
         Left            =   75
         Picture         =   "FrmLisPComp.frx":00FA
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Alterar Item "
         Top             =   1545
         Width           =   750
      End
      Begin VB.CommandButton CmdAlterar 
         Caption         =   "Receber"
         Height          =   615
         Left            =   75
         Picture         =   "FrmLisPComp.frx":026C
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Alterar Item "
         Top             =   855
         Width           =   750
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Novo"
         Height          =   600
         Left            =   75
         Picture         =   "FrmLisPComp.frx":03DE
         Style           =   1  'Graphical
         TabIndex        =   2
         Tag             =   "&Add"
         Top             =   180
         Width           =   750
      End
   End
   Begin MSFlexGridLib.MSFlexGrid MsflxOrca 
      Height          =   3375
      Left            =   270
      TabIndex        =   0
      Top             =   465
      Width           =   6915
      _ExtentX        =   12197
      _ExtentY        =   5953
      _Version        =   393216
      Cols            =   3
      FixedCols       =   0
      BackColorSel    =   65535
      ForeColorSel    =   0
      FocusRect       =   0
      HighLight       =   2
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
End
Attribute VB_Name = "FrmLisPComp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdAdd_Click()
  Unload Me
  'gnSequencia = Format(0, "000000000")
  gnSequencia = 0
  FrmCompras.Show vbModal
End Sub

Private Sub CmdAlterar_Click()
    MsflxOrca.Col = 0
    gnSequencia = MsflxOrca.Text
    Unload Me
    FrmCompras.Show vbModal
End Sub

Private Sub CmdExcluir_Click()
   MsflxOrca.Col = 0
   gnSequencia = Val(MsflxOrca.Text)
   If MsgBox("Deseja Realmente excluir o Pedido de no. " & Format(gnSequencia, "000000") & " ??? ", vbYesNo, "Atenção " & gOperador) = vbYes Then
      gSql = "DELETE FROM tab_compra WHERE numped = '" & Format(gnSequencia, "000000000") & "'"
      ConDb.Execute gSql
      gSql = "DELETE FROM tab_itemcompra WHERE numped = '" & Format(gnSequencia, "000000000") & "'"
      ConDb.Execute gSql
      MsgBox "Pedido de no. " & Format(gnSequencia, "000000") & " Foi Excluido ", vbOKOnly, " Olá " & gOperador
      MsflxOrca.Clear
      suCarregaDados
      
   End If

End Sub

Private Sub CmdSair_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
   suCarregaDados
End Sub

Private Sub suCarregaDados()

    gSql = "select numped,tab_fornece.nome,dataped from tab_compra,tab_fornece "
    gSql = gSql & "where tab_compra.notafisc = '' and tab_fornece.codfor = tab_compra.codfor"
    gSql = gSql & " ORDER BY dataped DESC"
    gRs.Open gSql, ConDb, adOpenKeyset
    If gRs.BOF And gRs.EOF Then
       MsgBox "Arquivo de Pedidos está vazio. Entre com um novo ", vbOKOnly, "Atenção " & gOperador
       Me.cmdAdd.SetFocus
       Me.CmdAlterar.Enabled = False
       Me.CmdExcluir.Enabled = False
    Else
       Carrega_Grid_pedidos
    End If
    gRs.Close
    
End Sub
Private Sub Carrega_Grid_pedidos()
'Teste do MsFlexgrid1
  
  MsflxOrca.Row = 0
  
  MsflxOrca.Col = 0
  MsflxOrca.Text = "Numero"
  MsflxOrca.ColWidth(0) = 1100
  MsflxOrca.Col = 1
  MsflxOrca.Text = "Nome do Fornecedor"
  MsflxOrca.ColWidth(1) = 4400
  MsflxOrca.Col = 2
  MsflxOrca.Text = "Dt.Ped"
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
         MsflxOrca.Col = 0: MsflxOrca.Text = f_nulo(!numped, "")
         MsflxOrca.Col = 1: MsflxOrca.Text = f_nulo(!nome, "")
         MsflxOrca.Col = 2: MsflxOrca.Text = Format(f_nulo(!dataped, ""), "dd/mm/yyyy")
         .MoveNext
       Loop
       MsflxOrca.FixedRows = 1
          
  End With

  MsflxOrca.Row = 1
  MsflxOrca.Col = 0
  
  End Sub


Private Sub Form_Load()
 
   Me.Move (Screen.Width - Me.Width) / 2, _
           (Screen.Height - Me.Height) / 2
    
End Sub
