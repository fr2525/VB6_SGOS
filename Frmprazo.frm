VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmPrazo 
   Caption         =   "A Prazo"
   ClientHeight    =   4920
   ClientLeft      =   3060
   ClientTop       =   1500
   ClientWidth     =   9450
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   4920
   ScaleWidth      =   9450
   Begin VB.Frame Fraaprazo 
      Height          =   2595
      Left            =   105
      TabIndex        =   9
      Top             =   2100
      Visible         =   0   'False
      Width           =   8310
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   1215
         Left            =   2355
         TabIndex        =   16
         Top             =   1050
         Width           =   3420
         _ExtentX        =   6033
         _ExtentY        =   2143
         _Version        =   393216
         Cols            =   3
         FixedCols       =   0
         FormatString    =   "No.Cheque      |^    Data        |>              Valor  "
      End
      Begin VB.TextBox TxtAgencia 
         Height          =   285
         Left            =   5820
         TabIndex        =   12
         Top             =   480
         Width           =   2295
      End
      Begin VB.TextBox TxtBanco 
         Height          =   285
         Left            =   3705
         TabIndex        =   11
         Top             =   480
         Width           =   2010
      End
      Begin VB.ComboBox CboClientes 
         Height          =   315
         Left            =   105
         TabIndex        =   10
         Top             =   465
         Width           =   3480
      End
      Begin VB.Label LblAgencia 
         Caption         =   "Agência"
         Height          =   225
         Left            =   5925
         TabIndex        =   15
         Top             =   195
         Width           =   885
      End
      Begin VB.Label LblBanco 
         Caption         =   "Banco"
         Height          =   240
         Left            =   3750
         TabIndex        =   14
         Top             =   195
         Width           =   795
      End
      Begin VB.Label LblCliente 
         Caption         =   "Cliente"
         Height          =   255
         Left            =   135
         TabIndex        =   13
         Top             =   165
         Width           =   1260
      End
   End
   Begin VB.Frame Fratipovenda 
      Height          =   1305
      Left            =   120
      TabIndex        =   5
      Top             =   660
      Visible         =   0   'False
      Width           =   8340
      Begin VB.TextBox TxtVlrentrada 
         Height          =   300
         Left            =   2520
         TabIndex        =   7
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Lbltipovenda 
         Caption         =   "Tipo de venda:"
         Height          =   285
         Left            =   90
         TabIndex        =   8
         Top             =   180
         Width           =   1170
      End
      Begin VB.Label LblVlrentrada 
         Caption         =   "Vlr.Entrada:"
         Height          =   240
         Left            =   2520
         TabIndex        =   6
         Top             =   195
         Width           =   1230
      End
   End
   Begin VB.CommandButton cmdfimprod 
      Height          =   480
      Left            =   8760
      Picture         =   "Frmprazo.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Finaliza o pedido"
      Top             =   2160
      Width           =   585
   End
   Begin VB.CommandButton CmdAlterar 
      Height          =   435
      Left            =   8775
      Picture         =   "Frmprazo.frx":0104
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Alterar Item "
      Top             =   1560
      Width           =   525
   End
   Begin VB.CommandButton CmdExcluir 
      Height          =   450
      Left            =   8760
      Picture         =   "Frmprazo.frx":0276
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Excluir Item"
      Top             =   945
      Width           =   540
   End
   Begin VB.Label LbltotaldoPedido 
      Alignment       =   1  'Right Justify
      Caption         =   "R$ 000.000,00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   315
      Left            =   2340
      TabIndex        =   3
      Top             =   180
      Width           =   1830
   End
   Begin VB.Label Label4 
      Caption         =   "Total do Pedido:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   270
      Left            =   240
      TabIndex        =   2
      Top             =   180
      Width           =   2025
   End
End
Attribute VB_Name = "FrmPrazo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim pRsProduto As New adodb.Recordset
Dim pRsFornec As New adodb.Recordset
Dim pRstipovenda As New adodb.Recordset
Dim pRsBalconista As New adodb.Recordset
Dim pRscliente As New adodb.Recordset

Private pnTotitem As Double
Private pnTotped As Double
Function f_pesqprod()
   Frmpesqprod.Show vbModal
   f_pesqprod = Frmpesqprod.pcCodprod
   Unload Frmpesqprod
End Function

Private Sub CmdExcluir_Click()
  MsflexgridItens.Enabled = True
  If MsflexgridItens.Rows <= 2 Then
     MsflexgridItens.Clear
  Else
     MsflexgridItens.RemoveItem MsflexgridItens.RowSel
  End If
  TxtReferencia.SetFocus
End Sub

Private Sub cmdfimprod_Click()
  Me.Height = 7650
  Me.Fratipovenda.Visible = True
  Me.Fratipovenda.Enabled = True
  Me.Fraaprazo.Visible = True
  Me.Fraaprazo.Enabled = True
  CboTipovenda.ListIndex = 0
  CboBalconista.ListIndex = 0
End Sub

Private Sub CmdPesquisaprod_Click()
  'FrmVendas.TxtReferencia = f_pesqprod()
  Frmpesqprod.Show vbModal
  TxtReferencia.SetFocus
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
   If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
  Me.Top = 1150
  Me.Height = 3420
  pnTotped = 0#
  LbltotaldoPedido.Caption = Format(pnTotped, "###,###,##0.00")
  'Centraliza a tela no video
  ' Me.Move (Screen.Width - Me.Width) / 2, _
  '         (Screen.Height - Me.Height) / 2
   gSql = "Select codfor,nome from cadfor"
   pRsFornec.Open gSql, ConDb, adOpenKeyset
   If pRsFornec.BOF And pRsFornec.EOF Then
      MsgBox "Não existem fornecedores no cadastro. Favor cadastrar", vbOKOnly, "Atenção"
      Unload Me
   End If
          
   gSql = "select * from produtos where produtos.ativo = 'S'"
   pRsProduto.Open gSql, ConDb, adOpenKeyset
      
   MsflexgridItens.Cols = 5
   MsflexgridItens.Rows = 1
   MsflexgridItens.Row = 0
   MsflexgridItens.Col = 0
   MsflexgridItens.Text = "Referencia"
   MsflexgridItens.Col = 1
   MsflexgridItens.ColWidth(1) = 4330
   MsflexgridItens.Text = "Descricao                      "
   MsflexgridItens.Col = 2
   MsflexgridItens.Text = "Qtde."
   MsflexgridItens.Col = 3
   MsflexgridItens.Text = "Preço Unit."
   MsflexgridItens.Col = 4
   MsflexgridItens.Text = "Total Item"
    
   suCarrega_Grids
   
    
End Sub

Private Sub MsflexgridItens_Click()
   MsflexgridItens.RowSel = 1
End Sub

Private Sub TxtPrecounit_LostFocus()
   pnTotitem = Val(TxtQtde.Text) * Val(TxtPrecounit.Text)
   MsflexgridItens.AddItem TxtReferencia.Text & vbTab _
                         & LblDescricao.Caption & vbTab _
                         & TxtQtde.Text & vbTab & TxtPrecounit.Text & _
                           vbTab & Format(pnTotitem, "###,##0.00")
   pnTotped = 0
   For i = 0 To MsflexgridItens.Rows - 1
       MsflexgridItens.Row = i
       MsflexgridItens.Col = 4
       pnTotped = pnTotped + Val(MsflexgridItens.Text)
   Next
   LbltotaldoPedido.Caption = Format(pnTotped, "###,###,##0.00")
   TxtReferencia.Text = ""
   TxtReferencia.SetFocus
End Sub

Private Sub TxtReferencia_KeyPress(KeyAscii As Integer)
   If KeyAscii = vbKeyEscape Then Unload Me
End Sub

Private Sub TxtReferencia_LostFocus()
   'If Len(TxtReferencia.Text) = 0 Then
   '   If (MsgBox("Deseja Finalizar o Pedido?", vbYesNo, "Atencao")) = vbNo Then
   '      TxtReferencia.SetFocus
   '   Else
   '      cmdfimprod_Click
   '   End If
   'Else
   If IsNumeric(TxtReferencia.Text) Then
      pRsProduto.Find "[codprod] = '" & TxtReferencia.Text & "'"
      If pRsProduto.EOF Then
         MsgBox "Produto não encontrado", vbOKOnly
         TxtReferencia.SetFocus
      Else
         If pRsProduto!prevenda1 = 0 Then
            MsgBox "Produto sem preço de venda. Verifique", vbOKOnly
            TxtReferencia.SetFocus
         Else
            LblDescricao.Caption = pRsProduto!descricao
            TxtPrecounit.Text = Format(pRsProduto!prevenda1, "###,##0.00")
            TxtQtde.Text = Format("1", "###,###")
            TxtQtde.SetFocus
         End If
      End If
   End If
   'End If
End Sub
Private Sub suCarrega_Grids()
   Abre_Le_rst_tipovenda
   Abre_Le_rst_Balconistas
   Abre_Le_rst_clientes
End Sub
Private Sub Abre_Le_rst_tipovenda()
   gSql = "select código,descricao "
   gSql = gSql & "FROM tipovend "
   pRstipovenda.Open gSql, ConDb, adOpenKeyset
   Carrega_Grid_tipovenda
    
End Sub
Private Sub Carrega_Grid_tipovenda()

 With pRstipovenda
      '.MoveLast
      'nItem = .RecordCount
      .MoveFirst
      Do While Not .EOF
        CboTipovenda.AddItem (pRstipovenda!descricao)
        CboTipovenda.ItemData(CboTipovenda.NewIndex) = pRstipovenda!código
        .MoveNext
      Loop
  End With
     
End Sub

Private Sub Abre_Le_rst_Balconistas()
   gSql = "select codvend,nome "
   gSql = gSql & "FROM cadvend "
   pRsBalconista.Open gSql, ConDb, adOpenKeyset
   Carrega_Grid_Balconista
    
End Sub
Private Sub Carrega_Grid_Balconista()

 With pRsBalconista
      '.MoveLast
      'nItem = .RecordCount
      .MoveFirst
      Do While Not .EOF
        CboBalconista.AddItem (pRsBalconista!nome)
        CboBalconista.ItemData(CboBalconista.NewIndex) = pRsBalconista!codvend
        .MoveNext
      Loop
  End With
     
End Sub
Private Sub Abre_Le_rst_clientes()
   gSql = "select codigo,nome "
   gSql = gSql & "FROM cadclie "
   pRscliente.Open gSql, ConDb, adOpenKeyset
   Carrega_Grid_clientes
    
End Sub
Private Sub Carrega_Grid_clientes()

 With pRscliente
      '.MoveLast
      'nItem = .RecordCount
      .MoveFirst
      Do While Not .EOF
        CboClientes.AddItem (pRscliente!nome)
        CboClientes.ItemData(CboClientes.NewIndex) = pRscliente!codigo
        .MoveNext
      Loop
  End With
     
End Sub

