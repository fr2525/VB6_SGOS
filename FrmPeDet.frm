VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form FrmPeDet 
   Caption         =   "Detalhes - <ESC> Sai "
   ClientHeight    =   4740
   ClientLeft      =   1365
   ClientTop       =   885
   ClientWidth     =   9120
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   4740
   ScaleWidth      =   9120
   Begin VB.ComboBox CboClientes 
      Enabled         =   0   'False
      Height          =   315
      Left            =   3450
      TabIndex        =   0
      Top             =   135
      Width           =   5160
   End
   Begin VB.ComboBox CboProdutos 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1365
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   600
      Width           =   4830
   End
   Begin VB.TextBox TxtSelecionado 
      Alignment       =   1  'Right Justify
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
      ForeColor       =   &H000040C0&
      Height          =   360
      Left            =   7275
      TabIndex        =   4
      Text            =   "0"
      Top             =   1125
      Width           =   1350
   End
   Begin VB.CommandButton CmdSair 
      Height          =   480
      Left            =   8175
      Picture         =   "FrmPeDet.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Sai "
      Top             =   3555
      Width           =   435
   End
   Begin VB.ComboBox CboPrecos 
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
      Height          =   360
      Left            =   1695
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1110
      Width           =   1995
   End
   Begin VB.CommandButton cmdimprimir 
      Height          =   480
      Left            =   8175
      Picture         =   "FrmPeDet.frx":00FA
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Imprime o orçamento"
      Top             =   3045
      Width           =   435
   End
   Begin MSFlexGridLib.MSFlexGrid MsflexgridItens 
      Height          =   2475
      Left            =   360
      TabIndex        =   5
      Top             =   1650
      Width           =   7680
      _ExtentX        =   13547
      _ExtentY        =   4366
      _Version        =   393216
      Rows            =   1
      Cols            =   6
      FixedCols       =   0
      Enabled         =   0   'False
      FormatString    =   "Codigo  |  Descrição                                                 |  Qtde.Pd | Qtde.At. |  Pço.Unit.    |    Total  Item  "
   End
   Begin VB.TextBox TxtQtde 
      Alignment       =   1  'Right Justify
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
      Height          =   360
      Left            =   7770
      TabIndex        =   2
      Text            =   "1"
      Top             =   600
      Width           =   840
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Cliente:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   2505
      TabIndex        =   17
      Top             =   180
      Width           =   795
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Produto:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   330
      TabIndex        =   16
      Top             =   660
      Width           =   885
   End
   Begin VB.Label LblNumero 
      Caption         =   "999999"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   270
      Left            =   1500
      TabIndex        =   15
      Top             =   210
      Width           =   1065
   End
   Begin VB.Label LblNumOrca 
      Caption         =   "Orçamento:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   270
      Left            =   135
      TabIndex        =   14
      Top             =   195
      Width           =   1290
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   30
      Left            =   630
      TabIndex        =   13
      Top             =   930
      Width           =   15
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Selecionado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   240
      Left            =   5805
      TabIndex        =   12
      Top             =   1155
      Width           =   1335
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
      Left            =   6150
      TabIndex        =   11
      Top             =   4245
      Width           =   1830
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
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
      Height          =   300
      Left            =   4065
      TabIndex        =   10
      Top             =   4245
      Width           =   1965
   End
   Begin VB.Label Label3 
      Caption         =   "Pço.Sugerido"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   300
      Left            =   150
      TabIndex        =   9
      Top             =   1155
      Width           =   1575
   End
   Begin VB.Label LblQtde 
      Caption         =   "Quantidade:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   270
      Left            =   6450
      TabIndex        =   7
      Top             =   645
      Width           =   1365
   End
End
Attribute VB_Name = "FrmPeDet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim prsProduto As New ADODB.Recordset
Dim prsClientes As New ADODB.Recordset
Dim pRsSequencia As New ADODB.Recordset
Dim prsLoja As New ADODB.Recordset
Dim pRsUnidade As New ADODB.Recordset
Private ControlVisible As Boolean     ' Se o controle esta visivel ou nao
Private LastRow As Long               ' Ultima linha em que se editou
Private LastCol As Long               ' ultima coluna em que se editou
Private pnQtde As Double
Private pcCodprod As String
Private pnPreco As Double
Private pnTotdivida As Double
Private pnNsu  As String
Private pnLinhas As Double
Private pnUnidade As Double
Private pnQtdeA As Double
Private pnQtdeP As Double
Private pnParcelas As Double

Public KeyAscii As Integer

' variaveis de interface

Public GarFlag As Integer
'Public db_dados As Database
Public gDatProc As String
Public item As String

Private pnTotitem As Double
Private pnTotped As Double

Private Sub CboPrecos_Change()
   pnTotitem = CDbl(TxtQtde.Text) * CDbl(CboPrecos.Text)
   Call sutotal
   Me.TxtSelecionado.Enabled = True
   Me.TxtSelecionado.Text = Format(CboPrecos.Text, "###,##0.00")
   Me.TxtSelecionado.SetFocus
End Sub

Private Sub CboPrecos_LostFocus()
 Me.TxtSelecionado.Enabled = True
 Me.TxtSelecionado.Text = CboPrecos.Text
 Me.TxtSelecionado.SetFocus
End Sub

Private Sub CboProdutos_LostFocus()
  Call Carrega_combo_precos
  'TxtPrecounit.Text = Format(IIf(pRsProduto!prevenda1 = 0, pRsProduto!prepromo, pRsProduto!prevenda1), "###,##0.00")
  
  'TxtSelecionado.Text = Format("0", "###,##0.00")
  TxtQtde.Enabled = True
  TxtQtde.Text = Format("1", "###,###")
  TxtQtde.SetFocus
    End Sub

Private Sub CmdAlterar_Click()
  With MsflexgridItens
     .Col = 0
     For i = 0 To CboProdutos.ListCount - 1
         If CboProdutos.ItemData(i) = .Text Then
            CboProdutos.ListIndex = i
            Exit For
         End If
     Next
     'Me.Txtreferencia = .Text
     .Col = 2
     Me.TxtQtde = .Text
     .Col = 3
     Me.CboPrecos.Text = .Text
     'MsflexgridItens.Enabled = True
     If .Rows <= 2 Then
        .Clear
        .Rows = 1
     Else
        .RemoveItem .RowSel
     End If
     Call sutotal
     Me.CboProdutos.SetFocus
  End With
End Sub

Private Sub CmdExcluir_Click()
  MsflexgridItens.Enabled = True
  If MsflexgridItens.Rows <= 2 Then
     'MSFlexGridItens.Clear
     MsflexgridItens.Rows = 1
  Else
     MsflexgridItens.RemoveItem MsflexgridItens.RowSel
  End If
  Call sutotal
  CboProdutos.SetFocus
  
End Sub

Private Sub CmdFecha_Click()
   gSql = "SELECT entrada,dias,parcelas "
   gSql = gSql & "FROM tipovend "
   'gSql = gSql & "WHERE código = " & CboTipovenda.ItemData
   pRstipovenda.Open gSql, ConDb, adOpenKeyset
   If pRstipovenda.EOF And pRstipovenda.BOF Then
      MsgBox "Problemas no arquivo de tipo de venda", vbCritical, "Atenção"
      Unload Me
   End If
  
End Sub

Private Sub CmdExcluir_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    CmdExcluir.Appearance = 1
End Sub

Private Sub cmdfimprod_Click()
Dim resposta
  If MsflexgridItens.Row = 0 Then
     MsgBox "Não digitou nenhum produto", vbOKOnly, "Atenção"
     CboProdutos.SetFocus
     Exit Sub
  End If
    
  With MsflexgridItens
     For i = 1 To .Rows - 1
        .Col = 0
        gSql = "SELECT descricao,estatual FROM tab_produtos "
        gSql = gSql & "Where codprod = '" & .Text & "'"
        pRsProd.Open gSql, ConDb, adOpenKeyset
        .Col = 2
        
'        If Val(.Text) > pRsProd!estatual Then
'           resposta = MsgBox("Estoque de " & Trim(pRsProd!descricao) & " vai ficar negativo. Confirma assim mesmo ?", vbYesNo, "Atenção " & gOperador)
'           If resposta = vbYes Then   ' User chose Yes.
'              'Continua a fazer a checagem
'           Else
'              pRsProd.Close
'              .Col = 0
'              For x = 0 To CboProdutos.ListCount - 1
'                  If CboTipovenda.ItemData(x) = .Text Then
'                     CboTipovenda.ListIndex = i
'                     Exit For
'                  End If
'              Next
'              .Col = 2
'              Me.TxtQtde = .Text
'              .Col = 3
'              Me.CboPrecos.Text = .Text
'              Me.TxtSelecionado.Text = .Text
'              Me.CboProdutos.SetFocus
'              Exit Sub
'           End If
'
'        End If
        pRsProd.Close
     Next
  End With
  
  'Me.Height = 6000
  'Me.Fratipovenda.Visible = True
  'Me.Fratipovenda.Enabled = True
  'Me.CmdTipovenda.Visible = True
  'Me.CmdTipovenda.Enabled = True
  'Me.Fraaprazo.Visible = True
  'Me.Fraaprazo.Enabled = True
   
  
  If pnParcelas > 0 Or (pnParcelas = 0 And _
                        pcEntrada = "N") Then
                                   
     'Venda a prazo então mostra os dados de prazo
     pnAprazo = True
'     Me.Height = 7890
     
'     Me.Fraaprazo.Visible = True
'     Me.Fraaprazo.Enabled = True
             
'     CmdVolta.Visible = True
'     Cmdfinaliza.Visible = True
'     CmdVolta.Enabled = True
     Cmdfinaliza.Enabled = True
     
  Else
     'Cmdfinaliza.Top = CmdTipovenda.Top
     
     CmdTipovenda.Enabled = False
     CmdTipovenda.Visible = False
     Cmdfinaliza.Visible = True
     Cmdfinaliza.Enabled = True
  End If
  
  'pRstipovenda.Close



End Sub

Private Sub CmdPesquisaprod_Click()
  'FrmVendas.TxtReferencia = f_pesqprod()
  Frmpesq.Show vbModal
  Txtreferencia.SetFocus
End Sub

Private Sub CmdFinaliza_Click()

End Sub

'Private Sub CmdGravar_Click()
'   Call suAtualizaOrcamento
'   FrmAPrazo.Show vbModal
'   Unload Me
'
'End Sub

Private Sub cmdimprimir_Click()
  
  gnCodcli = CboClientes.ItemData(CboClientes.ListIndex)
  MsgBox "Orçamento No. " & Format(gnSequencia, "000000") & " será impresso", vbOKOnly, "Atenção " & gOperador
  Call suImprime
  
  Unload Me
  FrmPedAten.Show vbModal
   
End Sub

Private Sub CmdSair_Click()
   
   Unload Me
   FrmPedAten.Show vbModal
   
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyReturn Then
      SendKeys "{TAB}"
      KeyAscii = 0
   End If
   
   If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
  'Centraliza a tela no video
   Me.Move (Screen.Width - Me.Width) / 2, _
           (Screen.Height - Me.Height) / 2
  
  If gnSequencia = 0 Then
     MsgBox "Houve problema ao selecionar o Pedido", vbOKOnly, "Atenção " & gOperador
     Unload Me
     FrmPedAten.Show vbModal
  Else
     Call suCarregaOrcamento
     Call sutotal
  End If
     
  Me.LblNumero.Caption = Format(gnSequencia, "#####0")
      
End Sub


Private Sub Image2_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
   Image2.Appearance = 1
End Sub

Private Sub suCarregaOrcamento()
   
   gSql = "select * from tab_vendas WHERE nsu = '" & Format(Str(gnSequencia), "000000000") & "'"
   gRs.Open gSql, ConDb, adOpenKeyset
   If gRs.BOF And gRs.EOF Then
      MsgBox "houve erro na carga do Orcamento. Programa será cancelado ", vbOKOnly, "Atenção, " & gOperador
      Unload Me
   End If
   
   Call Abre_Le_rst_clientes   'Carrega o combo de clientes
   
   For i = 0 To CboClientes.ListCount - 1
       If gRs!codcli = CboClientes.ItemData(i) Then
          CboClientes.ListIndex = i
          Exit For
       End If
   Next
   
   gRs.Close
   
   Call Abre_Le_rst_Produtos
   CboProdutos.ListIndex = 0
   TxtQtde.Text = 1
   
   gSql = "select tab_itemvenda.codprod,tab_produtos.descricao,qtdep,qtdea,precounit, qtdep * precounit as totalitem from tab_itemvenda,tab_produtos "
   gSql = gSql & " WHERE nsu = '" & Format(Str(gnSequencia), "000000000") & "'"
   gSql = gSql & " AND tab_itemvenda.codprod = tab_produtos.codprod "
   gRs.Open gSql, ConDb, adOpenKeyset
   If gRs.BOF And gRs.EOF Then
      MsgBox "houve erro na carga dos Itens do Orcamento. Programa será cancelado ", vbOKOnly, "Atenção, " & gOperador
      Unload Me
   End If
      

  MsflexgridItens.Row = 0
  MsflexgridItens.FontWidth = 1
  
  With gRs
      '.MoveLast
      'nItem = .RecordCount
      .MoveFirst
      MsflexgridItens.Rows = 1
      i = 0
      Do While Not .EOF
        
         pcCodprod = Format(!codprod, "000000")
         'Call suPegaUnidade
         'pnQtdeA = Round(Val(TxtQtde.Text) / pnUnidade, 0)
         'pnQtdeP = Round(Val(TxtQtde.Text) / pnUnidade, 0) * pnUnidade
         
         MsflexgridItens.Rows = MsflexgridItens.Rows + 1
         MsflexgridItens.Row = MsflexgridItens.Rows - 1
         MsflexgridItens.Col = 0: MsflexgridItens.Text = "" & !codprod
         MsflexgridItens.Col = 1: MsflexgridItens.Text = "" & !descricao
         'MsflexgridItens.Col = 2: MsflexgridItens.Text = f_nulo(pnQtdeA, 0)
         'MsflexgridItens.Col = 3: MsflexgridItens.Text = f_nulo(pnQtdeP, 0)
         MsflexgridItens.Col = 2: MsflexgridItens.Text = f_nulo(!QtdeP, 0)
         MsflexgridItens.Col = 3: MsflexgridItens.Text = f_nulo(!QtdeA, 0)
         MsflexgridItens.Col = 4: MsflexgridItens.Text = Format(f_nulo(!precounit, 0), "###,##0.00")
         MsflexgridItens.Col = 5: MsflexgridItens.Text = Format(f_nulo(!totalitem, 0), "###,##0.00")
        
         .MoveNext
         
       Loop
      MsflexgridItens.FixedRows = 1
          
  End With

  gRs.Close
  'CboClientes.SetFocus
   
    
End Sub
Private Sub Abre_Le_rst_clientes()
 
   gSql = "select codcli,nome "
   gSql = gSql & "FROM tab_clientes "
   gSql = gSql & " order by Nome "
   prsClientes.Open gSql, ConDb, adOpenKeyset
   Carrega_Grid_Clientes
   prsClientes.Close

End Sub
Private Sub Carrega_Grid_Clientes()

 CboClientes.Clear
 With prsClientes
      '.MoveLast
      'nItem = .RecordCount
      .MoveFirst
      Do While Not .EOF
         CboClientes.AddItem (prsClientes!Nome)
         CboClientes.ItemData(CboClientes.NewIndex) = prsClientes!codcli
        .MoveNext
      Loop
  End With
     
End Sub

Private Sub Abre_Le_rst_Produtos()
 
   gSql = "select codprod,descricao,prevenda1,prevenda2,prevenda3,prevenda4,prevenda5 "
   gSql = gSql & "FROM tab_produtos "
   gSql = gSql & " order by descricao "
   prsProduto.Open gSql, ConDb, adOpenKeyset
   Carrega_combo_Produtos
   prsProduto.Close

End Sub
Private Sub Carrega_combo_Produtos()

 CboProdutos.Clear
 With prsProduto
      '.MoveLast
      'nItem = .RecordCount
      .MoveFirst
      Do While Not .EOF
         CboProdutos.AddItem (prsProduto!descricao)
         CboProdutos.ItemData(CboProdutos.NewIndex) = prsProduto!codprod
        .MoveNext
      Loop
  End With
     
End Sub

Private Sub sutotal()
   gnTotPed = 0
   For i = 1 To MsflexgridItens.Rows - 1
       MsflexgridItens.Row = i
       MsflexgridItens.Col = 5
       gnTotPed = gnTotPed + Me.MsflexgridItens.Text
   Next
   LbltotaldoPedido.Caption = Format(gnTotPed, "###,###,##0.00")
End Sub


Private Sub Carrega_combo_precos()
   
   gSql = "select prevenda1,prevenda2,prevenda3,prevenda4,prevenda5 "
   gSql = gSql & "FROM tab_produtos "
   gSql = gSql & " WHERE codprod =  '" & Format(CboProdutos.ItemData(CboProdutos.ListIndex), "000000") & "'"
   prsProduto.Open gSql, ConDb, adOpenKeyset

   CboPrecos.Clear
   CboPrecos.AddItem Format(prsProduto!prevenda1, "###,##0.00")
   CboPrecos.AddItem Format(prsProduto!prevenda2, "###,##0.00")
   CboPrecos.AddItem Format(prsProduto!prevenda3, "###,##0.00")
   CboPrecos.AddItem Format(prsProduto!prevenda4, "###,##0.00")
   CboPrecos.AddItem Format(prsProduto!prevenda5, "###,##0.00")
   CboPrecos.ListIndex = 0
   CboPrecos.Enabled = True
   
   prsProduto.Close
    
End Sub

Private Sub suPegaUnidade()
   gSql = "select uni_qtd "
   gSql = gSql & " FROM tab_uni,tab_produtos  "
   gSql = gSql & " WHERE tab_produtos.codprod = '" & pcCodprod & "'"
   gSql = gSql & " AND tab_uni.uni_cod = tab_produtos.unidade"
   pRsUnidade.Open gSql, ConDb, adOpenKeyset
   If pRsUnidade.BOF And pRsUnidade.EOF Then
      pnUnidade = 1
   Else
      pnUnidade = pRsUnidade!uni_qtd
   End If
   pRsUnidade.Close
  
End Sub

