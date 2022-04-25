VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form FrmBalcao 
   ClientHeight    =   6585
   ClientLeft      =   4680
   ClientTop       =   2160
   ClientWidth     =   10800
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6585
   ScaleWidth      =   10800
   Begin VB.TextBox TxtCodBarras 
      Height          =   375
      Left            =   1860
      MaxLength       =   13
      TabIndex        =   18
      Top             =   180
      Width           =   1335
   End
   Begin VB.ComboBox CboTipovenda 
      Height          =   315
      Left            =   405
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   4830
      Width           =   4365
   End
   Begin VB.Frame Frame2 
      Height          =   1035
      Left            =   7545
      TabIndex        =   16
      Top             =   5295
      Width           =   2970
      Begin VB.CommandButton cmdimprimir 
         Caption         =   "(F8) Imprime"
         Height          =   810
         Left            =   60
         Picture         =   "FrmBalcao.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Imprime o orçamento"
         Top             =   165
         Width           =   825
      End
      Begin VB.CommandButton CmdSair 
         Caption         =   "Saida"
         Height          =   810
         Left            =   2070
         Picture         =   "FrmBalcao.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Sai "
         Top             =   165
         Width           =   825
      End
      Begin VB.CommandButton CmdGravar 
         Caption         =   "(F9) Finaliza"
         Height          =   810
         Left            =   1035
         Picture         =   "FrmBalcao.frx":01FC
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Finaliza os produtos"
         Top             =   165
         Width           =   870
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Itens"
      Height          =   2040
      Left            =   8310
      TabIndex        =   15
      Top             =   1365
      Width           =   1320
      Begin VB.CommandButton CmdExcluir 
         Appearance      =   0  'Flat
         Caption         =   "(F5) Exclui Item"
         Height          =   870
         Left            =   75
         Picture         =   "FrmBalcao.frx":02F6
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Excluir Item"
         Top             =   1110
         Width           =   1170
      End
      Begin VB.CommandButton CmdAlterar 
         Caption         =   "(F4) Altera Item"
         Height          =   855
         Left            =   60
         Picture         =   "FrmBalcao.frx":03F8
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Alterar Item "
         Top             =   195
         Width           =   1200
      End
   End
   Begin VB.ComboBox CboClientes 
      Height          =   315
      Left            =   5400
      TabIndex        =   4
      Top             =   4815
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
      Left            =   1830
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   165
      Visible         =   0   'False
      Width           =   7695
   End
   Begin MSFlexGridLib.MSFlexGrid MsflexgridItens 
      Height          =   3015
      Left            =   540
      TabIndex        =   8
      Top             =   870
      Width           =   7680
      _ExtentX        =   13547
      _ExtentY        =   5318
      _Version        =   393216
      Rows            =   1
      Cols            =   6
      FixedCols       =   0
      FormatString    =   "Codigo  |  Descrição                                                 |  Qtde.Pd | Qtde.At. |  Pço.Unit.    |    Total  Item  "
   End
   Begin VB.Label Lbltipovenda 
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo de venda:"
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
      Height          =   285
      Left            =   435
      TabIndex        =   17
      Top             =   4470
      Width           =   1980
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      Left            =   5415
      TabIndex        =   14
      Top             =   4515
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
      Left            =   735
      TabIndex        =   13
      Top             =   225
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
      Left            =   1860
      TabIndex        =   12
      Top             =   5880
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
      Left            =   390
      TabIndex        =   11
      Top             =   5880
      Width           =   1290
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
      TabIndex        =   10
      Top             =   4155
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
      Left            =   4125
      TabIndex        =   9
      Top             =   4155
      Width           =   1965
   End
End
Attribute VB_Name = "FrmBalcao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim prsProduto As New ADODB.Recordset
Dim prsClientes As New ADODB.Recordset
Dim pRsSequencia As New ADODB.Recordset
Dim pRstipovenda As New ADODB.Recordset
Dim prsLoja As New ADODB.Recordset
Dim pRsUnidade As New ADODB.Recordset
Dim prsCusto As New ADODB.Recordset
Dim prsEstoque As New ADODB.Recordset
Dim prsPrevisao As New ADODB.Recordset
Private ControlVisible As Boolean     ' Se o controle esta visivel ou nao
Private LastRow As Long               ' Ultima linha em que se editou
Private LastCol As Long               ' ultima coluna em que se editou
Private pnQtdeP As Double
Private pnQtdeA As Double
Private pcCodprod As String
Private pnPreco As Double
Private pnCusto As Double
Private pnTotdivida As Double
Private pnNsu  As String
Private pnLinhas As Double
Private pnUnidade As Double
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
  
 End Sub

Private Sub CboTipovenda_Click()

  gSql = "select código,descricao, entrada,dias,parcelas "
  gSql = gSql & "FROM tipovend WHERE código = " & CboTipovenda.ItemData(CboTipovenda.ListIndex)
  pRstipovenda.Open gSql, ConDb, adOpenKeyset
  If pRstipovenda.BOF And pRstipovenda.BOF Then
     MsgBox "Erro grave. Não achou o tipo de venda", vbOKOnly, "Atenção " & gOperador
     End
  End If
  pcEntrada = pRstipovenda!Entrada
  pnParcelas = pRstipovenda!parcelas
  pnDias = pRstipovenda!dias
  If pcEntrada = "S" Then
     'Me.TxtVlrentrada.Enabled = True
     If pnParcelas = 0 Then
        gnAPrazo = False
     '   Me.TxtVlrentrada.Text = Format(gnTotPed, "###,###,##0.00")
     Else
        gnAPrazo = True
     End If
     'Me.TxtVlrentrada.SetFocus
  Else
     'Me.TxtVlrentrada.Enabled = False
  End If
  
  pRstipovenda.Close

End Sub

Private Sub CmdAddItem_Click()
   
   pcCodprod = Format(CboProdutos.ItemData(CboProdutos.ListIndex), "000000")
   
   Call suPegaUnidade

   pnQtdeA = Round(Val(TxtQtde.Text) / pnUnidade, 0)
   pnQtdeP = Round(Val(TxtQtde.Text) / pnUnidade, 0) * pnUnidade

   pnTotitem = CDbl(pnQtdeP) * CDbl(Me.TxtSelecionado.Text)
   
   MsflexgridItens.AddItem Format(CboProdutos.ItemData(CboProdutos.ListIndex), "000000") & vbTab _
                         & CboProdutos.Text & vbTab _
                         & pnQtdeP & vbTab & pnQtdeA & vbTab & Me.TxtSelecionado.Text & _
                           vbTab & Format(pnTotitem, "###,##0.00")
   Call sutotal
   TxtQtde.Enabled = False
   CboPrecos.Enabled = False
   TxtSelecionado.Enabled = False
   CboProdutos.SetFocus

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
     'Me.CboPrecos.Text = .Text
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
  'FrmBalcao.TxtReferencia = f_pesqprod()
  Frmpesq.Show vbModal
  Txtreferencia.SetFocus
End Sub

Private Sub CmdGravar_Click()
   If fuVeEstoque Then
      Call suAtualizaOrcamento
      Unload Me
      If gnAPrazo Then
          FrmAPrazo.Show vbModal
      End If
   Else
   End If
   
End Sub

Private Sub cmdimprimir_Click()
Dim resposta
  If MsflexgridItens.Row = 0 Then
     MsgBox "Não digitou nenhum produto", vbOKOnly, "Atenção"
     CboProdutos.SetFocus
     Exit Sub
  End If
    
  Call suAtualizaOrcamento
 
  MsgBox "Orçamento No. " & Format(gnSequencia, "000000") & " será impresso", vbOKOnly, "Atenção " & gOperador
  
  Call suImprime
 
  'Unload Me
  'FrmListaOrc.Show vbModal
   
End Sub

Private Sub CmdSair_Click()
   '--> Apaga o numero de venda não usados
   gSql = "DELETE FROM tab_vendas WHERE tipovenda = 0 And codcli = 0 And codvend = 0 "
   ConDb.Execute gSql
   Unload Me
   'FrmListaOrc.Show vbModal
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyReturn Then
      SendKeys "{TAB}"
      KeyAscii = 0
   End If
   
   If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
   
  Call Centra(Me)
  
  'Me.Top = 1150
  'Me.Height = 4140
  If gnSequencia = 0 Then
     gSql = "SELECT MAX(VAL(NSU)) AS sequencia FROM tab_vendas "
     pRsSequencia.Open gSql, ConDb, adOpenKeyset
     If pRsSequencia.BOF And pRsSequencia.EOF Then
        gnSequencia = 1
     Else
        If IsNull(pRsSequencia!sequencia) Then
           gnSequencia = 1
        Else
           gnSequencia = Val(pRsSequencia!sequencia) + 1
        End If
     End If
     gSql = "INSERT INTO tab_vendas (nsu,tipovenda,codcli,codvend ) VALUES ('"
     gSql = gSql & Format(gnSequencia, "000000000") & "',0,0,0 ) "
     ConDb.Execute gSql
     pRsSequencia.Close
     
     Call Abre_Le_rst_clientes   'Carrega o combo de clientes
     CboClientes.ListIndex = 0
     
     Call Abre_Le_rst_Produtos
     CboProdutos.ListIndex = 0
     
     Call Abre_Le_rst_tipovend
     CboTipovenda.ListIndex = 0
     'CboClientes.SetFocus
  Else
     Call suCarregaOrcamento
     Call sutotal
  End If
     
  Me.LblNumero.Caption = Format(gnSequencia, "#####0")
      
 ' MSFlexGridItens.Cols = 5
 ' MSFlexGridItens.Rows = 1
 ' MSFlexGridItens.Row = 0
 ' MSFlexGridItens.Col = 0
 ' MSFlexGridItens.Text = "Referencia"
 ' MSFlexGridItens.Col = 1
 ' MSFlexGridItens.ColWidth(1) = 4330
 ' MSFlexGridItens.Text = "Descricao                      "
 ' MSFlexGridItens.Col = 2
 ' MSFlexGridItens.Text = "Qtde."
 ' MSFlexGridItens.Col = 3
 ' MSFlexGridItens.Text = "Preço Unit."
 ' MSFlexGridItens.Col = 4
 ' MSFlexGridItens.Text = "Total Item"
      
End Sub


Private Sub Image2_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
   Image2.Appearance = 1
End Sub

'Private Sub MSFlexGridCheques_Click()
'  Dim oldrow As Long
'  Dim lcColGrid As Double
'
'  oldrow = MSFlexGridCheques.Row
'
'  MSFlexGridCheques.Row = 0
'
'  With MSFlexGridCheques
'    .Redraw = False
'    Do While True
'       .Row = .Row + 1
'       For i = 0 To .Cols - 1
'           .Col = i: .CellBackColor = vbWhite
'       Next
'       If .Row = .Rows - 1 Then
'          Exit Do
'       End If
'    Loop
'    .Redraw = True
'
'    .Row = oldrow
'
'    .Col = 0:   .CellBackColor = vbYellow
'    .Col = 1:   .CellBackColor = vbYellow
'    .Col = 2:   .CellBackColor = vbYellow
'    '.Col = 3:   .CellBackColor = vbYellow
'    '.Col = 4:   .CellBackColor = vbYellow
'    '.Col = 5:   TxtUf.Text = .Text: .CellBackColor = vbYellow
'    '.Col = 6:   Txtcep.Text = .Text: .CellBackColor = vbYellow
'    '.Col = 7:   Txtcgc_cpf.Text = .Text: .CellBackColor = vbYellow
'    '.Col = 8:   TxtRG.Text = .Text: .CellBackColor = vbYellow
'
'    .TopRow = .Row
'
'    '.Refresh
'
'End With
'
'End Sub

Private Sub MsflexgridItens_Click()
  Dim oldrow As Long
  Dim lcColGrid As Double
  
  'If MsflexgridItens.Row = 1 Then
  '   lcColGrid = MsflexgridItens.Col
  '   MsflexgridItens.Col = lcColGrid
  '   MsflexgridItens.Sort = flexSortStringAscending
  'End If
 
  If MsflexgridItens.Rows = 1 Then
     Exit Sub
  End If
  
  oldrow = MsflexgridItens.Row
  
  MsflexgridItens.Row = 0
  
  With MsflexgridItens
    .Redraw = False
    Do While True
       .Row = .Row + 1
       For i = 0 To .Cols - 1
           .Col = i: .CellBackColor = vbWhite
       Next
       If .Row = .Rows - 1 Then
          Exit Do
       End If
    Loop
    .Redraw = True
    
    .Row = oldrow
    
    PintaGrid MsflexgridItens
    '.Col = 0:   .CellBackColor = vbYellow
    '.Col = 1:   .CellBackColor = vbYellow
    '.Col = 2:   .CellBackColor = vbYellow
    '.Col = 3:   .CellBackColor = vbYellow
    '.Col = 4:   .CellBackColor = vbYellow
    '.Col = 5:   .CellBackColor = vbYellow
    '.Col = 5:   TxtUf.Text = .Text: .CellBackColor = vbYellow
    '.Col = 6:   Txtcep.Text = .Text: .CellBackColor = vbYellow
    '.Col = 7:   Txtcgc_cpf.Text = .Text: .CellBackColor = vbYellow
    '.Col = 8:   TxtRG.Text = .Text: .CellBackColor = vbYellow
    
    '.TopRow = .Row
    
    '.Refresh
   
End With

End Sub

Private Sub MskDtaPara_GotFocus()
 With MskDtaPara
      .SelStart = 0
      .SelLength = Len(.Text)
   End With

End Sub

Private Sub MskDtaPara_Validate(Cancel As Boolean)
   If Not IsDate(MskDtaPara) Then
      MsgBox "Data Inválida", vbOKOnly, "Atenção " & gOperador
      Cancel = True
   End If
End Sub

Private Sub TxtNumCheque_GotFocus()
 With TxtNumCheque
      .SelStart = 0
      .SelLength = Len(.Text)
   End With
End Sub

Private Sub TxtNumCheque_Validate(Cancel As Boolean)
If Len(TxtNumCheque.Text) = 0 Then
      resposta = MsgBox("Já digitou todos os cheques ?", vbYesNo, "Atenção " & gOperador)
      If resposta - vbYes Then
         Cancel = False
         'Me.Cmdfinaliza.SetFocus
      Else
         Cancel = True
      End If
Else
    Cancel = False
End If
   
End Sub

Private Sub TxtQtde_GotFocus()
    With TxtQtde
      .SelStart = 0
      .SelLength = Len(.Text)
   End With
   If Len(Me.CboProdutos.Text) = 0 Then
      MsgBox "Não escolheu produto", vbOK, "Atencao" & gOperador
      Me.CboProdutos.SetFocus
      Exit Sub
   End If

End Sub

Private Sub suAtualizaOrcamento()
   
  gSql = "UPDATE tab_vendas SET "
  gSql = gSql & " tipovenda = 0, dta_venda  = Cdate('" & Date & "'),"
  gSql = gSql & " codcli = " & Me.CboClientes.ItemData(CboClientes.ListIndex)
  gSql = gSql & " where nsu = '" & Format(Str(gnSequencia), "000000000") & "'"
  ConDb.Execute gSql
     
  '*---> Apaga os itens de venda anteriores
  gSql = "DELETE FROM tab_itemvenda  "
  gSql = gSql & " WHERE nsu = '" & Format(Str(gnSequencia), "000000000") & "'"
  ConDb.Execute gSql
     
  '*---> E grava os atuais
  With MsflexgridItens
     For i = 1 To .Rows - 1
        .Col = 0
        pcCodprod = .Text
        .Col = 2
        pnQtdeP = Val(.Text)
        .Col = 3
        pnQtdeA = Val(.Text)
        .Col = 4
        pnPreco = CDbl(.Text)
        gSql = "select precocusto from tab_produtos where codprod = '" & pcCodprod & "'"
        prsCusto.Open gSql, ConDb
        pnCusto = IIf(IsNull(prsCusto!precocusto), 0, prsCusto!precocusto)
        prsCusto.Close
        '*---> Insere nos Itens de Venda
        gSql = "INSERT INTO tab_itemvenda (nsu,codprod,qtdep,qtdea,precocusto,precounit,valortot,operador,datatual) "
        gSql = gSql & " Values('" & Format(Str(gnSequencia), "000000000") & "','" & Format(pcCodprod, "000000") & "',"
        gSql = gSql & pnQtdeP & "," & pnQtdeA & ","
        gSql = gSql & Replace(pnCusto, ",", ".") & "," & Replace(pnPreco, ",", ".") & ","
        gSql = gSql & Replace((pnQtdeP * pnPreco), ",", ".")
        gSql = gSql & ",'" & gOperador & "',Cdate('" & Date & "'))"
        ConDb.Execute gSql
     Next
  End With

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
   
   gSql = "select tab_itemvenda.codprod,tab_produtos.descricao,qtdep,precounit, qtdep * precounit as totalitem from tab_itemvenda,tab_produtos "
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
        Call suPegaUnidade
        pnQtdeA = Round(Val(TxtQtde.Text) / pnUnidade, 0)
        pnQtdeP = Round(Val(TxtQtde.Text) / pnUnidade, 0) * pnUnidade
        
        MsflexgridItens.Rows = MsflexgridItens.Rows + 1
        MsflexgridItens.Row = MsflexgridItens.Rows - 1
        MsflexgridItens.Col = 0: MsflexgridItens.Text = "" & !codprod
        MsflexgridItens.Col = 1: MsflexgridItens.Text = "" & !descricao
        MsflexgridItens.Col = 2: MsflexgridItens.Text = f_nulo(pnQtdeA, 0)
        MsflexgridItens.Col = 3: MsflexgridItens.Text = f_nulo(pnQtdeP, 0)
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
         CboClientes.AddItem (prsClientes!nome)
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
Private Sub Abre_Le_rst_tipovend()
 
   gSql = "select código,descricao "
   gSql = gSql & "FROM tipovend "
   gSql = gSql & " order by descricao "
   pRstipovenda.Open gSql, ConDb, adOpenKeyset
   Carrega_Combo_Tipovenda
   pRstipovenda.Close

End Sub
Private Sub Carrega_Combo_Tipovenda()

 CboTipovenda.Clear
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

Private Sub sutotal()
   gnTotPed = 0
   For i = 1 To MsflexgridItens.Rows - 1
       MsflexgridItens.Row = i
       MsflexgridItens.Col = 5
       gnTotPed = gnTotPed + CDbl(Me.MsflexgridItens.Text)
   Next
   LbltotaldoPedido.Caption = Format(gnTotPed, "###,###,##0.00")
End Sub

Private Sub TxtSelecionado_GotFocus()
    With TxtQtde
      .SelStart = 0
      .SelLength = Len(.Text)
   End With

End Sub

Private Sub TxtValorCheque_GotFocus()
   With TxtQtde
      .SelStart = 0
      .SelLength = Len(.Text)
   End With

End Sub

Private Sub TxtValorCheque_LostFocus()
   TxtValorCheque.Text = Format(TxtValorCheque.Text, "###,###,##0.00")
   MSFlexGridCheques.AddItem TxtNumCheque.Text & vbTab _
                         & MskDtaPara.Text & vbTab _
                         & Format(TxtValorCheque.Text, "###,###,##0.00")
   TxtNumCheque.SetFocus
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
Private Function fuVeEstoque()
  
  fuVeEstoque = True
  '*---> Verifica se tem estoque
  With MsflexgridItens
     For i = 1 To .Rows - 1
        .Col = 0
        pcCodprod = .Text
        .Col = 2
        pnQtdeP = Val(.Text)
        .Col = 3
        pnQtdeA = Val(.Text)
        .Col = 4
        pnPreco = CDbl(.Text)
        gSql = "select estatual from tab_produtos where codprod = '" & pcCodprod & "'"
        prsEstoque.Open gSql, ConDb
        If pnQtdeP > f_nulo(prsEstoque!estatual, 0) Then
           gSql = "SELECT dtprevista FROM tab_compra,tab_itemcompra "
           gSql = gSql & " WHERE ISNULL(tab_compra.dtentrada) "
           gSql = gSql & " AND tab_compra.numped = tab_itemcompra.numped "
           gSql = gSql & " AND tab_itemcompra.codprod = '" & pcCodprod & "'"
           prsPrevisao.Open gSql, ConDb
           If prsPrevisao.BOF And prsPrevisao.EOF Then
              If MsgBox("Produto com estoque menor e sem data de previsao de entrada. Aceita ?", vbYesNo, "Atenção " & gOperadoe) = vbNo Then
                 fuVeEstoque = False
              End If
           Else
              If MsgBox("Produto com estoque menor e com data de previsao de entrada para " & _
                        Format(prsPrevisao!dtprevista, "dd/mm/yyyy") & ". Aceita ?", vbYesNo, "Atenção " & gOperadoe) = vbNo Then
                 fuVeEstoque = False
              End If
           End If
        End If
        prsPrevisao.Close
        prsEstoque.Close
      Next
  End With
 
End Function
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

Private Sub TxtVlrentrada_LostFocus()
   TxtVlrentrada.Text = Format(TxtVlrentrada, "###,###,##0.00")
End Sub

'Private Sub suImprimeCupom()
'
'  If gDesenv Then
'     Open "TESTE" For Output As #1 'Abre porta imp.
'  Else
'     Open "LPT1" For Output As #1 'Abre porta imp.
'  End If
'  'If pnTotped > 0 Then
'     suImprimeCabeCupom
'  'End If
'  With MsflexgridItens
'     For i = 1 To .Rows - 1
'         Print #1, Left(.TextMatrix(i, 1), 20); _
'                   Tab(22); Spc(f_conta(Format(.TextMatrix(i, 2), "000"))); Format(.TextMatrix(i, 2), "###"); _
'                   Spc(1); Spc(f_conta(Format(.TextMatrix(i, 3), "0000.00"))); Format(.TextMatrix(i, 3), "###0.00"); _
'                   Spc(1); Spc(f_conta(Format(CDbl(.TextMatrix(i, 2)) * CDbl(.TextMatrix(i, 3)), "0000.00"))); _
'                                       Format(CDbl(.TextMatrix(i, 2)) * CDbl(.TextMatrix(i, 3)), "###0.00")
'
'
'     Next
'  End With
'  If gnTotPed > 0 Then
'        Print #1, " "
'        Print #1, "      Total da Compra..........."; Spc(f_conta(Format(gnTotPed, "00000.00"))); Format(gnTotPed, "####0.00")
''           IF mDesc > 0
''                Print "      Desconto "
''                ?? Transform(pDesc,"99.99")
''                ?? "%.........."
''                ?? Transform(mDesc,"999999.99")
''                Print "      Total Geral (R$)........."
''                ?? Transform(m.totorca - mDesc,"999999.99")
''           End If
''            Print
''        If pnParcelas > 0 And ChkPre.Value = 0 Then 'Venda a Prazo Sem cheque pre-datado
''           Print #1, ""
''           Print #1, ""
''           Print #1, ""
''           Print #1, "----------------------------"
''           Print #1, "        ASSINATURA  "
''           If gCupom = "S" Then
''              gSql = "SELECT sum(qtde * preco) as divida from movcli "
''              gSql = gSql & " WHERE codcli = '" & pnCodcli & "'"
''              prsCliente.Open gSql, ConDb, adOpenKeyset
''              If prsCliente.BOF And prsCliente.EOF Then
''                 pnTotdivida = 0
''              Else
''                 pnTotdivida = prsCliente!divida
''              End If
''              If pnTotdivida > 0 Then
''                 Print #1, "Divida Anterior:" + Format(pnTotdivida, "##,##0.00")
''              End If
''             'Print #1, "Divida Atual..:" + TRANS(pnTotDIVIDA + (pntotped - mDesc), "##,##0.00")
''              Print #1, "Divida Atual..:" + Format(pnTotdivida + pnTotped, "##,##0.00")
''           End If
'''                IF !EMPT(cadclie.vencto)
'''                    IF pData - cadclie.vencto > 40
'''                        Print #1, "Vencimento: " + DtoC(cadclie.vencto) + " => CLIENTE EM ATRASO "
'''                    Else
'''                        Print #1, "Vencimento: " + DtoC(cadclie.vencto)
'''                    End If
'''                Else
'''                   Print #1, "Vencimento: " + DtoC(pData + 30)
'''                End If
'''
''        End If
'
'        Print #1, Replicate("-", 40)
'        Print #1, " ESTE CUPOM NAO TEM VALOR FISCAL "
'        Print #1, Replicate("-", 40)
'        If Len(gMensagem1) > 0 Then
'           Print #1, gMensagem1
'           Print #1, Replicate("-", 40)
'        End If
'        If Len(gMensagem2) > 0 Then
'           Print #1, gMensagem2
'           Print #1, Replicate("-", 40)
'        End If
'        Print #1, ""
'        Print #1, ""
'        Print #1, ""
'        Print #1, ""
'        Print #1, ""
'        Print #1, ""
'        Print #1, ""
'        Print #1, ""
'     End If
'     Close #1
''
'End Sub
'
'Private Sub suImprimeCabeCupom()
'
'   Print #1, "Cupom No.: " + Format(gnSequencia, "000000")
'   Print #1, "Data: "; Format(Date, "dd/mm/yyyy"); Spc(8); "Hora: " & Time()
'   Print #1, gNome        ' Plota o nome da empresa no cupom
'   '*? "Terminal: 0     Controle Interno"
'   Print #1, "Controle Interno"
'   Print #1, Replicate("-", 40)
'   If gnAPrazo Then
'      gSql = "select nome,endereco,bairro,cidade,estado,telefone,cgccpf,insc_est,contato "
'      gSql = gSql & "FROM tab_clientes WHERE codcli = " & pnCodcli
'      prsCliente.Open gSql, ConDb, adOpenKeyset
'      Print #1, Trim(gPalavra) + " A PRAZO"
'      Print #1, "Cliente: " & prsCliente!Nome
'      Print #1, Replicate("-", 40)
''*       IF m.pDivCupom = "S"
''*           ? "Saldo Acumulado: "+ TRANSF(cadclie.divida + m.totorca, "99,999.99")
'
''
'   Else
'      Print #1, gPalavra & " A VISTA"
'   End If
'   Print #1, "Atendente: " + Trim(CboBalconista.Text) + " Cod.: " & CboBalconista.ItemData(CboBalconista.ListIndex)
'   Print #1, Replicate("-", 40)
'   Print #1, "Produto              Qtd  V.Unit V.Total"
'
''    Case m.tipovenda = 4
''        Print "VENDA CONVENIO"
''        Print "Empr.: " + Left(CAdemp.nomempre, 32)
''        Print "Conv.: " + CAdconv.CONVENIADO
''        Print "Nome :" + CAdconv.Nome
''        Print Repl("-", 40)
''*
'
'
'End Sub
'
'Private Sub suImprimeCabePedido()
'
'Dim prsLoja As New ADODB.Recordset
'
''Ativar modo condensado => chr(27)&chr(15)
''Desativar => chr(18)
''
''Ativar modo expandido => chr(27) & chr(14)
''Desativar => chr(20)
''
''Ativar negrito >= Chr(27) & Chr(69)
''Desativar => chr(27) & chr(70)
''
''Ativar italico >= Chr(27) & Chr(52)
''Desativar => Chr(27) & chr(53)
''
''Avanço de linha e retorno de carro => chr(10) & chr(13)
'
'
'   gSql = "select nome, endereco,bairro,cidade,estado,cgc,telefone "
'   gSql = gSql & "FROM tab_lojas"
'   prsLoja.Open gSql, ConDb, adOpenKeyset
'   Print #1, Chr(27); Chr(14); Tab(10); Trim(prsLoja!Nome)
'   Print #1, Chr(27); Chr(14); Tab(10); Replicate("=", Len(prsLoja!Nome)); Chr(20)
'   Print #1, Tab((80 - Len("Telefone: " & prsLoja!Telefone)) / 2); "Telefone: " & prsLoja!Telefone
'   Print #1, "Orçamento No.:" & Format(gnSequencia, "000000"); Tab(60); "Data: " & Format(Now, "dd/mm/yyyy")
'   Print #1, "Vendedor:" & gOperador
'   Print #1, Replicate("-", 80)
'   Print #1, ""
'   prsLoja.Close
'
'   gSql = "select nome,endereco,bairro,cidade,estado,cep,telefone,celular,cgccpf,insc_est,contato "
'   gSql = gSql & "FROM tab_clientes WHERE codcli = " & gnCodcli
'   prsLoja.Open gSql, ConDb, adOpenKeyset
'   Print #1, "Cliente: ", prsLoja!Nome
'   Print #1, "Endereço:", f_nulo(prsLoja!endereco, " ")
'   Print #1, "CGC/CPF: " & f_nulo(Format(prsLoja!cgccpf, "##.###.###/####-##"), " "); Tab(50); "Insc. Est.: " & f_nulo(prsLoja!insc_est, " ")
'   Print #1, "Bairro: " & f_nulo(prsLoja!bairro, " "); Tab(50); "CEP: " & f_nulo(prsLoja!cep, " ")
'   Print #1, "Cidade: " & f_nulo(prsLoja!Cidade, " "); Tab(50); "Estado: " & f_nulo(prsLoja!estado, " ")
'   Print #1, "Contato: " & f_nulo(prsLoja!contato, " ")
'   Print #1, "Fone 1: " & f_nulo(prsLoja!Telefone, " "); Tab(50); "Fone 2: " & f_nulo(prsLoja!celular, " ")
'   prsLoja.Close
'   Print #1, ""
'   Print #1, Replicate("-", 80)
'   Print #1, "Codigo Desc.Produto                            Qtd   Qtd.A  Pço.Unit. Total Item"
'   'Print #1, "Codigo Desc.Produto                             Qtde        Pço.Unit. Total Item"
'   Print #1, Replicate("-", 80)
'
'End Sub
'
'Private Sub suImprimePedido()
'Dim pnLinhas, pnTotped
'
'  If gDesenv Then
'     Open App.Path & "\TESTE.txt" For Output As #1 'Abre porta imp.
'  Else
'     Open "LPT1" For Output As #1 'Abre porta imp.
'  End If
'  suImprimeCabePedido
'  pnLinhas = 20
''  SELECT uni_qtd FROM tab_uni,tab_prod ;
''               WHERE LEFT(cprodtemp.codprod,6) = tab_prod.codprod ;
''               AND tab_prod.prd_uni = tab_uni.uni_cod INTO CURSOR cUnidade
''IF RECCOUNT() = 0
''   lnUnidade = 1
''Else
''   lnUnidade = cUnidade.uni_qtd
''End If
'
''lnQtdea = CEILING(thisform.TxtQtde.Value / lnUnidade)
''lnQtdep = CEILING(thisform.TxtQtde.Value / lnUnidade) * lnUnidade
'
'   pnTotped = 0
'    ' Atenção -> a função f_conta, conta os zeros nao significativos da mascara do numero para poder
'    ' ajustar a direita, colocando a quantidade de espaços no lugar dos tais zeros a esquerda
'   With MsflexgridItens
'      For i = 1 To .Rows - 1
'          Print #1, .TextMatrix(i, 0); _
'                    Tab(8); .TextMatrix(i, 1); _
'                    Tab(44); _
'                    Spc(f_conta(Format(.TextMatrix(i, 2), "0000.00"))); _
'                                Format(.TextMatrix(i, 2), "###0.00"); _
'                    Tab(52); _
'                    Spc(f_conta(Format(.TextMatrix(i, 3), "0000.00"))); _
'                                Format(.TextMatrix(i, 3), "###0.00"); _
'                    Tab(60); _
'                    Spc(f_conta(Format(.TextMatrix(i, 4), "00,000.00"))); _
'                                Format(.TextMatrix(i, 4), "###,##0.00"); _
'                    Tab(71); _
'                    Spc(f_conta(Format(CDbl(.TextMatrix(i, 5)), "000,000.00"))); _
'                                Format(CDbl(.TextMatrix(i, 5)), "###,##0.00")
'
'          pnLinhas = pnLinhas + 1
'          If pnLinhas > 31 Then
'             suImprimeCabePedido
'          End If
'          pnTotped = pnTotped + CDbl(.TextMatrix(i, 5))
'
'      Next
'       If pnLinhas > 31 Then
'          suImprimeCabePedido
'       End If
'       Print #1, Replicate("-", 80)
'       Print #1, ""
'       Print #1, "Total de Itens: " & Format(.Rows - 1, "##,##0"); _
'                                      Tab(52); "Total do pedido: "; Tab(71) _
'                                      ; Spc(f_conta(Format(pnTotped, "000,000.00"))); Format(pnTotped, "###,##0.00")
'       For i = pnLinhas To 31
'           Print #1, ""
'       Next
'       Close #1 'Fecha comunicação com imp.
'   End With
'
'End Sub
'
'
'
'
Private Sub TxtCodBarras_GotFocus()
    With TxtCodBarras
      .SelStart = 0
      .SelLength = Len(.Text)
   End With
End Sub

Private Sub TxtCodBarras_LostFocus()
    If Len(TxtCodBarras.Text) = 0 Then
        Call suFinalizaVenda
        Exit Sub
        
    End If
    If IsNumeric(Me.TxtCodBarras) Then
        If f_PesquisaProd() Then
            Call suCarregaGrid
            Exit Sub
        End If
    Else
       If (Asc(Mid(TxtCodBarras.Text, 1, 1)) >= 65 And Asc(Mid(TxtCodBarras.Text, 1, 1)) <= 90) _
       Or (Asc(Mid(TxtCodBarras.Text, 1, 1)) >= 97 And Asc(Mid(TxtCodBarras.Text, 1, 1)) >= 122) Then
       '65 a 90 - 97 a 122
       'Ativa o combobox
       End If
    End If
    
    
End Sub
