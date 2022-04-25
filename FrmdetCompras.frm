VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FrmDetCompras 
   Caption         =   "Pedido de compra atendido"
   ClientHeight    =   8145
   ClientLeft      =   1080
   ClientTop       =   1365
   ClientWidth     =   10110
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8145
   ScaleWidth      =   10110
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   615
      Left            =   7965
      Picture         =   "FrmdetCompras.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   36
      Tag             =   "&Update"
      ToolTipText     =   "Cancelar Pedido de Compra"
      Top             =   4140
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.Frame FraNotafiscal 
      Caption         =   "Nota Fiscal"
      Enabled         =   0   'False
      Height          =   735
      Left            =   5475
      TabIndex        =   33
      Top             =   30
      Width           =   4515
      Begin VB.TextBox TxtNotafiscal 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1050
         TabIndex        =   1
         Top             =   240
         Width           =   1320
      End
      Begin MSMask.MaskEdBox MskDataNF 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   3
         EndProperty
         Height          =   345
         Left            =   3270
         TabIndex        =   2
         Top             =   240
         Width           =   1110
         _ExtentX        =   1958
         _ExtentY        =   609
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "99/99/9999"
         PromptChar      =   "_"
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "N.Fiscal:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   120
         TabIndex        =   35
         Top             =   270
         Width           =   780
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Data:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   2655
         TabIndex        =   34
         Top             =   270
         Width           =   480
      End
   End
   Begin VB.Frame Frapedido 
      Caption         =   "Pedido"
      Height          =   735
      Left            =   180
      TabIndex        =   29
      Top             =   30
      Width           =   5205
      Begin MSMask.MaskEdBox MskDtPedido 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   3
         EndProperty
         Height          =   360
         Left            =   4005
         TabIndex        =   0
         Top             =   210
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   635
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "99/99/9999"
         PromptChar      =   "_"
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Pedido No."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   300
         Left            =   105
         TabIndex        =   32
         Top             =   225
         Width           =   1320
      End
      Begin VB.Label lblNumero 
         AutoSize        =   -1  'True
         Caption         =   "999999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C000C0&
         Height          =   300
         Left            =   1545
         TabIndex        =   31
         Top             =   225
         Width           =   915
      End
      Begin VB.Label Label14 
         Caption         =   "Dt.Pedido:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   2925
         TabIndex        =   30
         Top             =   240
         Width           =   1050
      End
   End
   Begin VB.Frame FraDuplics 
      Caption         =   "Duplicatas"
      Height          =   2430
      Left            =   840
      TabIndex        =   23
      Top             =   5175
      Width           =   7170
      Begin MSMask.MaskEdBox TxtVencto 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   3
         EndProperty
         Height          =   300
         Left            =   3150
         TabIndex        =   11
         Top             =   240
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "99/99/9999"
         PromptChar      =   "_"
      End
      Begin VB.CommandButton CmdExcDup 
         Appearance      =   0  'Flat
         Height          =   480
         Left            =   6480
         Picture         =   "FrmdetCompras.frx":0532
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Excluir Item"
         Top             =   795
         Width           =   435
      End
      Begin VB.CommandButton CmdAltDup 
         Height          =   480
         Left            =   6480
         Picture         =   "FrmdetCompras.frx":0634
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Alterar Item "
         Top             =   1320
         Width           =   435
      End
      Begin VB.TextBox TxtValorDup 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   1
         EndProperty
         Height          =   300
         Left            =   5175
         TabIndex        =   12
         Text            =   "0.00"
         Top             =   270
         Width           =   1500
      End
      Begin VB.TextBox TxtNumdup 
         Height          =   300
         Left            =   975
         TabIndex        =   10
         Top             =   270
         Width           =   1245
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlxGridDup 
         Height          =   1215
         Left            =   1320
         TabIndex        =   13
         Top             =   720
         Width           =   4545
         _ExtentX        =   8017
         _ExtentY        =   2143
         _Version        =   393216
         Rows            =   1
         Cols            =   3
         FixedCols       =   0
         BackColorSel    =   8454143
         SelectionMode   =   1
         FormatString    =   "Número            |     Vencto      |>              Valor   "
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label LblTotdup 
         Alignment       =   1  'Right Justify
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   270
         Left            =   4620
         TabIndex        =   28
         Top             =   2025
         Width           =   1545
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Total Duplics.:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   240
         Left            =   3225
         TabIndex        =   27
         Top             =   2040
         Width           =   1515
      End
      Begin VB.Label Label5 
         Caption         =   "Vencto:"
         Height          =   240
         Left            =   2445
         TabIndex        =   26
         Top             =   300
         Width           =   735
      End
      Begin VB.Label Label10 
         Caption         =   "Valor"
         Height          =   285
         Left            =   4620
         TabIndex        =   25
         Top             =   330
         Width           =   540
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Numero"
         Height          =   195
         Left            =   270
         TabIndex        =   24
         Top             =   315
         Width           =   555
      End
   End
   Begin VB.CommandButton CmdOk 
      Caption         =   "Gravar"
      Height          =   615
      Left            =   8745
      MaskColor       =   &H00FF0000&
      Picture         =   "FrmdetCompras.frx":07A6
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   5985
      Width           =   705
   End
   Begin VB.ComboBox Cbofornece 
      Enabled         =   0   'False
      Height          =   315
      Left            =   1215
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   855
      Width           =   5175
   End
   Begin VB.CommandButton CmdSair 
      Caption         =   "&Sair"
      Height          =   615
      Left            =   8850
      Picture         =   "FrmdetCompras.frx":0CD8
      Style           =   1  'Graphical
      TabIndex        =   9
      Tag             =   "&Update"
      Top             =   4125
      Width           =   765
   End
   Begin VB.Frame FraItens 
      Caption         =   "Itens"
      ForeColor       =   &H00800000&
      Height          =   2865
      Left            =   135
      TabIndex        =   17
      Top             =   1200
      Width           =   9525
      Begin VB.ComboBox CboProdutos 
         Enabled         =   0   'False
         Height          =   315
         Left            =   930
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   210
         Width           =   4770
      End
      Begin VB.TextBox TxtPrecocusto 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0,000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   1
         EndProperty
         Enabled         =   0   'False
         Height          =   330
         Left            =   8010
         TabIndex        =   6
         Top             =   210
         Width           =   1305
      End
      Begin VB.TextBox TxtQtde 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   1
         EndProperty
         Enabled         =   0   'False
         Height          =   330
         Left            =   6420
         TabIndex        =   5
         Top             =   210
         Width           =   855
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGridItens 
         Height          =   2040
         Left            =   600
         TabIndex        =   7
         Top             =   720
         Width           =   8400
         _ExtentX        =   14817
         _ExtentY        =   3598
         _Version        =   393216
         Cols            =   5
         FixedCols       =   0
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Preço:"
         ForeColor       =   &H00800000&
         Height          =   225
         Left            =   7440
         TabIndex        =   21
         Top             =   270
         Width           =   465
      End
      Begin VB.Label Label8 
         Caption         =   "Qtde:"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   5955
         TabIndex        =   20
         Top             =   270
         Width           =   390
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Produto"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   240
         TabIndex        =   18
         Top             =   255
         Width           =   555
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Fornecedor"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   225
      TabIndex        =   22
      Top             =   855
      Width           =   810
   End
   Begin VB.Label LblTotaldoPedido 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   300
      Left            =   5445
      TabIndex        =   14
      Top             =   4215
      Width           =   2340
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Total:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   300
      Left            =   4425
      TabIndex        =   8
      Top             =   4230
      Width           =   690
   End
End
Attribute VB_Name = "FrmDetCompras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim prsProduto As New ADODB.Recordset
Dim pRsProd    As New ADODB.Recordset
Dim pRsFornec As New ADODB.Recordset
Dim prsEntrada As New ADODB.Recordset
Dim pRsSequencia As New ADODB.Recordset

Private pnCodfor As Double
Private pcNumDup As String
Private pdVencto As Date
Private pnValordup As Double
Private pnTotDup As Double

'Private LastRow As Long               ' Ultima linha em que se editou
'Private LastCol As Long               ' ultima coluna em que se editou
Private pnQtde As Double
Private pcCodprod As String
Private pnPreco As Double
Private pnprecusto As Double
Private pnprevenda1 As Double
Private pnprevenda2 As Double
Private pnprevenda3 As Double
Private pnprevenda4 As Double
Private pnprevenda5 As Double
Private pnPercentual As Double
Private pcNumcheque As String
Private pdDatapara As Date
Private pnValorcheque As Double
Private pncliente As Double
Private pnNsu  As String
Private pnLinhas As Double
Private pnTotitem As Double
Private pnTotped As Double

Private Sub CboFornece_Click()
  Call Abre_Le_rst_Produtos
End Sub

Private Sub CboProdutos_Click()
 gSql = "select precocusto FROM tab_produtos "
   gSql = gSql & "WHERE codprod = '" & Format(Str(CboProdutos.ItemData(CboProdutos.ListIndex)), "000000") & "'"
   gRs.Open gSql, ConDb, adOpenKeyset
   If gRs.BOF And gRs.EOF Then
      MsgBox "Não existem Produtos no cadastro. Favor cadastrar", vbOKOnly, "Atenção"
      Unload Me
   End If
   Me.TxtPrecocusto = Format(f_nulo(gRs!precocusto, 0), "###,##0.00")
   gRs.Close

End Sub

Private Sub CmdAltDup_Click()
  With MSFlxGridDup
     .Col = 0:    Me.TxtNumdup = .Text
     .Col = 1:    Me.TxtVencto = .Text
     .Col = 2:    Me.TxtValorDup = .Text
     
     'MsflexgridItens.Enabled = True
     If .Rows <= 2 Then
        .Clear
        .Rows = 1
     Else
        .RemoveItem .RowSel
     End If
     Me.TxtNumdup.SetFocus
  End With

     
End Sub

Private Sub CmdAlterar_Click()
  With MSFlexGridItens
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
     Me.TxtPrecocusto.Text = .Text
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


Private Sub CmdCancelar_Click()
   suCancelarCompra
   Unload Me
   FrmPCompaten.Show vbModal
End Sub

Private Sub CmdExcDup_Click()
  MSFlxGridDup.Enabled = True
  If MSFlxGridDup.Rows <= 2 Then
     'MSFlexGridItens.Clear
     MSFlxGridDup.Rows = 1
  Else
     MSFlxGridDup.RemoveItem MSFlxGridDup.RowSel
  End If
  TxtNumdup.SetFocus

End Sub

Private Sub CmdExcluir_Click()
  MSFlexGridItens.Enabled = True
  If MSFlexGridItens.Rows <= 2 Then
     MSFlexGridItens.Rows = 1
     'MSFlexGridItens.Clear
  Else
     MSFlexGridItens.RemoveItem MSFlexGridItens.RowSel
  End If
  Call sutotal
  'CboProdutos.SetFocus
End Sub

Private Sub cmdfimprod_Click()

pnCodfor = Me.Cbofornece.ItemData(ListIndex)
'
 If MSFlexGridItens.Row = 0 Then
     MsgBox "Não digitou nenhum produto", vbOKOnly, "Atenção"
     Me.CboProdutos.SetFocus
     Exit Sub
  End If
    
  If TxtNotafiscal = "" Then
     suAtualizaPedido
     Unload Me
     FrmLisPComp.Show vbModal
  Else
     suAtualizaNota
     Me.Height = 10230
     Centra Me
     Me.FraDuplics.Enabled = True
     Me.TxtNumdup.SetFocus
  End If
    
    
End Sub

Private Sub CmdPesquisaprod_Click()
  'FrmVendas.TxtReferencia = f_pesqprod()
  Frmpesq.Show vbModal
  'Frmpesqprod.Show vbModal
  Txtreferencia.SetFocus
End Sub

Private Sub CmdOk_Click()
     
  suAtualizaAPagar
  
  'limpa_tela Me
  'Me.LbltotaldoPedido.Caption = ""
  'MsflexgridItens.Rows = 1
  'MSFlxGridDup.Rows = 1
  Unload Me
  FrmLisPComp.Show vbModal

End Sub

Private Sub CmdSair_Click()
   Unload Me
   FrmPCompaten.Show vbModal
End Sub

Private Sub Form_Activate()
   Me.Height = 5350
   'Call Abre_Le_rst_Produtos
   MSFlexGridItens.Cols = 5
   MSFlexGridItens.Rows = 1
   MSFlexGridItens.Row = 0
   MSFlexGridItens.Col = 0
   MSFlexGridItens.Text = "Codigo"
   MSFlexGridItens.Col = 1
   MSFlexGridItens.ColWidth(1) = 4330
   MSFlexGridItens.Text = "Descricao                      "
   MSFlexGridItens.Col = 2
   MSFlexGridItens.Text = "Qtde."
   MSFlexGridItens.Col = 3
   MSFlexGridItens.Text = "Preço Unit."
   MSFlexGridItens.Col = 4
   MSFlexGridItens.Text = "Total Item"
     
   Call suCarregaPedido
   Call sutotal
     
  Me.lblNumero.Caption = Format(gnSequencia, "#####0")
   
  
End Sub

Private Sub Abre_Le_rst_fornec()
   gSql = "select codfor,nome FROM tab_fornece ORDER BY nome"
   pRsFornec.Open gSql, ConDb, adOpenKeyset
   If pRsFornec.BOF And pRsFornec.EOF Then
      MsgBox "Não existem fornecedores no cadastro. Favor cadastrar", vbOKOnly, "Atenção"
      Unload Me
   End If
   
   Carrega_combo_fornec
   pRsFornec.Close
End Sub

Public Sub Carrega_combo_fornec()

 With pRsFornec
      '.MoveLast
      'nItem = .RecordCount
      .MoveFirst
      Do While Not .EOF
        Cbofornece.AddItem (pRsFornec!nome)
        Cbofornece.ItemData(Cbofornece.NewIndex) = pRsFornec!codfor
        .MoveNext
      Loop
  End With
  Cbofornece.ListIndex = -1
End Sub
Private Sub Abre_Le_rst_Produtos()
   gSql = "select codprod,descricao FROM tab_produtos,sa4_prf "
   gSql = gSql & " WHERE prf_prd = tab_produtos.codprod AND "
   gSql = gSql & " prf_for = '" & Format(Me.Cbofornece.ItemData(Cbofornece.ListIndex), "000000") & "'"
   gSql = gSql & "ORDER BY descricao"
   prsProduto.Open gSql, ConDb, adOpenKeyset
   If prsProduto.BOF And prsProduto.EOF Then
      MsgBox "Não existem Produtos para este fornecedor. Favor cadastrar", vbOKOnly, "Atenção"
      'Unload Me
      Me.Cbofornece.SetFocus
   Else
      Carrega_combo_Produtos
 End If
   
prsProduto.Close

End Sub

Public Sub Carrega_combo_Produtos()

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
  CboProdutos.ListIndex = -1
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
  
End Sub



Private Sub MsflexgridItens_Click()
  Dim oldrow As Long
  Dim lcColGrid As Double
  
  'If MsflexgridItens.Row = 1 Then
  '   lcColGrid = MsflexgridItens.Col
  '   MsflexgridItens.Col = lcColGrid
  '   MsflexgridItens.Sort = flexSortStringAscending
  'End If
  If MSFlexGridItens.Rows = 1 Then
     Exit Sub
  End If
 
  oldrow = MSFlexGridItens.Row
  
  MSFlexGridItens.Row = 0
  
  With MSFlexGridItens
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
     
     .Col = 0:   .CellBackColor = vbYellow
     .Col = 1:   .CellBackColor = vbYellow
     .Col = 2:   .CellBackColor = vbYellow
     .Col = 3:   .CellBackColor = vbYellow
     .Col = 4:   .CellBackColor = vbYellow
     
     .TopRow = .Row
    
  End With

End Sub

Private Sub MskDataNf_gotfocus()
   With MskDataNF
      .SelStart = 0
      .SelLength = Len(.Text)
   End With

End Sub

Private Sub MskDataNf_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Or KeyAscii = 9 Then KeyAscii = 0
End Sub

Private Sub MskDataNf_Validate(Cancel As Boolean)
   If Not IsDate(MskDataNF) Then
      MsgBox "Data Inválida", vbOKOnly, "Atenção " & gOperador
      Cancel = True
   End If

End Sub

Private Sub MskVencto_GotFocus()
   With MskVencto
      .SelStart = 0
      .SelLength = Len(.Text)
   End With

End Sub

Private Sub MskVencto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Or KeyAscii = 9 Then KeyAscii = 0
End Sub

Private Sub MskVencto_ValidationError(InvalidText As String, StartPosition As Integer)
   If Not IsDate(MskVencto) Then
      MsgBox "Data Inválida", vbOKOnly, "Atenção " & gOperador
      Cancel = True
   End If

End Sub

Private Sub TxtDuplicata_GotFocus()
   With TxtDuplicata
      .SelStart = 0
      .SelLength = Len(.Text)
   End With

End Sub

Private Sub TxtDuplicata_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Or KeyAscii = 9 Then KeyAscii = 0
End Sub

Private Sub TxtNotafiscal_GotFocus()
  With TxtNotafiscal
      .SelStart = 0
      .SelLength = Len(.Text)
   End With

End Sub


Private Sub TxtNotafiscal_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Or KeyAscii = 9 Then KeyAscii = 0
End Sub

Private Sub TxtNumdup_Validate(Cancel As Boolean)
   If Len(TxtNumdup) = 0 Then
      MsgBox "Numero da duplicata deve ser preenchido", vbOKOnly, "ATenção " & gOperador
      Cancel = True
   End If
   
End Sub

Private Sub TxtPrecocusto_GotFocus()
  With TxtPrecocusto
      .SelStart = 0
      .SelLength = Len(.Text)
   End With

End Sub

Private Sub TxtPrecocusto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Or KeyAscii = 9 Then KeyAscii = 0
End Sub

Private Sub TxtPrecocusto_LostFocus()
   TxtPrecocusto.Text = Format(TxtPrecocusto.Text, "###,###,##0.000")
   If Len(TxtPrecocusto.Text) = 0 Then
      MsgBox "Favor preencher o preço", vbOKOnly, "Atenção " & gOperador
      TxtPrecocusto.SetFocus
      Exit Sub
   End If
   If CDbl(TxtPrecocusto.Text) = 0 Then
      MsgBox "Preço zerado.", vbOKOnly, "Atenção " & gOperador
      TxtPrecocusto.SetFocus
      Exit Sub
   End If
   
   pnTotitem = CDbl(TxtQtde.Text) * CDbl(TxtPrecocusto.Text)
   MSFlexGridItens.AddItem CboProdutos.ItemData(CboProdutos.ListIndex) & vbTab _
                         & CboProdutos.Text & vbTab _
                         & Format(TxtQtde.Text, "###0") & vbTab & TxtPrecocusto.Text & _
                           vbTab & Format(pnTotitem, "###,##0.000")
   Call sutotal
   'TxtReferencia.Text = ""
   CboProdutos.SetFocus
   TxtNotafiscal.Enabled = False
   Cbofornece.Enabled = False
   Me.MskDataNF.Enabled = False
   
   'TxtPrecocusto.Enabled = False

End Sub

Private Sub TxtQtde_GotFocus()
  With TxtQtde
      .SelStart = 0
      .SelLength = Len(.Text)
   End With

End Sub

Private Sub TxtQtde_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Or KeyAscii = 9 Then KeyAscii = 0
End Sub
Private Sub suAtualizaNota()
  
  With MSFlexGridItens
     For i = 1 To .Rows - 1
        .Col = 0
        pcCodprod = .Text
        .Col = 2
        pnQtde = Val(.Text)
        .Col = 3
        pnPreco = CDbl(.Text)
        '*--> Atualiza Preços de custo e de venda
        gSql = "Select codprod,precocusto,prevenda1,prevenda2,prevenda3,"
        gSql = gSql & "prevenda4,prevenda5 FROM tab_produtos"
        gSql = gSql & " Where codprod = '" & Format(pcCodprod, "000000") & "'"
        prsProduto.Open gSql, ConDb, adOpenKeyset
        pnPercentual = (pnPreco - prsProduto!precocusto) * 100 / pnPreco
        pnprecusto = prsProduto!precocusto
        pnprevenda1 = prsProduto!prevenda1
        pnprevenda2 = prsProduto!prevenda2
        pnprevenda3 = prsProduto!prevenda3
        pnprevenda4 = prsProduto!prevenda4
        pnprevenda5 = prsProduto!prevenda5
        prsProduto.Close
        If pnPercentual > 0 Then
           If MsgBox("Deseja atualizar o preço de " _
                     & Chr(13) & Chr(10) _
                     & MSFlexGridItens.TextMatrix(i, 1) _
                     & Chr(13) & Chr(10) & " de = R$ " & Format(pnprecusto, "##,###,##0.00") _
                     & Chr(13) & Chr(10) & " para = R$ " & Format(pnPreco, "##,###,##0.00") _
                     & " ??? ", vbYesNo, "Atenção " & gOperador) = 6 Then
              pnprecusto = Round(pnprecusto + (pnprecusto * pnPercentual / 100), 2)
              pnprevenda1 = Round(pnprevenda1 + (pnprevenda1 * pnPercentual / 100), 2)
              pnprevenda2 = Round(pnprevenda2 + (pnprevenda2 * pnPercentual / 100), 2)
              pnprevenda3 = Round(pnprevenda3 + (pnprevenda3 * pnPercentual / 100), 2)
              pnprevenda4 = Round(pnprevenda4 + (pnprevenda4 * pnPercentual / 100), 2)
              pnprevenda5 = Round(pnprevenda5 + (pnprevenda5 * pnPercentual / 100), 2)
              gSql = "UPDATE tab_produtos SET precocusto = " & Replace(pnPreco, ",", ".")
              gSql = gSql & ",prevenda1 = " & Replace(pnprevenda1, ",", ".")
              gSql = gSql & ",prevenda2 = " & Replace(pnprevenda2, ",", ".")
              gSql = gSql & ",prevenda3 = " & Replace(pnprevenda3, ",", ".")
              gSql = gSql & ",prevenda4 = " & Replace(pnprevenda4, ",", ".")
              gSql = gSql & ",prevenda5 = " & Replace(pnprevenda5, ",", ".")
              gSql = gSql & " Where codprod = '" & pcCodprod & "'"
              ConDb.Execute gSql
           End If
        End If
        
        '*--> Atualiza produtos
        gSql = "UPDATE tab_produtos SET estatual = estatual + " & pnQtde
        gSql = gSql & ",dtultvenda = " & "Cdate('" & Date & "')"
        gSql = gSql & " Where codprod = '" & pcCodprod & "'"
        ConDb.Execute gSql

        '*---> Insere nas Movimentacoes de Estoque
        gSql = "INSERT INTO tab_Movestoque (tipo,e_s,data,codprod,qtde,precounit,operador,datatual) "
        gSql = gSql & " Values('01','E'," & "Cdate('" & Date & "')" & ","
        gSql = gSql & "'" & pcCodprod & "'," & pnQtde & ","
        gSql = gSql & Replace(pnPreco, ",", ".")
        gSql = gSql & ",'" & gOperador & "',Cdate('" & Date & "'))"
        ConDb.Execute gSql
           
        gSql = "UPDATE tab_compra SET "
        gSql = gSql & " Notafisc = '" & Me.TxtNotafiscal & "',"
        gSql = gSql & " dtentrada  = Cdate('" & Me.MskDataNF & "'),"
        gSql = gSql & " codfor = " & Me.Cbofornece.ItemData(Cbofornece.ListIndex) & ","
        gSql = gSql & " Valor = " & Val(LblTotaldoPedido.Caption)
        gSql = gSql & " where numped = '" & Format(Str(gnSequencia), "000000000") & "'"
        ConDb.Execute gSql
             
        suAtualizaItens
           
     Next
  End With

End Sub

Private Sub suAtualizaPedido()
  
  gSql = "UPDATE tab_compra SET "
  gSql = gSql & " Notafisc = '', dataped  = Cdate('" & Now & "'),"
  gSql = gSql & " codfor = " & Me.Cbofornece.ItemData(Cbofornece.ListIndex) & ","
  gSql = gSql & " Valor = " & Val(LblTotaldoPedido.Caption)
  gSql = gSql & " where numped = '" & Format(Str(gnSequencia), "000000000") & "'"
  ConDb.Execute gSql
     
  suAtualizaItens
  
End Sub
Private Sub suAtualizaItens()
     
  '*---> Apaga os itens de Compra anteriores
  gSql = "DELETE FROM tab_itemcompra  "
  gSql = gSql & " WHERE numped = '" & Format(Str(gnSequencia), "000000000") & "'"
  gSql = gSql & " AND codfor = " & Me.Cbofornece.ItemData(Cbofornece.ListIndex)
  ConDb.Execute gSql
     
  '*---> E grava os atuais
  With MSFlexGridItens
     For i = 1 To .Rows - 1
        .Col = 0
        pcCodprod = .Text
        .Col = 2
        pnQtdeP = Val(.Text)
        .Col = 3
        pnPreco = CDbl(.Text)
        'gSql = "select precocusto from tab_produtos where codprod = '" & pcCodprod & "'"
        'prscusto.Open gSql, ConDb
        'pnCusto = IIf(IsNull(prscusto!precocusto), 0, prscusto!precocusto)
        'prscusto.Close
        '*---> Insere nos Itens de Compra
        gSql = "INSERT INTO tab_itemcompra (numped,notafisc,codprod,qtde,precounit,operador,datatual) "
        gSql = gSql & " Values('" & Format(Str(gnSequencia), "000000000") & "','"
        gSql = gSql & " " & Me.TxtNotafiscal & "','"
        gSql = gSql & Format(pcCodprod, "000000") & "',"
        gSql = gSql & pnQtdeP & ","
        gSql = gSql & Replace(pnPreco, ",", ".") & ",'"
        gSql = gSql & gOperador & "',Cdate('" & Now & "'))"
        ConDb.Execute gSql
     Next
  End With

End Sub
Private Sub suAtualizaAPagar()
Dim pcNumDup As String
Dim pdVencto As Date
Dim pnValor As Double
  With Me.MSFlxGridDup
     For i = 1 To .Rows - 1
        .Col = 0
        pcNumDup = .Text
        .Col = 2
        pdVencto = Val(.Text)
        .Col = 3
        pnValor = CDbl(.Text)
        
        '*---> Insere nas Movimentacoes de Estoque
        gSql = "INSERT INTO tab_Apagar (codfor,duplicata,datamov,vencto,valor,"
        gSql = gSql & " notafiscal, operador, datatual )"
        gSql = gSql & " Values(pnCodfor,pcNumdup, Cdate('" & Date & "')" & ","
        gSql = gSql & " Cdate('" & pdVencto & "'),"
        gSql = gSql & pnValor & ",'" & TxtNotafiscal
        gSql = gSql & "','" & gOperador & "',Cdate('" & Date & "'))"
        ConDb.Execute gSql
           
     Next
  End With

End Sub

Private Sub suCarregaPedido()
   
   gSql = "select * from tab_compra WHERE numped = '" & Format(Str(gnSequencia), "000000000") & "'"
   gRs.Open gSql, ConDb, adOpenKeyset
   If gRs.BOF And gRs.EOF Then
      MsgBox "Houve erro na carga do Pedido. Programa será cancelado ", vbOKOnly, "Atenção, " & gOperador
      Unload Me
   End If
      
   Call Abre_Le_rst_fornec   'Carrega o combo de Fornecedores
   
   For i = 0 To Cbofornece.ListCount - 1
       If gRs!codfor = Cbofornece.ItemData(i) Then
          Cbofornece.ListIndex = i
          Exit For
       End If
   Next
   
   gRs.Close
   
   Call Abre_Le_rst_Produtos
   CboProdutos.ListIndex = 0
   TxtQtde.Text = 1
   
   gSql = "select tab_itemcompra.codprod,tab_produtos.descricao,qtde,precounit, qtde * precounit as totalitem "
   gSql = gSql & " from tab_itemcompra,tab_produtos "
   gSql = gSql & " WHERE numped = '" & Format(Str(gnSequencia), "000000000") & "'"
   'gSql = gSql & " AND tab_itemcompra.codfor = tab_compra.codfor"
   gSql = gSql & " AND tab_itemcompra.codprod = tab_produtos.codprod "
   gRs.Open gSql, ConDb, adOpenKeyset
   If gRs.BOF And gRs.EOF Then
      MsgBox "houve erro na carga dos Itens do Pedido. Programa será cancelado ", vbOKOnly, "Atenção, " & gOperador
      gRs.Close
      CmdSair_Click
      
   End If
      

  MSFlexGridItens.Row = 0
  MSFlexGridItens.FontWidth = 1
  
  With gRs
     '.MoveLast
     'nItem = .RecordCount
     If .RecordCount = 0 Then
        MsgBox "Nao ha itens para o pedido  ", vbOKOnly, "Atenção, " & gOperador
     Else
        .MoveFirst
        MSFlexGridItens.Rows = 1
        i = 0
        Do While Not .EOF
          
          pcCodprod = Format(!codprod, "000000")
          'Call suPegaUnidade
          'pnQtdeP = Round(Val(gRs!qtde) / pnUnidade, 0) * pnUnidade
          pnQtdeP = gRs!qtde
          MSFlexGridItens.Rows = MSFlexGridItens.Rows + 1
          MSFlexGridItens.Row = MSFlexGridItens.Rows - 1
          MSFlexGridItens.Col = 0: MSFlexGridItens.Text = "" & !codprod
          MSFlexGridItens.Col = 1: MSFlexGridItens.Text = "" & !descricao
          MSFlexGridItens.Col = 2: MSFlexGridItens.Text = f_nulo(pnQtdeP, 0)
          MSFlexGridItens.Col = 3: MSFlexGridItens.Text = Format(f_nulo(!precounit, 0), "###,##0.000")
          MSFlexGridItens.Col = 4: MSFlexGridItens.Text = Format(f_nulo(!totalitem, 0), "###,##0.000")
        
          .MoveNext
         
        Loop
        MSFlexGridItens.FixedRows = 1
      End If
  End With

  gRs.Close
  'CboClientes.SetFocus
  
   
End Sub


Private Sub sutotal()
   pnTotped = 0
   If MSFlexGridItens.Rows > 1 Then
   For i = 1 To MSFlexGridItens.Rows - 1
       MSFlexGridItens.Row = i
       MSFlexGridItens.Col = 4
       pnTotped = pnTotped + CDbl(MSFlexGridItens.Text)
   Next
   LblTotaldoPedido.Caption = Format(pnTotped, "###,###,##0.00")
   End If
   
End Sub


Private Sub TxtValorDup_GotFocus()
  With TxtValorDup
      .SelStart = 0
      .SelLength = Len(.Text)
   End With

End Sub

Private Sub TxtVencto_GotFocus()
  With TxtVencto
      .SelStart = 0
      .SelLength = Len(.Text)
   End With

End Sub

Private Sub TxtVencto_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Or KeyAscii = 9 Then KeyAscii = 0
End Sub

Private Sub TxtVencto_Validate(Cancel As Boolean)
   If Not IsDate(TxtVencto) Then
      MsgBox "Data Inválida", vbOKOnly, "Atenção " & gOperador
      Cancel = True
   End If

End Sub
 
Private Sub suCancelarCompra()
    'Primeiro atualizar o estoque de cada produto do pedido
    'e apaga os itens do pedido de compra
    '
    gSql = "select  * FROM tab_itemcompra WHERE NUMPED = '" & Format(gnSequencia, "000000000") & "'"
    gRs.Open gSql, ConDb, adOpenKeyset
    If Not gRs.BOF And Not gRs.EOF Then
       gRs.MoveFirst
       Do While Not gRs.EOF
          gSql = "SELECT * FROM tab_produtos WHERE codprod = '" & gRs!codprod & "'"
          prsProduto.Open gSql, ConDb, adOpenKeyset, adLockOptimistic
          If Not prsProduto.BOF And Not prsProduto.EOF Then
             'prsProduto.EditMode
             prsProduto!estatual = prsProduto!estatual - gRs!qtde
             prsProduto.Update
          End If
          prsProduto.Close
          '**************************************************************
          '*---> Insere nas Movimentacoes de Estoque
          '**************************************************************
          gSql = "INSERT INTO tab_Movestoque (tipo,e_s,data,codvend,codprod,qtde,precounit,operador,datatual) "
          gSql = gSql & " Values(08,'S'," & "Cdate('" & Date & "')" & ","
          gSql = gSql & gnCodOperador
          gSql = gSql & ",'" & gRs!codprod & "'," & gRs!qtde & ","
          gSql = gSql & Replace(gRs!precounit, ",", ".")
          gSql = gSql & ",'" & gOperador & "',Cdate('" & Date & "'))"
          ConDb.Execute gSql
          gRs.Delete adAffectCurrent
          gRs.MoveNext
       Loop
    End If
    gRs.Close
    '
    'Agora vai procurar se tem contas a pagar e apaga
    '
    gSql = "SELECT * FROM tab_compra WHERE NUMPED = '" & Format(gnSequencia, "000000000") & "'"
    gRs.Open gSql, ConDb, adOpenKeyset, adLockOptimistic
    If Not gRs.BOF And Not gRs.EOF Then
       gSql = "DELETE * FROM tab_apagar WHERE "
       gSql = gSql & " codfor = " & gRs!codfor & " AND notafiscal = '" & gRs!notafisc & "'"
       ConDb.Execute gSql
       'e ai apaga o proprio pedido de compra
       gRs.Delete adAffectCurrent
    End If
    gRs.Close
End Sub
