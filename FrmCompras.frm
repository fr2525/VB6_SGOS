VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form FrmCompras 
   Caption         =   "Entrada de Mercadorias - <ESC> Sai"
   ClientHeight    =   7815
   ClientLeft      =   1080
   ClientTop       =   1365
   ClientWidth     =   10110
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7815
   ScaleWidth      =   10110
   Begin VB.Frame FraNotafiscal 
      Caption         =   "Nota Fiscal"
      Enabled         =   0   'False
      Height          =   750
      Left            =   180
      TabIndex        =   37
      Top             =   780
      Width           =   4410
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
         TabIndex        =   2
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
         Left            =   3120
         TabIndex        =   3
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
         TabIndex        =   39
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
         Left            =   2550
         TabIndex        =   38
         Top             =   270
         Width           =   480
      End
   End
   Begin VB.Frame Frapedido 
      Caption         =   "Pedido"
      Height          =   735
      Left            =   180
      TabIndex        =   33
      Top             =   30
      Width           =   8895
      Begin MSMask.MaskEdBox MskDtEntrega 
         Height          =   345
         Left            =   7530
         TabIndex        =   1
         Top             =   225
         Width           =   1140
         _ExtentX        =   2011
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
         Left            =   3945
         TabIndex        =   0
         Top             =   210
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   635
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
      Begin VB.Label Label13 
         Caption         =   "Previsão de Entrega:"
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
         Height          =   285
         Left            =   5430
         TabIndex        =   40
         Top             =   270
         Width           =   1980
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
         TabIndex        =   36
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
         TabIndex        =   35
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
         Left            =   2835
         TabIndex        =   34
         Top             =   240
         Width           =   1050
      End
   End
   Begin VB.Frame FraDuplics 
      Caption         =   "Duplicatas"
      Height          =   2430
      Left            =   840
      TabIndex        =   27
      Top             =   5130
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
         TabIndex        =   15
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
         Picture         =   "FrmCompras.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Excluir Item"
         Top             =   795
         Width           =   435
      End
      Begin VB.CommandButton CmdAltDup 
         Height          =   480
         Left            =   6480
         Picture         =   "FrmCompras.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   18
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
         TabIndex        =   16
         Text            =   "0.00"
         Top             =   270
         Width           =   1500
      End
      Begin VB.TextBox TxtNumdup 
         Height          =   300
         Left            =   975
         TabIndex        =   14
         Top             =   270
         Width           =   1245
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlxGridDup 
         Height          =   1215
         Left            =   1320
         TabIndex        =   17
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
         TabIndex        =   32
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
         TabIndex        =   31
         Top             =   2040
         Width           =   1515
      End
      Begin VB.Label Label5 
         Caption         =   "Vencto:"
         Height          =   240
         Left            =   2445
         TabIndex        =   30
         Top             =   300
         Width           =   735
      End
      Begin VB.Label Label10 
         Caption         =   "Valor"
         Height          =   285
         Left            =   4620
         TabIndex        =   29
         Top             =   330
         Width           =   540
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Numero"
         Height          =   195
         Left            =   270
         TabIndex        =   28
         Top             =   315
         Width           =   555
      End
   End
   Begin VB.CommandButton CmdOk 
      Caption         =   "Gravar"
      Height          =   615
      Left            =   8745
      MaskColor       =   &H00FF0000&
      Picture         =   "FrmCompras.frx":0274
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   5940
      Width           =   705
   End
   Begin VB.ComboBox Cbofornece 
      Height          =   315
      Left            =   4680
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1155
      Width           =   5040
   End
   Begin VB.CommandButton CmdSair 
      Caption         =   "&Sair"
      Height          =   615
      Left            =   9000
      Picture         =   "FrmCompras.frx":07A6
      Style           =   1  'Graphical
      TabIndex        =   13
      Tag             =   "&Update"
      Top             =   4125
      Width           =   675
   End
   Begin VB.CommandButton cmdfimprod 
      Caption         =   "Finalizar"
      Height          =   615
      Left            =   8280
      Picture         =   "FrmCompras.frx":08A0
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Finaliza os produtos"
      Top             =   4125
      Width           =   675
   End
   Begin VB.Frame FraItens 
      Caption         =   "Itens"
      ForeColor       =   &H00800000&
      Height          =   2565
      Left            =   180
      TabIndex        =   21
      Top             =   1530
      Width           =   9525
      Begin VB.CommandButton CmdAlterar 
         Caption         =   "Alterar"
         Height          =   570
         Left            =   8565
         Picture         =   "FrmCompras.frx":09A2
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Alterar Item "
         Top             =   945
         Width           =   615
      End
      Begin VB.CommandButton CmdExcluir 
         Appearance      =   0  'Flat
         Caption         =   "Excluir"
         Height          =   570
         Left            =   8565
         Picture         =   "FrmCompras.frx":0B14
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Excluir Item"
         Top             =   1605
         Width           =   615
      End
      Begin VB.ComboBox CboProdutos 
         Height          =   315
         Left            =   930
         Style           =   2  'Dropdown List
         TabIndex        =   5
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
         Height          =   330
         Left            =   8010
         TabIndex        =   7
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
         Height          =   330
         Left            =   6420
         TabIndex        =   6
         Top             =   210
         Width           =   855
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGridItens 
         Height          =   1830
         Left            =   120
         TabIndex        =   8
         Top             =   630
         Width           =   8205
         _ExtentX        =   14473
         _ExtentY        =   3228
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
         TabIndex        =   25
         Top             =   270
         Width           =   465
      End
      Begin VB.Label Label8 
         Caption         =   "Qtde:"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   5955
         TabIndex        =   24
         Top             =   270
         Width           =   390
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Produto"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   240
         TabIndex        =   22
         Top             =   255
         Width           =   555
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Fornecedor"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   4710
      TabIndex        =   26
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
      Left            =   5715
      TabIndex        =   19
      Top             =   4200
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
      Left            =   4935
      TabIndex        =   12
      Top             =   4215
      Width           =   690
   End
End
Attribute VB_Name = "FrmCompras"
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
  MsflexgridItens.Enabled = True
  If MsflexgridItens.Rows <= 2 Then
     MsflexgridItens.Rows = 1
     'MSFlexGridItens.Clear
  Else
     MsflexgridItens.RemoveItem MsflexgridItens.RowSel
  End If
  Call sutotal
  'CboProdutos.SetFocus
End Sub

Private Sub cmdfimprod_Click()

pnCodfor = Me.Cbofornece.ItemData(ListIndex)
'
 If MsflexgridItens.Row = 0 Then
     MsgBox "Não digitou nenhum produto", vbOKOnly, "Atenção"
     Me.CboProdutos.SetFocus
     Exit Sub
  End If
    
  If TxtNotafiscal = "" Then
     suAtualizaPedido
     If MsgBox("Deseja Imprimir o Pedido de Compra No. " & gnSequencia, vbYesNo, "Atenção " & gOperador) = vbYes Then
        suImprPedComp Val(gnSequencia)
     End If
     Unload Me
     FrmLisPComp.Show vbModal
  Else
     suAtualizaNota
     Me.Height = 7860
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
End Sub

Private Sub Command1_Click()
   suImprPedComp Val(gnSequencia)
End Sub

Private Sub Form_Activate()
   Me.Height = 5350
   'Call Abre_Le_rst_Produtos
   MsflexgridItens.Cols = 5
   MsflexgridItens.Rows = 1
   MsflexgridItens.Row = 0
   MsflexgridItens.Col = 0
   MsflexgridItens.Text = "Codigo"
   MsflexgridItens.Col = 1
   MsflexgridItens.ColWidth(1) = 4330
   MsflexgridItens.Text = "Descricao                      "
   MsflexgridItens.Col = 2
   MsflexgridItens.Text = "Qtde."
   MsflexgridItens.Col = 3
   MsflexgridItens.Text = "Preço Unit."
   MsflexgridItens.Col = 4
   MsflexgridItens.Text = "Total Item"

If gnSequencia = 0 Then
     gSql = "SELECT MAX(VAL(numped)) AS sequencia FROM tab_compra "
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
     gSql = "INSERT INTO tab_compra (numped,notafisc,operador,datatual ) VALUES ('"
     gSql = gSql & Format(gnSequencia, "000000000") & "','','" & gOperador & "',Cdate('" & Now & "'))"
     ConDb.Execute gSql
     pRsSequencia.Close
     
     Call Abre_Le_rst_fornec
     TxtQtde.Text = 1
     MskDtPedido.Text = Format(Date, "dd/mm/yyyy")
     MskDtPedido.SetFocus
    
     'Cbofornece.SetFocus
  Else
     
     Call suCarregaPedido
     Call sutotal
     FraNotafiscal.Enabled = True
     TxtNotafiscal.SetFocus
  End If
     
  Me.LblNumero.Caption = Format(gnSequencia, "#####0")
   
  
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
   'If prsProduto.BOF And prsProduto.EOF Then
   '   MsgBox "Não existem Produtos para este fornecedor. Favor cadastrar", vbOKOnly, "Atenção"
   '   'Unload Me
   '   Me.Cbofornece.SetFocus
   'Else
      Carrega_combo_Produtos
   'End If
   
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
   MsflexgridItens.AddItem CboProdutos.ItemData(CboProdutos.ListIndex) & vbTab _
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
  
  With MsflexgridItens
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
        pnprecusto = IIf(IsNull(prsProduto!precocusto), 0, pRsProd!precocusto)
        pnprevenda1 = IIf(IsNull(prsProduto!prevenda1), 0, prsProduto!prevenda1)
        pnprevenda2 = IIf(IsNull(prsProduto!prevenda2), 0, prsProduto!prevenda2)
        pnprevenda3 = IIf(IsNull(prsProduto!prevenda3), 0, prsProduto!prevenda3)
        pnprevenda4 = IIf(IsNull(prsProduto!prevenda4), 0, prsProduto!prevenda4)
        pnprevenda5 = IIf(IsNull(prsProduto!prevenda5), 0, prsProduto!prevenda5)
        prsProduto.Close
        If pnPercentual > 0 Then
           If MsgBox("Deseja atualizar o preço de " _
                     & Chr(13) & Chr(10) _
                     & MsflexgridItens.TextMatrix(i, 1) _
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
        gSql = gSql & " Values(02,'E'," & "Cdate('" & Date & "')" & ","
        gSql = gSql & "'" & pcCodprod & "'," & pnQtde & ","
        gSql = gSql & Replace(pnPreco, ",", ".")
        gSql = gSql & ",'" & gOperador & "',Cdate('" & Date & "'))"
        ConDb.Execute gSql
           
        gSql = "UPDATE tab_compra SET "
        gSql = gSql & " Notafisc = '" & Me.TxtNotafiscal & "',"
        gSql = gSql & " dtentrada  = Cdate('" & Me.MskDataNF & "'),"
        gSql = gSql & " codfor = " & Me.Cbofornece.ItemData(Cbofornece.ListIndex) & ","
        gSql = gSql & " Valor = " & Val(LbltotaldoPedido.Caption)
        gSql = gSql & " where numped = '" & Format(Str(gnSequencia), "000000000") & "'"
        ConDb.Execute gSql
             
        suAtualizaItens
           
     Next
  End With

End Sub

Private Sub suAtualizaPedido()
  
  gSql = "UPDATE tab_compra SET "
  gSql = gSql & " Notafisc = '', dataped  = Cdate('" & Me.MskDtPedido & "'),"
  gSql = gSql & " dtprevista = cdate('" & Me.MskDtEntrega & "'),"
  gSql = gSql & " codfor = " & Me.Cbofornece.ItemData(Cbofornece.ListIndex) & ","
  gSql = gSql & " Valor = " & Replace(Val(LbltotaldoPedido.Caption), ",", ".")
  gSql = gSql & " where numped = '" & Format(Str(gnSequencia), "000000000") & "'"
  ConDb.Execute gSql
     
  suAtualizaItens
  
End Sub
Private Sub suAtualizaItens()
  If TxtNotafiscal <> "" Then
     '*---> Apaga os itens de Compra anteriores
     gSql = "DELETE FROM tab_itemcompra  "
     gSql = gSql & " WHERE numped = '" & Format(Str(gnSequencia), "000000000") & "'"
     'gSql = gSql & " AND codfor = " & Me.Cbofornece.ItemData(Cbofornece.ListIndex)
     ConDb.Execute gSql
  End If
  '*---> E grava os atuais
  With MsflexgridItens
     
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
        '*---> Verifica se existe o fornecedor para o produto e caso não
        '*--->  exista insere no arquivo de relacionamento (sa4_prf)
        gSql = "select * from sa4_prf "
        gSql = gSql & " WHERE prf_prd = '" & pcCodprod & "' and "
        gSql = gSql & " prf_for = '" & Format(Me.Cbofornece.ItemData(Cbofornece.ListIndex), "000000") & "'"
        prsProduto.Open gSql, ConDb, adOpenKeyset
        If prsProduto.BOF And prsProduto.EOF Then
           gSql = "INSERT INTO sa4_prf (prf_prd, prf_for,operador,datatual) "
           gSql = gSql & " Values('" & Format(pcCodprod, "000000") & "','"
           gSql = gSql & Format(Me.Cbofornece.ItemData(Cbofornece.ListIndex), "000000") & "','"
           
           ConDb.Execute gSql
        End If
     Next
  End With

End Sub

Private Sub suCarregaPedido()
   
   gSql = "select * from tab_compra "
   gSql = gSql & " WHERE numped = '" & Format(Str(gnSequencia), "000000000") & "'"
   gRs.Open gSql, ConDb, adOpenKeyset
   If gRs.BOF And gRs.EOF Then
      MsgBox "Houve erro na carga do Pedido. Programa será cancelado ", vbOKOnly, "Atenção, " & gOperador
      Unload Me
   End If
      
   Me.MskDtPedido.Text = Format(gRs!dataped, "dd/mm/yyyy")
   Me.MskDtEntrega.Text = Format(f_nulo(gRs!dtprevista, "__/__/____"), "dd/mm/yyyy")
      
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
       'pnQtdeP = Round(Val(gRs!qtde) / pnUnidade, 0) * pnUnidade
       pnQtdeP = gRs!qtde
       MsflexgridItens.Rows = MsflexgridItens.Rows + 1
       MsflexgridItens.Row = MsflexgridItens.Rows - 1
       MsflexgridItens.Col = 0: MsflexgridItens.Text = "" & !codprod
       MsflexgridItens.Col = 1: MsflexgridItens.Text = "" & !descricao
       MsflexgridItens.Col = 2: MsflexgridItens.Text = f_nulo(pnQtdeP, 0)
       MsflexgridItens.Col = 3: MsflexgridItens.Text = Format(f_nulo(!precounit, 0), "###,##0.000")
       MsflexgridItens.Col = 4: MsflexgridItens.Text = Format(f_nulo(!totalitem, 0), "###,##0.000")
        
       .MoveNext
         
     Loop
     MsflexgridItens.FixedRows = 1
          
  End With

  gRs.Close
  'CboClientes.SetFocus
   
    
End Sub


Private Sub sutotal()
   pnTotped = 0
   For i = 1 To MsflexgridItens.Rows - 1
       MsflexgridItens.Row = i
       MsflexgridItens.Col = 4
       pnTotped = pnTotped + CDbl(MsflexgridItens.Text)
   Next
   LbltotaldoPedido.Caption = Format(pnTotped, "###,###,##0.00")
End Sub

Private Sub suAtualizaAPagar()
  With MSFlxGridDup
  
  For i = 1 To .Rows - 1
     .Col = 0
     pcNumDup = .Text
     .Col = 1
     pdVencto = .Text
     .Col = 2
     pnValordup = Val(.Text)
     
     gSql = "INSERT INTO tab_apagar (codfor,duplicata,datamov,vencto,valor,"
     gSql = gSql & "notafiscal,operador,datatual) "
     gSql = gSql & " Values('" & pnCodfor & "','"
     gSql = gSql & pcNumDup & "',Cdate('" & Date & "'),Cdate('" & pdVencto & "'),"
     gSql = gSql & Replace(pnValordup, ",", ".") & ",'" & TxtNotafiscal.Text & "'"
     gSql = gSql & ",'" & gOperador & "',Cdate('" & Date & "'))"
     ConDb.Execute gSql
  
  Next
  End With
  
End Sub

Private Sub TxtValorDup_GotFocus()
  With TxtValorDup
      .SelStart = 0
      .SelLength = Len(.Text)
   End With

End Sub

Private Sub TxtValorDup_LostFocus()
   MSFlxGridDup.AddItem Me.TxtNumdup & vbTab _
                      & Me.TxtVencto & vbTab _
                      & Format(Me.TxtValorDup, "###,##0.00")
   
   MSFlxGridDup.Col = 2
   pnTotDup = 0
   For i = 1 To MSFlxGridDup.Rows - 1
       MSFlxGridDup.Row = i
       pnTotDup = pnTotDup + CDbl(MSFlxGridDup.Text)
   Next
   Me.LblTotdup.Caption = Format(pnTotDup, "###,##0.00")
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

