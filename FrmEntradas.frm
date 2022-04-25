VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form FrmEntradas 
   Caption         =   "Entrada de Mercadorias - <ESC> Sai"
   ClientHeight    =   7320
   ClientLeft      =   1080
   ClientTop       =   1365
   ClientWidth     =   10095
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7320
   ScaleWidth      =   10095
   Begin VB.Frame Frame3 
      Caption         =   "Nota Fiscal"
      Enabled         =   0   'False
      Height          =   735
      Left            =   5475
      TabIndex        =   31
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
         TabIndex        =   32
         Text            =   "12"
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
         TabIndex        =   35
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
         TabIndex        =   34
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
         TabIndex        =   33
         Top             =   270
         Width           =   480
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Pedido"
      Height          =   735
      Left            =   180
      TabIndex        =   26
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
         TabIndex        =   27
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
         TabIndex        =   30
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
         TabIndex        =   29
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
         TabIndex        =   28
         Top             =   240
         Width           =   1050
      End
   End
   Begin VB.Frame FraDuplics 
      Caption         =   "Duplicatas"
      Height          =   2430
      Left            =   840
      TabIndex        =   20
      Top             =   4710
      Visible         =   0   'False
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
         TabIndex        =   8
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
         Picture         =   "FrmEntradas.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Excluir Item"
         Top             =   795
         Width           =   435
      End
      Begin VB.CommandButton CmdAltDup 
         Height          =   480
         Left            =   6480
         Picture         =   "FrmEntradas.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   13
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
         TabIndex        =   9
         Text            =   "0.00"
         Top             =   270
         Width           =   1500
      End
      Begin VB.TextBox TxtNumdup 
         Height          =   300
         Left            =   975
         TabIndex        =   7
         Top             =   270
         Width           =   1245
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlxGridDup 
         Height          =   1215
         Left            =   1320
         TabIndex        =   10
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
         TabIndex        =   25
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
         TabIndex        =   24
         Top             =   2040
         Width           =   1515
      End
      Begin VB.Label Label5 
         Caption         =   "Vencto:"
         Height          =   240
         Left            =   2445
         TabIndex        =   23
         Top             =   300
         Width           =   735
      End
      Begin VB.Label Label10 
         Caption         =   "Valor"
         Height          =   285
         Left            =   4620
         TabIndex        =   22
         Top             =   330
         Width           =   540
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Numero"
         Height          =   195
         Left            =   270
         TabIndex        =   21
         Top             =   315
         Width           =   555
      End
   End
   Begin VB.CommandButton CmdOk 
      Caption         =   "Gravar"
      Height          =   615
      Left            =   8745
      MaskColor       =   &H00FF0000&
      Picture         =   "FrmEntradas.frx":0274
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   5745
      Width           =   705
   End
   Begin VB.ComboBox Cbofornece 
      Height          =   315
      Left            =   1215
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   840
      Width           =   5175
   End
   Begin VB.CommandButton CmdSair 
      Caption         =   "&Sair"
      Height          =   615
      Left            =   9045
      Picture         =   "FrmEntradas.frx":07A6
      Style           =   1  'Graphical
      TabIndex        =   5
      Tag             =   "&Update"
      Top             =   4200
      Width           =   675
   End
   Begin VB.CommandButton cmdfimprod 
      Caption         =   "Finalizar"
      Height          =   615
      Left            =   8310
      Picture         =   "FrmEntradas.frx":08A0
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Finaliza os produtos"
      Top             =   4200
      Width           =   675
   End
   Begin VB.Frame Frame1 
      Caption         =   "Itens"
      ForeColor       =   &H00800000&
      Height          =   2910
      Left            =   135
      TabIndex        =   14
      Top             =   1200
      Width           =   9525
      Begin VB.CommandButton CmdAlterar 
         Caption         =   "Alterar"
         Height          =   615
         Left            =   8565
         Picture         =   "FrmEntradas.frx":09A2
         Style           =   1  'Graphical
         TabIndex        =   37
         ToolTipText     =   "Alterar Item "
         Top             =   945
         Width           =   615
      End
      Begin VB.CommandButton CmdExcluir 
         Appearance      =   0  'Flat
         Caption         =   "Excluir"
         Height          =   570
         Left            =   8565
         Picture         =   "FrmEntradas.frx":0B14
         Style           =   1  'Graphical
         TabIndex        =   36
         ToolTipText     =   "Excluir Item"
         Top             =   1605
         Width           =   615
      End
      Begin VB.ComboBox CboProdutos 
         Height          =   315
         Left            =   930
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   210
         Width           =   4770
      End
      Begin VB.TextBox TxtPrecocusto 
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
         Left            =   8010
         TabIndex        =   3
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
         TabIndex        =   2
         Top             =   210
         Width           =   855
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGridItens 
         Height          =   2040
         Left            =   120
         TabIndex        =   38
         Top             =   705
         Width           =   8205
         _ExtentX        =   14473
         _ExtentY        =   3598
         _Version        =   393216
         FixedCols       =   0
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Preço:"
         ForeColor       =   &H00800000&
         Height          =   225
         Left            =   7440
         TabIndex        =   18
         Top             =   270
         Width           =   465
      End
      Begin VB.Label Label8 
         Caption         =   "Qtde:"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   5955
         TabIndex        =   17
         Top             =   270
         Width           =   390
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Produto"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   240
         TabIndex        =   15
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
      TabIndex        =   19
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
      Left            =   5805
      TabIndex        =   11
      Top             =   4275
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
      Left            =   5025
      TabIndex        =   6
      Top             =   4290
      Width           =   690
   End
End
Attribute VB_Name = "FrmEntradas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim prsProduto As New ADODB.Recordset
Dim pRsProd    As New ADODB.Recordset
Dim pRsFornec As New ADODB.Recordset
Dim prsEntrada As New ADODB.Recordset

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
' If Len(MskCNPJ.Text) = 0 Then
'    MsgBox "Não escolheu fornecedor", vbOKOnly, "Atenção"
'    MskCNPJ.SetFocus
'    Exit Sub
'  End If
 
 If MSFlexGridItens.Row = 0 Then
     MsgBox "Não digitou nenhum produto", vbOKOnly, "Atenção"
     Me.CboProdutos.SetFocus
     Exit Sub
  End If
    
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
        '*---> Insere nos Itens de Compra
        gSql = "INSERT INTO tab_itemCompra (codfor,notafisc,item,codprod,qtde,precouni,operador,datatual) "
        gSql = gSql & " Values('" & pnCodfor & "','"
        gSql = gSql & TxtNotafiscal.Text & "','" & Format(i, "000") & "','" & pcCodprod & "',"
        gSql = gSql & pnQtde & "," & Replace(pnPreco, ",", ".")
        gSql = gSql & ",'" & gOperador & "',Cdate('" & Date & "'))"
        ConDb.Execute gSql
        
        '*---> Insere no Arquivo de Entrada
        gSql = "INSERT INTO entrada (codfor,datamov,valor,"
        gSql = gSql & "notafisc,operador,datatual) "
        gSql = gSql & " Values('" & pnCodfor & "'"
        gSql = gSql & ",Cdate('" & MskDataNota.Text & "'),"
        'If Len(MskVencto.Text) = 0 Then
        '   gSql = gSql & "NULL,"
        'Else
        '   gSql = gSql & "Cdate('" & MskVencto.Text & "'),"
        'End If
        gSql = gSql & Replace(pnTotped, ",", ".") & ",'" & TxtNotafiscal.Text & "'"
        gSql = gSql & ",'" & gOperador & "',Cdate('" & Date & "'))"
        ConDb.Execute gSql

        '*---> Insere nas Movimentacoes de Estoque
        gSql = "INSERT INTO tab_Movestoque (tipo,e_s,data,codprod,qtde,precounit,operador,datatual) "
        gSql = gSql & " Values('01','E'," & "Cdate('" & Date & "')" & ","
        gSql = gSql & "'" & pcCodprod & "'," & pnQtde & ","
        gSql = gSql & Replace(pnPreco, ",", ".")
        gSql = gSql & ",'" & gOperador & "',Cdate('" & Date & "'))"
        ConDb.Execute gSql
           
     Next
  End With
  
  Me.Height = 7560
  Me.TxtNumdup.SetFocus
  
End Sub

Private Sub CmdPesquisaprod_Click()
  'FrmVendas.TxtReferencia = f_pesqprod()
  Frmpesq.Show vbModal
  'Frmpesqprod.Show vbModal
  Txtreferencia.SetFocus
End Sub

Private Sub CmdOk_Click()
     
  suAtualizaAPagar
  
  limpa_tela Me
  
  'Me.LblDescricao.Caption = ""
  'Me.LblNomefor.Caption = ""
  Me.LblTotaldoPedido.Caption = ""
  MSFlexGridItens.Rows = 1
  MSFlxGridDup.Rows = 1
 '  MskDataNota.Text = ""
'  MskVencto.Text = ""
'  MskCNPJ.Text = ""
'  MskCNPJ.SetFocus
 Me.Height = 5100
 Me.Cbofornece.SetFocus
 
End Sub

Private Sub CmdSair_Click()
   Unload Me
End Sub

Private Sub Form_Activate()
   Me.Height = 5070
   Call Abre_Le_rst_fornec
   'Call Abre_Le_rst_Produtos
End Sub

Private Sub Abre_Le_rst_fornec()
   gSql = "select codfor,nome FROM tab_fornece ORDER BY nome"
   gRs.Open gSql, ConDb, adOpenKeyset
   If gRs.BOF And gRs.EOF Then
      MsgBox "Não existem fornecedores no cadastro. Favor cadastrar", vbOKOnly, "Atenção"
      Unload Me
   End If
   
   Carrega_combo_fornec
   gRs.Close
End Sub

Public Sub Carrega_combo_fornec()

 With gRs
      '.MoveLast
      'nItem = .RecordCount
      .MoveFirst
      Do While Not .EOF
        Cbofornece.AddItem (gRs!Nome)
        Cbofornece.ItemData(Cbofornece.NewIndex) = gRs!codfor
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
   gRs.Open gSql, ConDb, adOpenKeyset
   If gRs.BOF And gRs.EOF Then
      MsgBox "Não existem Produtos no cadastro. Favor cadastrar", vbOKOnly, "Atenção"
      Unload Me
   End If
   
   Carrega_combo_Produtos
   gRs.Close
End Sub

Public Sub Carrega_combo_Produtos()

 With gRs
      '.MoveLast
      'nItem = .RecordCount
      .MoveFirst
      Do While Not .EOF
        CboProdutos.AddItem (gRs!descricao)
        CboProdutos.ItemData(CboProdutos.NewIndex) = gRs!codprod
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
  
  Call centra(Me)
  
  'Me.Top = 1150
  'Me.Height = 4140
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
     
     TxtQtde.Text = 1
     
     'CboClientes.SetFocus
  Else
     Call suCarregaPedido
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
   
   
   
   
   
   MSFlexGridItens.Cols = 5
   MSFlexGridItens.Rows = 1
   MSFlexGridItens.Row = 0
   MSFlexGridItens.Col = 0
   MSFlexGridItens.Text = "Referencia"
   MSFlexGridItens.Col = 1
   MSFlexGridItens.ColWidth(1) = 4330
   MSFlexGridItens.Text = "Descricao                      "
   MSFlexGridItens.Col = 2
   MSFlexGridItens.Text = "Qtde."
   MSFlexGridItens.Col = 3
   MSFlexGridItens.Text = "Preço Unit."
   MSFlexGridItens.Col = 4
   MSFlexGridItens.Text = "Total Item"

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

Private Sub TxtNotafiscal_Validate(Cancel As Boolean)
   If Len(TxtNotafiscal.Text) = 0 Then
      MsgBox "Favor digitar o numero da Nota fiscal", vbOKOnly, "Atenção " & gOperador
      Cancel = True
   End If
   
   gSql = "SELECT * FROM ENTRADA  WHERE "
   gSql = gSql & " notafisc = '" & TxtNotafiscal.Text & "' AND "
   gSql = gSql & " codfor = '" & pnCodfor & "'"
   prsEntrada.Open gSql, ConDb, adOpenKeyset
   If Not prsEntrada.BOF And Not prsEntrada.EOF Then
      MsgBox "Nota fiscal já digitada", vbOKOnly
      Cancel = True
   End If
   prsEntrada.Close
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
   TxtPrecocusto.Text = Format(TxtPrecocusto.Text, "###,###,##0.00")
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
                           vbTab & Format(pnTotitem, "###,##0.00")
   Call sutotal
   'TxtReferencia.Text = ""
   CboProdutos.SetFocus
   TxtNotafiscal.Enabled = False
   Cbofornece.Enabled = False
   Me.MskDataNota.Enabled = False
   
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



Private Sub sutotal()
   pnTotped = 0
   For i = 1 To MSFlexGridItens.Rows - 1
       MSFlexGridItens.Row = i
       MSFlexGridItens.Col = 4
       pnTotped = pnTotped + CDbl(MSFlexGridItens.Text)
   Next
   LblTotaldoPedido.Caption = Format(pnTotped, "###,###,##0.00")
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
     
     gSql = "INSERT INTO tb_apagar (codfor,duplicata,datamov,vencto,valor,"
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
