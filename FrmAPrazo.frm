VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmAPrazo 
   Caption         =   "Fecha Venda a Prazo"
   ClientHeight    =   6975
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9495
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6975
   ScaleWidth      =   9495
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Fratroco 
      Caption         =   "Troco"
      Height          =   3660
      Left            =   330
      TabIndex        =   36
      Top             =   6375
      Visible         =   0   'False
      Width           =   6195
      Begin VB.TextBox TxtRecebido 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   495
         Left            =   3135
         TabIndex        =   41
         Top             =   465
         Width           =   2115
      End
      Begin VB.Label LblVlrTroco 
         Alignment       =   1  'Right Justify
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Left            =   3105
         TabIndex        =   42
         Top             =   1395
         Width           =   2100
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Troco:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C000C0&
         Height          =   375
         Left            =   1935
         TabIndex        =   40
         Top             =   2145
         Width           =   930
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Valor da Venda:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   585
         TabIndex        =   39
         Top             =   1410
         Width           =   2280
      End
      Begin VB.Label Label12 
         Caption         =   "Valor Recebido:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   405
         Left            =   630
         TabIndex        =   38
         Top             =   480
         Width           =   2310
      End
      Begin VB.Label LblTroco 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C000C0&
         Height          =   345
         Left            =   3255
         TabIndex        =   37
         Top             =   2085
         Width           =   1950
      End
   End
   Begin VB.CommandButton CmdOk 
      Caption         =   "OK!"
      Height          =   615
      Left            =   8640
      Picture         =   "FrmAPrazo.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   930
      Width           =   630
   End
   Begin VB.CommandButton Cmdfinaliza 
      Caption         =   "Finaliza"
      Enabled         =   0   'False
      Height          =   585
      Left            =   8640
      Picture         =   "FrmAPrazo.frx":0532
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Finaliza o pedido"
      Top             =   6165
      Visible         =   0   'False
      Width           =   630
   End
   Begin VB.CommandButton CmdSair 
      Caption         =   "Sair"
      Height          =   615
      Left            =   8655
      Picture         =   "FrmAPrazo.frx":0A64
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   150
      Width           =   630
   End
   Begin VB.Frame FraCheques 
      Caption         =   "Cheques"
      Height          =   4395
      Left            =   90
      TabIndex        =   21
      Top             =   1710
      Visible         =   0   'False
      Width           =   9195
      Begin VB.TextBox TxtNumCheque 
         Height          =   300
         Left            =   1260
         TabIndex        =   11
         Top             =   1110
         Width           =   1485
      End
      Begin VB.CommandButton CmdExcluir 
         Appearance      =   0  'Flat
         Height          =   480
         Left            =   8535
         Picture         =   "FrmAPrazo.frx":0B5E
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Excluir Item"
         Top             =   2265
         Width           =   435
      End
      Begin VB.CommandButton CmdAlterar 
         Height          =   480
         Left            =   8535
         Picture         =   "FrmAPrazo.frx":0C60
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Alterar Item "
         Top             =   2790
         Width           =   435
      End
      Begin VB.TextBox TxtValorCheque 
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
         Left            =   6615
         TabIndex        =   13
         Text            =   "0.00"
         Top             =   1110
         Width           =   1650
      End
      Begin VB.TextBox TxtPara 
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
         Left            =   4065
         TabIndex        =   12
         Top             =   1110
         Width           =   1245
      End
      Begin VB.TextBox TxtConta 
         Height          =   300
         Left            =   7245
         TabIndex        =   10
         Top             =   675
         Width           =   1545
      End
      Begin VB.TextBox TxtAgencia 
         Height          =   300
         Left            =   4050
         TabIndex        =   9
         Top             =   675
         Width           =   1920
      End
      Begin VB.TextBox TxtBanco 
         Height          =   300
         Left            =   1245
         TabIndex        =   8
         Top             =   675
         Width           =   1545
      End
      Begin VB.TextBox TxtEmitente 
         Height          =   300
         Left            =   1260
         TabIndex        =   7
         Top             =   255
         Width           =   5115
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGridCheques 
         Height          =   2490
         Left            =   255
         TabIndex        =   14
         Top             =   1680
         Width           =   8160
         _ExtentX        =   14393
         _ExtentY        =   4392
         _Version        =   393216
         Rows            =   1
         Cols            =   7
         FixedCols       =   0
         Enabled         =   -1  'True
         FormatString    =   "Emitente                                | Banco   |  Agencia     | C/Corrente   | No.Cheque    |  Bom para  |>          Valor  "
      End
      Begin VB.Label Label4 
         Caption         =   "No.Cheque:"
         Height          =   270
         Left            =   255
         TabIndex        =   35
         Top             =   1170
         Width           =   825
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Valor:"
         Height          =   195
         Left            =   6000
         TabIndex        =   31
         Top             =   1185
         Width           =   405
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Bom Para:"
         Height          =   195
         Left            =   3120
         TabIndex        =   30
         Top             =   1170
         Width           =   735
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "C/Corrente:"
         Height          =   195
         Left            =   6255
         TabIndex        =   29
         Top             =   705
         Width           =   825
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Agencia:"
         Height          =   195
         Left            =   3270
         TabIndex        =   28
         Top             =   690
         Width           =   630
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Banco:"
         Height          =   195
         Left            =   570
         TabIndex        =   27
         Top             =   705
         Width           =   510
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Emitente:"
         Height          =   195
         Left            =   465
         TabIndex        =   26
         Top             =   300
         Width           =   660
      End
   End
   Begin VB.Frame FraParcelas 
      Caption         =   "Parcelas"
      Height          =   4395
      Left            =   2895
      TabIndex        =   19
      Top             =   1740
      Width           =   3705
      Begin VB.TextBox Text1 
         BackColor       =   &H00C0FFFF&
         Height          =   300
         Left            =   15
         TabIndex        =   32
         Text            =   "Text1"
         Top             =   4080
         Visible         =   0   'False
         Width           =   675
      End
      Begin MSFlexGridLib.MSFlexGrid MSflxParcelas 
         Height          =   3615
         Left            =   345
         TabIndex        =   20
         Top             =   315
         Width           =   2970
         _ExtentX        =   5239
         _ExtentY        =   6376
         _Version        =   393216
         FixedCols       =   0
         FormatString    =   "  Vencto.:     |                      Valor   "
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
   Begin VB.Frame Fraaprazo 
      Height          =   1470
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   8235
      Begin VB.TextBox TxtOrcamento 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
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
         Height          =   300
         Left            =   1395
         TabIndex        =   23
         Text            =   "999999"
         Top             =   225
         Width           =   870
      End
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         Caption         =   "Cheque ?"
         Height          =   285
         Left            =   6900
         TabIndex        =   3
         Top             =   960
         Width           =   1020
      End
      Begin VB.TextBox TxtVlrentrada 
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
         Height          =   285
         Left            =   4290
         TabIndex        =   2
         Top             =   960
         Width           =   1635
      End
      Begin VB.ComboBox CboBalconista 
         Height          =   315
         Left            =   4305
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   225
         Width           =   3690
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Orçamento:"
         Height          =   195
         Left            =   240
         TabIndex        =   22
         Top             =   240
         Width           =   825
      End
      Begin VB.Label LblVlrentrada 
         AutoSize        =   -1  'True
         Caption         =   "Entrada? :"
         Height          =   195
         Left            =   4320
         TabIndex        =   18
         Top             =   705
         Width           =   735
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Balconista"
         Height          =   195
         Left            =   3300
         TabIndex        =   6
         Top             =   270
         Width           =   735
      End
   End
   Begin VB.Label LblTotParcelas 
      AutoSize        =   -1  'True
      Caption         =   "999.99"
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
      Left            =   7485
      TabIndex        =   34
      Top             =   6555
      Width           =   840
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Total das Parcelas / Cheques"
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
      Left            =   3585
      TabIndex        =   33
      Top             =   6540
      Width           =   3525
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Total do Orçamento"
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
      Left            =   4710
      TabIndex        =   25
      Top             =   6225
      Width           =   2385
   End
   Begin VB.Label LblTotOrca 
      AutoSize        =   -1  'True
      Caption         =   "999.99"
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
      Left            =   7500
      TabIndex        =   24
      Top             =   6225
      Width           =   840
   End
End
Attribute VB_Name = "FrmAPrazo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim pRstipovenda As New ADODB.Recordset
Dim prsOperador As New ADODB.Recordset
Dim prsCusto As New ADODB.Recordset
Dim prsCaixa As New ADODB.Recordset
Private pnQtdeA As Double
Private pnQtdeP As Double
Private pnPreco As Double
Private pnCusto As Double
Private pcCodprod As String
Private pcEntrada
Private pnParcelas As Double
Private pnDias As Double
Private pnValorParc As Double
Private pdVencto As Date
Private pcEmitente As String
Private pcBanco As String
Private pcAgencia As String
Private pcConta As String
Private pcNumcheque As String
Private pdPara As Date
Private pnValorcheque As String

Const NovaLinha = ">*"
Private LastCol As Double
Private LastRow As Double
Private pnTotparcelas


Private Sub CmdAlterar_Click()
  With MSFlexGridCheques
     .Col = 0:    Me.TxtEmitente = .Text
     .Col = 1:    Me.TxtBanco = .Text
     .Col = 2:    Me.TxtAgencia = .Text
     .Col = 3:    Me.TxtConta = .Text
     .Col = 4:    Me.TxtNumCheque = .Text
     .Col = 5:    Me.TxtPara = .Text
     .Col = 6:    Me.TxtValorCheque = .Text
     
     'MsflexgridItens.Enabled = True
     If .Rows <= 2 Then
        .Clear
        .Rows = 1
     Else
        .RemoveItem .RowSel
     End If
     Me.TxtEmitente.SetFocus
  End With

End Sub

Private Sub CmdExcluir_Click()
  MSFlexGridCheques.Enabled = True
  If MSFlexGridCheques.Rows <= 2 Then
     'MSFlexGridItens.Clear
     MSFlexGridCheques.Rows = 1
  Else
     MSFlexGridCheques.RemoveItem MSFlexGridCheques.RowSel
  End If
  TxtEmitente.SetFocus

End Sub

Private Sub CmdFinaliza_Click()

   gSql = "select código,descricao, entrada,dias,parcelas "
   gSql = gSql & "FROM tipovend WHERE código = " & CboTipovenda.ItemData(CboTipovenda.ListIndex)
   pRstipovenda.Open gSql, ConDb, adOpenKeyset
   If pRstipovenda.BOF And pRstipovenda.BOF Then
      MsgBox "Erro grave. Não achou o tipo de venda", vbOKOnly, "Atenção " & gOperador
      End
   End If
   
'**************************************************************
'*   Primeiro acertar o Estoque                               *
'**************************************************************
   With FrmVendas.MsflexgridItens
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
         '*--> Atualiza produtos
         gSql = "UPDATE tab_produtos SET estatual = estatual - " & pnQtdeP
         gSql = gSql & ",dtultvenda = " & "Cdate('" & Date & "')"
         gSql = gSql & " Where codprod = '" & pcCodprod & "'"
         ConDb.Execute gSql
         '*---> Insere nos Itens de Venda
         gSql = "INSERT INTO tab_itemvenda (nsu,codprod,qtdep,qtdea,precocusto,precounit,valortot,operador,datatual) "
         gSql = gSql & " Values('" & Format(Str(gnSequencia), "000000000") & "','" & Format(pcCodprod, "000000") & "',"
         gSql = gSql & pnQtdeP & "," & pnQtdeA & ","
         gSql = gSql & Replace(pnCusto, ",", ".") & "," & Replace(pnPreco, ",", ".") & ","
         gSql = gSql & Replace((pnQtdeP * pnPreco), ",", ".")
         gSql = gSql & ",'" & gOperador & "',Cdate('" & Date & "'))"
         ConDb.Execute gSql
         '**************************************************************
         '*---> Insere nas Movimentacoes de Estoque
         '**************************************************************
         gSql = "INSERT INTO tab_Movestoque (tipo,e_s,data,codvend,codprod,qtde,precounit,operador,datatual) "
         gSql = gSql & " Values(01,'S'," & "Cdate('" & Date & "')" & ","
         gSql = gSql & CboBalconista.ItemData(CboBalconista.ListIndex)
         gSql = gSql & ",'" & pcCodprod & "'," & pnQtdeP & ","
         gSql = gSql & Replace(pnPreco, ",", ".")
         gSql = gSql & ",'" & gOperador & "',Cdate('" & Date & "'))"
         ConDb.Execute gSql
        '**************************************************************
        '*---->  Vendas prazo - ver se tem
        '**************************************************************
        If pnParcelas > 0 Then 'Venda a Prazo
           If Check1.Value = 0 Then   'Venda a prazo sem cheque pre-datado
              gSql = "INSERT INTO movcli (nsu,codcli,dta_venda,codprod,qtde,preco,operador,datatual) "
              gSql = gSql & " Values('" & Format(Str(gnSequencia), "000000000") & "','" & FrmVendas.CboClientes.ItemData(FrmVendas.CboClientes.ListIndex) & "',"
              gSql = gSql & "Cdate('" & Now() & "')" & ",'" & Format(pcCodprod, "000000") & "'," & pnQtdeP & "," & Replace(pnPreco, ",", ".")
              gSql = gSql & ",'" & gOperador & "',Cdate('" & Date & "'))"
              ConDb.Execute gSql
           End If
        End If
     Next
  End With

  gSql = "Update tab_vendas set tipovenda = " & CboTipovenda.ItemData(CboTipovenda.ListIndex)
  gSql = gSql & ", codvend =  " & CboBalconista.ItemData(CboBalconista.ListIndex)
    gSql = gSql & ", dta_venda = Cdate('" & Now() & "')"
  gSql = gSql & ", datatual = Cdate('" & Now() & "')"
  gSql = gSql & ", operador = '" & gOperador & "'"
  gSql = gSql & " WHERE nsu = '" & Format(gnSequencia, "000000000") & "'"
  ConDb.Execute gSql

  pnCodcli = 1

  If pnParcelas > 0 Then 'Venda a Prazo
     gnAPrazo = True
     '*--> Atualiza o codcli na variavel
     pnCodcli = FrmVendas.CboClientes.ItemData(FrmVendas.CboClientes.ListIndex)
     '*--> Atualiza a data da venda no Cadaastro de cliente
     gSql = "UPDATE tab_clientes set ultcompra = " & "Cdate('" & Date & "')"
     gSql = gSql & " WHERE codcli = " & pnCodcli
     ConDb.Execute gSql
  End If
  
  '*--> Atualiza o codcli no arquivo de vendas
  gSql = "UPDATE tab_vendas set codCLI = " & pnCodcli
  gSql = gSql & " ,tipovenda = " & CboTipovenda.ItemData(CboTipovenda.ListIndex)
  gSql = gSql & " WHERE nsu = '" & Format(gnSequencia, "000000000") & "'"
  ConDb.Execute gSql
  '**************************************************************
  '*--->  Cheque pre datado
  '**************************************************************
  If Check1.Value = 1 Then   'Venda a prazo com cheque pre-datado
     With MSFlexGridCheques
        'For X = 1 To .Rows - 1
        .Col = 0: pcEmitente = .Text
        .Col = 1: pcBanco = .Text
        .Col = 2: pcAgencia = .Text
        .Col = 3: pcConta = .Text
        .Col = 4: pcNumcheque = .Text
        .Col = 5: pdPara = .Text
        .Col = 6: pnValorcheque = CDbl(.Text)
        gSql = "INSERT INTO chequepr (nomecli, pedido, banco,agencia, "
        gSql = gSql & " conta, numcheque, bompara, valor, operador,datatual) "
        gSql = gSql & " Values('" & pcEmitente & " ','"
        gSql = gSql & Format(gnSequencia, "000000000") & "','"
        gSql = gSql & pcBanco & "','"
        gSql = gSql & pcAgencia & "','" & pcConta & "','"
        gSql = gSql & pcNumcheque & "'," & "Cdate('" & pdPara & "')"
        gSql = gSql & "," & Replace(pnValorcheque, ",", ".") & ","
        gSql = gSql & "'" & gOperador & "', Cdate('" & Date & "')" & ")"
        ConDb.Execute gSql
        'Next
      End With

  End If
  '*--->
  If Val(TxtVlrentrada.Text) > 0 Then  'Tem entrada ou o valor total da venda
     gSql = "SELECT hoje FROM caixa WHERE Hoje = cdate('" & Date & "')"
     prsCaixa.Open gSql, ConDb, adOpenKeyset
     If prsCaixa.BOF And prsCaixa.EOF Then
        gSql = "INSERT INTO caixa (hoje, vvista, troco) "
        gSql = gSql & "VALUES ( Cdate('" & Date & "')" & "," & Replace(CDbl(TxtVlrentrada.Text), ",", ".") & ",0)"
        ConDb.Execute gSql
     Else
        gSql = "UPDATE caixa SET vvista = " & Replace(CDbl(TxtVlrentrada.Text), ",", ".")
        gSql = gSql & " WHERE Hoje = Cdate('" & Date & "')"
        ConDb.Execute gSql
     End If
     prsCaixa.Close
  End If
  '*--->>>>
  If MsgBox("Deseja Imprimir o cupom", vbYesNo, "Atenção " & gOperador) = vbNo Then
  Else
     suImprime
  End If
  Me.Fratroco.Visible = False
  Unload Me
  FrmListaOrc.Show vbModal
'
End Sub

Private Sub CmdOk_Click()
    Me.Height = 7710
    Me.Top = 100
    If pnParcelas > 0 Then
      If Check1.Value = 1 Then
         '--> Vai mostrar o frame de cheques
         FraCheques.Top = 1650
         FraCheques.Visible = True
         FraCheques.Enabled = True
         Me.TxtEmitente = FrmVendas.CboClientes.Text
         'MSFlexGridCheques.Clear
         MSFlexGridCheques.Rows = 1
         Fraaprazo.Enabled = False
         'Me.Height = 7000
         Cmdfinaliza.Visible = True
         Cmdfinaliza.Enabled = True
      Else
         '--> Senão mostra o frame de Fiado
         pnValorParc = (gnTotPed - Val(Me.TxtVlrentrada.Text)) / pnParcelas
         pdVencto = Date
         MSflxParcelas.Rows = pnParcelas + 1
         MSflxParcelas.Row = 0
         For i = 0 To pnParcelas - 1
             MSflxParcelas.Row = MSflxParcelas.Row + 1
             MSflxParcelas.Col = 0
             pdVencto = pdVencto + 30
             MSflxParcelas.Text = pdVencto
             MSflxParcelas.Col = 1
             MSflxParcelas.Text = Format(pnValorParc, "###,###,##0.00")
         Next
         FraParcelas.Top = 1650
         FraParcelas.Visible = True
         FraParcelas.Enabled = True
         Fraaprazo.Enabled = False
         Cmdfinaliza.Visible = True
         Cmdfinaliza.Enabled = True
      End If
      Cmdfinaliza.Visible = True
      Cmdfinaliza.Enabled = True
    End If
    If pnParcelas = 0 Then
       '*---> Colocar aqui um frame para troco
       Fratroco.Visible = True
       Fratroco.Enabled = True
       LblVlrTroco.Caption = Format(gnTotPed, "###,###,##0.00")
       Me.TxtRecebido.SetFocus
    End If
    Cmdfinaliza.Visible = True
    Cmdfinaliza.Enabled = True
    Me.LblTotParcelas = Format(0#, "###,##0.00")
End Sub

Private Sub CmdSair_Click()
   Unload Me
 
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
      SendKeys "{TAB}"
      KeyAscii = 0
   End If
   
End Sub

Private Sub Form_Load()
       
  Centra Me
       
  Me.Height = 2085
  TxtOrcamento = Right(gnSequencia, 6)
  CboTipovenda.ListIndex = -1
  CboBalconista.ListIndex = -1
  Me.FraParcelas.Visible = False
  Me.FraParcelas.Enabled = False
  Me.FraCheques.Visible = False
  Me.FraCheques.Enabled = False
  Me.Fratroco.Visible = False
  Me.Fratroco.Enabled = False
  Me.LblTotOrca.Caption = Format(gnTotPed, "##,##0.00")
  Me.LblTotParcelas.Caption = Format(gnTotPed, "##,##0.00")

  Abre_Le_rst_Balconista
  
  'CboClientes.ListIndex = 0

End Sub

Private Sub Abre_Le_rst_Balconista()
 
   gSql = "select codoperador,nome "
   gSql = gSql & "FROM tab_operador "
   gSql = gSql & " order by nome "
   prsOperador.Open gSql, ConDb, adOpenKeyset
   Carrega_Combo_Balconista
   prsOperador.Close

End Sub
Private Sub Carrega_Combo_Balconista()

 CboBalconista.Clear
 With prsOperador
      '.MoveLast
      'nItem = .RecordCount
      .MoveFirst
      Do While Not .EOF
         CboBalconista.AddItem (prsOperador!nome)
         CboBalconista.ItemData(CboBalconista.NewIndex) = prsOperador!codoperador
        .MoveNext
      Loop
  End With
     
End Sub

Private Sub Label15_Click()

End Sub

Private Sub MSFlexGridCheques_Click()
  Dim oldrow As Long
  Dim lcColGrid As Double
  With MSFlexGridCheques
  
     If .Rows = 1 Then
        Exit Sub
     End If
  
     oldrow = .Row
  
     .Row = 0
  
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
     .Col = 5:   .CellBackColor = vbYellow
    
     .TopRow = .Row
   
  End With

End Sub

Private Sub MSflxParcelas_Click()
   ' Quando clicar uma vez
    ' atribui o valor selecionado
   ' AtribuiValorCelula
End Sub

Private Sub MSflxParcelas_DblClick()
 'editar ao clicar duas vezes
    LastRow = MSflxParcelas.Row
    LastCol = MSflxParcelas.Col
    '
    OcultarControles
    '
    ExibirCelula

End Sub

Private Sub MSflxParcelas_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Editar ao pressionar F2
    If KeyCode = vbKeyF2 Then
        ExibirCelula
    ElseIf KeyCode = vbKeyDelete Then
        ' Excluir linhas selecionadas
        ExcluirLinhas
    End If
End Sub

Private Sub MSflxParcelas_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
    ' Editar ao teclar ENTER
    Case vbKeyReturn
        LastRow = MSflxParcelas.Row
        LastCol = MSflxParcelas.Col
        KeyAscii = 0
        ExibirCelula
    ' Cancelar ao pressionar  ESC
    Case vbKeyEscape
        KeyAscii = 0
        AtribuiValorCelula
    ' Editar ao pressinar qualquer tecla
    Case 32 To 255
        ExibirCelula
        With Text1
            If .Visible Then
                .Text = Chr$(KeyAscii)
                .SelStart = Len(.Text) + 1
            End If
        End With
    End Select

End Sub

Private Sub MSflxParcelas_Scroll()
' Ver se a coluna esta visivel
    ' entao ocultar os controles
    '
    If MSflxParcelas.ColIsVisible(LastCol) = False Then
        OcultarControles
        Exit Sub
    End If
    If MSflxParcelas.RowIsVisible(LastRow) = False Then
        OcultarControles
        Exit Sub
    End If
    ' ver se estava visivel antes de ocultar
    ' e posicionar na mesma celula
    If ControlVisible Then
        ExibirCelula
    End If
End Sub

Private Sub Text1_GotFocus()
    With Text1
        ' Posiciona o cursor no fim do texto
        .SelStart = Len(.Text)
    End With

End Sub
Private Sub ExibirCelula()

  Static OK As Boolean
    '
    ' Se for celula fixa , sair
   
    If MSflxParcelas.Col <= MSflxParcelas.FixedCols - 1 Or MSflxParcelas.Row <= MSflxParcelas.FixedRows - 1 Then
        Exit Sub
    End If
    '
    If OK Then Exit Sub
    OK = True
    '
    OcultarControles
    '
    LastRow = MSflxParcelas.Row
    LastCol = MSflxParcelas.Col
    '
    ' Nova Celula
    With MSflxParcelas
        If .TextMatrix(LastRow, 0) = NovaLinha Then
            .Rows = .Rows + 1
            .TextMatrix(LastRow, 0) = LastRow
            .TextMatrix(.Rows - 1, 0) = NovaLinha
       End If
    End With
    '
    Select Case LastCol
    Case Else
        Text1.Move MSflxParcelas.Left + MSflxParcelas.CellLeft, MSflxParcelas.Top + MSflxParcelas.CellTop, MSflxParcelas.ColWidth(MSflxParcelas.Col), MSflxParcelas.RowHeight(MSflxParcelas.Row)
        'Text1.Move Msflxparcelas.CellLeft - Screen.TwipsPerPixelX, Msflxparcelas.CellTop + 550 - Screen.TwipsPerPixelY, Msflxparcelas.CellWidth + Screen.TwipsPerPixelX * 2, Msflxparcelas.CellHeight + Screen.TwipsPerPixelY * 2
        Text1.Text = MSflxParcelas.Text
        If Len(MSflxParcelas.Text) = 0 Then
            If LastRow > 1 Then
                Text1.Text = MSflxParcelas.TextMatrix(LastRow - 1, LastCol)
            End If
        End If
        Text1.Visible = True
        If Text1.Visible Then
            Text1.ZOrder
            Text1.SetFocus
        End If
    End Select
    '
    ControlVisible = True
    '
    OK = False
End Sub
Private Sub ProximaCelula()
    If MSflxParcelas.Col < MSflxParcelas.Cols - 1 Then
        MSflxParcelas.Col = MSflxParcelas.Col + 1
    Else
        MSflxParcelas.Col = 1
        If MSflxParcelas.Row < MSflxParcelas.Rows - 1 Then
            MSflxParcelas.Row = MSflxParcelas.Row + 1
        End If
    End If
End Sub
Private Sub AtribuiValorCelula()
    Dim texto As String
    '
    OcultarControles
    ControlVisible = False
    '
    ' atribuir o texto anterior a celula
    Select Case LastCol
      Case 4 To 7
        'notas menores que 5 muda cor fonte para vermelho, demais azul
        texto = Text1.Text
        MSflxParcelas.TextMatrix(LastRow, LastCol) = texto
        If Val(MSflxParcelas.Text) < 5 Then
             MSflxParcelas.CellForeColor = vbRed
        Else
             MSflxParcelas.CellForeColor = vbBlue
        End If
      Case Else
        texto = Text1.Text
        MSflxParcelas.TextMatrix(LastRow, LastCol) = texto
        If LastCol = 1 Then
           pntotprazo = 0
           MSflxParcelas.Col = LastCol
           For i = 1 To MSflxParcelas.Rows - 1
               MSflxParcelas.Row = i
               pntotprazo = pntotprazo + Val(MSflxParcelas.Text)
           Next
           LblTotParcelas.Caption = Format(pntotprazo, "###,###,##0.00")
        End If
        
    End Select
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   ' ao pressionar ENTER aceitar a entrada de dados
   If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyTab Then
        KeyAscii = 0
      
        If Text1.Text = "" Then
           AtribuiValorCelula
           Text1.Visible = False
           ControlVisible = False
           Exit Sub
        End If
        If LastCol = 0 Then
           If Not ChkData(Text1.Text) Then
'           If Val(Text1.Text) > 10 Or Val(Text1.Text) < 0 Then
              MsgBox "Data Invalida !", vbInformation, "Atencao"
              Exit Sub
           End If
        End If
       AtribuiValorCelula
       ProximaCelula
    ' ESC, cancela a edição
    ElseIf KeyAscii = vbKeyEscape Then
        KeyAscii = 0
        Text1.Visible = False
        ControlVisible = False
    End If
End Sub

Private Sub ExcluirLinhas()
    ' Excluir linhas selecionadas
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim n As Long
    '
    ' Não excluir se for a ultima linha
    If MSflxParcelas.RowSel = MSflxParcelas.Rows - 1 Then
        Beep
        Exit Sub
    End If
    If MSflxParcelas.Row = MSflxParcelas.Rows - 1 Then
        Beep
        Exit Sub
    End If
    '
    ' Exclui sempre da linha maior par menor
    i = MSflxParcelas.Row
    j = MSflxParcelas.RowSel
    If i < j Then
        k = i
        i = j
        j = k
    End If
    For n = i To j Step -1
        MSflxParcelas.RemoveItem n
    Next
    LastRow = MSflxParcelas.Rows - 1
    LastCol = 1
    MSflxParcelas.Col = LastCol
    MSflxParcelas.Row = LastRow
    MSflxParcelas.RowSel = LastRow
    MSflxParcelas.ColSel = LastCol
End Sub
Private Sub OcultarControles()
    ' Ocultar o controle textbox
    Text1.Visible = False
End Sub
Private Sub GravarDados()
    ' Gravar os dados do grid
    Dim nFic As Long
    Dim r As Long
    Dim c As Long
    Dim strsql As String
    '
    Me.MousePointer = 11
    For r = 1 To MSflxParcelas.Rows - 2
       'r = r + 1
       
       strsql = "UPDATE tbl_poc_liberacao SET "
       strsql = strsql & " nro_libAnoFRO = " & IIf(MSflxParcelas.TextMatrix(r, 2) = "", "NUll", MSflxParcelas.TextMatrix(r, 2)) & ","
       strsql = strsql & " nro_libFRO = " & IIf(MSflxParcelas.TextMatrix(r, 3) = "", "NUll", MSflxParcelas.TextMatrix(r, 3)) & ","
       strsql = strsql & " dta_libFRO =       " & IIf(MSflxParcelas.TextMatrix(r, 4) = "", "Null", "convert(smalldatetime, '" & Format(MSflxParcelas.TextMatrix(r, 4), "dd/mm/yyyy") & "')") & ","
       strsql = strsql & " dta_libProtocolo = " & IIf(MSflxParcelas.TextMatrix(r, 5) = "", "Null", "convert(smalldatetime, '" & Format(MSflxParcelas.TextMatrix(r, 5), "dd/mm/yyyy") & "')")
       strsql = strsql & " WHERE nro_contrato = '" & MSflxParcelas.TextMatrix(r, 0) & "'"
       CnnPoc.Execute strsql
    Next
    Me.MousePointer = 1
End Sub
Private Function BoundedText(ByVal ptr As Object, ByVal txt As String, ByVal max_wid As Single) As String
    'Faz a string se ajustar a largura da celula
    Do While ptr.TextWidth(txt) > max_wid
        txt = Left$(txt, Len(txt) - 1)
    Loop
    BoundedText = txt
End Function

Private Sub TxtRecebido_GotFocus()
   With TxtRecebido
      .SelStart = 0
      .SelLength = Len(.Text)
   End With

End Sub

Private Sub TxtRecebido_LostFocus()
   Me.LblTroco.Caption = Format((f_nulo(Me.TxtRecebido, 0) - gnTotPed), "###,##0.00")
End Sub

Private Sub TxtValorCheque_GotFocus()
   With TxtValorCheque
      .SelStart = 0
      .SelLength = Len(.Text)
   End With

End Sub

Private Sub TxtValorCheque_LostFocus()
    MSFlexGridCheques.AddItem Me.TxtEmitente & vbTab _
                         & TxtBanco & vbTab _
                         & TxtAgencia & vbTab _
                         & TxtConta & vbTab _
                         & TxtNumCheque & vbTab _
                         & TxtPara & vbTab _
                         & Format(TxtValorCheque, "###,##0.00")
    Call sutotal
    TxtEmitente.SetFocus
      
End Sub
Private Sub sutotal()
   With MSFlexGridCheques
      .Row = 1
      .Col = 6
      pnTotparcelas = 0
      For i = 1 To .Rows - 1
          .Row = i
          pnTotparcelas = pnTotparcelas + CDbl(.Text)
      Next
      .Row = 1
   End With
   LblTotParcelas.Caption = Format(pnTotparcelas, "###,##0.00")
End Sub
