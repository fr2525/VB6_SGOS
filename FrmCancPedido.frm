VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmCancPedido 
   Caption         =   "Cancela Vendas por Pedido - <ESC> Sai"
   ClientHeight    =   5445
   ClientLeft      =   1065
   ClientTop       =   1350
   ClientWidth     =   8490
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5445
   ScaleWidth      =   8490
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   2295
      Left            =   720
      TabIndex        =   12
      Top             =   1560
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   4048
      _Version        =   393216
      Rows            =   5
      Cols            =   5
      FixedCols       =   0
      ScrollBars      =   2
      FormatString    =   $"FrmCancPedido.frx":0000
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Sair"
      Height          =   690
      Left            =   7335
      Picture         =   "FrmCancPedido.frx":0089
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Sai sem Cancelar o Pedido"
      Top             =   4590
      Width           =   765
   End
   Begin VB.CommandButton cmdconfirma 
      Caption         =   "Confirmar"
      Height          =   690
      Left            =   6300
      Picture         =   "FrmCancPedido.frx":0183
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Cancela o Pedido"
      Top             =   4560
      Width           =   765
   End
   Begin VB.TextBox TxtBalconista 
      Enabled         =   0   'False
      Height          =   285
      Left            =   3930
      TabIndex        =   7
      Top             =   585
      Width           =   4365
   End
   Begin VB.TextBox TxtDtaVenda 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1530
      TabIndex        =   5
      Top             =   615
      Width           =   1080
   End
   Begin VB.TextBox TxtCliente 
      Enabled         =   0   'False
      Height          =   285
      Left            =   3930
      TabIndex        =   4
      Top             =   165
      Width           =   4365
   End
   Begin VB.TextBox TxtPedido 
      Height          =   285
      Left            =   1260
      MaxLength       =   9
      TabIndex        =   1
      Top             =   165
      Width           =   1365
   End
   Begin VB.Label LblTotal 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   345
      Left            =   5910
      TabIndex        =   9
      Top             =   3960
      Width           =   2175
   End
   Begin VB.Label Label5 
      Caption         =   "Total:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   345
      Left            =   5115
      TabIndex        =   8
      Top             =   3960
      Width           =   795
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Balconista"
      Height          =   195
      Left            =   2895
      TabIndex        =   6
      Top             =   615
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "Data da Venda"
      Height          =   285
      Left            =   195
      TabIndex        =   3
      Top             =   615
      Width           =   1665
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Cliente"
      Height          =   195
      Left            =   3180
      TabIndex        =   2
      Top             =   225
      Width           =   480
   End
   Begin VB.Label Label1 
      Caption         =   "No.Pedido:"
      Height          =   315
      Left            =   210
      TabIndex        =   0
      Top             =   225
      Width           =   795
   End
End
Attribute VB_Name = "FrmCancPedido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdconfirma_Click()
   Dim lnResposta
   lnResposta = MsgBox("Deseja mesmo cancelar o pedido ? ", vbYesNo, "Atenção " & gOperador)
   If lnResposta = 6 Then
      suCancela
   End If
   limpa_tela Me
   VaSpread1.MaxCols = 5
   VaSpread1.MaxRows = 0
   TxtPedido.SetFocus
End Sub

Private Sub Command1_Click()
  Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
   If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub TxtBalconista_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Or KeyAscii = 9 Then KeyAscii = 0
End Sub

Private Sub TxtCliente_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Or KeyAscii = 9 Then KeyAscii = 0
End Sub

Private Sub TxtDtaVenda_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Or KeyAscii = 9 Then KeyAscii = 0
End Sub

Private Sub TxtPedido_GotFocus()
   With TxtPedido
      .SelStart = 0
      .SelLength = Len(.Text)
   End With

End Sub

Private Sub TxtPedido_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Or KeyAscii = 9 Then KeyAscii = 0
End Sub

Private Sub TxtPedido_LostFocus()
  If Len(TxtPedido.Text) = 0 Then
     Unload Me
     Exit Sub
  End If
  
  TxtPedido.Text = Format(TxtPedido.Text, "000000000")
  gSql = "Select nsu,A.codcli,dta_venda,A.codvend,B.nome as nomecli,C.nome as nomeope FROM "
  gSql = gSql & " tab_vendas A,tab_clientes B,tab_operador C "
  gSql = gSql & " WHERE A.nsu = '" & TxtPedido.Text & "'"
  gSql = gSql & " AND A.tipovenda > 0 "
  gSql = gSql & " AND A.codcli = B.codcli and (Val(A.codvend) = C.codoperador "
  gSql = gSql & " OR '" & gOperador & "' = 'Master')"
  gRs.Open gSql, ConDb, adOpenKeyset
  If gRs.BOF And gRs.EOF Then
     MsgBox "Pedido não encontrado", vbOKOnly, "Atenção " & gOperador
     TxtPedido.SetFocus
     gRs.Close
     Exit Sub
  End If
  TxtCliente.Text = gRs!nomecli
  TxtBalconista.Text = gRs!nomeope
  TxtDtaVenda.Text = Format(gRs!dta_venda, "dd/mm/yyyy")
  gRs.Close
  
  gSql = "Select nsu,A.codprod,qtdeP,precounit,valortot,B.descricao FROM "
  gSql = gSql & " tab_itemvenda A,tab_produtos B "
  gSql = gSql & " WHERE A.nsu = '" & Format(TxtPedido.Text, "000000000") & "'"
  gSql = gSql & " AND A.codprod = B.codprod "
  gRs.Open gSql, ConDb, adOpenKeyset
  If gRs.BOF And gRs.EOF Then
     MsgBox "Itens do pedido não encontrados", vbOKOnly, "Atenção " & gOperador
     TxtPedido.SetFocus
     gRs.Close
     Exit Sub
  End If
  
  MSFlexGrid1.Rows = gRs.RecordCount
  pnTotped = 0
  For i = 1 To gRs.RecordCount - 1
      With MSFlexGrid1
      .Row = i
      .Col = 0
      .Text = gRs!descricao
      .Col = 1
      .Text = gRs!QtdeP
      .Col = 2
      .Text = Format(gRs!precounit, "##,###,#0.00")
      .Col = 3
      .Text = Format(gRs!precounit * gRs!QtdeP, "##,###,#0.00")
      pnTotped = pnTotped + (gRs!precounit * gRs!QtdeP)
      .Col = 4 ' coluna escondida - Codigo do produto
      .Text = gRs!codprod
      End With
      gRs.MoveNext
  Next
  gRs.Close
  MSFlexGrid1.Row = 1
  MSFlexGrid1.Col = 0
  LblCodigo = VaSpread1.Text
  MSFlexGrid1.Col = 1
  lblproduto = MSFlexGrid1.Text
  LblTotal = Format(pnTotped, "###,###,##0.00")
End Sub

Private Sub suCancela()
  Dim resposta
  Dim lcCodprod As String
  Dim lnQtde As Double
  gSql = "Delete FROM "
  gSql = gSql & " tab_vendas   "
  gSql = gSql & " WHERE nsu = '" & TxtPedido.Text & "'"
  ConDb.Execute gSql
  
  resposta = MsgBox("Volta produtos para o Estoque ?", vbYesNo, "Atenção " & gOperador)
  If resposta = 6 Then
     For i = 1 To VaSpread1.MaxRows
        VaSpread1.Col = 5
        lcCodprod = VaSpread1.Text
        VaSpread1.Col = 2
        lnQtde = VaSpread1.Text
        gSql = "UPDATE tab_produtos set estatual = estatual + " & lnQtde
        gSql = gSql & " WHERE codprod =  '" & lcCodprod & "'"
        ConDb.Execute gSql
     Next
  End If
  gSql = "Delete FROM "
  gSql = gSql & " tab_itemvenda "
  gSql = gSql & " WHERE nsu = '" & TxtPedido.Text & "'"
  ConDb.Execute gSql

End Sub

Private Sub VaSpread1_Click()

End Sub
