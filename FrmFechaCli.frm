VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmFechaCli 
   Caption         =   "Fechamento de Clientes"
   ClientHeight    =   5115
   ClientLeft      =   795
   ClientTop       =   1065
   ClientWidth     =   9480
   LinkTopic       =   "Form1"
   ScaleHeight     =   5115
   ScaleWidth      =   9480
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   3015
      Left            =   240
      TabIndex        =   8
      Top             =   840
      Width           =   9045
      _ExtentX        =   15954
      _ExtentY        =   5318
      _Version        =   393216
      Cols            =   6
      FixedCols       =   0
      SelectionMode   =   1
      FormatString    =   "Pedido |   Dta.Venda  |  Produto                                         |   Qtde |          Preço  |         Total Item"
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
   Begin VB.CommandButton CmdSair 
      Caption         =   "&Sair"
      Height          =   630
      Left            =   8430
      Picture         =   "FrmFechaCli.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Tag             =   "&Update"
      Top             =   4170
      Width           =   735
   End
   Begin VB.CommandButton CmdFinaliza 
      Caption         =   "Finalizar"
      Height          =   630
      Left            =   7605
      Picture         =   "FrmFechaCli.frx":00FA
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4170
      Width           =   735
   End
   Begin VB.ComboBox CboClientes 
      Height          =   315
      Left            =   1065
      TabIndex        =   0
      Top             =   240
      Width           =   4575
   End
   Begin VB.Label LblSelecionado 
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
      ForeColor       =   &H000040C0&
      Height          =   285
      Left            =   1815
      TabIndex        =   7
      Top             =   4575
      Width           =   1755
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Selecionado:"
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
      Height          =   240
      Left            =   225
      TabIndex        =   6
      Top             =   4560
      Width           =   1395
   End
   Begin VB.Label LbltotalPed 
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
      Height          =   240
      Left            =   1815
      TabIndex        =   5
      Top             =   4200
      Width           =   1710
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
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
      Height          =   240
      Left            =   1050
      TabIndex        =   4
      Top             =   4185
      Width           =   615
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Cliente:"
      Height          =   195
      Left            =   210
      TabIndex        =   1
      Top             =   270
      Width           =   525
   End
End
Attribute VB_Name = "FrmFechaCli"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private prsCliente As New ADODB.Recordset
Private prsProduto As New ADODB.Recordset

Private Sub CboClientes_Click()
   Dim pnTotped As Double
   Dim nRow As Integer
   
   gSql = "SELECT distinct A.nsu,A.dta_venda,A.codprod,B.descricao,A.qtde,A.preco "
   gSql = gSql & " FROM movcli A, tab_produtos B WHERE "
   gSql = gSql & " A.codcli = " & CboClientes.ItemData(CboClientes.ListIndex)
   gSql = gSql & " AND A.codprod = B.codprod "
   prsProduto.Open gSql, ConDb, adOpenKeyset
   
   MSFlexGrid1.Rows = prsProduto.RecordCount + 1
   MSFlexGrid1.Cols = 8
   
   pnTotped = 0
   nRow = 1
   Do While Not prsProduto.EOF
      MSFlexGrid1.Row = nRow
      MSFlexGrid1.Col = 0
      MSFlexGrid1.Text = prsProduto!nsu
      MSFlexGrid1.Col = 1
      MSFlexGrid1.Text = Format(prsProduto!dta_venda, "dd/mm/yyyy")
      MSFlexGrid1.Col = 2
      MSFlexGrid1.Text = prsProduto!descricao
      MSFlexGrid1.Col = 3
      MSFlexGrid1.Text = Format(prsProduto!qtde, "##,000")
      MSFlexGrid1.Col = 4
      MSFlexGrid1.Text = Format(prsProduto!preco, "##,###,##0.00")
      MSFlexGrid1.Col = 5
      MSFlexGrid1.Text = Format(prsProduto!qtde * prsProduto!preco, "##,###,##0.00")
      MSFlexGrid1.Col = 6
      MSFlexGrid1.Text = prsProduto!codprod
      pnTotped = pnTotped + (prsProduto!qtde * prsProduto!preco)
      
      prsProduto.MoveNext
      nRow = nRow + 1
   Loop
   LbltotalPed = Format(pnTotped, "##,###,##0.00")
   LblSelecionado = Format(pnTotped, "##,###,##0.00")
   prsProduto.Close
      
End Sub

Private Sub Abre_Le_rst_clientes()
   
   gSql = "select DISTINCT B.codcli,nome "
   gSql = gSql & "FROM tab_clientes A,movcli B WHERE A.negativo = False"
   gSql = gSql & " AND A.codcli = VAL(B.codcli) order by nome "
   prsCliente.Open gSql, ConDb, adOpenKeyset
   Carrega_Combo_clientes
   prsCliente.Close

End Sub
Private Sub Carrega_Combo_clientes()

 CboClientes.Clear
 With prsCliente
      '.MoveLast
      'nItem = .RecordCount
      .MoveFirst
      Do While Not .EOF
         CboClientes.AddItem (prsCliente!nome)
         CboClientes.ItemData(CboClientes.NewIndex) = prsCliente!codcli
        .MoveNext
      Loop
  End With
     
End Sub

Private Sub CmdFinaliza_Click()
   
   For i = 1 To MSFlexGrid1.Rows
       MSFlexGrid1.Col = 1
       If MSFlexGrid1.Text = "1" Then
          gSql = "DELETE FROM movcli WHERE "
          gSql = gSql & "codcli = '" & CboClientes.ItemData(CboClientes.ListIndex) & "'"
          MSFlexGrid1.Col = 2
          gSql = gSql & " AND nsu = " & MSFlexGrid1.Text
          MSFlexGrid1.Col = 7
          gSql = gSql & " AND codprod = " & MSFlexGrid1.Text
          MSFlexGrid1.Col = 4
          gSql = gSql & " AND dta_venda = " & MSFlexGrid1.Text
          ConDb.Execute gSql
       End If
   Next
   
End Sub

Private Sub CmdSair_Click()
    Unload Me
    
End Sub

Private Sub Form_Activate()
   Call Abre_Le_rst_clientes
End Sub

Private Sub MSFlexGrid1_Click()
  
  'If MSFlexGrid1.BackColor = vbYellow Then
  '   MSFlexGrid1.Col = 1
    If MSFlexGrid1.CellBackColor = vbWhite Or MSFlexGrid1.CellBackColor = 0 Then
        For i = 0 To MSFlexGrid1.Cols - 1
          MSFlexGrid1.Col = i
          MSFlexGrid1.CellBackColor = vbYellow
        Next
      Else
        For i = 0 To MSFlexGrid1.Cols - 1
          MSFlexGrid1.Col = i
          MSFlexGrid1.CellBackColor = vbWhite
        Next
     End If
  'End If
  pnTotped = 0
  For i = 1 To MSFlexGrid1.Rows - 1
      MSFlexGrid1.Row = i
      MSFlexGrid1.Col = 1
      'If vaSpread1.Text = "1" Then
      If MSFlexGrid1.CellBackColor = vbYellow Then
         MSFlexGrid1.Col = 5
         pnTotped = pnTotped + Val(MSFlexGrid1.Text)
      End If
  Next
  LblSelecionado = Format(pnTotped, "##,###,##0.00")
End Sub

'End Sub

'Private Sub VaSpread1_Click(ByVal Col As Long, ByVal Row As Long)
  'vaSpread1.Row = vaSpread1.ActiveRow
  ''If vaSpread1.ActiveCol = 1 Then
  ''   vaSpread1.Col = 1
  ''   If vaSpread1.Text = "0" Then
  ''      vaSpread1.Text = "1"
  ''   Else
  ''      vaSpread1.Text = "0"
  ''   End If
  ''End If
  'pnTotped = 0
  'For i = 1 To vaSpread1.MaxRows
  '    vaSpread1.Row = i
  '    vaSpread1.Col = 1
  '    'If vaSpread1.Text = "1" Then
  '    If vaSpread1.IsBlockSelected Then
  '       vaSpread1.Col = 7
  '       pnTotped = pnTotped + Val(vaSpread1.Text)
  '    End If
  'Next
  'LblSelecionado = Format(pnTotped, "##,###,##0.00")
'End Sub
