VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "ss32x25.ocx"
Begin VB.Form Frmpesq 
   BackColor       =   &H00404000&
   Caption         =   "Pesquisa de Produtos"
   ClientHeight    =   4665
   ClientLeft      =   435
   ClientTop       =   1500
   ClientWidth     =   9840
   LinkTopic       =   "Form1"
   ScaleHeight     =   4665
   ScaleWidth      =   9840
   Begin FPSpread.vaSpread vaSpread1 
      Height          =   4275
      Left            =   30
      TabIndex        =   0
      Top             =   0
      Width           =   9795
      _Version        =   131077
      _ExtentX        =   17277
      _ExtentY        =   7541
      _StockProps     =   64
      ColsFrozen      =   7
      DisplayRowHeaders=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GrayAreaBackColor=   -2147483636
      MaxCols         =   7
      NoBeep          =   -1  'True
      OperationMode   =   2
      RowHeaderDisplay=   2
      ScrollBars      =   2
      ShadowColor     =   12632256
      ShadowDark      =   0
      ShadowText      =   0
      SpreadDesigner  =   "Frmpesq.frx":0000
      VisibleCols     =   3
      VisibleRows     =   500
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H00404000&
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9255
      Picture         =   "Frmpesq.frx":18FC
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Volta com o código escolhido"
      Top             =   4275
      Width           =   585
   End
   Begin VB.Label LblCodigo 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   315
      Left            =   240
      TabIndex        =   2
      Top             =   4320
      Width           =   1635
   End
   Begin VB.Label LblProduto 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   315
      Left            =   2115
      TabIndex        =   1
      Top             =   4320
      Width           =   3930
   End
End
Attribute VB_Name = "Frmpesq"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
   vaSpread1.Col = 1
   If FrmVendas.Visible = True Then
      FrmVendas.Txtreferencia = vaSpread1.Text
     'FrmVendas.TxtReferencia.SetFocus
   Else
      FrmEntradas.Txtreferencia = vaSpread1.Text
  '    FrmEntradas.TxtReferencia.SetFocus
   End If
   'pRsProduto.Close
   Unload Me

End Sub

Private Sub Form_Load()
 
 'Centraliza a tela no video
   Me.Move (Screen.Width - Me.Width) / 2, _
           (Screen.Height - Me.Height) / 2
   
   gSql = "SELECT A.codprod as codigo,A.descricao as descricao,A.prevenda1,"
   gSql = gSql & "A.prevenda2,A.prevenda3,A.prevenda4,A.prevenda5 "
   gSql = gSql & "FROM tab_produtos A ORDER BY A.descricao"
   gRs.Open gSql, ConDb, adOpenKeyset
   'vaSpread1.DataSource = gRs
   vaSpread1.MaxRows = gRs.RecordCount
   For i = 1 To gRs.RecordCount
       vaSpread1.Row = i
       vaSpread1.Col = 1
       vaSpread1.Text = gRs!codigo
       vaSpread1.Col = 2
       vaSpread1.Text = gRs!descricao
       vaSpread1.Col = 3
       vaSpread1.Text = Format(f_nulo(gRs!prevenda1, 0), "###,##0.00")
       vaSpread1.Col = 4
       vaSpread1.Text = Format(f_nulo(gRs!prevenda2, 0), "###,##0.00")
       vaSpread1.Col = 5
       vaSpread1.Text = Format(f_nulo(gRs!prevenda3, 0), "###,##0.00")
       vaSpread1.Col = 6
       vaSpread1.Text = Format(f_nulo(gRs!prevenda4, 0), "###,##0.00")
       vaSpread1.Col = 7
       vaSpread1.Text = Format(f_nulo(gRs!prevenda5, 0), "###,##0.00")
       gRs.MoveNext
   Next
   gRs.Close
   vaSpread1.Row = 1
   vaSpread1.Col = 1
   LblCodigo = vaSpread1.Text
   vaSpread1.Col = 2
   LblProduto = vaSpread1.Text

End Sub

Private Sub vaSpread1_Click(ByVal Col As Long, ByVal Row As Long)
   ' Select a block of cells
   Screen.MousePointer = 11
   If Row = 0 Then
      suSortSpread vaSpread1, Col, 1
      Screen.MousePointer = 1
      Exit Sub
  End If
  Screen.MousePointer = 1
  vaSpread1.Row = vaSpread1.ActiveRow
  vaSpread1.Col = 1
  LblCodigo = vaSpread1.Text
  vaSpread1.Col = 2
  LblProduto = vaSpread1.Text
End Sub

Private Sub vaSpread1_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyUp Then
      vaSpread1.Row = vaSpread1.ActiveRow - 1
      If vaSpread1.Row = 0 Then
         vaSpread1.Row = 1
      End If
      vaSpread1.Col = 1
      LblCodigo = vaSpread1.Text
      vaSpread1.Col = 2
      LblProduto = vaSpread1.Text
   End If
   If KeyCode = vbKeyDown Then
      vaSpread1.Row = vaSpread1.ActiveRow + 1
      If vaSpread1.Row > vaSpread1.MaxRows Then
         vaSpread1.Row = vaSpread1.Row - 1
      End If
      vaSpread1.Col = 1
      LblCodigo = vaSpread1.Text
      vaSpread1.Col = 2
      LblProduto = vaSpread1.Text
   End If

End Sub
