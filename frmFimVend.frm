VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form Frmfimvend 
   Caption         =   "Form1"
   ClientHeight    =   4320
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10125
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   4320
   ScaleWidth      =   10125
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Fratipovenda 
      Height          =   855
      Left            =   165
      TabIndex        =   23
      Top             =   105
      Visible         =   0   'False
      Width           =   8250
      Begin VB.ComboBox CboTipovenda 
         Height          =   315
         Left            =   3780
         TabIndex        =   25
         Top             =   360
         Width           =   4170
      End
      Begin VB.ComboBox CboBalconista 
         Height          =   315
         Left            =   165
         TabIndex        =   24
         Top             =   360
         Width           =   3300
      End
      Begin VB.Label Lbltipovenda 
         Caption         =   "Tipo de venda:"
         Height          =   285
         Left            =   3780
         TabIndex        =   27
         Top             =   150
         Width           =   1170
      End
      Begin VB.Label Lblbalconista 
         Caption         =   "Balconista"
         Height          =   210
         Left            =   180
         TabIndex        =   26
         Top             =   150
         Width           =   1380
      End
   End
   Begin VB.Frame Fraaprazo 
      Enabled         =   0   'False
      Height          =   2565
      Left            =   180
      TabIndex        =   3
      Top             =   1035
      Visible         =   0   'False
      Width           =   8235
      Begin VB.TextBox TxtAgencia 
         Height          =   285
         Left            =   2790
         TabIndex        =   17
         Top             =   645
         Width           =   1665
      End
      Begin VB.TextBox TxtBanco 
         Height          =   285
         Left            =   735
         TabIndex        =   16
         Top             =   645
         Width           =   1230
      End
      Begin VB.ComboBox CboClientes 
         Height          =   315
         Left            =   735
         TabIndex        =   15
         Top             =   255
         Width           =   3735
      End
      Begin VB.TextBox TxtVlrentrada 
         Enabled         =   0   'False
         Height          =   285
         Left            =   6645
         TabIndex        =   14
         Top             =   255
         Width           =   1320
      End
      Begin VB.CheckBox ChkPre 
         Alignment       =   1  'Right Justify
         Caption         =   "Ch.Pré "
         Height          =   345
         Left            =   4665
         TabIndex        =   13
         Top             =   225
         Width           =   825
      End
      Begin VB.Frame FraCheques 
         Caption         =   "Cheques"
         Height          =   1320
         Left            =   285
         TabIndex        =   6
         Top             =   1005
         Visible         =   0   'False
         Width           =   3585
         Begin VB.TextBox TxtNumCheque 
            Height          =   285
            Left            =   885
            TabIndex        =   9
            Top             =   330
            Width           =   855
         End
         Begin VB.TextBox TxtValorCheque 
            Height          =   285
            Left            =   855
            TabIndex        =   8
            Top             =   840
            Width           =   1215
         End
         Begin MSMask.MaskEdBox MskDtaPara 
            Height          =   285
            Left            =   2415
            TabIndex        =   7
            Top             =   330
            Width           =   1005
            _ExtentX        =   1773
            _ExtentY        =   503
            _Version        =   393216
            Format          =   "dd/mm/yyyy"
            PromptChar      =   "_"
         End
         Begin VB.Label Label6 
            Caption         =   "Número"
            Height          =   195
            Left            =   135
            TabIndex        =   12
            Top             =   315
            Width           =   615
         End
         Begin VB.Label Label7 
            Caption         =   "Para"
            Height          =   225
            Left            =   1890
            TabIndex        =   11
            Top             =   360
            Width           =   495
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Valor"
            Height          =   195
            Left            =   315
            TabIndex        =   10
            Top             =   870
            Width           =   360
         End
      End
      Begin VB.CommandButton CmdAltCheq 
         Height          =   495
         Left            =   7635
         Picture         =   "frmFimVend.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   1185
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.CommandButton CmdExcCheq 
         Height          =   435
         Left            =   7635
         Picture         =   "frmFimVend.frx":0172
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1815
         Visible         =   0   'False
         Width           =   465
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGridCheques 
         Height          =   1260
         Left            =   4080
         TabIndex        =   18
         Top             =   1110
         Visible         =   0   'False
         Width           =   3465
         _ExtentX        =   6112
         _ExtentY        =   2223
         _Version        =   393216
         Rows            =   4
         Cols            =   3
         FixedCols       =   0
         Enabled         =   0   'False
         FormatString    =   "No.Cheque      |^    Data        |>              Valor  "
      End
      Begin VB.Label LblAgencia 
         AutoSize        =   -1  'True
         Caption         =   "Agência:"
         Height          =   195
         Left            =   2115
         TabIndex        =   22
         Top             =   690
         Width           =   630
      End
      Begin VB.Label LblBanco 
         AutoSize        =   -1  'True
         Caption         =   "Banco:"
         Height          =   195
         Left            =   135
         TabIndex        =   21
         Top             =   705
         Width           =   510
      End
      Begin VB.Label LblCliente 
         AutoSize        =   -1  'True
         Caption         =   "Cliente:"
         Height          =   195
         Left            =   120
         TabIndex        =   20
         Top             =   300
         Width           =   525
      End
      Begin VB.Label LblVlrentrada 
         AutoSize        =   -1  'True
         Caption         =   "Vlr.Entrada:"
         Height          =   195
         Left            =   5685
         TabIndex        =   19
         Top             =   315
         Width           =   825
      End
   End
   Begin VB.CommandButton CmdVolta 
      Caption         =   "Volta"
      Enabled         =   0   'False
      Height          =   570
      Left            =   8880
      Picture         =   "frmFimVend.frx":0274
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Volta para alterar tipo de venda"
      Top             =   1410
      Visible         =   0   'False
      Width           =   630
   End
   Begin VB.CommandButton Cmdfinaliza 
      Caption         =   "Finaliza"
      Enabled         =   0   'False
      Height          =   570
      Left            =   8880
      Picture         =   "frmFimVend.frx":06B6
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Finaliza o pedido"
      Top             =   2310
      Visible         =   0   'False
      Width           =   630
   End
   Begin VB.CommandButton CmdTipovenda 
      Caption         =   "Seguir"
      Height          =   570
      Left            =   8865
      Picture         =   "frmFimVend.frx":0AF8
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Finaliza os produtos"
      Top             =   270
      Visible         =   0   'False
      Width           =   630
   End
End
Attribute VB_Name = "Frmfimvend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim pRstipovenda As New ADODB.Recordset
Dim pRsBalconista As New ADODB.Recordset
Dim prsCliente As New ADODB.Recordset
Private pcNumcheque As String
Private pdDatapara As Date
Private pnValorcheque As Double
Private pnCodcli As Double
Private pnNsu  As String
Private pnLinhas As Double
Public KeyAscii As Integer

' variaveis de interface

Public GarFlag As Integer
'Public db_dados As Database
Public gDatProc As String
Public item As String

Private pnTotitem As Double
Private pnTotped As Double
Private pnParcelas As Double
Private pcEntrada As String
Private pnDias As Double
Private pnAprazo As Boolean


Private Sub CboBalconista_KeyPress(KeyAscii As Integer)
  Dim CB As Long
  Dim FindString As String
  Const CB_ERR = (-1)
  Const CB_FINDSTRING = &H14C
   
  If KeyAscii < 32 Or KeyAscii > 127 Then Exit Sub
    
  If CboBalconista.SelLength = 0 Then
      FindString = CboBalconista.Text & Chr$(KeyAscii)
  Else
      FindString = Left$(CboBalconista.Text, CboBalconista.SelStart) & Chr$(KeyAscii)
  End If
  
  CB = SendMessage(CboBalconista.hWnd, CB_FINDSTRING, -1, ByVal FindString)
  
  If CB <> CB_ERR Then
      CboBalconista.ListIndex = CB
      CboBalconista.SelStart = Len(FindString)
      CboBalconista.SelLength = Len(CboBalconista.Text) - CboBalconista.SelStart
  End If
  KeyAscii = 0


End Sub

Private Sub CboClientes_KeyPress(KeyAscii As Integer)

  Dim CB As Long
  Dim FindString As String
  Const CB_ERR = (-1)
  Const CB_FINDSTRING = &H14C
   
  If KeyAscii < 32 Or KeyAscii > 127 Then Exit Sub
    
  If CboClientes.SelLength = 0 Then
      FindString = CboClientes.Text & Chr$(KeyAscii)
  Else
      FindString = Left$(CboClientes.Text, CboClientes.SelStart) & Chr$(KeyAscii)
  End If
  
  CB = SendMessage(CboClientes.hWnd, CB_FINDSTRING, -1, ByVal FindString)
  
  If CB <> CB_ERR Then
      CboClientes.ListIndex = CB
      CboClientes.SelStart = Len(FindString)
      CboClientes.SelLength = Len(CboClientes.Text) - CboClientes.SelStart
  End If
  KeyAscii = 0

End Sub

Private Sub ChkPre_Click()
   If ChkPre.Value = 1 Then
      Me.FraCheques.Visible = True
      Me.FraCheques.Enabled = True
      Me.MSFlexGridCheques.Visible = True
      Me.MSFlexGridCheques.Enabled = True
      MSFlexGridCheques.Cols = 3
      MSFlexGridCheques.Rows = 1
      'MSFlexGridCheques.Clear
      Me.CmdAltCheq.Enabled = True
      Me.CmdAltCheq.Visible = True
      Me.CmdExcCheq.Visible = True
      Me.CmdExcCheq.Enabled = True
      Me.LblBanco.Visible = True
      Me.LblAgencia.Visible = True
      Me.LblBanco.Enabled = True
      Me.LblAgencia.Enabled = True
   End If
   
End Sub

Private Sub CmdAltCheq_Click()
  With MSFlexGridCheques
     .Col = 0
     Me.TxtNumCheque = .Text
     .Col = 1
     Me.MskDtaPara = .Text
     .Col = 2
     Me.TxtValorCheque = .Text
     'MsflexgridItens.Enabled = True
     If .Rows <= 2 Then
        .Clear
        .Rows = 1
     Else
        .RemoveItem .RowSel
     End If
     Me.TxtNumCheque.SetFocus
End With

End Sub

Private Sub CmdExcCheq_Click()
  MSFlexGridCheques.Enabled = True
  If MSFlexGridCheques.Rows <= 2 Then
     'MSFlexGridCheques.Clear
     MSFlexGridCheques.Rows = 1
  Else
     MSFlexGridCheques.RemoveItem MSFlexGridCheques.RowSel
  End If
  TxtNumCheque.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
   If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()

  'Centraliza a tela no video
  Me.Move (Screen.Width - Me.Width) / 2, _
           (Screen.Height - Me.Height) / 2
  
suCarrega_Grids
  
End Sub
Private Sub suCarrega_Grids()
   Abre_Le_rst_tipovenda
   Abre_Le_rst_Balconistas
End Sub

Private Sub Abre_Le_rst_tipovenda()
   gSql = "select código,descricao "
   gSql = gSql & "FROM tipovend "
   pRstipovenda.Open gSql, ConDb, adOpenKeyset
   Carrega_Grid_tipovenda
   pRstipovenda.Close
End Sub
Private Sub Carrega_Grid_tipovenda()

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

Private Sub Abre_Le_rst_Balconistas()
   gSql = "select codoperador,nome "
   gSql = gSql & "FROM tab_operador "
   pRsBalconista.Open gSql, ConDb, adOpenKeyset
   Carrega_Grid_Balconista
    
End Sub
Private Sub Carrega_Grid_Balconista()

 CboBalconista.Clear
 With pRsBalconista
      '.MoveLast
      'nItem = .RecordCount
      .MoveFirst
      Do While Not .EOF
        CboBalconista.AddItem (pRsBalconista!nome)
        CboBalconista.ItemData(CboBalconista.NewIndex) = pRsBalconista!codoperador
        .MoveNext
      Loop
  End With
     
  For i = 0 To CboBalconista.ListCount - 1
      CboBalconista.ListIndex = i
      If CboBalconista.Text = gOperador Then
         
         Exit For
      End If
  Next
 'Else
 '     CboBalconista.ListIndex = -1
 '  End If
    
     
End Sub

