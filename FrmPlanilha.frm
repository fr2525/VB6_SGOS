VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form FrmPlanilha 
   Caption         =   "Planilha da Obra"
   ClientHeight    =   2790
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8205
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   2790
   ScaleWidth      =   8205
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Height          =   2325
      Left            =   720
      ScaleHeight     =   2265
      ScaleWidth      =   3345
      TabIndex        =   13
      Top             =   2640
      Visible         =   0   'False
      Width           =   3405
      Begin VB.Frame Frame2 
         Height          =   855
         Left            =   690
         TabIndex        =   21
         Top             =   1140
         Width           =   2235
         Begin VB.CommandButton Btcorrige 
            Caption         =   "&Corrige"
            Height          =   540
            Left            =   795
            Picture         =   "FrmPlanilha.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   18
            Tag             =   "&Update"
            Top             =   210
            Width           =   615
         End
         Begin VB.CommandButton Btsair 
            Caption         =   "&Sair"
            Height          =   540
            Left            =   1470
            Picture         =   "FrmPlanilha.frx":00FA
            Style           =   1  'Graphical
            TabIndex        =   19
            Tag             =   "&Update"
            Top             =   210
            Width           =   615
         End
         Begin VB.CommandButton BtSalvar 
            Caption         =   "&Salvar"
            Height          =   540
            Left            =   135
            Picture         =   "FrmPlanilha.frx":01F4
            Style           =   1  'Graphical
            TabIndex        =   17
            Tag             =   "&Update"
            Top             =   210
            Width           =   615
         End
      End
      Begin VB.TextBox TxtMedicao 
         Height          =   315
         Left            =   1800
         TabIndex        =   14
         Top             =   120
         Width           =   960
      End
      Begin VB.TextBox TxtDta_medicao 
         Height          =   315
         Left            =   1800
         TabIndex        =   16
         Top             =   615
         Width           =   1305
      End
      Begin VB.Label Label3 
         Caption         =   "Numero da Medição"
         Height          =   390
         Left            =   75
         TabIndex        =   20
         Top             =   150
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Data da Medição"
         Height          =   300
         Left            =   255
         TabIndex        =   15
         Top             =   660
         Width           =   1515
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " Medição "
      Height          =   1710
      Left            =   540
      TabIndex        =   9
      Top             =   690
      Width           =   3855
      Begin VB.CommandButton BtNovaMedicao 
         Height          =   345
         Left            =   3015
         Picture         =   "FrmPlanilha.frx":02EE
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   480
         Width           =   330
      End
      Begin VB.CommandButton BtExcluir 
         Height          =   345
         Left            =   3015
         Picture         =   "FrmPlanilha.frx":03D8
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   915
         Width           =   330
      End
      Begin MSFlexGridLib.MSFlexGrid GrdMedicao 
         Height          =   1170
         Left            =   315
         TabIndex        =   11
         Top             =   330
         Width           =   2430
         _ExtentX        =   4286
         _ExtentY        =   2064
         _Version        =   393216
         FixedCols       =   0
      End
   End
   Begin VB.CommandButton CmdSair 
      Caption         =   "&Sair"
      Height          =   510
      Left            =   7350
      Picture         =   "FrmPlanilha.frx":04DA
      Style           =   1  'Graphical
      TabIndex        =   7
      Tag             =   "&Update"
      Top             =   1920
      Width           =   615
   End
   Begin VB.CommandButton BtRecebimentos 
      Caption         =   "Recebimentos"
      Enabled         =   0   'False
      Height          =   495
      Left            =   5685
      TabIndex        =   5
      Top             =   1320
      Width           =   1155
   End
   Begin VB.CommandButton BtServicos 
      Caption         =   "Serviços"
      Enabled         =   0   'False
      Height          =   510
      Left            =   4530
      TabIndex        =   1
      Top             =   795
      Width           =   1155
   End
   Begin VB.CommandButton BtImpostos 
      Caption         =   "Impostos"
      Enabled         =   0   'False
      Height          =   510
      Left            =   6840
      TabIndex        =   3
      Top             =   795
      Width           =   1140
   End
   Begin VB.CommandButton Btconsulta 
      Caption         =   "Consultas"
      Enabled         =   0   'False
      Height          =   495
      Left            =   6810
      TabIndex        =   6
      Top             =   1320
      Width           =   1170
   End
   Begin VB.CommandButton BtNF 
      Caption         =   "N.F.Emitidas"
      Enabled         =   0   'False
      Height          =   510
      Left            =   4530
      TabIndex        =   4
      Top             =   1305
      Width           =   1155
   End
   Begin VB.CommandButton Btdespesas 
      Caption         =   "Despesas"
      Enabled         =   0   'False
      Height          =   510
      Left            =   5685
      TabIndex        =   2
      Top             =   795
      Width           =   1155
   End
   Begin VB.ComboBox CmbObra 
      Height          =   315
      Left            =   540
      TabIndex        =   0
      Top             =   150
      Width           =   3885
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Obra"
      Height          =   195
      Left            =   90
      TabIndex        =   8
      Top             =   180
      Width           =   345
   End
End
Attribute VB_Name = "FrmPlanilha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim pRsObra As New ADODB.Recordset
Dim pRsMedicao As New ADODB.Recordset
Dim lnOldHeight As Double
Public pcodObra As Double
Public pcodMedicao As Double
Public pdataMedicao As Date

Private Sub BtNova_Click()
   FrmObras.Show vbModal
End Sub

Private Sub Btconsulta_Click()
    frmplanconsulta.Show vbModal
End Sub

Private Sub Btcorrige_Click()
   TxtMedicao.Text = ""
   TxtDta_medicao = ""
   TxtMedicao.SetFocus
   
End Sub

Private Sub Btdespesas_Click()
  frmPlanDesp.Show vbModal
End Sub

Private Sub BtExcluir_Click()
  On Error GoTo ErroDelete
    
    'this may produce an error if you delete the last
    'record or the only record in the recordset
    If MsgBox("Deseja realmente apagar esta Medição ? ", vbYesNo, "Atenção") = vbYes Then
       gSql = "delete * from tab_med where cod_obra  = " & Me.CmbObra.ItemData(Me.CmbObra.ListIndex)
       Me.GrdMedicao.Col = 0
       gSql = gSql & " and medicao = " & Val(Me.GrdMedicao.Text) & ""
       Me.GrdMedicao.Col = 1
       gSql = gSql & " and dta_medicao = cdate('" & Me.GrdMedicao.Text & "')"
       ConDb.Execute gSql
       GrdMedicao.Clear
       suCarrega_Grid_Medicao
    End If
    Exit Sub
ErroDelete:
     MsgBox "Deu erro na exclusao da Medição " & Chr(13) & "Instrucao Sql = '" & _
            cSql & "'  "
End Sub



Private Sub BtImpostos_Click()
    frmPlanImpo.Show vbModal
End Sub

Private Sub BtNF_Click()
    frmPlanNotas.Show vbModal
End Sub

Private Sub BtNovaMedicao_Click()
   suDesabilita_botoes
   CmbObra.Enabled = False
   Frame1.Enabled = False
   Picture1.Visible = True
   BtNovaMedicao.Enabled = False
   BtExcluir.Enabled = False
   lnOldHeight = Me.Height
   Me.Height = 5835
   TxtMedicao.SetFocus
   'FrmMedicao.Show vbModal
End Sub

Private Sub BtNovaObra_Click()
   FrmObras.Show vbModal
End Sub


Private Sub BtRecebimentos_Click()
 frmplanreceb.Show vbModal
End Sub

Private Sub Btsair_Click()
   suHabilita_botoes
   CmbObra.Enabled = True
   Frame1.Enabled = True
   BtNovaMedicao.Enabled = True
   BtExcluir.Enabled = True
   CmdSair.Enabled = True
   Picture1.Visible = True
   Me.Height = lnOldHeight
End Sub

Private Sub BtSalvar_Click()
   gSql = "INSERT INTO tab_med (cod_obra,medicao,dta_medicao, operador, datatual) "
   gSql = gSql & " VALUES( " & FrmPlanilha.CmbObra.ItemData(FrmPlanilha.CmbObra.ListIndex) & ",'" & TxtMedicao.Text & "',cdate('" & TxtDta_medicao.Text & "'),'" & gOperador & "',Cdate('" & Date & "') )"
   ConDb.Execute gSql
   suHabilita_botoes
   CmbObra.Enabled = True
   Frame1.Enabled = True
   BtNovaMedicao.Enabled = True
   BtExcluir.Enabled = True
   CmdSair.Enabled = True
   suCarrega_Grid_Medicao
   Picture1.Visible = False
   Me.Height = lnOldHeight
   
End Sub

Private Sub BtServicos_Click()
   frmPlanServ.Show vbModal
End Sub

Private Sub CmbObra_Click()
   LblUnidade = CmbObra.Text
   'lclcontratante = CmbObra.Text
   suCarrega_Grid_Medicao
End Sub

Private Sub CmbObra_LostFocus()
  ' LblUnidade = CmbObra.Text
  ' 'lclcontratante = CmbObra.Text
  ' suCarrega_Grid_Medicao
End Sub

Private Sub CmbObra_Validate(Cancel As Boolean)
   'LblUnidade = CmbObra.Text
   ''lclcontratante = CmbObra.Text
   'suCarrega_Grid_Medicao
End Sub

Private Sub CmdSair_Click()
   Unload Me
End Sub

Private Sub Command5_Click()

End Sub

Private Sub Form_Activate()
   suCarregaGrids
End Sub

Private Sub suCarregaGrids()
   Abre_Le_rst_CadObras
   
End Sub
Private Sub Abre_Le_rst_CadObras()
   gSql = "select cod_obra,unidade,contratante "
   gSql = gSql & "FROM Tab_Obras "
   pRsObra.Open gSql, ConDb, adOpenKeyset
   suCarrega_Combo_obras
   pRsObra.Close
End Sub
Private Sub suCarrega_Combo_obras()
 
 CmbObra.Clear
 With pRsObra
      '.MoveLast
      'nItem = .RecordCount
      .MoveFirst
      Do While Not .EOF
        CmbObra.AddItem (pRsObra!unidade & " - " & pRsObra!contratante)
        CmbObra.ItemData(CmbObra.NewIndex) = pRsObra!cod_obra
        .MoveNext
      Loop
  End With
     
End Sub

Private Sub suCarrega_Grid_Medicao()

 gSql = "Select cod_obra,medicao,dta_medicao from tab_med "
 gSql = gSql & " WHERE cod_obra = " & CmbObra.ItemData(CmbObra.ListIndex)
 pRsMedicao.Open gSql, ConDb, adOpenForwardOnly
 If pRsMedicao.BOF And pRsMedicao.EOF Then
    MsgBox "Não existem medicoes no cadastro. Favor cadastrar", vbOKOnly, "Atenção"
    suDesabilita_botoes
    pRsMedicao.Close
    CmbObra.Enabled = False
    Frame1.Enabled = False
    Picture1.Visible = True
    BtNovaMedicao.Enabled = False
    BtExcluir.Enabled = False
    lnOldHeight = Me.Height
    Me.Height = 5835
    TxtMedicao.SetFocus
    Exit Sub
 End If
          
   GrdMedicao.Clear
   GrdMedicao.Cols = 2
   GrdMedicao.Rows = 1
   GrdMedicao.Row = 0
   GrdMedicao.Col = 0
   GrdMedicao.Text = "No.Medição"
   GrdMedicao.Col = 1
   'GrdMedicao.ColWidth(1) = 4330
   GrdMedicao.Text = "Data    "
   'GrdMedicao.Col = 2
    

  'Teste do MsHFlexgrid1 - eh eh eh
  GrdMedicao.Row = 0
  GrdMedicao.FontWidth = 1
  
  With pRsMedicao
      '.MoveLast
      'nItem = .RecordCount
      .MoveFirst
      GrdMedicao.Rows = 1
      i = 0
      Do While Not .EOF
        GrdMedicao.Rows = GrdMedicao.Rows + 1
        
        GrdMedicao.Row = GrdMedicao.Rows - 1
        GrdMedicao.Col = 0: GrdMedicao.Text = "" & !medicao
        GrdMedicao.Col = 1: GrdMedicao.Text = "" & !Dta_medicao
        
        .MoveNext
         
      Loop
      GrdMedicao.FixedRows = 1
          
  End With
       
  pRsMedicao.Close

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
   
End Sub

Private Sub Form_Load()
  'Centraliza a tela no video
   Me.Move (Screen.Width - Me.Width) / 2, _
           (Screen.Height - Me.Height) / 2
   

End Sub

Private Sub Form_Unload(Cancel As Integer)
  'Habilita_menu frmMain
  
End Sub

Private Sub GrdMedicao_Click()
  Dim oldrow As Long
  
  oldrow = GrdMedicao.Row
  
  GrdMedicao.Row = 0
  
  With GrdMedicao
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
     
     gUnidade = CmbObra.Text
     
     .Col = 0: .CellBackColor = vbYellow
     gNumMedicao = Val(.Text)

     .Col = 1: .CellBackColor = vbYellow
     gDataMedicao = .Text
     .TopRow = .Row
     gCodObra = CmbObra.ItemData(CmbObra.ListIndex)
     
  End With
  suHabilita_botoes
  
End Sub

Private Sub TxtDta_medicao_Validate(Cancel As Boolean)
   If Not TxtDta_medicao = "" Then
   If Not IsDate(TxtDta_medicao.Text) Then
      MsgBox " Data Invalida !! ", vbCritical, " Erro na Data "
      Cancel = True
   End If
   End If
End Sub

Private Sub suHabilita_botoes()

  Me.BtServicos.Enabled = True
  BtImpostos.Enabled = True
  Btdespesas.Enabled = True
  BtRecebimentos.Enabled = True
  Me.BtNF.Enabled = True
  Me.BtRecebimentos.Enabled = True
  Me.Btconsulta.Enabled = True

End Sub

Private Sub suDesabilita_botoes()

  Me.BtServicos.Enabled = False
  BtImpostos.Enabled = False
  Btdespesas.Enabled = False
  BtRecebimentos.Enabled = False
  Me.BtNF.Enabled = False
  Me.BtRecebimentos.Enabled = False
  Me.Btconsulta.Enabled = False
  CmdSair.Enabled = False
End Sub

