VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.MDIForm frmMain 
   BackColor       =   &H00004000&
   Caption         =   "Vibe Informatica"
   ClientHeight    =   7425
   ClientLeft      =   180
   ClientTop       =   1770
   ClientWidth     =   11055
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Height          =   1275
      Left            =   0
      ScaleHeight     =   1215
      ScaleWidth      =   10995
      TabIndex        =   1
      Top             =   0
      Width           =   11055
      Begin VB.CommandButton btnVendas 
         Caption         =   "Vendas"
         Height          =   1155
         Left            =   60
         Picture         =   "frmMain.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   60
         Width           =   1395
      End
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   7170
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   450
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            TextSave        =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   150
      Top             =   1110
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   4
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":0F4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":1B9E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":27F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":2902
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu MnArquivos 
      Caption         =   "&Arquivos"
      Begin VB.Menu mnclientes 
         Caption         =   "&Clientes"
      End
      Begin VB.Menu Mnfornecedores 
         Caption         =   "&Fornecedores"
      End
      Begin VB.Menu mnBalconistas 
         Caption         =   "Operadores"
      End
      Begin VB.Menu MnGrupos 
         Caption         =   "&Grupos"
      End
      Begin VB.Menu MnProdutos 
         Caption         =   "&Produtos"
      End
      Begin VB.Menu mnTiposUnidade 
         Caption         =   "Tipos de &Unidades"
      End
      Begin VB.Menu Mntipovend 
         Caption         =   "Tipos de &Venda"
      End
      Begin VB.Menu MnTipomov 
         Caption         =   "Tipos de &Movimentação"
      End
      Begin VB.Menu mntracoOS 
         Caption         =   "-"
      End
      Begin VB.Menu mnAparelhos 
         Caption         =   "Aparelhos"
      End
      Begin VB.Menu mnMarcas 
         Caption         =   "Marcas"
      End
      Begin VB.Menu mnModelos 
         Caption         =   "Modelos"
      End
   End
   Begin VB.Menu mnMovimenta 
      Caption         =   "Movimentações"
      Begin VB.Menu mnvendas 
         Caption         =   "&Vendas"
      End
      Begin VB.Menu mnCancvenda 
         Caption         =   "Cancelar Vendas"
      End
      Begin VB.Menu mntraco1 
         Caption         =   "-"
      End
      Begin VB.Menu Mnfechacli 
         Caption         =   "&Fechamento de Clientes"
      End
      Begin VB.Menu Mntraco2 
         Caption         =   "-"
      End
      Begin VB.Menu MnCompras 
         Caption         =   "&Compras"
         Begin VB.Menu mnPedcompra 
            Caption         =   "&Pedidos de Compra"
         End
         Begin VB.Menu MPedAtend 
            Caption         =   "&Pedidos Recebidos"
         End
      End
      Begin VB.Menu Mntraco3 
         Caption         =   "-"
      End
      Begin VB.Menu mnoutrasmov 
         Caption         =   "Outras Movimentações"
      End
      Begin VB.Menu mnuApagar 
         Caption         =   "Contas a &Pagar"
      End
      Begin VB.Menu Mnuhifen1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuLancdesp 
         Caption         =   "Lançto.Despesas"
      End
   End
   Begin VB.Menu mnRelato 
      Caption         =   "&Relatórios"
      Begin VB.Menu MnuRelLisPreco 
         Caption         =   "Lista de Preços"
      End
      Begin VB.Menu mnuRelAbaixoMini 
         Caption         =   "Produtos abaixo do Mínimo"
      End
      Begin VB.Menu MnuRelzero 
         Caption         =   "Produtos com Estoque Zero"
      End
      Begin VB.Menu MnuRelPosfisica 
         Caption         =   "Posição Física"
      End
      Begin VB.Menu MnuRelInventa 
         Caption         =   "Lista Para Inventário"
      End
      Begin VB.Menu Mnu1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuRelvendas 
         Caption         =   "Vendas do Período"
         Begin VB.Menu mnuRVenPedido 
            Caption         =   "Por Pedido de Venda"
         End
         Begin VB.Menu MnuRelVenTipo 
            Caption         =   "Por Tipo de Venda"
         End
         Begin VB.Menu MnuRVenBalco 
            Caption         =   "Por Balconista"
         End
      End
      Begin VB.Menu MnuRelMaisvend 
         Caption         =   "Produtos Mais Vendidos"
      End
      Begin VB.Menu MnuRelMenosVend 
         Caption         =   "Produtos Menos Vendidos"
      End
      Begin VB.Menu MnuRelMovim 
         Caption         =   "Movimentações do Período"
      End
      Begin VB.Menu mnu2 
         Caption         =   "-"
      End
      Begin VB.Menu MnuListaNegra 
         Caption         =   "Lista Negra"
      End
      Begin VB.Menu MnuRelclisaldomaior 
         Caption         =   "Clientes com saldo maior que limite"
      End
      Begin VB.Menu traco1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuRelctasareceber 
         Caption         =   "Contas a Receber "
      End
      Begin VB.Menu mnuCheqARec 
         Caption         =   "Cheques a Receber"
      End
      Begin VB.Menu MnureCctaAPagar 
         Caption         =   "Contas a Pagar no período"
      End
      Begin VB.Menu traco2 
         Caption         =   "-"
      End
      Begin VB.Menu MnuRelcaixa 
         Caption         =   "Caixa do dia"
      End
   End
   Begin VB.Menu mnutilitarios 
      Caption         =   "&Utilitários"
      Begin VB.Menu mncopia 
         Caption         =   "Cópia de segurança"
      End
      Begin VB.Menu MnParametros 
         Caption         =   "&Parametros"
      End
      Begin VB.Menu MnPermissoes 
         Caption         =   "&Permissões"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpContents 
         Caption         =   "&Conteúdo"
      End
      Begin VB.Menu mnuHelpBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&Sistema  Sgl..."
      End
   End
   Begin VB.Menu mnsaida 
      Caption         =   "&Saida"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function OSWinHelp% Lib "User32" Alias "WinHelpA" (ByVal hwnd&, ByVal HelpFile$, ByVal wCommand%, dwData As Any)
Private LocalFoto

Private Sub btnVendas_Click()
    FrmBalcao.Show vbModal
End Sub

Private Sub MDIForm_Load()
    
    Left = GetSetting(App.Title, "Settings", "MainLeft", 1000)
    Top = GetSetting(App.Title, "Settings", "MainTop", 1000)
    Width = GetSetting(App.Title, "Settings", "MainWidth", 6500)
    Height = GetSetting(App.Title, "Settings", "MainHeight", 6500)
    LocalFoto = App.Path & "\soapple.jpg"
    If gNivel > 1 Then
        mnclientes.Enabled = False
        mnBalconistas.Enabled = False
        Mnfornecedores.Enabled = False
        MnGrupos.Enabled = False
        mnAparelhos.Enabled = False
        Mnfornecedores.Enabled = False
        mnModelos.Enabled = False
        mnMarcas.Enabled = False
        MnParametros.Enabled = False
        MnProdutos.Enabled = False
        MnTipomov.Enabled = False
        Mntipovend.Enabled = False
        mnTiposUnidade.Enabled = False
    End If
    Me.Picture = LoadPicture(LocalFoto)
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
 If UnloadMode <> 1 Then
    If MsgBox("Vai embora mesmo ?", 32 + 4 + 256) <> 6 Then
       Cancel = True
    End If
End If
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    If Me.WindowState <> vbMinimized Then
        SaveSetting App.Title, "Settings", "MainLeft", Me.Left
        SaveSetting App.Title, "Settings", "MainTop", Me.Top
        SaveSetting App.Title, "Settings", "MainWidth", Me.Width
        SaveSetting App.Title, "Settings", "MainHeight", Me.Height
    End If
End Sub

Private Sub mnuViewDatacadvend_Click()
    'Dim F As New frmcadvend
    'F.Show
End Sub

Private Sub mnAparelhos_Click()
    frmcadAparelhos.Show vbModal
End Sub

Private Sub mnBalconistas_Click()
   frmBalco.Show vbModal
End Sub

Private Sub MnCadImpostos_Click()
   frmImpostos.Show vbModal
End Sub

Private Sub mnCadObras_Click()
   FrmObras.Show vbModal
End Sub

Private Sub mnCadservicos_Click()
   frmServicos.Show vbModal
End Sub

Private Sub mnCancvenda_Click()
   FrmCancPedido.Show vbModal
End Sub

Private Sub MNCcusto_Click()
    frmCCusto.Show vbModal
End Sub

Private Sub mnclientes_Click()
   FrmClientes.Show vbModal
End Sub

Private Sub mncopia_Click()
   FrmBackup.Show vbModal
End Sub

Private Sub Mnentrada_Click()
   FrmLisPComp.Show vbModal
End Sub

Private Sub Mnfechacli_Click()
   FrmFechaCli.Show vbModal
End Sub

Private Sub Mnfornecedores_Click()
   frmfornec.Show vbModal
End Sub

Private Sub MnGrupos_Click()
   frmgrupos.Show vbModal
End Sub

Private Sub MnLojas_Click()
   frmlojas.Show vbModal
End Sub

Private Sub MnManutAPagar_Click()
   FrmApagar.Show vbModal
End Sub

Private Sub mnlistapreco_Click()
    Call suListaprecos
End Sub

Private Sub mnMoedas_Click()
   frmcadmoe.Show vbModal
End Sub

Private Sub mnObrasPlanilha_Click()
   FrmPlanilha.Show vbModal
End Sub

Private Sub MnOrca_Click()
   FrmListaOrc.Show vbModal
End Sub

Private Sub mnoutrasmov_Click()
   frmOutMo.Show vbModal
End Sub

Private Sub MnParametros_Click()
   frmlojas.Show vbModal
End Sub

Private Sub mnPlanilhaPedidos_Click()

End Sub

Private Sub mnPedBaixa_Click()
   FrmPedAten.Show vbModal
End Sub

Private Sub mnPedcompra_Click()
   FrmLisPComp.Show vbModal
End Sub

Private Sub MnProdutos_Click()
   FrmProdutos.Show vbModal
End Sub

Private Sub mnsaida_Click()
   If MsgBox("Vai embora mesmo ?", 32 + 4 + 256) <> 6 Then
      Cancel = True
   Else
      End
   End If
End Sub

Private Sub MnTipomov_Click()
   FrmTipomov.Show vbModal
End Sub

Private Sub Mntipovend_Click()
   Frmtipovend.Show vbModal
End Sub

Private Sub mnuApagar_Click()
   FrmApagar.Show vbModal
End Sub

Private Sub mnuCheqARec_Click()
FrmDataCArec.PassRel = "rChArec.rpt"
FrmDataCArec.Show vbModal

End Sub

Private Sub mnuHelpAbout_Click()
    frmAbout.Show vbModal, Me
End Sub


Private Sub mnuHelpContents_Click()

    Dim nRet As Integer


    'if there is no helpfile for this project display a message to the user
    'you can set the HelpFile for your application in the
    'Project Properties dialog
    If Len(App.HelpFile) = 0 Then
        MsgBox "Incapaz de mostrar Conteúdo do Help. Não há Help associado a essa aplicação.", vbInformation, Me.Caption
    Else
        On Error Resume Next
        nRet = OSWinHelp(Me.hwnd, App.HelpFile, 3, 0)
        If Err Then
            MsgBox Err.Description
        End If
    End If
End Sub


Private Sub mnuHelpSearch_Click()

    Dim nRet As Integer


    'if there is no helpfile for this project display a message to the user
    'you can set the HelpFile for your application in the
    'Project Properties dialog
    If Len(App.HelpFile) = 0 Then
        MsgBox "Incapaz de mostrar Conteúdo do Help. Não há Help associado a essa aplicação.", vbInformation, Me.Caption
    Else
        On Error Resume Next
        nRet = OSWinHelp(Me.hwnd, App.HelpFile, 261, 0)
        If Err Then
            MsgBox Err.Description
        End If
    End If
End Sub

Private Sub Mnuusua_Click()

End Sub

Private Sub mnvendasperiodo_Click()

End Sub

Private Sub MnuLancdesp_Click()
   frmLancDesp.Show vbModal
End Sub

Private Sub MnuListaNegra_Click()
  FrmCompras.CrRelcomp.DataFiles(0) = ConDb.DefaultDatabase & ".mdb"
  FrmCompras.CrRelcomp.Destination = 0 'Vídeo
  CristalSelect = "{tab_Clientes.negativo} = True"
  FrmCompras.CrRelcomp.SelectionFormula = CristalSelect
  FrmCompras.CrRelcomp.Formulas(0) = "nomeloja = '" & gNome & "'"
  FrmCompras.CrRelcomp.ReportFileName = gPathRel & "\rlisnegra.rpt"
  FrmCompras.CrRelcomp.Action = 1
End Sub

Private Sub MnureCctaAPagar_Click()
MsgBox "Em Desenvolvimento ", vbOKOnly, "Atenção " & gOperador
End Sub

Private Sub mnuRelAbaixoMini_Click()
  FrmCompras.CrRelcomp.DataFiles(0) = ConDb.DefaultDatabase & ".mdb"
  FrmCompras.CrRelcomp.Destination = 0 'Vídeo
  CristalSelect = "{tab_produtos.estatual} < {tab_produtos.minimo}"
  FrmCompras.CrRelcomp.SelectionFormula = CristalSelect
  FrmCompras.CrRelcomp.Formulas(0) = "nomeloja = '" & gNome & "'"
  FrmCompras.CrRelcomp.ReportFileName = gPathRel & "\relminim.rpt"
  FrmCompras.CrRelcomp.Action = 1

End Sub

Private Sub MnuRelcaixa_Click()
MsgBox "Em Desenvolvimento ", vbOKOnly, "Atenção " & gOperador
End Sub

Private Sub MnuRelclisaldomaior_Click()
  FrmCompras.CrRelcomp.DataFiles(0) = ConDb.DefaultDatabase & ".mdb"
  FrmCompras.CrRelcomp.Destination = 0 'Vídeo
  CristalSelect = "{tab_Clientes.limite} < {tab_clientes.saldo}"
  FrmCompras.CrRelcomp.SelectionFormula = CristalSelect
  FrmCompras.CrRelcomp.Formulas(0) = "nomeloja = '" & gNome & "'"
  FrmCompras.CrRelcomp.ReportFileName = gPathRel & "\rsaldomaior.rpt"
  FrmCompras.CrRelcomp.Action = 1

End Sub

Private Sub MnuRelctasareceber_Click()
FrmDataArec.PassRel = "rarecebe.rpt"
FrmDataArec.Show vbModal
End Sub

Private Sub MnuRelInventa_Click()
MsgBox "Em Desenvolvimento ", vbOKOnly, "Atenção " & gOperador
End Sub

Private Sub MnuRelLisPreco_Click()
   Call suListaprecos
End Sub

Private Sub MnuRelMaisvend_Click()
FrmDatas1.PassRel = "rMaisVen.rpt"
FrmDatas1.Show vbModal

End Sub

Private Sub MnuRelMenosVend_Click()
  FrmDatasMenos.PassRel = "rMenosVen.rpt"
  FrmDatasMenos.Show vbModal

End Sub

Private Sub MnuRelMovim_Click()
  FrmDataMovi.PassRel = "rMovPer.rpt"
  FrmDataMovi.Show vbModal

End Sub

Private Sub MnuRelPosfisica_Click()
    Call suPosicaoFisica
End Sub


Private Sub MnuRelVenTipo_Click()
   FrmDatas.PassRel = "rVenTipo.rpt"
   FrmDatas.Show vbModal
   
End Sub

Private Sub MnuRelzero_Click()
  FrmCompras.CrRelcomp.DataFiles(0) = ConDb.DefaultDatabase & ".mdb"
  FrmCompras.CrRelcomp.Destination = 0 'Vídeo
  CristalSelect = "{tab_produtos.estatual} <= 0 "
  FrmCompras.CrRelcomp.SelectionFormula = CristalSelect
  FrmCompras.CrRelcomp.Formulas(0) = "nomeloja = '" & gNome & "'"
  FrmCompras.CrRelcomp.ReportFileName = gPathRel & "relzero.rpt"
  FrmCompras.CrRelcomp.Action = 1

End Sub

Private Sub MnuRVenBalco_Click()
   FrmDatas.PassRel = "rVenBalco.rpt"
   FrmDatas.Show vbModal
   
End Sub

Private Sub mnuRVenPedido_Click()
   FrmDatas.PassRel = "rVenPedido.rpt"
   FrmDatas.Show vbModal
   
End Sub

Private Sub mnvendas_Click()
    FrmVendas.Show vbModal
End Sub

Private Sub MPedAtend_Click()
   FrmPCompaten.Show vbModal
End Sub

Private Sub mUnidades_Click()
    FrmUnid.Show vbModal
End Sub

Private Sub suListaprecos()
    Dim CristalSelect As String
    'Relatório de reajuste de principal
    'Dim Diai$, Mesi$, Anoi$
    'Dim Diaf$, Mesf$, Anof$
    'Diai = Str(Day(txt_dtent.Text)): Mesi = Str(Month(txt_dtent.Text)): Anoi = Str(Year(txt_dtent.Text))
    'Diaf = Str(Day(txt_dtfim.Text)): Mesf = Str(Month(txt_dtfim.Text)): Anof = Str(Year(txt_dtfim.Text))
    
    'crtRelReaj.Connect = "DSN=poc;UID=usr_poc_opcredito;PWD=opcredito;DSQ=dbs_poc_opcredito"
    'If Tipo_Tela = "A" Then
       CrRep1.DataFiles(0) = ConDb.DefaultDatabase & ".mdb"
       CrRep1.Destination = 0 'Vídeo
       'CristalSelect = "{viw_poc_entradaativo.dta_libinicio} >= Date(" + Anoi + "," + Mesi + "," + Diai + ") and {viw_poc_entradaativo.dta_libinicio} <= Date(" + Anof + "," + Mesf + "," + Diaf + ") "
       'CrRep1.SelectionFormula = CristalSelect
       CrRep1.Formulas(0) = "nomeloja = '" & gNome & "'"
       CrRep1.ReportFileName = gPathRel & "listpreco.rpt"
       CrRep1.Action = 1
    'Else
    '   crtRelReaj.Destination = 0 'Vídeo
    '   CristalSelect = "{viw_poc_entradapassivo.dta_libinicio} >= Date(" + Anoi + "," + Mesi + "," + Diai + ") and {viw_poc_entradapassivo.dta_libinicio} <= Date(" + Anof + "," + Mesf + "," + Diaf + ") "
    '   crtRelReaj.SelectionFormula = CristalSelect
    '   crtRelReaj.ReportFileName = App.Path + "\report\entradapassivo.rpt"
    '   crtRelReaj.Action = 1
    'End If
    

End Sub
Private Sub suPosicaoFisica()
    Dim CristalSelect As String
       CrRep1.DataFiles(0) = ConDb.DefaultDatabase & ".mdb"
       CrRep1.Destination = 0 'Vídeo
       'CristalSelect = "{viw_poc_entradaativo.dta_libinicio} >= Date(" + Anoi + "," + Mesi + "," + Diai + ") and {viw_poc_entradaativo.dta_libinicio} <= Date(" + Anof + "," + Mesf + "," + Diaf + ") "
       'CrRep1.SelectionFormula = CristalSelect
       CrRep1.Formulas(0) = "nomeloja = '" & gNome & "'"
       CrRep1.ReportFileName = gPathRel & "Posfisic.rpt"
       CrRep1.Action = 1
    'Else
    '   crtRelReaj.Destination = 0 'Vídeo
    '   CristalSelect = "{viw_poc_entradapassivo.dta_libinicio} >= Date(" + Anoi + "," + Mesi + "," + Diai + ") and {viw_poc_entradapassivo.dta_libinicio} <= Date(" + Anof + "," + Mesf + "," + Diaf + ") "
    '   crtRelReaj.SelectionFormula = CristalSelect
    '   crtRelReaj.ReportFileName = App.Path + "\report\entradapassivo.rpt"
    '   crtRelReaj.Action = 1
    'End If
    

End Sub


Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
'    Select Case Button.Key
'    Case Is = "Vendas"
'         mnvendas_Click
'    Case Is = "Compras"
'         'mnEntrada_click
'    Case Is = "Relatorios"
'         'mnRelato_click
'    Case Is = "Saida"
'         mnsaida_Click
'    End Select
    
End Sub

