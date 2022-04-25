VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmProdutos 
   Caption         =   "FrmProdutos"
   ClientHeight    =   6615
   ClientLeft      =   1590
   ClientTop       =   945
   ClientWidth     =   8850
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6615
   ScaleWidth      =   8850
   Begin TabDlg.SSTab SSTab1 
      Height          =   6465
      Left            =   30
      TabIndex        =   0
      Top             =   60
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   11404
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Lista"
      TabPicture(0)   =   "FrmProdutos.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "MSFlexGrdProd"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Detalhe"
      TabPicture(1)   =   "FrmProdutos.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label6"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label5"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label3"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label2"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label1"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Label15"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Label16"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Label8"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Label9"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Frame1"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "CmbMoeda"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "TxtCodBarra"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "TxtDescricao"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "CmbGrupo"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "TxtCodigo"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "Frame3"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "Frame4"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "TxtUltCompra"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "TxtUltVenda"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "TxtEstatual"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "TxtEstminimo"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).Control(21)=   "FraFornecedores"
      Tab(1).Control(21).Enabled=   0   'False
      Tab(1).Control(22)=   "MSFlxGridaux"
      Tab(1).Control(22).Enabled=   0   'False
      Tab(1).ControlCount=   23
      Begin MSFlexGridLib.MSFlexGrid MSFlxGridaux 
         Height          =   990
         Left            =   -74865
         TabIndex        =   46
         Top             =   5235
         Visible         =   0   'False
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   1746
         _Version        =   393216
         Rows            =   1
         Cols            =   1
         FixedRows       =   0
         FixedCols       =   0
         Enabled         =   0   'False
         FocusRect       =   0
         HighLight       =   0
         GridLinesFixed  =   0
         ScrollBars      =   0
      End
      Begin VB.Frame FraFornecedores 
         Caption         =   "Fornecedores"
         Height          =   1725
         Left            =   -73500
         TabIndex        =   44
         Top             =   3645
         Width           =   6240
         Begin MSFlexGridLib.MSFlexGrid MsflxgridForne 
            Height          =   1410
            Left            =   120
            TabIndex        =   45
            Top             =   195
            Width           =   5955
            _ExtentX        =   10504
            _ExtentY        =   2487
            _Version        =   393216
            FixedCols       =   0
            BackColor       =   16777215
            BackColorSel    =   8454143
            ForeColorSel    =   0
            Enabled         =   0   'False
            SelectionMode   =   1
            FormatString    =   " Código    |  Nome do Fornecedor                                                                         "
         End
      End
      Begin VB.TextBox TxtEstminimo 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   -71595
         TabIndex        =   16
         Text            =   "0"
         Top             =   3015
         Width           =   690
      End
      Begin VB.TextBox TxtEstatual 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   -73665
         TabIndex        =   15
         Text            =   "0"
         Top             =   3015
         Width           =   675
      End
      Begin VB.TextBox TxtUltVenda 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "d/M/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   3
         EndProperty
         Enabled         =   0   'False
         Height          =   315
         Left            =   -67455
         TabIndex        =   18
         Text            =   "99/99/9999"
         Top             =   3015
         Width           =   1065
      End
      Begin VB.TextBox TxtUltCompra 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "d/M/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   3
         EndProperty
         Enabled         =   0   'False
         Height          =   315
         Left            =   -69660
         TabIndex        =   17
         Text            =   "99/99/9999"
         Top             =   3015
         Width           =   1050
      End
      Begin VB.Frame Frame4 
         Height          =   735
         Left            =   -72825
         TabIndex        =   39
         Top             =   5460
         Width           =   4245
         Begin VB.CommandButton cmddesfaz 
            Caption         =   "&Desfaz"
            Enabled         =   0   'False
            Height          =   540
            Left            =   2865
            Picture         =   "FrmProdutos.frx":0038
            Style           =   1  'Graphical
            TabIndex        =   27
            Tag             =   "&Update"
            Top             =   150
            Width           =   615
         End
         Begin VB.CommandButton CmdSair 
            Caption         =   "&Sair"
            Height          =   540
            Left            =   3555
            Picture         =   "FrmProdutos.frx":0132
            Style           =   1  'Graphical
            TabIndex        =   28
            Tag             =   "&Update"
            Top             =   150
            Width           =   615
         End
         Begin VB.CommandButton cmdAdd 
            Caption         =   "&Incluir"
            Height          =   540
            Left            =   75
            Picture         =   "FrmProdutos.frx":022C
            Style           =   1  'Graphical
            TabIndex        =   29
            Tag             =   "&Add"
            Top             =   150
            Width           =   615
         End
         Begin VB.CommandButton cmdDelete 
            Caption         =   "&Excluir"
            Height          =   540
            Left            =   1470
            Picture         =   "FrmProdutos.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   30
            Tag             =   "&Delete"
            Top             =   150
            Width           =   615
         End
         Begin VB.CommandButton cmdEditar 
            Caption         =   "&Editar"
            Height          =   540
            Left            =   765
            Picture         =   "FrmProdutos.frx":0488
            Style           =   1  'Graphical
            TabIndex        =   37
            Tag             =   "&Refresh"
            Top             =   150
            Width           =   615
         End
         Begin VB.CommandButton cmdUpdate 
            Caption         =   "&Salvar"
            Enabled         =   0   'False
            Height          =   540
            Left            =   2190
            Picture         =   "FrmProdutos.frx":05FA
            Style           =   1  'Graphical
            TabIndex        =   21
            Tag             =   "&Update"
            Top             =   135
            Width           =   615
         End
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrdProd 
         Height          =   5355
         Left            =   195
         TabIndex        =   38
         Top             =   750
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   9446
         _Version        =   393216
         Rows            =   13
         Cols            =   16
         FixedCols       =   0
         FormatString    =   $"FrmProdutos.frx":06F4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Frame Frame3 
         Caption         =   "Preços: "
         Height          =   1245
         Left            =   -74505
         TabIndex        =   26
         Top             =   1665
         Width           =   7755
         Begin VB.TextBox TxtIndice 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   2865
            TabIndex        =   9
            Text            =   "0,00"
            Top             =   375
            Width           =   495
         End
         Begin VB.TextBox TxtPrevenda3 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#.##0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   1
            EndProperty
            Height          =   315
            Left            =   795
            TabIndex        =   12
            Text            =   "0,00"
            Top             =   795
            Width           =   1245
         End
         Begin VB.TextBox TxtPrevenda4 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#.##0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   1
            EndProperty
            Height          =   315
            Left            =   4200
            TabIndex        =   13
            Text            =   "0,00"
            Top             =   780
            Width           =   1245
         End
         Begin VB.TextBox TxtPreVenda5 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#.##0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   1
            EndProperty
            Height          =   315
            Left            =   6300
            TabIndex        =   14
            Text            =   "0,00"
            Top             =   780
            Width           =   1245
         End
         Begin VB.TextBox TxtPreVenda2 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#.##0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   1
            EndProperty
            Height          =   315
            Left            =   6285
            TabIndex        =   11
            Text            =   "0,00"
            Top             =   330
            Width           =   1245
         End
         Begin VB.TextBox TxtPrevenda1 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#.##0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   1
            EndProperty
            Height          =   315
            Left            =   4200
            TabIndex        =   10
            Text            =   "0,00"
            Top             =   360
            Width           =   1245
         End
         Begin VB.TextBox TxtPcocusto 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#.##0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   1
            EndProperty
            Height          =   315
            Left            =   795
            TabIndex        =   8
            Text            =   "0,00"
            Top             =   345
            Width           =   1245
         End
         Begin VB.Label Label4 
            Caption         =   "% Lucro"
            Height          =   285
            Left            =   2175
            TabIndex        =   47
            Top             =   375
            Width           =   570
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "Venda 4"
            Height          =   195
            Left            =   3495
            TabIndex        =   41
            Top             =   825
            Width           =   600
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Venda 3"
            Height          =   195
            Left            =   135
            TabIndex        =   40
            Top             =   825
            Width           =   600
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Venda 5"
            Height          =   195
            Left            =   5610
            TabIndex        =   34
            Top             =   840
            Width           =   600
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Venda 2"
            Height          =   195
            Left            =   5580
            TabIndex        =   33
            Top             =   405
            Width           =   600
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Venda 1"
            Height          =   195
            Left            =   3510
            TabIndex        =   32
            Top             =   405
            Width           =   600
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Custo:"
            Height          =   195
            Left            =   315
            TabIndex        =   31
            Top             =   375
            Width           =   420
         End
      End
      Begin VB.TextBox TxtCodigo 
         Height          =   285
         Left            =   -73845
         MaxLength       =   6
         TabIndex        =   1
         Top             =   495
         Width           =   975
      End
      Begin VB.ComboBox CmbGrupo 
         Height          =   315
         Left            =   -73860
         TabIndex        =   3
         Top             =   885
         Width           =   4215
      End
      Begin VB.TextBox TxtDescricao 
         Height          =   315
         Left            =   -71730
         TabIndex        =   2
         Top             =   480
         Width           =   5145
      End
      Begin VB.TextBox TxtCodBarra 
         Height          =   315
         Left            =   -73860
         TabIndex        =   6
         Top             =   1290
         Width           =   1755
      End
      Begin VB.ComboBox CmbMoeda 
         Height          =   315
         Left            =   -71085
         TabIndex        =   7
         Top             =   1275
         Width           =   1980
      End
      Begin VB.Frame Frame1 
         Caption         =   "Ativo ? "
         Height          =   615
         Left            =   -68010
         TabIndex        =   19
         Top             =   855
         Width           =   1425
         Begin VB.OptionButton Optsim 
            Caption         =   "Sim"
            Height          =   345
            Index           =   1
            Left            =   90
            TabIndex        =   4
            Top             =   210
            Width           =   585
         End
         Begin VB.OptionButton OptNao 
            Caption         =   "Não"
            Height          =   375
            Index           =   1
            Left            =   690
            TabIndex        =   5
            Top             =   210
            Width           =   645
         End
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Estoque Minimo:"
         Height          =   195
         Left            =   -72855
         TabIndex        =   43
         Top             =   3060
         Width           =   1170
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Estoque Atual:"
         Height          =   195
         Left            =   -74790
         TabIndex        =   42
         Top             =   3060
         Width           =   1035
      End
      Begin VB.Label Label16 
         Caption         =   "Ultima Venda"
         Height          =   195
         Left            =   -68475
         TabIndex        =   36
         Top             =   3060
         Width           =   1065
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Ultima compra"
         Height          =   195
         Left            =   -70755
         TabIndex        =   35
         Top             =   3060
         Width           =   1005
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Código:"
         Height          =   195
         Left            =   -74580
         TabIndex        =   25
         Top             =   540
         Width           =   540
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Grupo:"
         Height          =   195
         Left            =   -74460
         TabIndex        =   24
         Top             =   945
         Width           =   480
      End
      Begin VB.Label Label3 
         Caption         =   "Descrição:"
         Height          =   255
         Left            =   -72645
         TabIndex        =   23
         Top             =   510
         Width           =   795
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Cod.Barra:"
         Height          =   195
         Left            =   -74760
         TabIndex        =   22
         Top             =   1350
         Width           =   750
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Moeda"
         Height          =   195
         Left            =   -71820
         TabIndex        =   20
         Top             =   1335
         Width           =   495
      End
   End
End
Attribute VB_Name = "FrmProdutos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim pRsProd As New ADODB.Recordset
Private lIncluir As Boolean
Private lPrimeiro As Boolean
Private pnCodProd As Double
Dim lcCodgrupo, lcCodfor, lcCodmoeda As String

Private Sub cmdAdd_Click()
   lIncluir = True
   Habilita Me
   limpa_tela Me
   gSql = "Select Max(Val(codprod)) as ultcod from tab_produtos"
   gRs.Open gSql, ConDb, adOpenKeyset
   If gRs.BOF And gRs.EOF Then
      pnCodProd = 1
   Else
      pnCodProd = gRs!ultcod + 1
   End If
   
   gRs.Close
   
   suPreencheGridFornece
   suGridForneceEspecifico
   
   cmdAdd.Enabled = False
   cmdEditar.Enabled = False
   cmdDelete.Enabled = False
   cmdUpdate.Enabled = True
   cmddesfaz.Enabled = True
   CmdSair.Enabled = False
   TxtUltCompra.Enabled = False
   TxtUltVenda.Enabled = False
   
   Me.MsflxgridForne.Enabled = True
   
   CmbGrupo.ListIndex = 0
   CmbMoeda.ListIndex = 0
   TxtIndice.Text = Format(55#, "###,00")
   TxtCodigo.Text = Format(Str(pnCodProd), "000000")
   TxtCodigo.SetFocus
End Sub

Private Sub cmdDelete_Click()
 On Error GoTo ErroDelete
    
    'this may produce an error if you delete the last
    'record or the only record in the recordset
    If MsgBox("Deseja realmente apagar este Produto ? ", vbYesNo, "Atenção") = vbYes Then
       gSql = "delete * from tab_produtos where codprod = '" & Me.TxtCodigo.Text & "'"
       ConDb.Execute gSql
       On Error GoTo 0
   
       'pRsProd.Close
       Abre_Le_rst_prod
       Carrega_Grid_prod
       pRsProd.Close
       Carrega_tela
       Desabilita Me
     End If
     Me.cmdUpdate.Enabled = False
     Me.cmddesfaz.Enabled = False
     Me.cmdEditar.Enabled = True
     Me.cmdAdd.Enabled = True
     Me.CmdSair.Enabled = True
     Me.cmdDelete.Enabled = True
     Exit Sub
     
ErroDelete:
     MsgBox "Deu erro na exclusao do Fornecedor " & Chr(13) & "Instrucao Sql = '" & _
            gSql & "'  "
End Sub

Private Sub cmddesfaz_Click()
  
  lIncluir = False
 
  'Carrega_tela
  suGridForneceEspecifico
  Desabilita Me
  'MSFlexGridprod_Click
  Me.cmdUpdate.Enabled = False
  Me.cmddesfaz.Enabled = False
  Me.cmdEditar.Enabled = True
  Me.cmdAdd.Enabled = True
  Me.CmdSair.Enabled = True
  Me.cmdDelete.Enabled = True

End Sub

Private Sub cmdEditar_Click()

   suPreencheGridFornece
   suGridForneceEspecifico
   
   Habilita Me
   cmdAdd.Enabled = False
   cmdEditar.Enabled = False
   cmdDelete.Enabled = False
   cmdUpdate.Enabled = True
   cmddesfaz.Enabled = True
   CmdSair.Enabled = False
   TxtUltCompra.Enabled = False
   TxtUltVenda.Enabled = False
   
   TxtCodigo.SetFocus
End Sub

Private Sub CmdSair_Click()
   Unload Me
End Sub

Private Sub cmdUpdate_Click()
   
   'Dim lcCodgrupo, lcCodfor, lcCodmoeda As String
   
   'pRsProd.Close
   gSql = "select codgrupo,descricao FROM tab_grupos where descricao = '" & Me.CmbGrupo.Text & "'"
   gRs.Open gSql, ConDb, adOpenKeyset
   If gRs.BOF And gRs.EOF Then
      MsgBox "Grupo não cadastrado.", vbOKOnly, "Atenção"
      Unload Me
   Else
      lcCodgrupo = gRs!codgrupo
   End If
   
   gRs.Close

   '---> Os fornecedores desse produto -> Primeiro apaga
   gSql = "DELETE  FROM sa4_prf where prf_prd = '" & Me.TxtCodigo.Text & "'"
   ConDb.Execute gSql
   
   '*---> Agora inclue
   For i = 1 To Me.MSFlxGridaux.Rows - 1
       MSFlxGridaux.Col = 0
       MSFlxGridaux.Row = i
       If Len(MSFlxGridaux.Text) > 0 Then
          gSql = "INSERT INTO sa4_prf (prf_prd,prf_for) "
          gSql = gSql & " VALUES ('" & Me.TxtCodigo.Text & "','"
          gSql = gSql & MSFlxGridaux.Text & "')"
          ConDb.Execute gSql
       End If
   Next
   
   'gRs.Close
   If CmbMoeda.ListIndex = -1 Then
      MsgBox "Favor escolher uma moeda.", vbOKOnly, "Atenção"
      CmbMoeda.SetFocus
      Exit Sub
   End If
   gSql = "select codigo,nome FROM cadmoe where nome = '" & Me.CmbMoeda.Text & "'"
   gRs.Open gSql, ConDb, adOpenKeyset
   If gRs.BOF And gRs.EOF Then
      MsgBox "Não existem moedas no cadastro. Favor cadastrar", vbOKOnly, "Atenção"
      Unload Me
   Else
      lcCodmoeda = gRs!codigo
   End If
   
   gRs.Close
    
   'On Error GoTo ErroProduto
   
   'ConDb.BeginTrans
   
   If lIncluir Then
      suIncluir_Produto
   Else
      gSql = "UPDATE tab_produtos SET descricao = '" & Me.TxtDescricao.Text & "',"
      gSql = gSql & "codgrupo = " & lcCodgrupo & ","
      gSql = gSql & "codbar = '" & IIf(Len(Me.TxtCodBarra.Text) = 0, " ", Me.TxtCodBarra.Text) & "',"
      gSql = gSql & "codigo_m = " & lcCodmoeda & ","
      gSql = gSql & "ativo = " & IIf(Me.Optsim(1) = True, True, False) & ","
      gSql = gSql & "estatual = " & Val(Me.TxtEstatual.Text) & ","
      gSql = gSql & "minimo = " & Val(Me.TxtEstminimo.Text) & ","
      gSql = gSql & "precocusto = " & Replace(IIf(Len(Me.TxtPcocusto.Text) = 0, 0, Me.TxtPcocusto.Text), ",", ".") & ","
      gSql = gSql & "indlucro = " & Replace(IIf(Len(Me.TxtIndice.Text) = 0, 0, Me.TxtIndice.Text), ",", ".") & ","
      gSql = gSql & "prevenda1 = " & Replace(IIf(Len(Me.TxtPrevenda1.Text) = 0, 0, Me.TxtPrevenda1.Text), ",", ".") & ","
      gSql = gSql & "prevenda2 = " & Replace(IIf(Len(Me.TxtPreVenda2.Text) = 0, 0, Me.TxtPreVenda2.Text), ",", ".") & ","
      gSql = gSql & "prevenda3 = " & Replace(IIf(Len(Me.TxtPrevenda3.Text) = 0, 0, Me.TxtPrevenda3.Text), ",", ".") & ","
      gSql = gSql & "prevenda4 = " & Replace(IIf(Len(Me.TxtPrevenda4.Text) = 0, 0, Me.TxtPrevenda4.Text), ",", ".") & ","
      gSql = gSql & "prevenda5 = " & Replace(IIf(Len(Me.TxtPreVenda5.Text) = 0, 0, Me.TxtPreVenda5.Text), ",", ".") & ","
      If Me.TxtUltCompra.Text <> "" Then
         gSql = gSql & "dtultcomp = Cdate('" & Me.TxtUltCompra.Text & "')"
      Else
         gSql = gSql & "dtultcomp = NULL"
      End If
      'gSql = gSql & "dtultcomp = '" & Format(Me.TxtUltCompra.Text, "dd/mm/yyyy") & "'"
      gSql = gSql & " ,operador = '" & gOperador & "', datatual = Cdate('" & Date & "')"
      gSql = gSql & " WHERE codprod = '" & Me.TxtCodigo.Text & "'"
      ConDb.Execute gSql
      
   End If
       
   'ConDb.CommitTrans
       
   Abre_Le_rst_prod
   
   Carrega_Grid_prod
   pRsProd.MoveFirst
   pRsProd.Close
   
   Carrega_tela
   
   Desabilita Me
      
   Me.cmdUpdate.Enabled = False
   Me.cmddesfaz.Enabled = False
   Me.MsflxgridForne.Enabled = False
   Me.cmdEditar.Enabled = True
   Me.cmdAdd.Enabled = True
   Me.CmdSair.Enabled = True
   Me.cmdDelete.Enabled = True
   Exit Sub
ErroProduto:
     
   'ConDb.RollbackTrans
   If lIncluir Then
      MsgBox " Erro na inclusao do Produto = " & ConDb.State, vbOKOnly, "Atenção"
   Else
      MsgBox " Erro na Alteração do Produto = " & ConDb.State, vbOKOnly, "Atenção"
   End If
End Sub

Private Sub Form_Activate()
 'Procedimento para carregar as grids dos tab_prod, fornecedores e grupos
   
   Abre_Le_rst_prod
   
   If pRsProd.BOF And pRsProd.EOF Then
      If MsgBox("Arquivo vazio. Incluir dados agora ?", vbYesNo, "Atenção ") = vbYes Then
         'suIncluir_Produto
         pRsProd.Close
         'Abre_Le_rst_prod
         Abre_Le_rst_grupos
         Abre_Le_rst_moeda
         lPrimeiro = True
         lIncluir = True
         cmdEditar_Click
         Exit Sub
      Else
         'Desabilita Me
         Unload Me
      End If
   Else
      pRsProd.MoveFirst
      'Carrega_tela
    
      Desabilita Me
      'lIncluir = False
      'lPrimeiro = False
   End If
   
   suCarrega_Grids
   pRsProd.Close
      
   Me.SSTab1.Tab = 0

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub Form_Load()
   
   Call Centra(Me)
   
End Sub
Private Sub suAchaFornece()

   'Acha os fornecedores
   Me.MSFlexGrdProd.Col = 0
   gSql = "select sa4_prf.prf_for as codfor,tab_fornece.nome"
   gSql = gSql & " FROM sa4_prf,tab_fornece "
   gSql = gSql & " Where sa4_prf.prf_prd = '" & Me.TxtCodigo.Text & "'"
   gSql = gSql & " AND tab_fornece.codfor = Val(sa4_prf.prf_for)"
   gRs.Open gSql, ConDb, adOpenKeyset
      
   MsflxgridForne.Rows = 1
   
   If Not gRs.EOF And Not gRs.BOF Then
      
      MsflxgridForne.Rows = 1
          
      i = 0
      Do While Not gRs.EOF
        MsflxgridForne.Rows = MsflxgridForne.Rows + 1
        MsflxgridForne.Row = MsflxgridForne.Rows - 1
        Me.MsflxgridForne.Col = 0: Me.MsflxgridForne.Text = f_nulo(gRs!codfor, "0")
        Me.MsflxgridForne.Col = 1: Me.MsflxgridForne.Text = f_nulo(gRs!nome, "Nao Cadastrado")
        gRs.MoveNext
        
      Loop
      MsflxgridForne.FixedRows = 1
   End If
   gRs.Close
End Sub

Private Sub suGridForneceEspecifico()
   
   MsflxgridForne.Redraw = True
   MsflxgridForne.Col = 0
   MSFlxGridaux.Clear
   MSFlxGridaux.Redraw = True
   MSFlxGridaux.Col = 0
   MSFlxGridaux.Rows = 0
   
   For i = 1 To MsflxgridForne.Rows - 1
      MsflxgridForne.Row = i
      MsflxgridForne.Col = 0
      gSql = "SELECT * FROM SA4_PRF where prf_for =  '" & Format(MsflxgridForne.Text, "000000") & "'"
      gSql = gSql & " AND prf_prd = '" & Me.TxtCodigo & "'"
      gRs.Open gSql, ConDb, adOpenKeyset
   
      If gRs.BOF And gRs.EOF Then
      Else
         
         MsflxgridForne.Col = 0:   MsflxgridForne.CellBackColor = vbYellow
         MSFlxGridaux.Rows = MSFlxGridaux.Rows + 1
         MSFlxGridaux.Row = MSFlxGridaux.Rows - 1
         MSFlxGridaux.AddItem MsflxgridForne.Text
         
         MsflxgridForne.Col = 1:   MsflxgridForne.CellBackColor = vbYellow
         
      End If
      gRs.Close
   Next
   
   MsflxgridForne.Refresh
   
End Sub
Private Sub suPreencheGridFornece()

'Preenche os fornecedores
   Me.MSFlexGrdProd.Col = 0
   gSql = "select  codfor,nome"
   gSql = gSql & " FROM tab_fornece "
   gRs.Open gSql, ConDb, adOpenKeyset
   If Not gRs.EOF And Not gRs.BOF Then
      MsflxgridForne.Rows = 1
      i = 0
      MsflxgridForne.Redraw = False
      Do While Not gRs.EOF
         MsflxgridForne.Rows = MsflxgridForne.Rows + 1
         MsflxgridForne.Row = MsflxgridForne.Rows - 1
         Me.MsflxgridForne.Col = 0: Me.MsflxgridForne.Text = f_nulo(gRs!codfor, "0")
         Me.MsflxgridForne.Col = 1: Me.MsflxgridForne.Text = f_nulo(gRs!nome, "Nao Cadastrado")
         gRs.MoveNext
      Loop
      MsflxgridForne.Redraw = True
      MsflxgridForne.FixedRows = 1
   End If
   gRs.Close
  
End Sub

Private Sub suCarrega_Grids()
   'Abre_Le_rst_prod
   Carrega_Grid_prod
   Abre_Le_rst_grupos
   Abre_Le_rst_moeda
End Sub

Private Sub Abre_Le_rst_prod()
   gSql = "select tab_produtos.codprod,tab_grupos.descricao as nome_grupo,tab_produtos.descricao as desprod, "
   gSql = gSql & " estinicial,estatual,minimo,precocusto,indlucro,prevenda1,prevenda2,prevenda3,prevenda4,prevenda5,dtultcomp,dtultvenda,codigo_m "
   gSql = gSql & " FROM tab_produtos,tab_grupos "
   gSql = gSql & " Where tab_produtos.codgrupo = tab_grupos.codgrupo "
   pRsProd.Open gSql, ConDb, adOpenKeyset
    
End Sub

Private Sub Abre_Le_rst_grupos()
   gSql = "select codgrupo,descricao FROM tab_grupos ORDER BY descricao"
   gRs.Open gSql, ConDb, adOpenKeyset
   If gRs.BOF And gRs.EOF Then
      MsgBox "Não existem grupos no cadastro. Favor cadastrar", vbOKOnly, "Atenção"
      Unload Me
   End If
   
   Carrega_combo_grupos
   gRs.Close
End Sub

Private Sub Abre_Le_rst_moeda()
   
   gSql = "select codigo,nome FROM cadmoe"
   gRs.Open gSql, ConDb, adOpenKeyset
   If gRs.BOF And gRs.EOF Then
      MsgBox "Não existem moedas no cadastro. Favor cadastrar", vbOKOnly, "Atenção"
      Unload Me
   End If
   Carrega_combo_moeda
   gRs.Close
End Sub

Private Sub Carrega_Grid_prod()

  'Teste do MsHFlexgrid1 - eh eh eh
  MSFlexGrdProd.Row = 0
  MSFlexGrdProd.FontWidth = 1
  
  With pRsProd
      '.MoveLast
      'nItem = .RecordCount
      .MoveFirst
      MSFlexGrdProd.Rows = 1
      i = 0
      MSFlexGrdProd.Redraw = False
      Do While Not .EOF
        MSFlexGrdProd.Rows = MSFlexGrdProd.Rows + 1
        MSFlexGrdProd.Row = MSFlexGrdProd.Rows - 1
        MSFlexGrdProd.Col = 0: MSFlexGrdProd.Text = "" & !codprod
        MSFlexGrdProd.Col = 1: MSFlexGrdProd.Text = "" & !desprod
        MSFlexGrdProd.Col = 2: MSFlexGrdProd.Text = "" & !nome_grupo
        MSFlexGrdProd.Col = 3: MSFlexGrdProd.Text = f_nulo(!estinicial, 0)
        MSFlexGrdProd.Col = 4: MSFlexGrdProd.Text = f_nulo(!estatual, 0)
        MSFlexGrdProd.Col = 5: MSFlexGrdProd.Text = f_nulo(!minimo, 0)
        MSFlexGrdProd.Col = 6: MSFlexGrdProd.Text = Format(f_nulo(!precocusto, 0), "###,###,##0.00")
        MSFlexGrdProd.Col = 7: MSFlexGrdProd.Text = Format(f_nulo(!indlucro, 0), "##0.00")
        MSFlexGrdProd.Col = 8: MSFlexGrdProd.Text = Format(f_nulo(!prevenda1, 0), "###,###,##0.00")
        MSFlexGrdProd.Col = 9: MSFlexGrdProd.Text = Format(f_nulo(!prevenda2, 0), "###,###,##0.00")
        MSFlexGrdProd.Col = 10: MSFlexGrdProd.Text = Format(f_nulo(!prevenda3, 0), "###,###,##0.00")
        MSFlexGrdProd.Col = 11: MSFlexGrdProd.Text = Format(f_nulo(!prevenda4, 0), "###,###,##0.00")
        MSFlexGrdProd.Col = 12: MSFlexGrdProd.Text = Format(f_nulo(!prevenda5, 0), "###,###,##0.00")
        MSFlexGrdProd.Col = 13: MSFlexGrdProd.Text = IIf(IsNull(!dtultcomp), Format("", "dd/mm/yyyy"), !dtultcomp)
        MSFlexGrdProd.Col = 14: MSFlexGrdProd.Text = IIf(IsNull(!dtultvenda), Format("", "dd/mm/yyyy"), !dtultvenda)
        MSFlexGrdProd.Col = 15: MSFlexGrdProd.Text = "" & !codigo_m
        
        .MoveNext
         
       Loop
      MSFlexGrdProd.Redraw = True
      MSFlexGrdProd.FixedRows = 1
          
  End With

  End Sub

Public Sub Carrega_combo_grupos()
 With gRs
      '.MoveLast
      'nItem = .RecordCount
      .MoveFirst
      Do While Not .EOF
        CmbGrupo.AddItem (gRs!descricao)
        CmbGrupo.ItemData(CmbGrupo.NewIndex) = gRs!codgrupo
        .MoveNext
      Loop
  End With
  CmbGrupo.ListIndex = -1
End Sub
Public Sub Carrega_combo_moeda()
 With gRs
      '.MoveLast
      'nItem = .RecordCount
      .MoveFirst
      Do While Not .EOF
        CmbMoeda.AddItem (gRs!nome)
        CmbMoeda.ItemData(CmbMoeda.NewIndex) = gRs!codigo
        .MoveNext
      Loop
  End With
  CmbMoeda.ListIndex = -1
End Sub

Private Sub Carrega_tela()
   'Limpa as variaveis da tela se caso ficarem com dados da outra tela
   limpa_tela Me
   'Carrega a tela com os dados do registro
   Me.MSFlexGrdProd.Col = 0
   Me.TxtCodigo.Text = Me.MSFlexGrdProd.Text
   Me.MSFlexGrdProd.Col = 1
   Me.TxtDescricao.Text = Me.MSFlexGrdProd.Text
   
   '---> Acha os Fornecedores
   suAchaFornece
   
   'Acha o grupo para por no combo
   Me.MSFlexGrdProd.Col = 2
   gSql = "select codgrupo,descricao "
   gSql = gSql & " FROM tab_grupos "
   gSql = gSql & " Where descricao = '" & Me.MSFlexGrdProd.Text & "'"
   gRs.Open gSql, ConDb, adOpenKeyset
   If Not gRs.EOF And Not gRs.BOF Then
      For i = 0 To CmbGrupo.ListCount - 1
         If CmbGrupo.ItemData(i) = gRs!codgrupo Then
            CmbGrupo.ListIndex = i
            Exit For
         End If
      Next
   Else
      CmbGrupo.ListIndex = -1
   End If
   gRs.Close
      
   'Acha a moeda
   Me.MSFlexGrdProd.Col = Me.MSFlexGrdProd.Cols - 1
   gSql = "select codigo,nome "
   gSql = gSql & " FROM cadmoe "
   gSql = gSql & " Where cadmoe.codigo = " & Val(Me.MSFlexGrdProd.Text)
   gRs.Open gSql, ConDb, adOpenKeyset
   If Not gRs.EOF And Not gRs.BOF Then
      For i = 0 To CmbMoeda.ListCount
         If CmbMoeda.ItemData(i) = gRs!codigo Then
            CmbMoeda.ListIndex = i
            Exit For
         End If
      Next
   Else
      CmbMoeda.ListIndex = -1
   End If
   gRs.Close
   'pRsProd.Close
   
   Desabilita Me
   gSql = "select * "
   gSql = gSql & " FROM tab_produtos"
   gSql = gSql & " Where tab_produtos.codprod = '" & Me.TxtCodigo.Text & "'"
   pRsProd.Open gSql, ConDb, adOpenKeyset
   If Not pRsProd.EOF And Not pRsProd.BOF Then
      Me.TxtCodBarra.Text = "" & pRsProd!codbar
      If pRsProd!ativo = "S" Then
         Me.Optsim.item(1).Value = 1
         Me.OptNao.item(1).Value = 0
      Else
         Me.Optsim.item(1).Value = 0
         Me.OptNao.item(1).Value = 1
      End If
      'Me.TxtEstinicial.Text = IIf(IsNull(pRsProd!estinicial), 0, pRsProd!estinicial)
      Me.TxtEstatual.Text = IIf(IsNull(pRsProd!estatual), 0, pRsProd!estatual)
      Me.TxtEstminimo.Text = IIf(IsNull(pRsProd!minimo), 0, pRsProd!minimo)
      Me.TxtPcocusto.Text = Format(IIf(IsNull(pRsProd!precocusto), 0, pRsProd!precocusto), "###,###,##0.00")
      Me.TxtIndice.Text = Format(IIf(IsNull(pRsProd!indlucro), 0, pRsProd!indlucro), "##0.00")
      Me.TxtPrevenda1.Text = Format(IIf(IsNull(pRsProd!prevenda1), 0, pRsProd!prevenda1), "###,###,##0.00")
      Me.TxtPreVenda2.Text = Format(IIf(IsNull(pRsProd!prevenda2), 0, pRsProd!prevenda2), "###,###,##0.00")
      Me.TxtPrevenda3.Text = Format(IIf(IsNull(pRsProd!prevenda3), 0, pRsProd!prevenda1), "###,###,##0.00")
      Me.TxtPrevenda4.Text = Format(IIf(IsNull(pRsProd!prevenda4), 0, pRsProd!prevenda2), "###,###,##0.00")
      Me.TxtPreVenda5.Text = Format(IIf(IsNull(pRsProd!prevenda5), 0, pRsProd!prevenda1), "###,###,##0.00")
      Me.TxtUltVenda.Text = "" & pRsProd!dtultvenda
      Me.TxtUltCompra.Text = "" & pRsProd!dtultcomp
   Else
      MsgBox "Erro grave. Produto nao encontrado ", vbOKOnly, "Atenção"
      Unload Me
      End
   End If
   pRsProd.Close
   
   
End Sub
 
Private Sub MSFlexGrdProd_Click()
  
  Dim oldrow As Long
  Dim lcColGrid As Double
  
  If MSFlexGrdProd.Row = 1 Then
     lcColGrid = MSFlexGrdProd.Col
     MSFlexGrdProd.Col = lcColGrid
     MSFlexGrdProd.Sort = flexSortStringAscending
  End If
 
  oldrow = MSFlexGrdProd.Row
  
  MSFlexGrdProd.Row = 0
  
  With MSFlexGrdProd
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
    .Col = 6:   .CellBackColor = vbYellow
    .Col = 7:   .CellBackColor = vbYellow
    .Col = 8:   .CellBackColor = vbYellow
    .Col = 9:   .CellBackColor = vbYellow
    .Col = 10:  .CellBackColor = vbYellow
    .Col = 11:  .CellBackColor = vbYellow
    .Col = 12:  .CellBackColor = vbYellow
    .Col = 13:  .CellBackColor = vbYellow
    .Col = 14:  .CellBackColor = vbYellow
    .Col = 15:  .CellBackColor = vbYellow
    .TopRow = .Row
     
    '.Refresh
   
End With


End Sub

Private Sub MsflxgridForne_Click()
Dim Jatem
Jatem = False
   MSFlxGridaux.Redraw = True
   MSFlxGridaux.Col = 0
   MsflxgridForne.Col = 0
   For i = 1 To MSFlxGridaux.Rows - 1
       If MsflxgridForne.Text = MSFlxGridaux.Text Then
          Jatem = True
          Exit For
       End If
   Next
   If Not Jatem Then
      MSFlxGridaux.Rows = MSFlxGridaux.Rows + 1
      MSFlxGridaux.Row = MSFlxGridaux.Rows - 1
      MSFlxGridaux.Text = MsflxgridForne.Text
      MSFlxGridaux.Refresh
   End If
     
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
   If SSTab1.Tab > 0 Then
      Carrega_tela
   End If
   
End Sub

Private Sub suIncluir_Produto()
   gSql = "INSERT INTO tab_produtos (codprod, descricao, codgrupo,"
   gSql = gSql & " codbar, codigo_m, ativo, estatual, minimo,"
   gSql = gSql & "precocusto, indlucro,prevenda1, prevenda2, prevenda3,"
   gSql = gSql & "prevenda4,prevenda5,dtultcomp,dtultvenda, operador, datatual ) "
   gSql = gSql & "VALUES ('" & Me.TxtCodigo.Text & "','"
   gSql = gSql & Me.TxtDescricao.Text & "',"
   gSql = gSql & lcCodgrupo & ",'"
   gSql = gSql & Me.TxtCodBarra.Text & "',"
   gSql = gSql & lcCodmoeda & ","
   gSql = gSql & IIf(Me.Optsim(1) = True, True, False) & ","
   gSql = gSql & Val(Me.TxtEstatual.Text) & ","
   gSql = gSql & Val(Me.TxtEstminimo.Text) & ","
   gSql = gSql & Replace(IIf(Len(Me.TxtPcocusto.Text) = 0, 0, Me.TxtPcocusto.Text), ",", ".") & ","
   gSql = gSql & Replace(IIf(Len(Me.TxtIndice.Text) = 0, 0, Me.TxtIndice.Text), ",", ".") & ","
   gSql = gSql & Replace(IIf(Len(Me.TxtPrevenda1.Text) = 0, 0, Me.TxtPrevenda1.Text), ",", ".") & ","
   gSql = gSql & Replace(IIf(Len(Me.TxtPreVenda2.Text) = 0, 0, Me.TxtPreVenda2.Text), ",", ".") & ","
   gSql = gSql & Replace(IIf(Len(Me.TxtPrevenda3.Text) = 0, 0, Me.TxtPrevenda3.Text), ",", ".") & ","
   gSql = gSql & Replace(IIf(Len(Me.TxtPrevenda4.Text) = 0, 0, Me.TxtPrevenda4.Text), ",", ".") & ","
   gSql = gSql & Replace(IIf(Len(Me.TxtPreVenda5.Text) = 0, 0, Me.TxtPreVenda5.Text), ",", ".") & ","
   If Me.TxtUltCompra.Text <> "" Then
      gSql = gSql & "'" & CDate(Me.TxtUltCompra.Text) & "',"
   Else
      gSql = gSql & "NULL,"
   End If
   If Me.TxtUltVenda.Text <> "" Then
      gSql = gSql & "'" & CDate(Me.TxtUltVenda.Text) & "',"
   Else
      gSql = gSql & "NULL,"
   End If
   gSql = gSql & "'" & gOperador & "',Cdate('" & Date & "'))"
   ConDb.Execute gSql
   lIncluir = False
   

End Sub

Private Sub TxtEstatual_GotFocus()
 With TxtEstatual
      .SelStart = 0
      .SelLength = Len(.Text)
   End With

End Sub


Private Sub TxtEstminimo_GotFocus()
 With TxtEstminimo
      .SelStart = 0
      .SelLength = Len(.Text)
   End With

End Sub

Private Sub TxtIndice_GotFocus()
 With TxtIndice
      .SelStart = 0
      .SelLength = Len(.Text)
   End With


End Sub

Private Sub TxtIndice_LostFocus()
    TxtIndice.Text = Format(TxtIndice.Text, "###,###,##0.00")
    TxtPrevenda1.Text = Format(Val(TxtPcocusto.Text) + (Val(TxtPcocusto.Text) * Val(TxtIndice.Text) / 100), "###,###,##0.00")
End Sub

Private Sub TxtPcocusto_GotFocus()
 With TxtPcocusto
      .SelStart = 0
      .SelLength = Len(.Text)
   End With

End Sub

Private Sub TxtPcocusto_LostFocus()
    TxtPcocusto.Text = Format(TxtPcocusto.Text, "###,###,##0.00")
End Sub

Private Sub TxtPrevenda1_GotFocus()
 With TxtPrevenda1
      .SelStart = 0
      .SelLength = Len(.Text)
   End With

End Sub

Private Sub TxtPrevenda1_LostFocus()
   TxtPrevenda1.Text = Format(TxtPrevenda1.Text, "#0.00")
End Sub

Private Sub TxtPrevenda2_GotFocus()
 With TxtPreVenda2
      .SelStart = 0
      .SelLength = Len(.Text)
   End With

End Sub

Private Sub TxtPrevenda2_LostFocus()
   TxtPreVenda2.Text = Format(TxtPreVenda2.Text, "#0.00")
End Sub

Private Sub TxtPrevenda3_GotFocus()
 With TxtPrevenda3
      .SelStart = 0
      .SelLength = Len(.Text)
   End With

End Sub

Private Sub TxtPrevenda3_LostFocus()
   TxtPrevenda3.Text = Format(TxtPrevenda3.Text, "#0.00")
End Sub

Private Sub TxtPrevenda4_GotFocus()
 With TxtPrevenda4
      .SelStart = 0
      .SelLength = Len(.Text)
   End With

End Sub

Private Sub TxtPrevenda4_LostFocus()
   TxtPrevenda4.Text = Format(TxtPrevenda4.Text, "#0.00")
End Sub

Private Sub TxtPrevenda5_GotFocus()
 With TxtPreVenda5
      .SelStart = 0
      .SelLength = Len(.Text)
   End With

End Sub

Private Sub TxtPrevenda5_LostFocus()
   TxtPreVenda5.Text = Format(TxtPreVenda5.Text, "#0.00")
End Sub


