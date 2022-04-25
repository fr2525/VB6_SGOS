VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form FrmClientes 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5985
   ClientLeft      =   1980
   ClientTop       =   1350
   ClientWidth     =   8220
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5985
   ScaleWidth      =   8220
   ShowInTaskbar   =   0   'False
   Tag             =   "cadvend"
   Begin TabDlg.SSTab SSTab1 
      Height          =   5775
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   10186
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Lista"
      TabPicture(0)   =   "FrmClientes.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "MSFlexGrid1"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Detalhe"
      TabPicture(1)   =   "FrmClientes.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "lblCep"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label1"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "LblUltimacompra"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label11"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "LblAniver"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Label9"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "LblCelular"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "LblTelefone"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Lblrg"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "Lblcgc"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "LblUf"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "LblCidade"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "LblBairro"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "LblEnder"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "LblCodclie"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "LblNome(1)"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "lblLabels(0)"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "Label4"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "Label5"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "Label12"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).Control(21)=   "Txtcep"
      Tab(1).Control(21).Enabled=   0   'False
      Tab(1).Control(22)=   "TxtUltimaCompra"
      Tab(1).Control(22).Enabled=   0   'False
      Tab(1).Control(23)=   "TxtContato"
      Tab(1).Control(23).Enabled=   0   'False
      Tab(1).Control(24)=   "TxtEmail"
      Tab(1).Control(24).Enabled=   0   'False
      Tab(1).Control(25)=   "TxtAnoAniver"
      Tab(1).Control(25).Enabled=   0   'False
      Tab(1).Control(26)=   "TxtMesAniver"
      Tab(1).Control(26).Enabled=   0   'False
      Tab(1).Control(27)=   "TxtDiaAniver"
      Tab(1).Control(27).Enabled=   0   'False
      Tab(1).Control(28)=   "TxtCelular"
      Tab(1).Control(28).Enabled=   0   'False
      Tab(1).Control(29)=   "TxtTelefone"
      Tab(1).Control(29).Enabled=   0   'False
      Tab(1).Control(30)=   "TxtRG"
      Tab(1).Control(30).Enabled=   0   'False
      Tab(1).Control(31)=   "Txtcgc_cpf"
      Tab(1).Control(31).Enabled=   0   'False
      Tab(1).Control(32)=   "TxtUf"
      Tab(1).Control(32).Enabled=   0   'False
      Tab(1).Control(33)=   "TxtCidade"
      Tab(1).Control(33).Enabled=   0   'False
      Tab(1).Control(34)=   "TxtBairro"
      Tab(1).Control(34).Enabled=   0   'False
      Tab(1).Control(35)=   "TxtEndereco"
      Tab(1).Control(35).Enabled=   0   'False
      Tab(1).Control(36)=   "TxtNome"
      Tab(1).Control(36).Enabled=   0   'False
      Tab(1).Control(37)=   "Frame1"
      Tab(1).Control(37).Enabled=   0   'False
      Tab(1).Control(38)=   "Frame2"
      Tab(1).Control(38).Enabled=   0   'False
      Tab(1).Control(39)=   "TxtLimite"
      Tab(1).Control(39).Enabled=   0   'False
      Tab(1).Control(40)=   "TxtSaldo"
      Tab(1).Control(40).Enabled=   0   'False
      Tab(1).Control(41)=   "ChkNegativo"
      Tab(1).Control(41).Enabled=   0   'False
      Tab(1).Control(42)=   "TxtInsc_est"
      Tab(1).Control(42).Enabled=   0   'False
      Tab(1).ControlCount=   43
      Begin VB.TextBox TxtInsc_est 
         Height          =   285
         Left            =   5925
         TabIndex        =   9
         Top             =   1545
         Width           =   1860
      End
      Begin VB.CheckBox ChkNegativo 
         Alignment       =   1  'Right Justify
         Caption         =   "Neg.?"
         Height          =   255
         Left            =   7020
         TabIndex        =   20
         Top             =   2565
         Width           =   750
      End
      Begin VB.TextBox TxtSaldo 
         Height          =   285
         Left            =   5715
         TabIndex        =   19
         Top             =   2565
         Width           =   1245
      End
      Begin VB.TextBox TxtLimite 
         Height          =   285
         Left            =   3420
         TabIndex        =   18
         Top             =   2580
         Width           =   1545
      End
      Begin VB.Frame Frame2 
         Caption         =   "Cobrança"
         Height          =   1455
         Left            =   555
         TabIndex        =   52
         Top             =   3075
         Width           =   6810
         Begin VB.TextBox TxtCepcobra 
            Height          =   285
            Left            =   2160
            TabIndex        =   25
            Top             =   1005
            Width           =   990
         End
         Begin VB.TextBox TxtUFCobra 
            Height          =   285
            Left            =   1020
            TabIndex        =   24
            Top             =   1005
            Width           =   450
         End
         Begin VB.TextBox TxtCidaCobra 
            Height          =   285
            Left            =   3960
            TabIndex        =   23
            Top             =   645
            Width           =   2460
         End
         Begin VB.TextBox TxtBairCobra 
            Height          =   285
            Left            =   1020
            TabIndex        =   22
            Top             =   630
            Width           =   2130
         End
         Begin VB.TextBox TxtEndCobra 
            Height          =   285
            Left            =   1035
            TabIndex        =   21
            Top             =   255
            Width           =   5385
         End
         Begin VB.Label Label10 
            Caption         =   "CEP"
            Height          =   225
            Left            =   1695
            TabIndex        =   59
            Top             =   1035
            Width           =   465
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Estado"
            Height          =   195
            Left            =   315
            TabIndex        =   58
            Top             =   1020
            Width           =   495
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Cidade"
            Height          =   195
            Left            =   3285
            TabIndex        =   57
            Top             =   690
            Width           =   495
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Bairro"
            Height          =   195
            Left            =   420
            TabIndex        =   56
            Top             =   705
            Width           =   405
         End
         Begin VB.Label Label3 
            Caption         =   "Endereço"
            Height          =   195
            Left            =   165
            TabIndex        =   53
            Top             =   300
            Width           =   975
         End
      End
      Begin VB.Frame Frame1 
         Height          =   735
         Left            =   1695
         TabIndex        =   45
         Top             =   4710
         Width           =   4245
         Begin VB.CommandButton cmdUpdate 
            Caption         =   "&Salvar"
            Enabled         =   0   'False
            Height          =   540
            Left            =   2190
            Picture         =   "FrmClientes.frx":0038
            Style           =   1  'Graphical
            TabIndex        =   46
            Tag             =   "&Update"
            Top             =   150
            Width           =   615
         End
         Begin VB.CommandButton cmdEditar 
            Caption         =   "&Editar"
            Height          =   540
            Left            =   795
            Picture         =   "FrmClientes.frx":0132
            Style           =   1  'Graphical
            TabIndex        =   51
            Tag             =   "&Refresh"
            Top             =   150
            Width           =   615
         End
         Begin VB.CommandButton cmdDelete 
            Caption         =   "&Excluir"
            Height          =   540
            Left            =   1485
            Picture         =   "FrmClientes.frx":02A4
            Style           =   1  'Graphical
            TabIndex        =   50
            Tag             =   "&Delete"
            Top             =   150
            Width           =   615
         End
         Begin VB.CommandButton cmdAdd 
            Caption         =   "&Incluir"
            Height          =   540
            Left            =   120
            Picture         =   "FrmClientes.frx":0416
            Style           =   1  'Graphical
            TabIndex        =   49
            Tag             =   "&Add"
            Top             =   120
            Width           =   615
         End
         Begin VB.CommandButton CmdSair 
            Caption         =   "&Sair"
            Height          =   540
            Left            =   3555
            Picture         =   "FrmClientes.frx":0500
            Style           =   1  'Graphical
            TabIndex        =   48
            Tag             =   "&Update"
            Top             =   150
            Width           =   615
         End
         Begin VB.CommandButton cmddesfaz 
            Caption         =   "&Desfaz"
            Enabled         =   0   'False
            Height          =   540
            Left            =   2865
            Picture         =   "FrmClientes.frx":05FA
            Style           =   1  'Graphical
            TabIndex        =   47
            Tag             =   "&Update"
            Top             =   150
            Width           =   615
         End
      End
      Begin VB.TextBox TxtNome 
         Height          =   285
         Left            =   990
         TabIndex        =   1
         Top             =   525
         Width           =   3600
      End
      Begin VB.TextBox TxtEndereco 
         Height          =   285
         Left            =   975
         TabIndex        =   2
         Top             =   840
         Width           =   3600
      End
      Begin VB.TextBox TxtBairro 
         Height          =   285
         Left            =   5385
         TabIndex        =   3
         Top             =   855
         Width           =   2400
      End
      Begin VB.TextBox TxtCidade 
         Height          =   285
         Left            =   960
         TabIndex        =   4
         Top             =   1185
         Width           =   2280
      End
      Begin VB.TextBox TxtUf 
         Height          =   285
         Left            =   5370
         TabIndex        =   5
         Top             =   1185
         Width           =   450
      End
      Begin VB.TextBox Txtcgc_cpf 
         Height          =   285
         Left            =   960
         TabIndex        =   7
         Top             =   1545
         Width           =   1740
      End
      Begin VB.TextBox TxtRG 
         Height          =   285
         Left            =   3585
         TabIndex        =   8
         Top             =   1545
         Width           =   1320
      End
      Begin VB.TextBox TxtTelefone 
         Height          =   285
         Left            =   960
         TabIndex        =   10
         Top             =   1875
         Width           =   1680
      End
      Begin VB.TextBox TxtCelular 
         Height          =   285
         Left            =   3585
         TabIndex        =   11
         Top             =   1905
         Width           =   1320
      End
      Begin VB.TextBox TxtDiaAniver 
         Height          =   285
         Left            =   960
         TabIndex        =   15
         Top             =   2580
         Width           =   285
      End
      Begin VB.TextBox TxtMesAniver 
         Height          =   285
         Left            =   1425
         TabIndex        =   16
         Top             =   2580
         Width           =   285
      End
      Begin VB.TextBox TxtAnoAniver 
         Height          =   285
         Left            =   1935
         TabIndex        =   17
         Top             =   2580
         Width           =   285
      End
      Begin VB.TextBox TxtEmail 
         Enabled         =   0   'False
         Height          =   285
         Left            =   960
         MaxLength       =   50
         TabIndex        =   13
         Top             =   2235
         Width           =   2925
      End
      Begin VB.TextBox TxtContato 
         Height          =   285
         Left            =   5715
         MaxLength       =   50
         TabIndex        =   14
         Top             =   2235
         Width           =   2055
      End
      Begin MSMask.MaskEdBox TxtUltimaCompra 
         Height          =   285
         Left            =   6690
         TabIndex        =   12
         Top             =   1905
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Txtcep 
         Height          =   285
         Left            =   6600
         TabIndex        =   6
         Top             =   1185
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   9
         Format          =   "#####-###"
         PromptChar      =   "_"
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   4065
         Left            =   -74760
         TabIndex        =   44
         Top             =   870
         Width           =   7500
         _ExtentX        =   13229
         _ExtentY        =   7170
         _Version        =   393216
         Rows            =   5
         Cols            =   10
         FixedCols       =   0
         FormatString    =   $"FrmClientes.frx":06F4
      End
      Begin VB.Label Label12 
         Caption         =   "Insc.Est."
         Height          =   315
         Left            =   5115
         TabIndex        =   60
         Top             =   1560
         Width           =   825
      End
      Begin VB.Label Label5 
         Caption         =   "Saldo"
         Height          =   195
         Left            =   5130
         TabIndex        =   55
         Top             =   2625
         Width           =   735
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Limite"
         Height          =   195
         Left            =   2835
         TabIndex        =   54
         Top             =   2640
         Width           =   405
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Código:"
         Height          =   195
         Index           =   0
         Left            =   6405
         TabIndex        =   43
         Tag             =   "CODVEND:"
         Top             =   540
         Width           =   540
      End
      Begin VB.Label LblNome 
         Alignment       =   1  'Right Justify
         Caption         =   "Nome:"
         Height          =   255
         Index           =   1
         Left            =   345
         TabIndex        =   42
         Tag             =   "NOME:"
         Top             =   555
         Width           =   555
      End
      Begin VB.Label LblCodclie 
         Caption         =   "codclie"
         ForeColor       =   &H000000FF&
         Height          =   270
         Left            =   7200
         TabIndex        =   41
         Top             =   540
         Width           =   615
      End
      Begin VB.Label LblEnder 
         Caption         =   "Endereço"
         Height          =   210
         Left            =   105
         TabIndex        =   40
         Top             =   885
         Width           =   750
      End
      Begin VB.Label LblBairro 
         Caption         =   "Bairro"
         Height          =   240
         Left            =   4740
         TabIndex        =   39
         Top             =   885
         Width           =   420
      End
      Begin VB.Label LblCidade 
         Caption         =   "Cidade"
         Height          =   240
         Left            =   285
         TabIndex        =   38
         Top             =   1215
         Width           =   540
      End
      Begin VB.Label LblUf 
         Caption         =   "Estado"
         Height          =   195
         Left            =   4620
         TabIndex        =   37
         Top             =   1245
         Width           =   555
      End
      Begin VB.Label Lblcgc 
         Caption         =   "CPF/CNPJ"
         Height          =   210
         Left            =   135
         TabIndex        =   36
         Top             =   1575
         Width           =   720
      End
      Begin VB.Label Lblrg 
         Caption         =   "R.G."
         Height          =   210
         Left            =   3105
         TabIndex        =   35
         Top             =   1575
         Width           =   360
      End
      Begin VB.Label LblTelefone 
         Caption         =   "Telefone"
         Height          =   240
         Left            =   210
         TabIndex        =   34
         Top             =   1905
         Width           =   630
      End
      Begin VB.Label LblCelular 
         Caption         =   "Celular"
         Height          =   195
         Left            =   2925
         TabIndex        =   33
         Top             =   1935
         Width           =   570
      End
      Begin VB.Label Label9 
         Caption         =   "/"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   1275
         TabIndex        =   32
         Top             =   2640
         Width           =   90
      End
      Begin VB.Label LblAniver 
         Caption         =   "Aniversário"
         Height          =   240
         Left            =   75
         TabIndex        =   31
         Top             =   2610
         Width           =   810
      End
      Begin VB.Label Label11 
         Caption         =   "/"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   1770
         TabIndex        =   30
         Top             =   2610
         Width           =   90
      End
      Begin VB.Label LblUltimacompra 
         Caption         =   "Ultima compra"
         Height          =   240
         Left            =   5430
         TabIndex        =   29
         Top             =   1935
         Width           =   1125
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "e-mail"
         Height          =   285
         Left            =   315
         TabIndex        =   28
         Top             =   2205
         Width           =   525
      End
      Begin VB.Label lblCep 
         Alignment       =   1  'Right Justify
         Caption         =   "CEP"
         Height          =   255
         Left            =   6030
         TabIndex        =   27
         Top             =   1245
         Width           =   465
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Contato"
         Height          =   195
         Left            =   5055
         TabIndex        =   26
         Top             =   2265
         Width           =   555
      End
   End
End
Attribute VB_Name = "FrmClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private prsCliente As New ADODB.Recordset
'Private pQd As QueryDef

Private lIncluir As Boolean
Private lPrimeiro As Boolean

Private Sub Carrega_tela()
   limpa_tela Me
   Me.LblCodclie = gRs("idCli")
   Me.TxtNome.text = gRs("nome")
   If Not IsNull(gRs("endereco")) Then Me.TxtEndereco.text = gRs("endereco")
   If Not IsNull(gRs("bairro")) Then Me.TxtBairro.text = gRs("bairro")
   If Not IsNull(gRs("Cidade")) Then Me.TxtCidade.text = gRs("Cidade")
   If Not IsNull(gRs("estado")) Then Me.TxtUf.text = gRs("estado")
   If Not IsNull(gRs("cep")) Then Me.Txtcep.text = gRs("cep")
   If Not IsNull(gRs("cgccpf")) Then Me.Txtcgc_cpf.text = gRs("cgccpf")
   If Not IsNull(gRs("rg")) Then Me.TxtRG.text = gRs("rg")
   If Not IsNull(gRs("Telefone")) Then Me.TxtTelefone.text = gRs("Telefone")
   If Not IsNull(gRs("celular")) Then Me.TxtCelular.text = gRs("celular")
   If Not IsNull(gRs("diaAniver")) Then Me.TxtDiaAniver.text = gRs("diaAniver")
   If Not IsNull(gRs("MesAniver")) Then Me.TxtMesAniver.text = gRs("MesAniver")
   If Not IsNull(gRs("AnoAniver")) Then Me.TxtAnoAniver.text = gRs("AnoAniver")
   If Not IsNull(gRs("Ultcompra")) Then Me.TxtUltimaCompra.text = Format(gRs("Ultcompra"), "dd/mm/YYYY")
   If Not IsNull(gRs("email")) Then Me.TxtContato.text = gRs("email")
   'If Not IsNull(gRs("limite")) Then Me.TxtLimite.text = Format(gRs("Limite"), "###,###,##0.00")
   'If Not IsNull(gRs("Saldo")) Then Me.TxtSaldo.text = Format(gRs("saldo"), "###,###,##0.00")
   'If Not IsNull(gRs("EndCobra")) Then Me.TxtEndCobra.text = gRs("EndCobra")
   'If Not IsNull(gRs("BairCobra")) Then Me.TxtBairCobra.text = gRs("BairCobra")
   'If Not IsNull(gRs("CidaCobra")) Then Me.TxtCidaCobra.text = gRs("cidaCobra")
   
   'If Not IsNull(gRs("UFCobra")) Then Me.TxtUFCobra.text = gRs("UFCobra")
   'If Not IsNull(gRs("CepCobra")) Then Me.TxtCepcobra.text = gRs("CepCobra")
   'If Not IsNull(gRs("Insc_est")) Then Me.TxtInsc_est.text = gRs("Insc_est")
   'If IsNull(gRs!negativo) Then
   '   Me.ChkNegativo.Value = 0
   'Else
   '   Me.ChkNegativo.Value = IIf(Len(gRs!negativo) = 0, 0, gRs!negativo)
   'End If
End Sub
Private Sub cmdAdd_Click()

   Me.LblCodclie.Caption = ""
   limpa_tela Me
   Me.TxtNome.SetFocus
   Me.cmdUpdate.Enabled = True
   Me.cmddesfaz.Enabled = True
   Me.cmdEditar.Enabled = False
   Me.cmdAdd.Enabled = False
   Me.CmdSair.Enabled = False
   Me.cmdDelete.Enabled = False
 
   lIncluir = True
End Sub

Private Sub cmdDelete_Click()
    
    On Error GoTo ErroDelete
    
    'this may produce an error if you delete the last
    'record or the only record in the recordset
    If MsgBox("Deseja realmente apagar este Cliente ? ", vbYesNo, "Atenção") = vbYes Then
        gSql = "delete * from tab_clientes where id = " & Val(Me.LblCodclie.Caption)
        ConDb.Execute gSql
        gRs.Close
        Abre_Le_rst
        Carrega_Grid
        gRs.MoveFirst
        Carrega_tela
        Desabilita Me
     End If
     Exit Sub
     
ErroDelete:
     MsgBox "Deu erro na exclusao do Cliente" & Chr(13) & "Instrucao Sql = '" & _
            cSql & "'  "
End Sub


Private Sub cmddesfaz_Click()
  
  lIncluir = False
  
  Desabilita Me
   
  Me.cmdUpdate.Enabled = False
  Me.cmddesfaz.Enabled = False
  Me.cmdEditar.Enabled = True
  Me.cmdAdd.Enabled = True
  Me.CmdSair.Enabled = True
  Me.cmdDelete.Enabled = True
 
End Sub

Private Sub cmdEditar_Click()
   
   Habilita Me
   Me.TxtNome.SetFocus
   Me.cmdUpdate.Enabled = True
   Me.cmddesfaz.Enabled = True
   Me.cmdEditar.Enabled = False
   Me.cmdAdd.Enabled = False
   Me.CmdSair.Enabled = False
   Me.cmdDelete.Enabled = False
  
End Sub

Private Sub CmdSair_Click()
   Unload Me
End Sub

Private Sub cmdUpdate_Click()
   
   gRs.Close
   If lIncluir Then
      gSql = "SELECT nome,cgccpf from tab_clientes WHERE cgccpf = '" & Txtcgc_cpf.text & "'"
      prsCliente.Open gSql, ConDb, adOpenKeyset
      If prsCliente.BOF And prsCliente.EOF Then
         prsCliente.Close
         suInsert
         ConDb.Execute gSql
      Else
         MsgBox "Cliente com CGC/CPF já cadastrado", vbOKOnly, "Atenção " & gOperador
         prsCliente.Close
         Txtcgc_cpf.SetFocus
         Exit Sub
      End If
      lIncluir = False
   Else
      gSql = "UPDATE  tab_clientes SET Nome = '" & Me.TxtNome.text & "',"
      gSql = gSql & " endereco = '" & Me.TxtEndereco.text & "',"
      gSql = gSql & " bairro = '" & Me.TxtBairro.text & " ',"
      gSql = gSql & " Cidade = '" & Me.TxtCidade.text & " ',"
      gSql = gSql & " estado = '" & Me.TxtUf.text & "',"
      gSql = gSql & " CEP  = '" & Me.Txtcep.text & "',"
      gSql = gSql & " CgcCpf = '" & Me.Txtcgc_cpf.text & "',"
      gSql = gSql & " rg = '" & Me.TxtRG.text & " ',"
      gSql = gSql & " telefone = '" & Me.TxtTelefone.text & "',"
      gSql = gSql & " celular = '" & Me.TxtCelular.text & "',"
      gSql = gSql & " diaaniver = '" & f_nulo(Me.TxtDiaAniver.text, " ") & "',"
      gSql = gSql & " mesaniver = '" & f_nulo(Me.TxtMesAniver.text, " ") & "',"
      gSql = gSql & " anoaniver = '" & f_nulo(Me.TxtAnoAniver.text, " ") & "',"
      gSql = gSql & " email = '" & f_nulo(Me.TxtEmail.text, " ") & "',"
      If Me.TxtUltimaCompra.text <> "" Then
         gSql = gSql & " ultcompra =  '" & Format(CDate(Me.TxtUltimaCompra.text), "yyyy-mm-dd") & "',"
      Else
         gSql = gSql & " ultcompra =  '',"
      End If
      gSql = gSql & " operador = " & gnCodOperador & ", datatual = '" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "'"
      gSql = gSql & " WHERE idCli = " & Me.LblCodclie.Caption
      ConDb.Execute gSql
      
   End If
     
   Abre_Le_rst
   
   Carrega_Grid
   
   gRs.MoveFirst
   
   Carrega_tela
   'Deixa os textbox desabilitados
   Desabilita Me
   
   Me.cmdUpdate.Enabled = False
   Me.cmddesfaz.Enabled = False
   Me.cmdEditar.Enabled = True
   Me.cmdAdd.Enabled = True
   Me.CmdSair.Enabled = True
   Me.cmdDelete.Enabled = True
     
End Sub



Private Sub Form_Activate()
   Abre_Le_rst
   
   Me.LblCodclie.Caption = ""

   If gRs.BOF And gRs.EOF Then
      If MsgBox("Arquivo vazio. Incluir dados agora ?", vbYesNo, "Atenção ") = vbYes Then
         '**--> função para dar o INSERT --->
         suInsert
         ConDb.Execute gSql
         gRs.Close
         Abre_Le_rst
         Me.LblCodclie.Caption = gRs!idCli
         cmdEditar_Click
         lPrimeiro = True
      Exit Sub
    Else
         Desabilita Me
      End If
      
   Else
      gRs.MoveFirst
      Me.LblCodclie.Caption = gRs!idCli
      gRs.Close
      gSql = "Select * from tab_clientes "
      gSql = gSql & " WHERE idCli = " & Val(Me.LblCodclie.Caption)
      gRs.Open gSql, ConDb, adOpenForwardOnly
      Carrega_tela
      Desabilita Me
      gRs.Close
      lIncluir = False
      lPrimeiro = False
   End If
   
   Abre_Le_rst
   Carrega_Grid
      
   lIncluir = False

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then Sendkeys "{TAB}"
End Sub

Private Sub Form_Load()
 
   'Centraliza a tela no video
   Me.Move (Screen.Width - Me.Width) / 2, _
           (Screen.Height - Me.Height) / 2
 
End Sub

Private Sub Form_Unload(Cancel As Integer)
    gRs.Close
    Screen.MousePointer = vbDefault
End Sub
Private Sub Abre_Le_rst()
   gSql = "select * "
   gSql = gSql & " FROM tab_clientes"
   gRs.Open gSql, ConDb, adOpenKeyset
   
End Sub
Private Sub Carrega_Grid()

'Teste do MsHFlexgrid1 - eh eh eh
  MSFlexGrid1.Row = 0
  
  With gRs
      '.MoveLast
      'nItem = .RecordCount
      '.MoveFirst
      MSFlexGrid1.Rows = 1
      MSFlexGrid1.Redraw = False
      
      Do While Not .EOF
         MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
         MSFlexGrid1.Row = MSFlexGrid1.Rows - 1
            
         MSFlexGrid1.Col = 0: MSFlexGrid1.text = f_nulo(!idCli, "")
         MSFlexGrid1.Col = 1: MSFlexGrid1.text = f_nulo(!nome, "")
         MSFlexGrid1.Col = 2: MSFlexGrid1.text = f_nulo(!Telefone, "")
         MSFlexGrid1.Col = 3: MSFlexGrid1.text = f_nulo(!endereco, "")
         MSFlexGrid1.Col = 4: MSFlexGrid1.text = f_nulo(!bairro, "")
         MSFlexGrid1.Col = 5: MSFlexGrid1.text = f_nulo(!Cidade, "")
         MSFlexGrid1.Col = 6: MSFlexGrid1.text = f_nulo(!estado, "")
         MSFlexGrid1.Col = 7: MSFlexGrid1.text = f_nulo(!cep, "")
        ' MSFlexGrid1.Col = 8: MSFlexGrid1.Text = f_nulo(!cgccpf, "")
        ' MSFlexGrid1.Col = 9: MSFlexGrid1.Text = f_nulo(!rg, "")
        ' MSFlexGrid1.Col = 10: MSFlexGrid1.Text = f_nulo(!diaAniver, "")
        ' MSFlexGrid1.Col = 11: MSFlexGrid1.Text = f_nulo(!MesAniver, "")
        ' MSFlexGrid1.Col = 12: MSFlexGrid1.Text = f_nulo(!AnoAniver, "")
         MSFlexGrid1.Col = 8: MSFlexGrid1.text = f_nulo(!celular, "")

         MSFlexGrid1.Col = 9: MSFlexGrid1.text = f_nulo(!Ultcompra, "01/01/2002")
        ' MSFlexGrid1.Col = 15: MSFlexGrid1.Text = f_nulo(!email, "")
        ' MSFlexGrid1.Col = 16: MSFlexGrid1.Text = f_nulo(!contato, "")
        ' MSFlexGrid1.Col = 17: MSFlexGrid1.Text = Format(f_nulo(!limite, 0), "###,###,##0.00")
        ' MSFlexGrid1.Col = 18: MSFlexGrid1.Text = Format(f_nulo(!Saldo, 0), "###,###,##0.00")
        ' MSFlexGrid1.Col = 19: MSFlexGrid1.Text = f_nulo(!EndCobra, "")
        ' MSFlexGrid1.Col = 20: MSFlexGrid1.Text = f_nulo(!BairCobra, "")
        ' MSFlexGrid1.Col = 21: MSFlexGrid1.Text = f_nulo(!CidaCobra, "")
        ' MSFlexGrid1.Col = 22: MSFlexGrid1.Text = f_nulo(!UFCobra, "")
        ' MSFlexGrid1.Col = 23: MSFlexGrid1.Text = f_nulo(!CepCobra, "")
        ' MSFlexGrid1.Col = 24: MSFlexGrid1.Text = f_nulo(!insc_est, "")
        ' MSFlexGrid1.Col = 25: MSFlexGrid1.Text = f_nulo(!Negativo, "")
        
         
         .MoveNext
         
       Loop
       MSFlexGrid1.Redraw = True
       
       MSFlexGrid1.FixedRows = 1
          
  End With
  
  End Sub

Private Sub MSFlexGrid1_Click()
  Dim oldrow As Long
  Dim lcColGrid As Double
  
  If MSFlexGrid1.Row = 1 Then
     lcColGrid = MSFlexGrid1.Col
     MSFlexGrid1.Col = lcColGrid
     MSFlexGrid1.Sort = flexSortStringAscending
  End If
 
  oldrow = MSFlexGrid1.Row
  
  MSFlexGrid1.Row = 0
  
  With MSFlexGrid1
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
    
    .Col = 0:   LblCodclie.Caption = .text: .CellBackColor = vbYellow
    .Col = 1:   TxtNome.text = .text: .CellBackColor = vbYellow
    .Col = 2:   TxtTelefone.text = .text: .CellBackColor = vbYellow
    .Col = 3:   TxtEndereco.text = .text: .CellBackColor = vbYellow
    .Col = 4:   TxtBairro.text = .text: .CellBackColor = vbYellow
    .Col = 5:   TxtCidade.text = .text: .CellBackColor = vbYellow
    .Col = 6:   TxtUf.text = .text: .CellBackColor = vbYellow
    .Col = 7:   Txtcep.text = .text: .CellBackColor = vbYellow
    '.Col = 8:   Txtcgc_cpf.Text = .Text: .CellBackColor = vbYellow
    '.Col = 9:   TxtRG.Text = .Text: .CellBackColor = vbYellow
       
    '.Col = 10:  TxtDiaAniver.Text = .Text: .CellBackColor = vbYellow
    '.Col = 11:  TxtMesAniver.Text = .Text: .CellBackColor = vbYellow
    '.Col = 12:  TxtAnoAniver.Text = Right(.Text, 2): .CellBackColor = vbYellow
    .Col = 9:  TxtCelular.text = .text: .CellBackColor = vbYellow
    '.Col = 14:  TxtUltimaCompra.Text = .Text: .CellBackColor = vbYellow
    '.Col = 15:  TxtEmail.Text = .Text: .CellBackColor = vbYellow
    '.Col = 16:  TxtContato.Text = .Text: .CellBackColor = vbYellow
    '.Col = 17:  TxtLimite.Text = .Text: .CellBackColor = vbYellow
    '.Col = 18:  TxtSaldo.Text = .Text: .CellBackColor = vbYellow
    '.Col = 19:  TxtEndCobra.Text = .Text: .CellBackColor = vbYellow
    '.Col = 20:  TxtBairCobra.Text = .Text: .CellBackColor = vbYellow
    '.Col = 21:  TxtCidaCobra.Text = .Text: .CellBackColor = vbYellow
    '.Col = 22:  TxtUFCobra.Text = .Text: .CellBackColor = vbYellow
    '.Col = 23:  TxtCepcobra.Text = .Text: .CellBackColor = vbYellow
    '.Col = 24:  TxtInsc_est.Text = .Text: .CellBackColor = vbYellow
    '.Col = 25:  ChkNegativo.Value = IIf(.Text = False, 0, 1): .CellBackColor = vbYellow
    
    .TopRow = .Row
    
    '.Refresh
   
End With
gRs.Close
gSql = "Select * from tab_clientes "
gSql = gSql & " WHERE idCli = " & Val(Me.LblCodclie.Caption)
gRs.Open gSql, ConDb, adOpenForwardOnly
Carrega_tela
Desabilita Me
'gRs.Close

End Sub

Private Sub suInsert()

    gSql = "INSERT INTO tab_clientes (Nome,endereco,bairro,cidade,estado,"
    gSql = gSql & "CEP,CgcCpf,rg,telefone,celular,diaaniver,mesaniver,anoaniver,"
    gSql = gSql & "email,ultcompra,operador,datatual) "
    gSql = gSql & "VALUES ('" & Me.TxtNome.text & "','"
    gSql = gSql & Me.TxtEndereco.text & "','"
    gSql = gSql & Me.TxtBairro.text & "','"
    gSql = gSql & Me.TxtCidade.text & "','"
    gSql = gSql & Me.TxtUf.text & "','"
    gSql = gSql & Me.Txtcep.text & "','"
    gSql = gSql & Me.Txtcgc_cpf.text & "','"
    gSql = gSql & Me.TxtRG.text & "','"
    gSql = gSql & Me.TxtTelefone.text & "','"
    gSql = gSql & Me.TxtCelular.text & "','"
    gSql = gSql & Me.TxtDiaAniver.text & "','"
    gSql = gSql & Me.TxtMesAniver.text & "','"
    gSql = gSql & Me.TxtAnoAniver.text & "','"
    gSql = gSql & Me.TxtEmail.text & "',"
    If Me.TxtUltimaCompra.text <> "" Then
       gSql = gSql & "'" & CDate(Me.TxtUltimaCompra.text) & "','"
    Else
       gSql = gSql & "NULL,"
    End If
    'gSql = gSql & Replace(Me.TxtLimite.text, ",", ".") & "','"
    'gSql = gSql & Replace(Me.TxtSaldo.text, ",", ".") & "','"
    'gSql = gSql & Me.TxtEndCobra.text & "','"
    'gSql = gSql & Me.TxtBairCobra.text & "','"
    'gSql = gSql & Me.TxtCidaCobra.text & "','"
    'gSql = gSql & Me.TxtUFCobra.text & "','"
    'gSql = gSql & Me.TxtCepcobra.text & "','"
    'gSql = gSql & Me.ChkNegativo.Value & "','"
    'gSql = gSql & Me.TxtInsc_est.text & "','"
    gSql = gSql & gnCodOperador & ",'" & fuDateSQL() & "')"

End Sub

Private Sub TxtAnoAniver_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Or KeyAscii = 9 Then KeyAscii = 0
End Sub

Private Sub TxtBairCobra_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Or KeyAscii = 9 Then KeyAscii = 0
End Sub

Private Sub TxtBairro_GotFocus()
 With TxtBairro
      .SelStart = 0
      .SelLength = Len(.text)
   End With
End Sub

Private Sub TxtBairro_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Or KeyAscii = 9 Then KeyAscii = 0
End Sub

Private Sub TxtCelular_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Or KeyAscii = 9 Then KeyAscii = 0
End Sub

Private Sub Txtcep_GotFocus()
   With Txtcep
      .SelStart = 0
      .SelLength = Len(.text)
   End With

End Sub

Private Sub Txtcep_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Or KeyAscii = 9 Then KeyAscii = 0
End Sub

Private Sub TxtCepcobra_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Or KeyAscii = 9 Then KeyAscii = 0
End Sub

Private Sub Txtcgc_cpf_GotFocus()
   With Txtcgc_cpf
      .SelStart = 0
      .SelLength = Len(.text)
   End With
End Sub

Private Sub Txtcgc_cpf_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Or KeyAscii = 9 Then KeyAscii = 0
End Sub

Private Sub TxtCidaCobra_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Or KeyAscii = 9 Then KeyAscii = 0
End Sub

Private Sub TxtCidade_GotFocus()
   With TxtCidade
      .SelStart = 0
      .SelLength = Len(.text)
   End With
End Sub

Private Sub TxtCidade_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Or KeyAscii = 9 Then KeyAscii = 0
End Sub

Private Sub TxtContato_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Or KeyAscii = 9 Then KeyAscii = 0
End Sub

Private Sub TxtDiaAniver_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Or KeyAscii = 9 Then KeyAscii = 0
End Sub

Private Sub TxtEmail_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Or KeyAscii = 9 Then KeyAscii = 0
End Sub

Private Sub TxtEndCobra_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Or KeyAscii = 9 Then KeyAscii = 0
End Sub

Private Sub TxtEndereco_GotFocus()
 With TxtEndereco
      .SelStart = 0
      .SelLength = Len(.text)
   End With
End Sub

Private Sub TxtEndereco_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Or KeyAscii = 9 Then KeyAscii = 0
End Sub

Private Sub TxtInsc_est_GotFocus()
   With TxtInsc_est
      .SelStart = 0
      .SelLength = Len(.text)
   End With

End Sub

Private Sub TxtInsc_est_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Or KeyAscii = 9 Then KeyAscii = 0
End Sub

Private Sub TxtLimite_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Or KeyAscii = 9 Then KeyAscii = 0
End Sub

Private Sub TxtMesAniver_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Or KeyAscii = 9 Then KeyAscii = 0
End Sub

Private Sub TxtNome_GotFocus()
   With TxtNome
      .SelStart = 0
      .SelLength = Len(.text)
   End With
End Sub

Private Sub TxtNome_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Or KeyAscii = 9 Then KeyAscii = 0
End Sub

Private Sub TxtRG_GotFocus()
   With TxtRG
      .SelStart = 0
      .SelLength = Len(.text)
   End With

End Sub

Private Sub TxtRG_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Or KeyAscii = 9 Then KeyAscii = 0
End Sub

Private Sub TxtSaldo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Or KeyAscii = 9 Then KeyAscii = 0
End Sub

Private Sub TxtTelefone_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Or KeyAscii = 9 Then KeyAscii = 0
End Sub

Private Sub TxtUf_GotFocus()
 With TxtUf
      .SelStart = 0
      .SelLength = Len(.text)
   End With
End Sub

Private Sub TxtUf_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Or KeyAscii = 9 Then KeyAscii = 0
End Sub

Private Sub TxtUf_LostFocus()
   TxtUf.text = UCase(TxtUf.text)
End Sub

Private Sub TxtUFCobra_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Or KeyAscii = 9 Then KeyAscii = 0
End Sub

Private Sub TxtUFCobra_LostFocus()
   TxtUFCobra.text = UCase(TxtUFCobra.text)
End Sub

Private Sub TxtUltimaCompra_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Or KeyAscii = 9 Then KeyAscii = 0
End Sub
