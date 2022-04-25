VERSION 5.00
Begin VB.Form frmlojas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Lojas"
   ClientHeight    =   5130
   ClientLeft      =   2820
   ClientTop       =   1560
   ClientWidth     =   7500
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5130
   ScaleWidth      =   7500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox ChkDivCupom 
      Alignment       =   1  'Right Justify
      Caption         =   "Dívida no Cupom?"
      Height          =   330
      Left            =   5220
      TabIndex        =   13
      Top             =   2160
      Width           =   1695
   End
   Begin VB.CheckBox ChkSenha 
      Alignment       =   1  'Right Justify
      Caption         =   "Senha?"
      Height          =   285
      Left            =   465
      TabIndex        =   12
      Top             =   2220
      Width           =   915
   End
   Begin VB.TextBox txtPalavra 
      Height          =   285
      Left            =   1185
      TabIndex        =   11
      Top             =   1875
      Width           =   2085
   End
   Begin VB.Frame Frame2 
      Caption         =   "Mensagens:"
      Height          =   930
      Left            =   885
      TabIndex        =   28
      Top             =   2835
      Width           =   5445
      Begin VB.TextBox TxtMens2 
         Enabled         =   0   'False
         Height          =   300
         Left            =   150
         TabIndex        =   15
         Top             =   495
         Width           =   5175
      End
      Begin VB.TextBox TxtMens1 
         Enabled         =   0   'False
         Height          =   300
         Left            =   150
         TabIndex        =   14
         Top             =   225
         Width           =   5175
      End
   End
   Begin VB.TextBox TxtCelular 
      Enabled         =   0   'False
      Height          =   300
      Left            =   4860
      TabIndex        =   8
      Top             =   1140
      Width           =   1185
   End
   Begin VB.TextBox TxtTelefone 
      Enabled         =   0   'False
      Height          =   300
      Left            =   1200
      TabIndex        =   7
      Top             =   1140
      Width           =   945
   End
   Begin VB.TextBox TxtInsc_est 
      Enabled         =   0   'False
      Height          =   300
      Left            =   4860
      TabIndex        =   10
      Top             =   1500
      Width           =   2070
   End
   Begin VB.TextBox TxtCNPJ 
      Enabled         =   0   'False
      Height          =   300
      Left            =   1200
      TabIndex        =   9
      Top             =   1500
      Width           =   2070
   End
   Begin VB.TextBox TxtCep 
      Enabled         =   0   'False
      Height          =   300
      Left            =   6180
      TabIndex        =   6
      Top             =   780
      Width           =   990
   End
   Begin VB.TextBox TxtUf 
      Enabled         =   0   'False
      Height          =   300
      Left            =   4860
      TabIndex        =   5
      Top             =   780
      Width           =   375
   End
   Begin VB.TextBox TxtCidade 
      Enabled         =   0   'False
      Height          =   300
      Left            =   1200
      TabIndex        =   4
      Top             =   780
      Width           =   2010
   End
   Begin VB.TextBox TxtBairro 
      Enabled         =   0   'False
      Height          =   300
      Left            =   4860
      TabIndex        =   3
      Top             =   420
      Width           =   2310
   End
   Begin VB.TextBox TxtEndereco 
      Enabled         =   0   'False
      Height          =   300
      Left            =   1200
      TabIndex        =   2
      Top             =   420
      Width           =   2880
   End
   Begin VB.TextBox TxtNome 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1200
      MaxLength       =   45
      TabIndex        =   1
      Top             =   60
      Width           =   4140
   End
   Begin VB.Frame Frame1 
      Height          =   840
      Left            =   2250
      TabIndex        =   17
      Top             =   3945
      Width           =   2865
      Begin VB.CommandButton cmddesfaz 
         Caption         =   "&Desfaz"
         Enabled         =   0   'False
         Height          =   540
         Left            =   1455
         Picture         =   "frmlojas.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   29
         Tag             =   "&Update"
         Top             =   180
         Width           =   615
      End
      Begin VB.CommandButton CmdSair 
         Caption         =   "&Sair"
         Height          =   540
         Left            =   2145
         Picture         =   "frmlojas.frx":00FA
         Style           =   1  'Graphical
         TabIndex        =   32
         Tag             =   "&Update"
         Top             =   180
         Width           =   615
      End
      Begin VB.CommandButton cmdEditar 
         Caption         =   "&Editar"
         Height          =   540
         Left            =   90
         Picture         =   "frmlojas.frx":01F4
         Style           =   1  'Graphical
         TabIndex        =   18
         Tag             =   "&Refresh"
         Top             =   180
         Width           =   615
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Salvar"
         Enabled         =   0   'False
         Height          =   540
         Left            =   765
         Picture         =   "frmlojas.frx":0366
         Style           =   1  'Graphical
         TabIndex        =   27
         Tag             =   "&Update"
         Top             =   180
         Width           =   615
      End
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Palavra"
      Height          =   195
      Left            =   390
      TabIndex        =   33
      Top             =   1905
      Width           =   540
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Código:"
      Height          =   255
      Index           =   0
      Left            =   5700
      TabIndex        =   31
      Tag             =   "CODVEND:"
      Top             =   60
      Width           =   555
   End
   Begin VB.Label LblCodLoja 
      Caption         =   "codloja"
      ForeColor       =   &H000000FF&
      Height          =   270
      Left            =   6360
      TabIndex        =   30
      Top             =   60
      Width           =   615
   End
   Begin VB.Label lblCelular 
      Alignment       =   1  'Right Justify
      Caption         =   "Celular:"
      Height          =   255
      Left            =   4200
      TabIndex        =   26
      Top             =   1200
      Width           =   585
   End
   Begin VB.Label LblTelefone 
      Alignment       =   1  'Right Justify
      Caption         =   "Telefone:"
      Height          =   255
      Left            =   420
      TabIndex        =   25
      Top             =   1140
      Width           =   705
   End
   Begin VB.Label lblInscest 
      Alignment       =   1  'Right Justify
      Caption         =   "Insc.Est.:"
      Height          =   210
      Left            =   4020
      TabIndex        =   24
      Top             =   1560
      Width           =   795
   End
   Begin VB.Label lblcgc 
      Alignment       =   1  'Right Justify
      Caption         =   "CNPJ:"
      Height          =   255
      Left            =   660
      TabIndex        =   23
      Top             =   1500
      Width           =   495
   End
   Begin VB.Label lblCep 
      Alignment       =   1  'Right Justify
      Caption         =   "CEP:"
      Height          =   240
      Left            =   5580
      TabIndex        =   22
      Top             =   840
      Width           =   450
   End
   Begin VB.Label lbluf 
      Alignment       =   1  'Right Justify
      Caption         =   "Estado:"
      Height          =   180
      Left            =   4200
      TabIndex        =   21
      Top             =   840
      Width           =   600
   End
   Begin VB.Label lblcidade 
      Alignment       =   1  'Right Justify
      Caption         =   "Cidade"
      Height          =   225
      Left            =   540
      TabIndex        =   20
      Top             =   810
      Width           =   540
   End
   Begin VB.Label lblBairro 
      Alignment       =   1  'Right Justify
      Caption         =   "Bairro:"
      Height          =   225
      Left            =   4260
      TabIndex        =   19
      Top             =   480
      Width           =   540
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Endereço:"
      Height          =   255
      Index           =   2
      Left            =   300
      TabIndex        =   16
      Tag             =   "SALARIO:"
      Top             =   420
      Width           =   855
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Nome:"
      Height          =   195
      Index           =   1
      Left            =   690
      TabIndex        =   0
      Tag             =   "NOME:"
      Top             =   60
      Width           =   465
   End
End
Attribute VB_Name = "frmlojas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private lIncluir As Boolean
Private lPrimeiro As Boolean

Private Sub Carrega_tela()
   'Limpa as variaveis da tela se caso ficarem com dados da outra tela
   limpa_tela Me
   'Carrega a tela com os dados do registro
   Me.LblCodLoja = gRs("loja")
   Me.TxtNome.text = gRs("nome")
   If Not IsNull(gRs!endereco) Then Me.TxtEndereco.text = gRs!endereco
   If Not IsNull(gRs!bairro) Then Me.TxtBairro.text = "" & gRs!bairro
   If Not IsNull(gRs!Cidade) Then Me.TxtCidade.text = "" & gRs("cidade")
   If Not IsNull(gRs!estado) Then Me.TxtUf.text = "" & gRs("estado")
   If Not IsNull(gRs!cep) Then Me.TxtCep.text = "" & gRs("cep")
   If Not IsNull(gRs("CNPJ")) Then Me.TxtCNPJ.text = gRs("CNPJ")
   If Not IsNull(gRs("Insc_est")) Then Me.TxtInsc_est.text = gRs("Insc_est")
   If Not IsNull(gRs("telefone")) Then Me.TxtTelefone.text = gRs("telefone")
   If Not IsNull(gRs("celular")) Then Me.TxtCelular.text = gRs("celular")
   If Not IsNull(gRs("senha")) Then Me.ChkSenha.Value = IIf(gRs("senha") = True, 1, 0)
   If Not IsNull(gRs("divcupom")) Then Me.ChkDivCupom.Value = IIf(gRs("divcupom") = True, 1, 0)
   If Not IsNull(gRs("mensagem1")) Then Me.TxtMens1.text = gRs("mensagem1")
   If Not IsNull(gRs("mensagem2")) Then Me.TxtMens2.text = gRs("mensagem2")
   
End Sub

Private Sub cmddesfaz_Click()
  lIncluir = False
  ' Carrega_tela
  cmdEditar.Enabled = True
  CmdSair.Enabled = True
  cmdUpdate.Enabled = False
  cmddesfaz.Enabled = False
  Desabilita Me
End Sub

Private Sub cmdEditar_Click()
   cmdEditar.Enabled = False
   CmdSair.Enabled = False
   cmdUpdate.Enabled = True
   cmddesfaz.Enabled = True
   Habilita Me
   'Me.TxtNome.SetFocus
End Sub

Private Sub CmdSair_Click()
   Unload Me
End Sub

Private Sub cmdUpdate_Click()
   Call Abre_Le_rst
   If gRs.BOF And gRs.EOF Then
       gSql = "INSERT INTO tab_lojas (Nome,endereco,bairro,cidade,"
       gSql = gSql & " estado,cep,CNPJ,insc_est,telefone,celular,"
       gSql = gSql & " senha,divcupom,mensagem1,mensagem2,operador,datatual)"
       gSql = gSql & " Values ('" & Me.TxtNome.text & "','"
       gSql = gSql & Me.TxtEndereco.text & "',' "
       gSql = gSql & Me.TxtBairro.text & "','"
       gSql = gSql & Me.TxtCidade.text & "','"
       gSql = gSql & Me.TxtUf.text & "','"
       gSql = gSql & Me.TxtCep.text & "','"
       gSql = gSql & Me.TxtCNPJ.text & "','"
       gSql = gSql & Me.TxtInsc_est.text & "','"
       gSql = gSql & Me.TxtTelefone.text & "','"
       gSql = gSql & Me.TxtCelular.text & "','"
       gSql = gSql & IIf(Me.ChkSenha.Value = 0, "N", "S") & "','"
       gSql = gSql & IIf(Me.ChkDivCupom.Value = 0, "N", "S") & "','"
       gSql = gSql & Me.TxtMens1.text & "','"
       gSql = gSql & Me.TxtMens2.text & "',"
       gSql = gSql & gnCodOperador & ",'" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "')"
          
   Else
   
        gSql = "UPDATE tab_lojas SET Nome = '" & Me.TxtNome.text & "',"
        gSql = gSql & " endereco = '" & Me.TxtEndereco.text & "', "
        gSql = gSql & " bairro = '" & Me.TxtBairro.text & "',"
        gSql = gSql & " cidade = '" & Me.TxtCidade.text & "',"
        gSql = gSql & " estado = '" & Me.TxtUf.text & "',"
        gSql = gSql & " cep =  '" & Me.TxtCep.text & "',"
        gSql = gSql & " CNPJ = '" & Me.TxtCNPJ.text & "',"
        gSql = gSql & " insc_est = '" & Me.TxtInsc_est.text & "',"
        gSql = gSql & " telefone = '" & Me.TxtTelefone.text & "',"
        gSql = gSql & " celular = '" & Me.TxtCelular.text & "',"
        gSql = gSql & " senha = " & IIf(Me.ChkSenha.Value = 0, "N", "S") & ","
        gSql = gSql & " divcupom = " & IIf(Me.ChkSenha.Value = 0, "N", "S") & ","
        gSql = gSql & " mensagem1 = '" & Me.TxtMens1.text & "',"
        gSql = gSql & " mensagem2 = '" & Me.TxtMens2.text & "',"
        gSql = gSql & " operador = " & gnCodOperador & ", datatual = " & Format(Now, "yyyy-mm-dd hh:mm:ss") & "'"
        gSql = gSql & " WHERE loja = " & Val(Me.LblCodLoja.Caption)
   End If
   ConDb.Execute gSql
   cmdEditar.Enabled = True
   CmdSair.Enabled = True
   cmdUpdate.Enabled = False
   cmddesfaz.Enabled = False
   Desabilita Me
   ConDb.Close
   
End Sub

Private Sub Form_Activate()
   Abre_Le_rst
    
   limpa_tela Me
   
   frmlojas.LblCodLoja.Caption = ""
   If gRs.BOF And gRs.EOF Then
       cmdEditar_Click
       lPrimeiro = True
       TxtNome.SetFocus
       Exit Sub
   Else
       cmdEditar.Enabled = True
       CmdSair.Enabled = True
       cmdUpdate.Enabled = False
       cmddesfaz.Enabled = False
   End If
   
   gRs.MoveFirst
   Carrega_tela
   cmdEditar.Enabled = True
   CmdSair.Enabled = True
   cmdUpdate.Enabled = False
   cmddesfaz.Enabled = False
   lIncluir = False
   lPrimeiro = False
   gRs.Close
   ConDb.Close
   
   Desabilita Me
   Screen.MousePointer = vbDefault
   
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then Sendkeys "{TAB}"
     
End Sub

Private Sub Form_Load()
   'Centraliza a tela no video
   Me.Move (Screen.Width - Me.Width) / 2, _
           (Screen.Height - Me.Height) / 2
   Screen.MousePointer = vbDefault
   
   End Sub

Private Sub Form_Unload(Cancel As Integer)
    Screen.MousePointer = vbDefault
End Sub

Private Sub Abre_Le_rst()

    Call sConectaBanco
       
    gSql = "select * FROM tab_lojas"
    If gRs.State = adStateOpen Then
        gRs.Close
    End If
   
    gRs.Open gSql, ConDb, adOpenKeyset, adLockOptimistic
    
End Sub

Private Sub TxtBairro_GotFocus()
    SelText TxtBairro
End Sub

Private Sub TxtBairro_KeyPress(KeyAscii As Integer)
    Char = Chr(KeyAscii)
    KeyAscii = Asc(UCase(Char))
End Sub

Private Sub TxtCidade_GotFocus()
    SelText TxtCidade
End Sub

Private Sub TxtCidade_KeyPress(KeyAscii As Integer)
    Char = Chr(KeyAscii)
    KeyAscii = Asc(UCase(Char))
End Sub

Private Sub TxtEndereco_GotFocus()
    SelText TxtEndereco
End Sub

Private Sub TxtEndereco_KeyPress(KeyAscii As Integer)
    Char = Chr(KeyAscii)
    KeyAscii = Asc(UCase(Char))
End Sub

Private Sub TxtNome_GotFocus()
    SelText TxtNome
End Sub

Private Sub TxtNome_KeyPress(KeyAscii As Integer)
    Char = Chr(KeyAscii)
    KeyAscii = Asc(UCase(Char))
End Sub

Private Sub TxtUf_KeyPress(KeyAscii As Integer)
    Char = Chr(KeyAscii)
    KeyAscii = Asc(UCase(Char))
End Sub
