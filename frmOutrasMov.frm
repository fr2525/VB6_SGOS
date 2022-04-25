VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmOutrasMov 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Outras Movimentações"
   ClientHeight    =   4185
   ClientLeft      =   2670
   ClientTop       =   900
   ClientWidth     =   7125
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4185
   ScaleWidth      =   7125
   ShowInTaskbar   =   0   'False
   Tag             =   "cadvend"
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   6360
      TabIndex        =   23
      Text            =   "Text1"
      Top             =   120
      Width           =   285
   End
   Begin VB.ComboBox CmbTipo 
      Height          =   315
      Left            =   1275
      TabIndex        =   32
      Text            =   "Combo1"
      Top             =   120
      Width           =   3540
   End
   Begin VB.TextBox TxtObserv 
      Height          =   285
      Left            =   1200
      TabIndex        =   10
      Top             =   2370
      Width           =   5445
   End
   Begin VB.TextBox TxtInsc_est 
      Height          =   285
      Left            =   1215
      TabIndex        =   8
      Top             =   1980
      Width           =   1335
   End
   Begin MSMask.MaskEdBox Txtcep 
      Height          =   300
      Left            =   5415
      TabIndex        =   5
      Top             =   1230
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   529
      _Version        =   393216
      MaxLength       =   9
      Format          =   "00000-000"
      PromptChar      =   "_"
   End
   Begin VB.Frame Frame1 
      Height          =   795
      Left            =   1560
      TabIndex        =   22
      Top             =   3045
      Width           =   4245
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Salvar"
         Enabled         =   0   'False
         Height          =   540
         Left            =   2190
         Picture         =   "frmOutrasMov.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   14
         Tag             =   "&Update"
         Top             =   180
         Width           =   615
      End
      Begin VB.CommandButton cmdEditar 
         Caption         =   "&Editar"
         Height          =   540
         Left            =   780
         Picture         =   "frmOutrasMov.frx":00FA
         Style           =   1  'Graphical
         TabIndex        =   28
         Tag             =   "&Refresh"
         Top             =   180
         Width           =   615
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Excluir"
         Height          =   540
         Left            =   1500
         Picture         =   "frmOutrasMov.frx":026C
         Style           =   1  'Graphical
         TabIndex        =   27
         Tag             =   "&Delete"
         Top             =   180
         Width           =   615
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Incluir"
         Height          =   540
         Left            =   120
         Picture         =   "frmOutrasMov.frx":03DE
         Style           =   1  'Graphical
         TabIndex        =   26
         Tag             =   "&Add"
         Top             =   180
         Width           =   615
      End
      Begin VB.CommandButton CmdSair 
         Caption         =   "&Sair"
         Height          =   540
         Left            =   3540
         Picture         =   "frmOutrasMov.frx":04C8
         Style           =   1  'Graphical
         TabIndex        =   25
         Tag             =   "&Update"
         Top             =   180
         Width           =   615
      End
      Begin VB.CommandButton cmddesfaz 
         Caption         =   "&Desfaz"
         Enabled         =   0   'False
         Height          =   540
         Left            =   2880
         Picture         =   "frmOutrasMov.frx":05C2
         Style           =   1  'Graphical
         TabIndex        =   24
         Tag             =   "&Update"
         Top             =   180
         Width           =   615
      End
   End
   Begin VB.TextBox TxtContato 
      Enabled         =   0   'False
      Height          =   285
      Left            =   3555
      TabIndex        =   9
      Top             =   1995
      Width           =   3105
   End
   Begin VB.TextBox TxtBairro 
      Height          =   285
      Left            =   4260
      TabIndex        =   2
      Top             =   855
      Width           =   2400
   End
   Begin VB.TextBox TxtCidade 
      Height          =   285
      Left            =   1230
      TabIndex        =   3
      Top             =   1230
      Width           =   2280
   End
   Begin VB.TextBox TxtUf 
      Height          =   285
      Left            =   4260
      TabIndex        =   4
      Top             =   1230
      Width           =   450
   End
   Begin VB.TextBox TxtTelefone 
      Height          =   285
      Left            =   1230
      TabIndex        =   6
      Top             =   1605
      Width           =   1890
   End
   Begin VB.TextBox TxtCelular 
      Height          =   285
      Left            =   4770
      TabIndex        =   7
      Top             =   1635
      Width           =   1890
   End
   Begin VB.TextBox TxtEndereco 
      Height          =   285
      Left            =   1245
      TabIndex        =   1
      Top             =   855
      Width           =   2310
   End
   Begin VB.TextBox TxtNome 
      Height          =   285
      Left            =   1260
      TabIndex        =   0
      Top             =   495
      Width           =   5400
   End
   Begin VB.Label Label4 
      Caption         =   "Observ.:"
      Height          =   225
      Left            =   495
      TabIndex        =   31
      Top             =   2430
      Width           =   645
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Insc.Est."
      Height          =   195
      Left            =   495
      TabIndex        =   30
      Top             =   1980
      Width           =   615
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "(E)ntrada/(S)aida"
      Height          =   195
      Left            =   5025
      TabIndex        =   29
      Top             =   165
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Contato"
      Height          =   210
      Left            =   2790
      TabIndex        =   13
      Top             =   2040
      Width           =   645
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Tipo"
      Height          =   195
      Index           =   0
      Left            =   795
      TabIndex        =   21
      Tag             =   "CODVEND:"
      Top             =   165
      Width           =   315
   End
   Begin VB.Label LblBairro 
      Caption         =   "Bairro"
      Height          =   240
      Left            =   3720
      TabIndex        =   20
      Top             =   885
      Width           =   420
   End
   Begin VB.Label LblCidade 
      Caption         =   "Cidade"
      Height          =   240
      Left            =   570
      TabIndex        =   19
      Top             =   1260
      Width           =   540
   End
   Begin VB.Label LblUf 
      Caption         =   "Estado"
      Height          =   195
      Left            =   3600
      TabIndex        =   18
      Top             =   1290
      Width           =   555
   End
   Begin VB.Label LblTelefone 
      Caption         =   "Telefone"
      Height          =   240
      Left            =   480
      TabIndex        =   17
      Top             =   1635
      Width           =   630
   End
   Begin VB.Label LblCelular 
      Alignment       =   1  'Right Justify
      Caption         =   "Celular"
      Height          =   195
      Left            =   4170
      TabIndex        =   16
      Top             =   1695
      Width           =   480
   End
   Begin VB.Label lblCep 
      Alignment       =   1  'Right Justify
      Caption         =   "CEP"
      Height          =   255
      Left            =   4800
      TabIndex        =   15
      Top             =   1290
      Width           =   465
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Endereço:"
      Height          =   195
      Index           =   2
      Left            =   375
      TabIndex        =   12
      Tag             =   "NOME:"
      Top             =   855
      Width           =   735
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Nome:"
      Height          =   195
      Index           =   1
      Left            =   645
      TabIndex        =   11
      Tag             =   "NOME:"
      Top             =   540
      Width           =   465
   End
End
Attribute VB_Name = "frmOutrasmov"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private lIncluir As Boolean
Private pRsFornec As New ADODB.Recordset

Private Sub Abre_Le_rst()
  gSql = "select * FROM tab_fornece"
  gRs.Open gSql, ConDb, adOpenKeyset
  
End Sub

Private Sub cmdAdd_Click()
   
   limpa_tela Me
   
   Me.LblCodfor.Caption = ""
   Me.MskCnpj.SetFocus
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
    If MsgBox("Deseja realmente apagar este Fornecedor ? ", vbYesNo, "Atenção") = vbYes Then
       gSql = "delete * from tab_fornece where codfor = " & Me.LblCodfor.Caption
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
     MsgBox "Deu erro na exclusao do Fornecedor " & Chr(13) & "Instrucao Sql = '" & _
            cSql & "'  "
End Sub


Private Sub cmddesfaz_Click()
  
  lIncluir = False
  
  ' Carrega_tela
  Desabilita Me
  MSFlexGrid1_Click
  Me.cmdUpdate.Enabled = False
  Me.cmddesfaz.Enabled = False
  Me.cmdEditar.Enabled = True
  Me.cmdAdd.Enabled = True
  Me.CmdSair.Enabled = True
  Me.cmdDelete.Enabled = True

End Sub

Private Sub cmdEditar_Click()
   ' Carrega_tela
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
      gSql = "SELECT cnpj from tab_fornece "
      gSql = gSql & " WHERE cnpj1 = '" & Me.MskCnpj.Text & "'"
      pRsFornec.Open gSql, ConDb, adOpenKeyset
      If pRsFornec.BOF And pRsFornec.EOF Then
      Else
         MsgBox "CNPJ já cadastrado", vbOKOnly, "Atenção " & gOperador
         Exit Sub
      End If
      pRsFornec.Close

      gSql = "INSERT INTO tab_fornece (Nome,endereco,bairro,cidade,estado,"
      gSql = gSql & "CEP,telefone,fax,contato, cnpj1,insc_est,observacao,operador, datatual"
      gSql = gSql & ") "
      gSql = gSql & "VALUES ('" & Me.TxtNome.Text & "','"
      gSql = gSql & Me.TxtEndereco.Text & "','"
      gSql = gSql & Me.TxtBairro.Text & "','"
      gSql = gSql & Me.TxtCidade.Text & "','"
      gSql = gSql & Me.TxtUf.Text & "','"
      gSql = gSql & Me.Txtcep.Text & "','"
      gSql = gSql & Me.TxtTelefone.Text & "','"
      gSql = gSql & Me.TxtCelular.Text & "','"
      gSql = gSql & Me.TxtContato.Text & "','"
      gSql = gSql & Me.MskCnpj.Text & "','"
      gSql = gSql & Me.TxtInsc_est.Text & "','"
      gSql = gSql & Me.TxtObserv.Text & "',"
      gSql = gSql & ",'" & gOperador & "',Cdate('" & Date & "') )"
      ConDb.Execute gSql
      lIncluir = False
   Else
      gSql = "UPDATE tab_fornece SET Nome = '" & Me.TxtNome.Text & "',"
      gSql = gSql & "endereco = '" & Me.TxtEndereco.Text & "',"
      gSql = gSql & "bairro = '" & Me.TxtBairro.Text & "',"
      gSql = gSql & "cidade = '" & Me.TxtCidade.Text & "',"
      gSql = gSql & "estado = '" & Me.TxtUf.Text & "',"
      gSql = gSql & "CEP = '" & Me.Txtcep.Text & "',"
      gSql = gSql & "Telefone = '" & Me.TxtTelefone.Text & "',"
      gSql = gSql & "Fax = '" & Me.TxtCelular.Text & "',"
      gSql = gSql & "contato = '" & Me.TxtContato.Text & "',"
      gSql = gSql & "CNPJ = '" & Me.MskCnpj.Text & "',"
      gSql = gSql & "Insc_est = '" & Me.TxtInsc_est.Text & "',"
      gSql = gSql & "Observacao = '" & Me.TxtObserv.Text & "'"
      gSql = gSql & " ,operador = '" & gOperador & "', datatual = cDate('" & Date & "')"
      gSql = gSql & " WHERE codfor = " & Me.LblCodfor.Caption
      ConDb.Execute gSql
      
   End If
       
   Abre_Le_rst
   
   Carrega_Grid
   gRs.MoveFirst
   Carrega_tela
   Desabilita Me
      
   Me.cmdUpdate.Enabled = False
   Me.cmddesfaz.Enabled = False
   Me.cmdEditar.Enabled = True
   Me.cmdAdd.Enabled = True
   Me.CmdSair.Enabled = True
   Me.cmdDelete.Enabled = True
     
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub Form_Load()
   Abre_Le_rst
   'Centraliza a tela no video
   Me.Move (Screen.Width - Me.Width) / 2, _
           (Screen.Height - Me.Height) / 2
   
   Me.LblCodfor.Caption = ""
    If gRs.BOF And gRs.EOF Then
      If MsgBox("Arquivo vazio. Incluir dados agora ?", vbYesNo, "Atenção ") = vbYes Then
         'rst.AddNew
         With gRs
           .AddNew
           !nome = ""
           .Update
         End With
         cmdEditar_Click
         lPrimeiro = True
      Else
         Desabilita Me
      End If
      
   Else
      gRs.MoveFirst
      Carrega_tela
    
      Desabilita Me
      lIncluir = False
      lPrimeiro = False
   End If
   Carrega_Grid
   
   lIncluir = False

End Sub

Private Sub Form_Unload(Cancel As Integer)
    gRs.Close
    Screen.MousePointer = vbDefault
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
    
    .Col = 0:   LblCodfor.Caption = .Text: .CellBackColor = vbYellow
    .Col = 1:   TxtNome.Text = .Text: .CellBackColor = vbYellow
    .Col = 2:   MskCnpj.Text = .Text: .CellBackColor = vbYellow
    .Col = 3:   TxtEndereco.Text = .Text: .CellBackColor = vbYellow
    .Col = 4:   TxtBairro.Text = .Text: .CellBackColor = vbYellow
    .Col = 5:   TxtCidade.Text = .Text: .CellBackColor = vbYellow
    .Col = 6:   TxtUf.Text = .Text: .CellBackColor = vbYellow
    .Col = 7:   Txtcep.Text = .Text: .CellBackColor = vbYellow
    .Col = 8:   TxtTelefone.Text = .Text: .CellBackColor = vbYellow
    .Col = 9:   TxtCelular.Text = .Text: .CellBackColor = vbYellow
    .Col = 10:  TxtContato.Text = .Text: .CellBackColor = vbYellow
    .Col = 11:  TxtInsc_est.Text = .Text: .CellBackColor = vbYellow
    .Col = 12:  TxtObserv.Text = .Text: .CellBackColor = vbYellow
    .TopRow = .Row
    
    
End With


End Sub


Private Sub Carrega_tela()
   'Limpa as variaveis da tela se caso ficarem com dados da outra tela
   limpa_tela Me
   'Carrega a tela com os dados do registro
   Me.LblCodfor.Caption = gRs("codfor")
   Me.TxtNome.Text = "" & gRs("Nome")
   Me.TxtEndereco.Text = "" & gRs("endereco")
   Me.TxtBairro.Text = "" & gRs("bairro").Value
   Me.TxtCidade.Text = "" & gRs("Cidade")
   Me.TxtUf.Text = gRs("estado")
   Me.Txtcep.Text = gRs("cep")
   Me.TxtTelefone.Text = "" & gRs("Telefone")
   Me.TxtCelular.Text = "" & gRs("Fax")
   Me.TxtContato.Text = "" & gRs("Contato")
   Me.MskCnpj.Text = "" & gRs("cnpj")
   Me.TxtInsc_est.Text = "" & gRs("Insc_est")
   Me.TxtObserv.Text = "" & gRs("Observacao")
   
End Sub

Private Sub Carrega_Grid()

'Teste do MsHFlexgrid1 - eh eh eh
  MSFlexGrid1.Row = 0
  
  With gRs
      '.MoveLast
      'nItem = .RecordCount
      '.MoveFirst
      MSFlexGrid1.Rows = 1
      Do While Not .EOF
         MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
         MSFlexGrid1.Row = MSFlexGrid1.Rows - 1
         MSFlexGrid1.Col = 0:  MSFlexGrid1.Text = f_nulo(!codfor, "")
         MSFlexGrid1.Col = 1:  MSFlexGrid1.Text = f_nulo(!nome, "")
         MSFlexGrid1.Col = 2:  MSFlexGrid1.Text = f_nulo(!CNPJ, "")
         MSFlexGrid1.Col = 3:  MSFlexGrid1.Text = f_nulo(!endereco, "")
         MSFlexGrid1.Col = 4:  MSFlexGrid1.Text = f_nulo(!bairro, "")
         MSFlexGrid1.Col = 5:  MSFlexGrid1.Text = f_nulo(!Cidade, "")
         MSFlexGrid1.Col = 6:  MSFlexGrid1.Text = f_nulo(!estado, "")
         MSFlexGrid1.Col = 7:  MSFlexGrid1.Text = f_nulo(!cep, "")
         MSFlexGrid1.Col = 8:  MSFlexGrid1.Text = f_nulo(!Telefone, "")
         MSFlexGrid1.Col = 9:  MSFlexGrid1.Text = f_nulo(!Fax, "")
         MSFlexGrid1.Col = 10: MSFlexGrid1.Text = f_nulo(!contato, "")
         MSFlexGrid1.Col = 11: MSFlexGrid1.Text = f_nulo(!insc_est, "")
         MSFlexGrid1.Col = 12: MSFlexGrid1.Text = f_nulo(!Observacao, "")
         .MoveNext
         
       Loop
       MSFlexGrid1.FixedRows = 1
          
  End With
  
  End Sub


Private Sub MskCNPJ_GotFocus()
   MskCnpj.Mask = "##############"
End Sub

Private Sub MskCNPJ_LostFocus()
   If Len(MskCnpj.Text) > 0 Then
      Select Case Len(MskCnpj.Text)
       Case Is = 11
         MskCnpj.Mask = "###.###.###-##"
         If Not calculacpf(MskCnpj.Text) Then
            MsgBox "CPF com DV incorreto !!!"
            MskCnpj = ""
            MskCnpj.Mask = "##############"
            MskCnpj.SetFocus
         End If
       Case Is = 14
         MskCnpj.Mask = "##.###.###/####-##"
         If Not ValidaCGC(MskCnpj.Text) Then
            MsgBox "CGC com DV incorreto !!! "
            MskCnpj = ""
            MskCnpj.Mask = "##############"
            MskCnpj.SetFocus
         End If
      End Select
    End If

End Sub

