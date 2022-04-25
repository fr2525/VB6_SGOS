VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmfornec 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Fornecedores"
   ClientHeight    =   6135
   ClientLeft      =   2670
   ClientTop       =   900
   ClientWidth     =   7125
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6135
   ScaleWidth      =   7125
   ShowInTaskbar   =   0   'False
   Tag             =   "cadvend"
   Begin VB.TextBox TxtObserv 
      Height          =   285
      Left            =   1200
      TabIndex        =   12
      Top             =   2415
      Width           =   5445
   End
   Begin MSMask.MaskEdBox MskCNPJ 
      Height          =   300
      Left            =   3060
      TabIndex        =   1
      Top             =   120
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   529
      _Version        =   393216
      PromptInclude   =   0   'False
      AutoTab         =   -1  'True
      MaxLength       =   18
      Mask            =   "##.###.###/####-##"
      PromptChar      =   "_"
   End
   Begin VB.TextBox TxtInsc_est 
      Height          =   285
      Left            =   1215
      TabIndex        =   10
      Top             =   1980
      Width           =   1335
   End
   Begin MSMask.MaskEdBox Txtcep 
      Height          =   300
      Left            =   5415
      TabIndex        =   7
      Top             =   1230
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   529
      _Version        =   393216
      MaxLength       =   9
      Format          =   "00000-000"
      PromptChar      =   "_"
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   2025
      Left            =   450
      TabIndex        =   25
      Top             =   2940
      Width           =   6225
      _ExtentX        =   10980
      _ExtentY        =   3572
      _Version        =   393216
      Rows            =   5
      Cols            =   13
      FixedCols       =   0
      FormatString    =   $"frmfornece.frx":0000
   End
   Begin VB.Frame Frame1 
      Height          =   795
      Left            =   1470
      TabIndex        =   24
      Top             =   5085
      Width           =   4245
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Salvar"
         Enabled         =   0   'False
         Height          =   540
         Left            =   2190
         Picture         =   "frmfornece.frx":014F
         Style           =   1  'Graphical
         TabIndex        =   16
         Tag             =   "&Update"
         Top             =   180
         Width           =   615
      End
      Begin VB.CommandButton cmdEditar 
         Caption         =   "&Editar"
         Height          =   540
         Left            =   780
         Picture         =   "frmfornece.frx":0249
         Style           =   1  'Graphical
         TabIndex        =   30
         Tag             =   "&Refresh"
         Top             =   180
         Width           =   615
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Excluir"
         Height          =   540
         Left            =   1500
         Picture         =   "frmfornece.frx":03BB
         Style           =   1  'Graphical
         TabIndex        =   29
         Tag             =   "&Delete"
         Top             =   180
         Width           =   615
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Incluir"
         Height          =   540
         Left            =   120
         Picture         =   "frmfornece.frx":052D
         Style           =   1  'Graphical
         TabIndex        =   28
         Tag             =   "&Add"
         Top             =   180
         Width           =   615
      End
      Begin VB.CommandButton CmdSair 
         Caption         =   "&Sair"
         Height          =   540
         Left            =   3540
         Picture         =   "frmfornece.frx":0617
         Style           =   1  'Graphical
         TabIndex        =   27
         Tag             =   "&Update"
         Top             =   180
         Width           =   615
      End
      Begin VB.CommandButton cmddesfaz 
         Caption         =   "&Desfaz"
         Enabled         =   0   'False
         Height          =   540
         Left            =   2880
         Picture         =   "frmfornece.frx":0711
         Style           =   1  'Graphical
         TabIndex        =   26
         Tag             =   "&Update"
         Top             =   180
         Width           =   615
      End
   End
   Begin VB.TextBox TxtContato 
      Enabled         =   0   'False
      Height          =   285
      Left            =   3555
      TabIndex        =   11
      Top             =   1995
      Width           =   3105
   End
   Begin VB.TextBox TxtBairro 
      Height          =   285
      Left            =   4260
      TabIndex        =   4
      Top             =   855
      Width           =   2400
   End
   Begin VB.TextBox TxtCidade 
      Height          =   285
      Left            =   1230
      TabIndex        =   5
      Top             =   1230
      Width           =   2280
   End
   Begin VB.TextBox TxtUf 
      Height          =   285
      Left            =   4260
      TabIndex        =   6
      Top             =   1230
      Width           =   450
   End
   Begin VB.TextBox TxtTelefone 
      Height          =   285
      Left            =   1230
      TabIndex        =   8
      Top             =   1605
      Width           =   1890
   End
   Begin VB.TextBox TxtCelular 
      Height          =   285
      Left            =   4770
      TabIndex        =   9
      Top             =   1635
      Width           =   1890
   End
   Begin VB.TextBox TxtEndereco 
      Height          =   285
      Left            =   1245
      TabIndex        =   3
      Top             =   855
      Width           =   2310
   End
   Begin VB.TextBox TxtNome 
      Height          =   285
      Left            =   1260
      TabIndex        =   2
      Top             =   480
      Width           =   5400
   End
   Begin VB.Label Label4 
      Caption         =   "Observ.:"
      Height          =   225
      Left            =   495
      TabIndex        =   33
      Top             =   2430
      Width           =   765
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Insc.Est."
      Height          =   195
      Left            =   465
      TabIndex        =   32
      Top             =   1980
      Width           =   615
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "CNPJ"
      Height          =   195
      Left            =   2430
      TabIndex        =   31
      Top             =   180
      Width           =   405
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Contato"
      Height          =   210
      Left            =   2790
      TabIndex        =   15
      Top             =   2040
      Width           =   645
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Código:"
      Height          =   195
      Index           =   0
      Left            =   705
      TabIndex        =   23
      Tag             =   "CODVEND:"
      Top             =   165
      Width           =   540
   End
   Begin VB.Label LblBairro 
      Caption         =   "Bairro"
      Height          =   240
      Left            =   3720
      TabIndex        =   22
      Top             =   885
      Width           =   420
   End
   Begin VB.Label LblCidade 
      Caption         =   "Cidade"
      Height          =   240
      Left            =   645
      TabIndex        =   21
      Top             =   1260
      Width           =   540
   End
   Begin VB.Label LblUf 
      Caption         =   "Estado"
      Height          =   195
      Left            =   3600
      TabIndex        =   20
      Top             =   1290
      Width           =   555
   End
   Begin VB.Label LblTelefone 
      Caption         =   "Telefone"
      Height          =   240
      Left            =   450
      TabIndex        =   19
      Top             =   1635
      Width           =   630
   End
   Begin VB.Label LblCelular 
      Alignment       =   1  'Right Justify
      Caption         =   "Celular"
      Height          =   195
      Left            =   4170
      TabIndex        =   18
      Top             =   1695
      Width           =   480
   End
   Begin VB.Label lblCep 
      Alignment       =   1  'Right Justify
      Caption         =   "CEP"
      Height          =   255
      Left            =   4800
      TabIndex        =   17
      Top             =   1290
      Width           =   465
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Endereço:"
      Height          =   195
      Index           =   2
      Left            =   450
      TabIndex        =   14
      Tag             =   "NOME:"
      Top             =   855
      Width           =   735
   End
   Begin VB.Label LblCodfor 
      Caption         =   "codfor"
      ForeColor       =   &H000000FF&
      Height          =   210
      Left            =   1410
      TabIndex        =   0
      Top             =   165
      Width           =   975
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Nome:"
      Height          =   195
      Index           =   1
      Left            =   660
      TabIndex        =   13
      Tag             =   "NOME:"
      Top             =   540
      Width           =   465
   End
End
Attribute VB_Name = "frmfornec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private lIncluir As Boolean
Private pRsFornec As ADODB.Recordset

Private Sub Abre_Le_rst()
  gSql = "select * FROM tab_fornece"
  gRs.Open gSql, ConDb, adOpenKeyset
  
End Sub

Private Sub cmdAdd_Click()
   
   limpa_tela Me
   
   Me.LblCodfor.Caption = ""
   'Me.MskCNPJ.text = "##.###.###/####-##"
   Me.cmdUpdate.Enabled = True
   Me.cmddesfaz.Enabled = True
   Me.cmdEditar.Enabled = False
   Me.cmdAdd.Enabled = False
   Me.CmdSair.Enabled = False
   Me.cmdDelete.Enabled = False
   Me.MskCNPJ.Mask = "##.###.###/####-##"
   lIncluir = True
End Sub

Private Sub cmdDelete_Click()

    On Error GoTo ErroDelete
    
    'this may produce an error if you delete the last
    'record or the only record in the recordset
    If MsgBox("Deseja realmente apagar este Fornecedor ? ", vbYesNo, "Atenção") = vbYes Then
       gSql = "delete * from tab_fornece where idFor = " & Me.LblCodfor.Caption
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
   'Me.TxtNome.SetFocus
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
      gSql = gSql & " WHERE cnpj1 = '" & Me.MskCNPJ.text & "'"
      pRsFornec.Open gSql, ConDb, adOpenForwardOnly, adLockOptimistic
      
      If pRsFornec.BOF And pRsFornec.EOF Then
      Else
         MsgBox "CNPJ já cadastrado", vbOKOnly, "Atenção " & gOperador
         Exit Sub
      End If
      pRsFornec.Close

      gSql = "INSERT INTO tab_fornece (Nome,endereco,bairro,cidade,estado,"
      gSql = gSql & "CEP,telefone,celular,contato, cnpj,insc_est,observacao,operador, datatual"
      gSql = gSql & ") "
      gSql = gSql & "VALUES ('" & Me.TxtNome.text & "','"
      gSql = gSql & Me.TxtEndereco.text & "','"
      gSql = gSql & Me.TxtBairro.text & "','"
      gSql = gSql & Me.TxtCidade.text & "','"
      gSql = gSql & Me.TxtUf.text & "','"
      gSql = gSql & Me.Txtcep.text & "','"
      gSql = gSql & Me.TxtTelefone.text & "','"
      gSql = gSql & Me.TxtCelular.text & "','"
      gSql = gSql & Me.TxtContato.text & "','"
      gSql = gSql & Me.MskCNPJ.text & "','"
      gSql = gSql & Me.TxtInsc_est.text & "','"
      gSql = gSql & Me.TxtObserv.text & "'"
      gSql = gSql & "," & gnCodOperador & ",'" & fuDateSQL() & "')"
      ConDb.Execute gSql
      lIncluir = False
   Else
      gSql = "UPDATE tab_fornece SET Nome = '" & Me.TxtNome.text & "',"
      gSql = gSql & "endereco = '" & Me.TxtEndereco.text & "',"
      gSql = gSql & "bairro = '" & Me.TxtBairro.text & "',"
      gSql = gSql & "cidade = '" & Me.TxtCidade.text & "',"
      gSql = gSql & "estado = '" & Me.TxtUf.text & "',"
      gSql = gSql & "CEP = '" & Me.Txtcep.text & "',"
      gSql = gSql & "Telefone = '" & Me.TxtTelefone.text & "',"
      gSql = gSql & "celular = '" & Me.TxtCelular.text & "',"
      gSql = gSql & "contato = '" & Me.TxtContato.text & "',"
      gSql = gSql & "CNPJ = '" & Me.MskCNPJ.text & "',"
      gSql = gSql & "Insc_est = '" & Me.TxtInsc_est.text & "',"
      gSql = gSql & "Observacao = '" & Me.TxtObserv.text & "'"
      gSql = gSql & " ,operador = " & gnCodOperador & ", datatual = '" & fuDateSQL() & "'"
      gSql = gSql & " WHERE idFor = " & Me.LblCodfor.Caption
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
  If KeyCode = vbKeyReturn Then Sendkeys "{TAB}"
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
         'With gRs
         '  .AddNew
         '  !nome = ""
         '  .Update
         'End With
         cmdEditar_Click
         lPrimeiro = True
         lIncluir = True
      Else
         Desabilita Me
      End If
      
   Else
      gRs.MoveFirst
      Carrega_tela
    
      Desabilita Me
      lIncluir = False
      lPrimeiro = False
      Carrega_Grid
   End If
   
   'lIncluir = False

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
    
    .Col = 0:   LblCodfor.Caption = .text: .CellBackColor = vbYellow
    .Col = 1:   TxtNome.text = .text: .CellBackColor = vbYellow
    .Col = 2:   MskCNPJ.text = .text: .CellBackColor = vbYellow
    .Col = 3:   TxtEndereco.text = .text: .CellBackColor = vbYellow
    .Col = 4:   TxtBairro.text = .text: .CellBackColor = vbYellow
    .Col = 5:   TxtCidade.text = .text: .CellBackColor = vbYellow
    .Col = 6:   TxtUf.text = .text: .CellBackColor = vbYellow
    .Col = 7:   Txtcep.text = .text: .CellBackColor = vbYellow
    .Col = 8:   TxtTelefone.text = .text: .CellBackColor = vbYellow
    .Col = 9:   TxtCelular.text = .text: .CellBackColor = vbYellow
    .Col = 10:  TxtContato.text = .text: .CellBackColor = vbYellow
    .Col = 11:  TxtInsc_est.text = .text: .CellBackColor = vbYellow
    .Col = 12:  TxtObserv.text = .text: .CellBackColor = vbYellow
    .TopRow = .Row
    
    
End With


End Sub


Private Sub Carrega_tela()
   'Limpa as variaveis da tela se caso ficarem com dados da outra tela
   limpa_tela Me
   'Carrega a tela com os dados do registro
   Me.LblCodfor.Caption = gRs("idFor")
   Me.TxtNome.text = "" & gRs("Nome")
   Me.TxtEndereco.text = "" & gRs("endereco")
   Me.TxtBairro.text = "" & gRs("bairro").Value
   Me.TxtCidade.text = "" & gRs("Cidade")
   Me.TxtUf.text = gRs("estado")
   Me.Txtcep.text = gRs("cep")
   Me.TxtTelefone.text = "" & gRs("Telefone")
   Me.TxtCelular.text = "" & gRs("celular")
   Me.TxtContato.text = "" & gRs("Contato")
   Me.MskCNPJ.text = "" & gRs("cnpj")
   Me.TxtInsc_est.text = "" & gRs("Insc_est")
   Me.TxtObserv.text = "" & gRs("Observacao")
   
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
         MSFlexGrid1.Col = 0:  MSFlexGrid1.text = f_nulo(!idFor, "")
         MSFlexGrid1.Col = 1:  MSFlexGrid1.text = f_nulo(!nome, "")
         MSFlexGrid1.Col = 2:  MSFlexGrid1.text = f_nulo(!CNPJ, "")
         MSFlexGrid1.Col = 3:  MSFlexGrid1.text = f_nulo(!endereco, "")
         MSFlexGrid1.Col = 4:  MSFlexGrid1.text = f_nulo(!bairro, "")
         MSFlexGrid1.Col = 5:  MSFlexGrid1.text = f_nulo(!Cidade, "")
         MSFlexGrid1.Col = 6:  MSFlexGrid1.text = f_nulo(!estado, "")
         MSFlexGrid1.Col = 7:  MSFlexGrid1.text = f_nulo(!cep, "")
         MSFlexGrid1.Col = 8:  MSFlexGrid1.text = f_nulo(!Telefone, "")
         MSFlexGrid1.Col = 9:  MSFlexGrid1.text = f_nulo(!celular, "")
         MSFlexGrid1.Col = 10: MSFlexGrid1.text = f_nulo(!contato, "")
         MSFlexGrid1.Col = 11: MSFlexGrid1.text = f_nulo(!insc_est, "")
         MSFlexGrid1.Col = 12: MSFlexGrid1.text = f_nulo(!Observacao, "")
         .MoveNext
         
       Loop
       If MSFlexGrid1.Rows > 1 Then
            MSFlexGrid1.FixedRows = 1
        End If
          
  End With
  
  End Sub


Private Sub MskCNPJ_GotFocus()
   MskCNPJ.Mask = "##############"
End Sub

Private Sub MskCnpj_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Or KeyAscii = 9 Then KeyAscii = 0
End Sub

Private Sub MskCNPJ_LostFocus()
   If Len(MskCNPJ.text) > 0 Then
      Select Case Len(MskCNPJ.text)
       Case Is = 11
         MskCNPJ.Mask = "###.###.###-##"
         If Not calculacpf(MskCNPJ.text) Then
            MsgBox "CPF com DV incorreto !!!"
            MskCNPJ = ""
            MskCNPJ.Mask = "##############"
            MskCNPJ.SetFocus
         End If
       Case Is = 14
         MskCNPJ.Mask = "##.###.###/####-##"
         If Not ValidaCGC(MskCNPJ.text) Then
            MsgBox "CGC com DV incorreto !!! "
            MskCNPJ = ""
            MskCNPJ.Mask = "##############"
            MskCNPJ.SetFocus
         End If
      End Select
    End If

End Sub

Private Sub TxtBairro_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Or KeyAscii = 9 Then KeyAscii = 0
End Sub

Private Sub TxtCelular_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Or KeyAscii = 9 Then KeyAscii = 0
End Sub

Private Sub Txtcep_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Or KeyAscii = 9 Then KeyAscii = 0
End Sub

Private Sub TxtCidade_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Or KeyAscii = 9 Then KeyAscii = 0
End Sub

Private Sub TxtContato_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Or KeyAscii = 9 Then KeyAscii = 0
End Sub

Private Sub TxtEndereco_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Or KeyAscii = 9 Then KeyAscii = 0
End Sub

Private Sub TxtInsc_est_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Or KeyAscii = 9 Then KeyAscii = 0
End Sub

Private Sub TxtNome_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Or KeyAscii = 9 Then KeyAscii = 0
End Sub

Private Sub TxtObserv_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Or KeyAscii = 9 Then KeyAscii = 0
End Sub

Private Sub TxtTelefone_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Or KeyAscii = 9 Then KeyAscii = 0
End Sub

Private Sub TxtUf_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Or KeyAscii = 9 Then KeyAscii = 0
End Sub
