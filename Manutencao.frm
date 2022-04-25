VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmManut 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Manutenção"
   ClientHeight    =   5655
   ClientLeft      =   4545
   ClientTop       =   2445
   ClientWidth     =   7245
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5655
   ScaleWidth      =   7245
   ShowInTaskbar   =   0   'False
   Tag             =   "cadvend"
   Begin VB.TextBox TxtCep 
      Height          =   300
      Left            =   5505
      TabIndex        =   7
      Top             =   720
      Width           =   1245
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   2085
      Left            =   570
      TabIndex        =   22
      Top             =   1920
      Width           =   6225
      _ExtentX        =   10980
      _ExtentY        =   3678
      _Version        =   393216
      Rows            =   5
      Cols            =   10
      FixedCols       =   0
      FormatString    =   $"Manutencao.frx":0000
   End
   Begin VB.Frame Frame1 
      Height          =   795
      Left            =   1560
      TabIndex        =   21
      Top             =   4320
      Width           =   4245
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Salvar"
         Enabled         =   0   'False
         Height          =   540
         Left            =   2220
         Picture         =   "Manutencao.frx":00D6
         Style           =   1  'Graphical
         TabIndex        =   11
         Tag             =   "&Update"
         Top             =   180
         Width           =   615
      End
      Begin VB.CommandButton cmdEditar 
         Caption         =   "&Editar"
         Height          =   540
         Left            =   780
         Picture         =   "Manutencao.frx":01D0
         Style           =   1  'Graphical
         TabIndex        =   27
         Tag             =   "&Refresh"
         Top             =   180
         Width           =   615
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Excluir"
         Height          =   540
         Left            =   1500
         Picture         =   "Manutencao.frx":0342
         Style           =   1  'Graphical
         TabIndex        =   26
         Tag             =   "&Delete"
         Top             =   180
         Width           =   615
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Incluir"
         Height          =   540
         Left            =   120
         Picture         =   "Manutencao.frx":04B4
         Style           =   1  'Graphical
         TabIndex        =   25
         Tag             =   "&Add"
         Top             =   180
         Width           =   615
      End
      Begin VB.CommandButton CmdSair 
         Caption         =   "&Sair"
         Height          =   540
         Left            =   3540
         Picture         =   "Manutencao.frx":059E
         Style           =   1  'Graphical
         TabIndex        =   24
         Tag             =   "&Update"
         Top             =   180
         Width           =   615
      End
      Begin VB.CommandButton cmddesfaz 
         Caption         =   "&Desfaz"
         Enabled         =   0   'False
         Height          =   540
         Left            =   2880
         Picture         =   "Manutencao.frx":0698
         Style           =   1  'Graphical
         TabIndex        =   23
         Tag             =   "&Update"
         Top             =   180
         Width           =   615
      End
   End
   Begin VB.TextBox TxtContato 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1320
      TabIndex        =   10
      Top             =   1440
      Width           =   3105
   End
   Begin VB.TextBox TxtBairro 
      Height          =   285
      Left            =   4350
      TabIndex        =   4
      Top             =   360
      Width           =   2400
   End
   Begin VB.TextBox TxtCidade 
      Height          =   285
      Left            =   1350
      TabIndex        =   5
      Top             =   720
      Width           =   2280
   End
   Begin VB.TextBox TxtUf 
      Height          =   285
      Left            =   4380
      TabIndex        =   6
      Top             =   720
      Width           =   450
   End
   Begin VB.TextBox TxtTelefone 
      Height          =   285
      Left            =   1350
      TabIndex        =   8
      Top             =   1080
      Width           =   1890
   End
   Begin VB.TextBox TxtCelular 
      Height          =   285
      Left            =   4860
      TabIndex        =   9
      Top             =   1110
      Width           =   1890
   End
   Begin VB.TextBox TxtEndereco 
      Height          =   285
      Left            =   1365
      TabIndex        =   3
      Top             =   360
      Width           =   2310
   End
   Begin VB.TextBox TxtNome 
      Height          =   285
      Left            =   3360
      TabIndex        =   2
      Top             =   30
      Width           =   3375
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Contato"
      Height          =   210
      Left            =   540
      TabIndex        =   13
      Top             =   1470
      Width           =   645
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Código:"
      Height          =   195
      Index           =   0
      Left            =   705
      TabIndex        =   20
      Tag             =   "CODVEND:"
      Top             =   45
      Width           =   540
   End
   Begin VB.Label LblBairro 
      Caption         =   "Bairro"
      Height          =   240
      Left            =   3840
      TabIndex        =   19
      Top             =   390
      Width           =   420
   End
   Begin VB.Label LblCidade 
      Caption         =   "Cidade"
      Height          =   240
      Left            =   660
      TabIndex        =   18
      Top             =   750
      Width           =   540
   End
   Begin VB.Label LblUf 
      Caption         =   "Estado"
      Height          =   195
      Left            =   3720
      TabIndex        =   17
      Top             =   780
      Width           =   555
   End
   Begin VB.Label LblTelefone 
      Caption         =   "Telefone"
      Height          =   240
      Left            =   570
      TabIndex        =   16
      Top             =   1110
      Width           =   630
   End
   Begin VB.Label LblCelular 
      Alignment       =   1  'Right Justify
      Caption         =   "Celular"
      Height          =   195
      Left            =   4290
      TabIndex        =   15
      Top             =   1170
      Width           =   480
   End
   Begin VB.Label lblCep 
      Alignment       =   1  'Right Justify
      Caption         =   "CEP"
      Height          =   255
      Left            =   4920
      TabIndex        =   14
      Top             =   780
      Width           =   465
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Endereço:"
      Height          =   195
      Index           =   2
      Left            =   570
      TabIndex        =   12
      Tag             =   "NOME:"
      Top             =   360
      Width           =   735
   End
   Begin VB.Label LblCodfor 
      Caption         =   "codfor"
      ForeColor       =   &H000000FF&
      Height          =   210
      Left            =   1410
      TabIndex        =   0
      Top             =   60
      Width           =   975
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Nome:"
      Height          =   195
      Index           =   1
      Left            =   2700
      TabIndex        =   1
      Tag             =   "NOME:"
      Top             =   60
      Width           =   465
   End
End
Attribute VB_Name = "frmManut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private lIncluir As Boolean
Private lPrimeiro As Boolean

Private Sub Abre_Le_rst()
  gSql = "select * FROM tab_fornece"
  gRs.Open gSql, ConDb, adOpenKeyset
  
End Sub

Private Sub cmdAdd_Click()
   
   limpa_tela Me
   
   Me.LblCodfor.Caption = ""
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
      gSql = "INSERT INTO tab_fornece (Nome,endereco,bairro,cidade,estado,"
      gSql = gSql & "CEP,telefone,fax,contato, operador, datatual"
      gSql = gSql & ") "
      gSql = gSql & "VALUES ('" & Me.TxtNome.Text & "','"
      gSql = gSql & Me.TxtEndereco.Text & "','"
      gSql = gSql & Me.TxtBairro.Text & "','"
      gSql = gSql & Me.TxtCidade.Text & "','"
      gSql = gSql & Me.TxtUf.Text & "','"
      gSql = gSql & Me.TxtCep.Text & "','"
      gSql = gSql & Me.TxtTelefone.Text & "','"
      gSql = gSql & Me.TxtCelular.Text & "','"
      gSql = gSql & Me.TxtContato.Text & "'"
      gSql = gSql & ",'" & gOperador & "'," & Date & " )"
      ConDb.Execute gSql
      lIncluir = False
   Else
      gSql = "UPDATE tab_fornece SET Nome = '" & Me.TxtNome.Text & "',"
      gSql = gSql & "endereco = '" & Me.TxtEndereco.Text & "',"
      gSql = gSql & "bairro = '" & Me.TxtBairro.Text & "',"
      gSql = gSql & "cidade = '" & Me.TxtCidade.Text & "',"
      gSql = gSql & "estado = '" & Me.TxtUf.Text & "',"
      gSql = gSql & "CEP = '" & Me.TxtCep.Text & "',"
      gSql = gSql & "Telefone = '" & Me.TxtTelefone.Text & "',"
      gSql = gSql & "Fax = '" & Me.TxtCelular.Text & "',"
      gSql = gSql & "contato = '" & Me.TxtContato.Text & "'"
      gSql = gSql & " ,operador = '" & gOperador & "', datatual = " & Date
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

Private Sub Form_Activate()
   Abre_Le_rst
   
   Me.LblCodfor.Caption = ""
    If gRs.BOF And gRs.EOF Then
      If MsgBox("Arquivo vazio. Incluir dados agora ?", vbYesNo, "Atenção ") = vbYes Then
         gSql = "INSERT INTO tab_fornece (Nome,endereco,bairro,cidade,estado,"
         gSql = gSql & "CEP,telefone,fax,contato, operador, datatual"
         gSql = gSql & ") "
         gSql = gSql & "VALUES ('" & Me.TxtNome.Text & "','"
         gSql = gSql & Me.TxtEndereco.Text & "','"
         gSql = gSql & Me.TxtBairro.Text & "','"
         gSql = gSql & Me.TxtCidade.Text & "','"
         gSql = gSql & Me.TxtUf.Text & "','"
         gSql = gSql & Me.TxtCep.Text & "','"
         gSql = gSql & Me.TxtTelefone.Text & "','"
         gSql = gSql & Me.TxtCelular.Text & "','"
         gSql = gSql & Me.TxtContato.Text & "'"
         gSql = gSql & ",'" & gOperador & "'," & Date & " )"
         ConDb.Execute gSql
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

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
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

Private Sub MSFlexGrid1_Click()
 Dim oldrow As Long
  
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
    .Col = 2:   TxtEndereco.Text = .Text: .CellBackColor = vbYellow
    .Col = 3:   TxtBairro.Text = .Text: .CellBackColor = vbYellow
    .Col = 4:   TxtCidade.Text = .Text: .CellBackColor = vbYellow
    .Col = 5:   TxtUf.Text = .Text: .CellBackColor = vbYellow
    .Col = 6:   TxtCep.Text = .Text: .CellBackColor = vbYellow
    .Col = 7:   TxtTelefone.Text = .Text: .CellBackColor = vbYellow
    .Col = 8:   TxtCelular.Text = .Text: .CellBackColor = vbYellow
    .Col = 9:   TxtContato.Text = .Text: .CellBackColor = vbYellow
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
   Me.TxtCep.Text = gRs("cep")
   Me.TxtTelefone.Text = "" & gRs("Telefone")
   Me.TxtCelular.Text = "" & gRs("Fax")
   Me.TxtContato.Text = "" & gRs("Contato")
   
   
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
         MSFlexGrid1.Col = 2:  MSFlexGrid1.Text = f_nulo(!endereco, "")
         MSFlexGrid1.Col = 3:  MSFlexGrid1.Text = f_nulo(!bairro, "")
         MSFlexGrid1.Col = 4:  MSFlexGrid1.Text = f_nulo(!Cidade, "")
         MSFlexGrid1.Col = 5:  MSFlexGrid1.Text = f_nulo(!estado, "")
         MSFlexGrid1.Col = 6:  MSFlexGrid1.Text = f_nulo(!cep, "")
         MSFlexGrid1.Col = 7:  MSFlexGrid1.Text = f_nulo(!Telefone, "")
         MSFlexGrid1.Col = 8: MSFlexGrid1.Text = f_nulo(!Fax, "")
         MSFlexGrid1.Col = 9: MSFlexGrid1.Text = f_nulo(!contato, "")
                  .MoveNext
         
       Loop
       MSFlexGrid1.FixedRows = 1
          
  End With
  
  End Sub

