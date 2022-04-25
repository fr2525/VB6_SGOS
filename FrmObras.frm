VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form FrmObras 
   Caption         =   "Obras"
   ClientHeight    =   4605
   ClientLeft      =   3990
   ClientTop       =   3225
   ClientWidth     =   5970
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   4605
   ScaleWidth      =   5970
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   780
      TabIndex        =   9
      Top             =   3600
      Width           =   4245
      Begin VB.CommandButton cmddesfaz 
         Caption         =   "&Desfaz"
         Enabled         =   0   'False
         Height          =   540
         Left            =   2865
         Picture         =   "FrmObras.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   15
         Tag             =   "&Update"
         Top             =   150
         Width           =   615
      End
      Begin VB.CommandButton CmdSair 
         Caption         =   "&Sair"
         Height          =   540
         Left            =   3555
         Picture         =   "FrmObras.frx":00FA
         Style           =   1  'Graphical
         TabIndex        =   14
         Tag             =   "&Update"
         Top             =   150
         Width           =   615
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Incluir"
         Height          =   540
         Left            =   105
         Picture         =   "FrmObras.frx":01F4
         Style           =   1  'Graphical
         TabIndex        =   13
         Tag             =   "&Add"
         Top             =   150
         Width           =   615
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Excluir"
         Height          =   540
         Left            =   1485
         Picture         =   "FrmObras.frx":02DE
         Style           =   1  'Graphical
         TabIndex        =   12
         Tag             =   "&Delete"
         Top             =   150
         Width           =   615
      End
      Begin VB.CommandButton cmdEditar 
         Caption         =   "&Editar"
         Height          =   540
         Left            =   795
         Picture         =   "FrmObras.frx":0450
         Style           =   1  'Graphical
         TabIndex        =   11
         Tag             =   "&Refresh"
         Top             =   150
         Width           =   615
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Salvar"
         Enabled         =   0   'False
         Height          =   540
         Left            =   2190
         Picture         =   "FrmObras.frx":05C2
         Style           =   1  'Graphical
         TabIndex        =   10
         Tag             =   "&Update"
         Top             =   150
         Width           =   615
      End
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   1395
      Left            =   240
      TabIndex        =   8
      Top             =   1950
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   2461
      _Version        =   393216
      Rows            =   5
      Cols            =   4
      FixedCols       =   0
      FormatString    =   $"FrmObras.frx":06BC
   End
   Begin VB.TextBox TxtObservacoes 
      Height          =   585
      Left            =   1260
      MultiLine       =   -1  'True
      TabIndex        =   7
      Top             =   1170
      Width           =   4365
   End
   Begin VB.TextBox TxtContratante 
      Height          =   285
      Left            =   1260
      TabIndex        =   5
      Top             =   810
      Width           =   4335
   End
   Begin VB.TextBox TxtUnidade 
      Height          =   285
      Left            =   1260
      TabIndex        =   3
      Top             =   420
      Width           =   4335
   End
   Begin VB.Label Label5 
      Caption         =   "Observ.:"
      Height          =   255
      Left            =   390
      TabIndex        =   6
      Top             =   1200
      Width           =   705
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Contratante:"
      Height          =   195
      Left            =   240
      TabIndex        =   4
      Top             =   840
      Width           =   870
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Unidade:"
      Height          =   195
      Left            =   450
      TabIndex        =   2
      Top             =   480
      Width           =   645
   End
   Begin VB.Label LblCodObra 
      AutoSize        =   -1  'True
      Caption         =   "Label2"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   1290
      TabIndex        =   1
      Top             =   180
      Width           =   600
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Cod.Obra:"
      Height          =   195
      Left            =   390
      TabIndex        =   0
      Top             =   180
      Width           =   720
   End
End
Attribute VB_Name = "FrmObras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private lIncluir As Boolean
Private lPrimeiro As Boolean

Private Sub Abre_Le_rst()
   gSql = "select * FROM tab_obras"
   gRs.Open gSql, ConDb, adOpenKeyset, adLockOptimistic
   
End Sub
Private Sub Carrega_Grid()

'Teste do MsHFlexgrid1
  MSFlexGrid1.Row = 0
  
  With gRs
      '.MoveLast
      'nItem = .RecordCount
      '.MoveFirst
      MSFlexGrid1.Rows = 1
      Do While Not .EOF
         MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
         MSFlexGrid1.Row = MSFlexGrid1.Rows - 1
            
         MSFlexGrid1.Col = 0:  MSFlexGrid1.Text = f_nulo(!cod_obra, "")
         MSFlexGrid1.Col = 1:  MSFlexGrid1.Text = f_nulo(!unidade, "")
         MSFlexGrid1.Col = 2:  MSFlexGrid1.Text = f_nulo(!contratante, "")
         MSFlexGrid1.Col = 3:  MSFlexGrid1.Text = f_nulo(!observacoes, "")
         .MoveNext
         
       Loop
       MSFlexGrid1.FixedRows = 1
          
  End With
  
  End Sub
Private Sub Carrega_tela()
   'Limpa as variaveis da tela se caso ficarem com dados da outra tela
   limpa_tela Me
   'Carrega a tela com os dados do registro
   Me.LblCodObra = gRs("cod_obra")
   If Not IsNull(gRs("unidade")) Then Me.TxtUnidade.Text = gRs("unidade")
   If Not IsNull(gRs("contratante")) Then Me.TxtContratante.Text = gRs!contratante
   If Not IsNull(gRs("observacoes")) Then Me.TxtObservacoes.Text = "" & gRs!observacoes
      
End Sub

Private Sub cmdAdd_Click()
   lIncluir = True
   limpa_tela Me
   
   Me.LblCodObra.Caption = ""
   Me.cmdUpdate.Enabled = True
   Me.cmddesfaz.Enabled = True
   Me.cmdEditar.Enabled = False
   Me.cmdAdd.Enabled = False
   Me.CmdSair.Enabled = False
   Me.cmdDelete.Enabled = False
   Me.TxtUnidade.SetFocus

End Sub

Private Sub cmdDelete_Click()
   On Error GoTo ErroDelete
    
    'this may produce an error if you delete the last
    'record or the only record in the recordset
    If MsgBox("Deseja realmente apagar esta Obra? ", vbYesNo, "Atenção") = vbYes Then
        gSql = "delete * from tab_obras where cod_obra = " & Val(Me.LblCodObra.Caption)
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
     MsgBox "Deu erro na exclusao da Obra " & Chr(13) & "Instrucao Sql = '" & _
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
   Habilita Me
   Me.cmdUpdate.Enabled = True
   Me.cmddesfaz.Enabled = True
   Me.cmdEditar.Enabled = False
   Me.cmdAdd.Enabled = False
   Me.CmdSair.Enabled = False
   Me.cmdDelete.Enabled = False
   'Me.TxtUnidade.SetFocus
   
End Sub

Private Sub CmdSair_Click()
   Unload Me
End Sub

Private Sub cmdUpdate_Click()
   gRs.Close
   If lIncluir Then
      gSql = "INSERT INTO Tab_obras (unidade,contratante,observacoes,operador,datatual) "
      gSql = gSql & "VALUES ('" & Me.TxtUnidade.Text & "','"
      gSql = gSql & Me.TxtContratante.Text & "','"
      gSql = gSql & Me.TxtObservacoes.Text & "','"
      gSql = gSql & gOperador & "',Cdate('" & Date & "')) "
      ConDb.Execute gSql
      lIncluir = False
      
   Else
      gSql = "UPDATE Tab_obras SET unidade = '" & Me.TxtUnidade.Text
      gSql = gSql & "', Contratante = '" & Me.TxtContratante
      gSql = gSql & "', observacoes = '" & Me.TxtObservacoes
      gSql = gSql & "', operador = '" & gOperador & "', datatual = Cdate('" & Date & "')"
      gSql = gSql & " WHERE cod_obra = " & Val(Me.LblCodObra.Caption)
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
    
   limpa_tela Me
   
   Me.LblCodObra.Caption = ""
   If gRs.BOF And gRs.EOF Then
      If MsgBox("Arquivo vazio. Incluir dados agora ?", vbYesNo, "Atenção ") = vbYes Then
         'rst.AddNew
         gSql = "INSERT INTO Tab_obras (unidade,contratante,observacoes,operador,datatual) "
         gSql = gSql & "VALUES ('" & Me.TxtUnidade.Text & "','"
         gSql = gSql & Me.TxtContratante.Text & "','"
         gSql = gSql & Me.TxtObservacoes.Text & "','"
         gSql = gSql & gOperador & "'," & Date & " ) "
         ConDb.Execute gSql
         gRs.Close
         Abre_Le_rst
         Me.LblCodObra.Caption = gRs!cod_obra
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
  
    .Refresh
    .Row = oldrow
    
    .Col = 0:   LblCodObra.Caption = .Text:  .CellBackColor = vbYellow
    .Col = 1:   TxtUnidade.Text = .Text:        .CellBackColor = vbYellow
    .Col = 2:   TxtContratante.Text = .Text:    .CellBackColor = vbYellow
    .Col = 3:   TxtObservacoes.Text = .Text:      .CellBackColor = vbYellow
    
    .Redraw = True
    
  End With

End Sub

