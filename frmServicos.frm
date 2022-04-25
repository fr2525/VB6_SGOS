VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmServicos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Manutenção de Serviços"
   ClientHeight    =   4575
   ClientLeft      =   4545
   ClientTop       =   2445
   ClientWidth     =   7245
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4575
   ScaleWidth      =   7245
   ShowInTaskbar   =   0   'False
   Tag             =   "cadvend"
   Begin VB.TextBox TxtPcovenda 
      Height          =   315
      Left            =   5295
      TabIndex        =   5
      Top             =   585
      Width           =   1440
   End
   Begin VB.TextBox TxtPcocusto 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """R$ ""#.##0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   2
      EndProperty
      Enabled         =   0   'False
      Height          =   300
      Left            =   2775
      TabIndex        =   4
      Top             =   585
      Width           =   1440
   End
   Begin VB.TextBox TxtUnidade 
      Enabled         =   0   'False
      Height          =   285
      Left            =   870
      TabIndex        =   3
      Top             =   585
      Width           =   885
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   2085
      Left            =   510
      TabIndex        =   9
      Top             =   1230
      Width           =   6225
      _ExtentX        =   10980
      _ExtentY        =   3678
      _Version        =   393216
      Rows            =   10
      Cols            =   5
      FixedCols       =   0
      FormatString    =   "Código|Descrição                                                              |Unidade| Pço.custo |Pço.Venda   "
   End
   Begin VB.Frame Frame1 
      Height          =   795
      Left            =   1410
      TabIndex        =   8
      Top             =   3510
      Width           =   4245
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Salvar"
         Enabled         =   0   'False
         Height          =   540
         Left            =   2220
         Picture         =   "frmServicos.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   6
         Tag             =   "&Update"
         Top             =   180
         Width           =   615
      End
      Begin VB.CommandButton cmdEditar 
         Caption         =   "&Editar"
         Height          =   540
         Left            =   780
         Picture         =   "frmServicos.frx":00FA
         Style           =   1  'Graphical
         TabIndex        =   14
         Tag             =   "&Refresh"
         Top             =   180
         Width           =   615
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Excluir"
         Height          =   540
         Left            =   1500
         Picture         =   "frmServicos.frx":026C
         Style           =   1  'Graphical
         TabIndex        =   13
         Tag             =   "&Delete"
         Top             =   180
         Width           =   615
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Incluir"
         Height          =   540
         Left            =   120
         Picture         =   "frmServicos.frx":03DE
         Style           =   1  'Graphical
         TabIndex        =   12
         Tag             =   "&Add"
         Top             =   180
         Width           =   615
      End
      Begin VB.CommandButton CmdSair 
         Caption         =   "&Sair"
         Height          =   540
         Left            =   3540
         Picture         =   "frmServicos.frx":04C8
         Style           =   1  'Graphical
         TabIndex        =   11
         Tag             =   "&Update"
         Top             =   180
         Width           =   615
      End
      Begin VB.CommandButton cmddesfaz 
         Caption         =   "&Desfaz"
         Enabled         =   0   'False
         Height          =   540
         Left            =   2880
         Picture         =   "frmServicos.frx":05C2
         Style           =   1  'Graphical
         TabIndex        =   10
         Tag             =   "&Update"
         Top             =   180
         Width           =   615
      End
   End
   Begin VB.TextBox TxtDescricao 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2775
      TabIndex        =   2
      Top             =   195
      Width           =   3975
   End
   Begin VB.Label Label3 
      Caption         =   "Pço.Venda"
      Height          =   285
      Left            =   4350
      TabIndex        =   17
      Top             =   615
      Width           =   840
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Pço.custo:"
      Height          =   195
      Left            =   1935
      TabIndex        =   16
      Top             =   660
      Width           =   765
   End
   Begin VB.Label Label1 
      Caption         =   "Unidade:"
      Height          =   195
      Left            =   105
      TabIndex        =   15
      Top             =   630
      Width           =   675
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Código:"
      Height          =   195
      Index           =   0
      Left            =   210
      TabIndex        =   7
      Tag             =   "CODVEND:"
      Top             =   225
      Width           =   540
   End
   Begin VB.Label LblCod_serv 
      Caption         =   "cod_servico"
      ForeColor       =   &H000000FF&
      Height          =   210
      Left            =   900
      TabIndex        =   0
      Top             =   225
      Width           =   975
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Descrição"
      Height          =   195
      Index           =   1
      Left            =   1905
      TabIndex        =   1
      Tag             =   "NOME:"
      Top             =   240
      Width           =   720
   End
End
Attribute VB_Name = "frmServicos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private lIncluir As Boolean
Private lPrimeiro As Boolean

Private Sub Abre_Le_rst()
  gSql = "select * FROM tab_servicos"
  gRs.Open gSql, ConDb, adOpenDynamic
 
End Sub

Private Sub cmdAdd_Click()
   
   limpa_tela Me
   
   Me.LblCod_serv.Caption = ""
   Me.TxtDescricao.SetFocus
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
    If MsgBox("Deseja realmente apagar este Serviço ? ", vbYesNo, "Atenção") = vbYes Then
       gSql = "delete * from tab_servicos where cod_servico = " & Me.LblCod_serv.Caption
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
     MsgBox "Deu erro na exclusao do Serviço " & Chr(13) & "Instrucao Sql = '" & _
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
   Me.TxtDescricao.SetFocus
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
      gSql = "INSERT INTO tab_servicos (descricao,unidade,preco_custo,"
      gSql = gSql & "preco_venda,operador, datatual"
      gSql = gSql & ") "
      gSql = gSql & "VALUES ('" & Me.TxtDescricao.Text & "','"
      gSql = gSql & Me.TxtUnidade.Text & "',"
      gSql = gSql & Val(Me.TxtPcocusto.Text) & ","
      gSql = gSql & Val(Me.TxtPcovenda.Text)
      gSql = gSql & ",'" & gOperador & "',Cdate('" & Date & "'))"
      ConDb.Execute gSql
      lIncluir = False
   Else
      gSql = "UPDATE tab_servicos SET descricao = '" & Me.TxtDescricao.Text & "',"
      gSql = gSql & "unidade = " & Val(Me.TxtUnidade.Text) & ","
      gSql = gSql & "preco_custo = " & Val(Me.TxtPcocusto.Text) & ","
      gSql = gSql & "preco_custo = " & Val(Me.TxtPcovenda.Text) & ","
      gSql = gSql & " operador = '" & gOperador & "', datatual = Cdate('" & Date & "'))"
      gSql = gSql & " WHERE cod_servico = " & Val(Me.LblCod_serv.Caption)
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
   
   Me.LblCod_serv.Caption = ""
   If gRs.BOF And gRs.EOF Then
      If MsgBox("Arquivo vazio. Incluir dados agora ?", vbYesNo, "Atenção ") = vbYes Then
         gSql = "INSERT INTO tab_servicos (descricao,unidade,preco,"
         gSql = gSql & "operador, datatual"
         gSql = gSql & ") "
         gSql = gSql & "VALUES ('" & f_nulo(Me.TxtDescricao.Text, " ") & "','"
         gSql = gSql & f_nulo(Me.TxtUnidade.Text, " ") & "',"
         gSql = gSql & Val(Me.TxtPcocusto.Text) & ","
         gSql = gSql & Val(Me.TxtPcovenda.Text) & ",'"
         gSql = gSql & gOperador & "'," & Date & " )"
         ConDb.Execute gSql
         gRs.Close
         Abre_Le_rst
         Me.LblCod_serv.Caption = gRs!cod_serv
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
    
    .Col = 0:   LblCod_serv.Caption = .Text: .CellBackColor = vbYellow
    .Col = 1:   TxtDescricao.Text = .Text: .CellBackColor = vbYellow
    .Col = 2:   TxtUnidade.Text = .Text: .CellBackColor = vbYellow
    .Col = 3:   TxtPcocusto.Text = .Text: .CellBackColor = vbYellow
    .Col = 4:   TxtPcovenda.Text = .Text: .CellBackColor = vbYellow
    .TopRow = .Row
    
    
End With


End Sub


Private Sub Carrega_tela()
   'Limpa as variaveis da tela se caso ficarem com dados da outra tela
   limpa_tela Me
   'Carrega a tela com os dados do registro
   Me.LblCod_serv.Caption = gRs("cod_servico")
   Me.TxtDescricao.Text = "" & gRs("descricao")
   Me.TxtUnidade.Text = "" & gRs("unidade")
   Me.TxtPcocusto.Text = "" & Format(gRs("preco_custo"), "R$#,000.00")
   Me.TxtPcovenda.Text = "" & Format(gRs("preco_venda"), "R$#,000.00")
   
   
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
            
         MSFlexGrid1.Col = 0:  MSFlexGrid1.Text = !cod_servico
         MSFlexGrid1.Col = 1:  MSFlexGrid1.Text = f_nulo(!descricao, "")
         MSFlexGrid1.Col = 2:  MSFlexGrid1.Text = f_nulo(!unidade, "")
         MSFlexGrid1.Col = 3:  MSFlexGrid1.Text = Format(!preco_custo, "R$#,##0.00;(R$#,##0.00)")
         MSFlexGrid1.Col = 4:  MSFlexGrid1.Text = Format(!preco_venda, "R$#,##0.00;(R$#,##0.00)")
         .MoveNext
         
       Loop
       MSFlexGrid1.FixedRows = 1
          
  End With

End Sub

