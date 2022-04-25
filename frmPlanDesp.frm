VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmPlanDesp 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Planilha de Despesas"
   ClientHeight    =   5505
   ClientLeft      =   4170
   ClientTop       =   2445
   ClientWidth     =   6345
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5505
   ScaleWidth      =   6345
   ShowInTaskbar   =   0   'False
   Tag             =   "cadvend"
   Begin VB.TextBox TxtDataDespesa 
      Height          =   315
      Left            =   855
      TabIndex        =   20
      Top             =   1635
      Width           =   990
   End
   Begin MSMask.MaskEdBox MskValor 
      Height          =   315
      Left            =   4245
      TabIndex        =   18
      Top             =   1680
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   556
      _Version        =   393216
      Format          =   "R$#,##0.00;(R$#,##0.00)"
      PromptChar      =   "_"
   End
   Begin VB.TextBox Txtdescricao 
      Height          =   285
      Left            =   870
      TabIndex        =   17
      Top             =   1290
      Width           =   4905
   End
   Begin VB.Frame Frame2 
      Caption         =   "Dados da Obra"
      Height          =   1005
      Left            =   0
      TabIndex        =   8
      Top             =   90
      Width           =   5865
      Begin VB.Label LblObra 
         Caption         =   "Unidade"
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   1035
         TabIndex        =   14
         Top             =   285
         Width           =   4725
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Obra"
         Height          =   195
         Left            =   450
         TabIndex        =   13
         Top             =   285
         Width           =   345
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Medição"
         Height          =   195
         Left            =   210
         TabIndex        =   12
         Top             =   585
         Width           =   615
      End
      Begin VB.Label LblMedicao 
         Caption         =   "Num.Medição"
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   1050
         TabIndex        =   11
         Top             =   615
         Width           =   1245
      End
      Begin VB.Label Label4 
         Caption         =   "Dta.Medição"
         Height          =   225
         Left            =   3435
         TabIndex        =   10
         Top             =   645
         Width           =   1005
      End
      Begin VB.Label LblDta_medicao 
         Caption         =   "99/99/9999"
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   4515
         TabIndex        =   9
         Top             =   645
         Width           =   1245
      End
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   2085
      Left            =   210
      TabIndex        =   7
      Top             =   2190
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   3678
      _Version        =   393216
      Rows            =   5
      Cols            =   3
      FixedCols       =   0
      FormatString    =   "Descrição                                                                               |Valor            |"
   End
   Begin VB.Frame Frame1 
      Height          =   795
      Left            =   990
      TabIndex        =   1
      Top             =   4455
      Width           =   4245
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Salvar"
         Enabled         =   0   'False
         Height          =   540
         Left            =   2220
         Picture         =   "frmPlanDesp.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   0
         Tag             =   "&Update"
         Top             =   180
         Width           =   615
      End
      Begin VB.CommandButton cmdEditar 
         Caption         =   "&Editar"
         Height          =   540
         Left            =   780
         Picture         =   "frmPlanDesp.frx":00FA
         Style           =   1  'Graphical
         TabIndex        =   6
         Tag             =   "&Refresh"
         Top             =   180
         Width           =   615
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Excluir"
         Height          =   540
         Left            =   1500
         Picture         =   "frmPlanDesp.frx":026C
         Style           =   1  'Graphical
         TabIndex        =   5
         Tag             =   "&Delete"
         Top             =   180
         Width           =   615
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Incluir"
         Height          =   540
         Left            =   120
         Picture         =   "frmPlanDesp.frx":03DE
         Style           =   1  'Graphical
         TabIndex        =   4
         Tag             =   "&Add"
         Top             =   180
         Width           =   615
      End
      Begin VB.CommandButton CmdSair 
         Caption         =   "&Sair"
         Height          =   540
         Left            =   3540
         Picture         =   "frmPlanDesp.frx":04C8
         Style           =   1  'Graphical
         TabIndex        =   3
         Tag             =   "&Update"
         Top             =   180
         Width           =   615
      End
      Begin VB.CommandButton cmddesfaz 
         Caption         =   "&Desfaz"
         Enabled         =   0   'False
         Height          =   540
         Left            =   2880
         Picture         =   "frmPlanDesp.frx":05C2
         Style           =   1  'Graphical
         TabIndex        =   2
         Tag             =   "&Update"
         Top             =   180
         Width           =   615
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Data:"
      Height          =   195
      Left            =   315
      TabIndex        =   19
      Top             =   1710
      Width           =   390
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Descricao"
      Height          =   195
      Left            =   60
      TabIndex        =   16
      Top             =   1320
      Width           =   720
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Valor R$"
      Height          =   195
      Left            =   3495
      TabIndex        =   15
      Top             =   1755
      Width           =   615
   End
End
Attribute VB_Name = "frmPlanDesp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private lIncluir As Boolean
Dim prsServico As New ADODB.Recordset

Private Sub Abre_Le_rst()
  gSql = "select descricao,valor,dta_despesa FROM plandesp as A"
  gSql = gSql & " WHERE cod_obra = " & gCodObra
  gSql = gSql & " AND medicao = " & gNumMedicao
  gSql = gSql & " AND dta_medicao = " & gDataMedicao
  gRs.Open gSql, ConDb, adOpenKeyset
  
End Sub

Private Sub cmdAdd_Click()
   suBotao_add
End Sub

Private Sub cmdDelete_Click()

    On Error GoTo ErroDelete
    
    'this may produce an error if you delete the last
    'record or the only record in the recordset
    If MsgBox("Deseja realmente apagar este Item de despesa ? ", vbYesNo, "Atenção") = vbYes Then
       gSql = "delete * from plandesp "
       gSql = gSql & " WHERE cod_obra = " & gCodObra
       gSql = gSql & " AND medicao = " & gNumMedicao
       gSql = gSql & " AND dta_medicao = gdatamedicao"
       gSql = gSql & " AND descricao = '" & TxtDescricao & "'"
       gSql = gSql & " AND dta_despesa = " & TxtDataDespesa
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
     MsgBox "Deu erro na exclusao do Item de Despesa " & Chr(13) & "Instrucao Sql = '" & _
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
      gSql = "INSERT INTO Plandesp (cod_obra,medicao,dta_medicao,"
      gSql = gSql & "descricao,valor,dta_despesa,operador, datatual"
      gSql = gSql & ") "
      gSql = gSql & "VALUES ( " & gCodObra & ","
      gSql = gSql & gNumMedicao & ","
      gSql = gSql & gDataMedicao & ",'"
      gSql = gSql & Me.TxtDescricao & "',"
      gSql = gSql & Me.MskValor.Text & ","
      gSql = gSql & Me.TxtDataDespesa.Text & ",'"
      gSql = gSql & gOperador & "'," & Date & " )"
      ConDb.Execute gSql
      lIncluir = False
   Else
      gSql = "UPDATE Plandesp SET descricao = '" & Me.TxtDescricao & "',"
      gSql = gSql & "valor = " & Val(Me.MskValor.Text) & ","
      gSql = gSql & "dta_despesa = " & Me.TxtDataDespesa.Text & ","
      gSql = gSql & "operador = '" & gOperador & "', datatual = " & Date
      gSql = gSql & " WHERE cod_obra  = " & gCodObra
      gSql = gSql & " AND medicao = " & gNumMedicao
      gSql = gSql & " AND dta_medicao = " & gDataMedicao
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
       
   If gRs.BOF And gRs.EOF Then
      If MsgBox("Arquivo vazio. Incluir dados agora ?", vbYesNo, "Atenção ") = vbYes Then
         cmdEditar_Click
         lIncluir = True
         lPrimeiro = True
         Me.TxtDescricao.SetFocus
      Else
         Desabilita Me
         Unload Me
      End If
   Else
      gRs.MoveFirst
      Carrega_tela
      Desabilita Me
      lIncluir = False
      lPrimeiro = False
      Carrega_Grid
   End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub Form_Load()
  'Centraliza a tela no video
   Me.Move (Screen.Width - Me.Width) / 2, _
           (Screen.Height - Me.Height) / 2
    LblObra.Caption = gUnidade
    LblMedicao = gNumMedicao
    LblDta_medicao = gDataMedicao

End Sub

Private Sub Form_Unload(Cancel As Integer)
    gRs.Close
    Screen.MousePointer = vbDefault
    'Unload FrmPlanilha
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
     
    .Col = 0:  Me.TxtDescricao.Text = .Text: .CellBackColor = vbYellow
    .Col = 1:  Me.MskValor.Text = Format(.Text, "R$#,##0.00;(R$#,##0.00)"): .CellBackColor = vbYellow
    .Col = 2:  Me.TxtDataDespesa.Text = .Text: .CellBackColor = vbYellow
    .TopRow = .Row
    
 End With

End Sub

Private Sub Carrega_tela()
   'Limpa as variaveis da tela se caso ficarem com dados da outra tela
   limpa_tela Me
   'Carrega a tela com os dados do registro
   With MSFlexGrid1
'      .Col = 0
'      gSql = "select cod_servico,descricao,preco "
'      gSql = gSql & " FROM tab_servicos "
'      gSql = gSql & " Where descricao = '" & Me.MSFlexGrid1.Text & "'"
'      prsServico.Open gSql, ConDb, adOpenKeyset
'      If Not prsServico.EOF And Not prsServico.BOF Then
'         For i = 0 To CmbServico.ListCount
'            If Me.CmbServico.ItemData(i) = prsServico!cod_servico Then
'               CmbServico.ListIndex = i
'               Exit For
'            End If
'         Next
'      Else
'         CmbServico.ListIndex = -1
'      End If
      
'      prsServico.Close
     .Col = 0:  TxtDescricao.Text = .Text: .CellBackColor = vbYellow
     .Col = 1:  MskValor.Text = Format(.Text, "R$#,##0.00;(R$#,##0.00)"):  .CellBackColor = vbYellow
     .Col = 2:  TxtDataDespesa.Text = .Text: .CellBackColor = vbYellow
     .TopRow = .Row

   End With
   
   'Me.TxtValor.Text = "" & gRs("valor")
    
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
            
         MSFlexGrid1.Col = 0:  MSFlexGrid1.Text = !descricao
         MSFlexGrid1.Col = 1:  MSFlexGrid1.Text = !Valor
         .MoveNext
         
       Loop
       MSFlexGrid1.FixedRows = 1
          
  End With
  
End Sub

Private Sub suBotao_add()
    
   Habilita Me
   limpa_tela Me
   'Me.CmbServico.SetFocus
   Me.cmdUpdate.Enabled = True
   Me.cmddesfaz.Enabled = True
   Me.cmdEditar.Enabled = False
   Me.cmdAdd.Enabled = False
   Me.CmdSair.Enabled = False
   Me.cmdDelete.Enabled = False
   lIncluir = True

End Sub

Public Sub Carrega_combo_servicos()
 With prsServico
      '.MoveLast
      'nItem = .RecordCount
      .MoveFirst
      Do While Not .EOF
        CmbServico.AddItem (prsServico!descricao)
        CmbServico.ItemData(CmbServico.NewIndex) = prsServico!cod_servico
        .MoveNext
      Loop
  End With
     
End Sub

Private Sub MskValor_KeyPress(KeyAscii As Integer)

'
'
' Na propriedade Format do controle MaskedBox informa o seguinte valor : #,##0.00;($#,##0.00)
'

If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = Asc(",") Then
    'na faz nada
Else
   If KeyAscii = Asc(".") Then
      KeyAscii = Asc(",")
   Else
      KeyAscii = 0
   End If
End If

End Sub

Private Sub TxtDataDespesa_Validate(Cancel As Boolean)
If Not f_ValidaData(TxtDataDespesa) Then
   MsgBox "Data Inválida", vbOKOnly, "Atenção " & gOperador
   Cancel = True
End If
   
End Sub
