VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmPlanImpo 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Planilha de Impostos"
   ClientHeight    =   5505
   ClientLeft      =   4545
   ClientTop       =   2445
   ClientWidth     =   6030
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5505
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   Tag             =   "cadvend"
   Begin VB.TextBox TxtDataImposto 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd/MM/yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   3
      EndProperty
      Height          =   315
      Left            =   945
      TabIndex        =   5
      Top             =   1740
      Width           =   990
   End
   Begin MSMask.MaskEdBox MskValor 
      Height          =   315
      Left            =   4245
      TabIndex        =   6
      Top             =   1725
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   556
      _Version        =   393216
      Format          =   "$#,##0.00;($#,##0.00)"
      PromptChar      =   "_"
   End
   Begin VB.Frame Frame2 
      Caption         =   "Dados da Obra"
      Height          =   1005
      Left            =   0
      TabIndex        =   14
      Top             =   90
      Width           =   5865
      Begin VB.Label LblObra 
         Caption         =   "Unidade"
         DataSource      =   "gcodobra"
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   1035
         TabIndex        =   1
         Top             =   285
         Width           =   4725
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Obra"
         Height          =   195
         Left            =   450
         TabIndex        =   17
         Top             =   285
         Width           =   345
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Medição"
         Height          =   195
         Left            =   210
         TabIndex        =   16
         Top             =   585
         Width           =   615
      End
      Begin VB.Label LblMedicao 
         Caption         =   "Num.Medição"
         DataSource      =   "gnummedicao"
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   1050
         TabIndex        =   2
         Top             =   615
         Width           =   1245
      End
      Begin VB.Label Label4 
         Caption         =   "Dta.Medição"
         Height          =   225
         Left            =   3435
         TabIndex        =   15
         Top             =   645
         Width           =   1005
      End
      Begin VB.Label LblDta_medicao 
         Caption         =   "99/99/9999"
         DataSource      =   "gdatamedicao"
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   4515
         TabIndex        =   3
         Top             =   645
         Width           =   1245
      End
   End
   Begin VB.ComboBox CmbImposto 
      Enabled         =   0   'False
      Height          =   315
      Left            =   960
      TabIndex        =   4
      Top             =   1305
      Width           =   4845
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   2085
      Left            =   210
      TabIndex        =   13
      Top             =   2220
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   3678
      _Version        =   393216
      Rows            =   5
      Cols            =   3
      FixedCols       =   0
      FormatString    =   "Imposto                                                                   |Data         |Valor          "
   End
   Begin VB.Frame Frame1 
      Height          =   795
      Left            =   795
      TabIndex        =   7
      Top             =   4455
      Width           =   4245
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Salvar"
         Enabled         =   0   'False
         Height          =   540
         Left            =   2220
         Picture         =   "frmPlanImpo.frx":0000
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
         Picture         =   "frmPlanImpo.frx":00FA
         Style           =   1  'Graphical
         TabIndex        =   12
         Tag             =   "&Refresh"
         Top             =   180
         Width           =   615
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Excluir"
         Height          =   540
         Left            =   1500
         Picture         =   "frmPlanImpo.frx":026C
         Style           =   1  'Graphical
         TabIndex        =   11
         Tag             =   "&Delete"
         Top             =   180
         Width           =   615
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Incluir"
         Height          =   540
         Left            =   120
         Picture         =   "frmPlanImpo.frx":03DE
         Style           =   1  'Graphical
         TabIndex        =   10
         Tag             =   "&Add"
         Top             =   180
         Width           =   615
      End
      Begin VB.CommandButton CmdSair 
         Caption         =   "&Sair"
         Height          =   540
         Left            =   3540
         Picture         =   "frmPlanImpo.frx":04C8
         Style           =   1  'Graphical
         TabIndex        =   9
         Tag             =   "&Update"
         Top             =   180
         Width           =   615
      End
      Begin VB.CommandButton cmddesfaz 
         Caption         =   "&Desfaz"
         Enabled         =   0   'False
         Height          =   540
         Left            =   2880
         Picture         =   "frmPlanImpo.frx":05C2
         Style           =   1  'Graphical
         TabIndex        =   8
         Tag             =   "&Update"
         Top             =   180
         Width           =   615
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Data"
      Height          =   195
      Left            =   450
      TabIndex        =   20
      Top             =   1785
      Width           =   345
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Imposto"
      Height          =   195
      Left            =   180
      TabIndex        =   19
      Top             =   1320
      Width           =   555
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Valor R$"
      Height          =   195
      Left            =   3420
      TabIndex        =   18
      Top             =   1755
      Width           =   615
   End
End
Attribute VB_Name = "frmPlanImpo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private lIncluir As Boolean
Dim prsImposto As New ADODB.Recordset

Private Sub Abre_Le_rst()
  gSql = "select *,B.descricao FROM planimpo as A,tab_impostos as B"
  gSql = gSql & " WHERE A.cod_obra = " & gCodObra
  gSql = gSql & " AND A.medicao = " & gNumMedicao
  gSql = gSql & " AND A.dta_medicao = " & gDataMedicao
  gSql = gSql & " AND A.cod_imposto = B.cod_imposto "
  gRs.Open gSql, ConDb, adOpenKeyset
  
End Sub

Private Sub cmdAdd_Click()
   suBotao_add
End Sub

Private Sub cmdDelete_Click()

    'On Error GoTo ErroDelete
    
    'this may produce an error if you delete the last
    'record or the only record in the recordset
    If MsgBox("Deseja realmente apagar este Item de imposto ? ", vbYesNo, "Atenção") = vbYes Then
       gSql = "delete * from planimpo "
       gSql = gSql & " WHERE cod_obra = " & gCodObra
       gSql = gSql & " AND medicao = " & gNumMedicao
       gSql = gSql & " AND dta_medicao = " & gDataMedicao
       gSql = gSql & " AND dta_imposto = " & TxtDataImposto.Text
       gSql = gSql & " AND cod_imposto = " & Val(Me.CmbImposto.ItemData(Me.CmbImposto.ListIndex))
       ConDb.Execute gSql
       gRs.Close
       Abre_Le_rst
       Carrega_tela
       suTesta_Vazio
    End If
'     Exit Sub
     
'ErroDelete:
'     MsgBox "Deu erro na exclusao do Fornecedor " & Chr(13) & "Instrucao Sql = '" & _
'            cSql & "'  "
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
   Me.CmbImposto.SetFocus
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
      gSql = "INSERT INTO Planimpo (cod_obra,medicao,dta_medicao,"
      gSql = gSql & "cod_imposto,valor,dta_imposto,operador, datatual"
      gSql = gSql & ") "
      gSql = gSql & "VALUES ( " & gCodObra & ","
      gSql = gSql & gNumMedicao & ","
      gSql = gSql & "Cdate('" & gDataMedicao & "'),'"
      gSql = gSql & Me.CmbImposto.ItemData(Me.CmbImposto.ListIndex) & "',"
      gSql = gSql & Val(Me.MskValor.Text) & ","
      gSql = gSql & "Cdate('" & Me.TxtDataImposto.Text & "'),'"
      gSql = gSql & gOperador & "',Cdate('" & Date & "')"
      ConDb.Execute gSql
      lIncluir = False
   Else
      gSql = "UPDATE Planimpo SET cod_imposto = " & Val(Me.CmbImposto.ItemData(Me.CmbImposto.ListIndex)) & ","
      gSql = gSql & "valor = " & Val(Me.MskValor.Text) & ","
      gSql = gSql & "dta_imposto = Cdate('" & Me.TxtDataImposto.Text & "'),"
      gSql = gSql & "operador = '" & gOperador & "', datatual = Cdate('" & Date & "')"
      gSql = gSql & " WHERE cod_obra  = " & gCodObra
      gSql = gSql & " AND medicao = " & gNumMedicao
      gSql = gSql & " AND dta_medicao = Cdate('" & gDataMedicao & "')"
      gSql = gSql & " AND cod_imposto = " & Val(Me.CmbImposto.ItemData(Me.CmbImposto.ListIndex))
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
       
   gSql = "select * FROM tab_impostos"
   prsImposto.Open gSql, ConDb, adOpenForwardOnly
   If prsImposto.BOF And prsImposto.EOF Then
      MsgBox "Não existem Impostos Cadastrados ! Favor Cadastrar. ", vbOK, "Atenção"
      prsImposto.Close
      Unload Me
      'Unload FrmPlanilha
      Exit Sub
   End If
   Carrega_combo_Impostos
   prsImposto.Close
   
   suTesta_Vazio
   
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
   ' Unload FrmPlanilha
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
     
     'Acha o Serviço para por no combo
     .Col = 0
     gSql = "select cod_imposto,descricao "
     gSql = gSql & " FROM tab_impostos "
     gSql = gSql & " Where descricao = '" & Me.MSFlexGrid1.Text & "'"
     prsImposto.Open gSql, ConDb, adOpenKeyset
     If Not prsImposto.EOF And Not prsImposto.BOF Then
        For i = 0 To CmbImposto.ListCount
           If Me.CmbImposto.ItemData(i) = prsImposto!cod_Imposto Then
              CmbImposto.ListIndex = i
              Exit For
           End If
        Next
     Else
        CmbImposto.ListIndex = -1
     End If
     prsImposto.Close

    .Col = 0:  .CellBackColor = vbYellow
    .Col = 1:  TxtDataImposto.Text = .Text: .CellBackColor = vbYellow
    .Col = 2:  MskValor.Text = .Text: .CellBackColor = vbYellow
    
    .TopRow = .Row
    
 End With

End Sub

Private Sub Carrega_tela()
   'Limpa as variaveis da tela se caso ficarem com dados da outra tela
   limpa_tela Me
   'Carrega a tela com os dados do registro
   With MSFlexGrid1
      .Col = 0
      gSql = "select cod_imposto,descricao,valor "
      gSql = gSql & " FROM tab_Impostos "
      gSql = gSql & " Where descricao = '" & Me.MSFlexGrid1.Text & "'"
      prsImposto.Open gSql, ConDb, adOpenKeyset
      If Not prsImposto.EOF And Not prsImposto.BOF Then
         For i = 0 To CmbImposto.ListCount
            If Me.CmbImposto.ItemData(i) = prsImposto!cod_Imposto Then
               CmbImposto.ListIndex = i
               Exit For
            End If
         Next
      Else
         CmbImposto.ListIndex = -1
      End If
      
      prsImposto.Close
     .Col = 0:  .CellBackColor = vbYellow
     .Col = 1: TxtDataImposto.Text = .Text: .CellBackColor = vbYellow
     .Col = 2:  MskValor.Text = .Text: .CellBackColor = vbYellow
     .TopRow = .Row

   End With
   
   'Me.MskValor.Text = "" & Format(gRs!Valor, "R$#,##0.00;(R$#,##0.00)")
    
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
         MSFlexGrid1.Col = 1:  MSFlexGrid1.Text = Format(!dta_imposto, "dd/mm/yyyy")
         MSFlexGrid1.Col = 2:  MSFlexGrid1.Text = Format(!Valor, "R$#,##0.00;(R$#,##0.00)")
         .MoveNext
         
       Loop
       'MSFlexGrid1.FixedRows = 1
          
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

Public Sub Carrega_combo_Impostos()
 With prsImposto
      '.MoveLast
      'nItem = .RecordCount
      .MoveFirst
      Do While Not .EOF
        CmbImposto.AddItem (prsImposto!descricao)
        CmbImposto.ItemData(CmbImposto.NewIndex) = prsImposto!cod_Imposto
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

Private Sub TxtDataImposto_Validate(Cancel As Boolean)
If Not IsDate(TxtDataImposto.Text) Then
   MsgBox "Data Invalida ", vbOKCancel, "Atenção " & gOperador
   Cancel = True
End If

End Sub

Private Sub suTesta_Vazio()
   If gRs.BOF And gRs.EOF Then
      If MsgBox("Arquivo vazio. Incluir dados agora ?", vbYesNo, "Atenção ") = vbYes Then
         cmdEditar_Click
         lIncluir = True
         lPrimeiro = True
         Me.CmbImposto.SetFocus
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
