VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmLancDesp 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4920
   ClientLeft      =   3315
   ClientTop       =   2460
   ClientWidth     =   6735
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   6735
   ShowInTaskbar   =   0   'False
   Tag             =   "cadvend"
   Begin VB.TextBox TxtValor 
      Alignment       =   1  'Right Justify
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   1
      EndProperty
      Height          =   315
      Left            =   4410
      TabIndex        =   2
      Text            =   "0,00"
      Top             =   90
      Width           =   1560
   End
   Begin VB.ComboBox CboCCusto 
      Height          =   315
      Left            =   1305
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   525
      Width           =   4695
   End
   Begin MSMask.MaskEdBox MskData 
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
      Left            =   1335
      TabIndex        =   1
      Top             =   60
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   10
      Mask            =   "99/99/9999"
      PromptChar      =   "_"
   End
   Begin VB.TextBox TxtDescricao 
      Height          =   315
      Left            =   1290
      MaxLength       =   50
      TabIndex        =   4
      Top             =   960
      Width           =   4665
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   1845
      Left            =   210
      TabIndex        =   15
      Top             =   1560
      Width           =   6285
      _ExtentX        =   11086
      _ExtentY        =   3254
      _Version        =   393216
      Rows            =   5
      Cols            =   4
      FixedCols       =   0
      BackColorSel    =   16777215
      SelectionMode   =   1
      FormatString    =   "Data            |> Valor          | C.Custo                                  | Descrição                                 "
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   1155
      TabIndex        =   0
      Top             =   3675
      Width           =   4245
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Salvar"
         Enabled         =   0   'False
         Height          =   540
         Left            =   2190
         Picture         =   "frmLancDesp.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   8
         Tag             =   "&Update"
         Top             =   150
         Width           =   615
      End
      Begin VB.CommandButton cmdEditar 
         Caption         =   "&Editar"
         Height          =   540
         Left            =   795
         Picture         =   "frmLancDesp.frx":00FA
         Style           =   1  'Graphical
         TabIndex        =   6
         Tag             =   "&Refresh"
         Top             =   150
         Width           =   615
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Excluir"
         Height          =   540
         Left            =   1485
         Picture         =   "frmLancDesp.frx":026C
         Style           =   1  'Graphical
         TabIndex        =   7
         Tag             =   "&Delete"
         Top             =   150
         Width           =   615
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Incluir"
         Height          =   540
         Left            =   105
         Picture         =   "frmLancDesp.frx":03DE
         Style           =   1  'Graphical
         TabIndex        =   5
         Tag             =   "&Add"
         Top             =   150
         Width           =   615
      End
      Begin VB.CommandButton CmdSair 
         Caption         =   "&Sair"
         Height          =   540
         Left            =   3555
         Picture         =   "frmLancDesp.frx":04C8
         Style           =   1  'Graphical
         TabIndex        =   10
         Tag             =   "&Update"
         Top             =   150
         Width           =   615
      End
      Begin VB.CommandButton cmddesfaz 
         Caption         =   "&Desfaz"
         Enabled         =   0   'False
         Height          =   540
         Left            =   2865
         Picture         =   "frmLancDesp.frx":05C2
         Style           =   1  'Graphical
         TabIndex        =   9
         Tag             =   "&Update"
         Top             =   150
         Width           =   615
      End
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Valor:"
      Height          =   195
      Left            =   3840
      TabIndex        =   14
      Top             =   135
      Width           =   405
   End
   Begin VB.Label Label1 
      Caption         =   "Data:"
      Height          =   210
      Left            =   810
      TabIndex        =   13
      Top             =   105
      Width           =   540
   End
   Begin VB.Label lblfixo 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Descrição:"
      Height          =   195
      Index           =   2
      Left            =   405
      TabIndex        =   12
      Tag             =   "SALARIO:"
      Top             =   990
      Width           =   765
   End
   Begin VB.Label lblDescricao 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "C.Custo:"
      Height          =   195
      Index           =   1
      Left            =   585
      TabIndex        =   11
      Tag             =   "NOME:"
      Top             =   570
      Width           =   600
   End
End
Attribute VB_Name = "frmLancDesp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private lIncluir As Boolean
Private lPrimeiro As Boolean
Private prsCCusto As New ADODB.Recordset

Private Sub Carrega_tela()
   'Limpa as variaveis da tela se caso ficarem com dados da outra tela
   limpa_tela Me
   'Carrega a tela com os dados do registro
   Me.MskData = gRs!datalanc
   Me.TxtDescricao.Text = gRs!descricao
   Me.TxtValor.Text = f_nulo(gRs!Valor, "")
   
   'Acha o C.Custo para por no combo
   Me.MSFlexGrid1.Col = 2
   gSql = "select codigo,tab_ccusto.descricao "
   gSql = gSql & " FROM tab_ccusto "
   gSql = gSql & " Where tab_ccusto.descricao = '" & Me.MSFlexGrid1.Text & "'"
   prsCCusto.Open gSql, ConDb, adOpenKeyset
   If Not prsCCusto.EOF And Not prsCCusto.BOF Then
      For i = 0 To CboCCusto.ListCount - 1
         If CboCCusto.ItemData(i) = prsCCusto!codigo Then
            CboCCusto.ListIndex = i
            Exit For
         End If
      Next
   Else
      CboCCusto.ListIndex = -1
   End If
   prsCCusto.Close
   
   
End Sub
Private Sub Carrega_Grid()

'Teste do MsHFlexgrid1 - eh eh eh
  MSFlexGrid1.Row = 0
  
  If gRs.BOF And gRs.EOF Then
     Exit Sub
  End If
  
  With gRs
      '.MoveLast
      'nItem = .RecordCount
      '.MoveFirst
      MSFlexGrid1.Rows = 1
      Do While Not .EOF
         MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
         MSFlexGrid1.Row = MSFlexGrid1.Rows - 1
            
         MSFlexGrid1.Col = 0: MSFlexGrid1.Text = f_nulo(!datalanc, "")
         MSFlexGrid1.Col = 1: MSFlexGrid1.Text = Format(f_nulo(!Valor, 0), "###,##0.00")
         MSFlexGrid1.Col = 2: MSFlexGrid1.Text = f_nulo(!desccc, "")
         MSFlexGrid1.Col = 3: MSFlexGrid1.Text = f_nulo(!descricao, "")
         .MoveNext
         
       Loop
       MSFlexGrid1.FixedRows = 1
          
  End With
  
  End Sub

Private Sub cmdAdd_Click()
   
   lIncluir = True
   limpa_tela Me
   MskData.Text = Date
   'Me.lblcodigo.Caption = ""
   Habilita Me
   suCmdAdd Me  'Habilita e desabilita botoes
   Me.cmdEditar.Enabled = False
   Me.MskData.SetFocus
End Sub

Private Sub cmdDelete_Click()
    'this may produce an error if you delete the last
    'record or the only record in the recordset
    If MsgBox("Deseja realmente apagar este Lançamento ? ", vbYesNo, "Atenção") = vbYes Then
       gSql = "delete * from tab_ldes where datalanc = cdate ('" & Me.MskData.Text & "')"
       gSql = gSql & " AND valor = " & CDbl(Me.TxtValor)
       gSql = gSql & " AND descricao = '" & Me.TxtDescricao & "'"
       ConDb.Execute gSql
       gRs.Close
       Abre_Le_rst
       Carrega_Grid
       gRs.MoveFirst
       Carrega_tela
       Desabilita Me
     End If
End Sub

Private Sub cmddesfaz_Click()
  lIncluir = False
  'Carrega_tela
  Desabilita Me
  MSFlexGrid1_Click
  suCmdDesfaz Me
  Me.cmdEditar.Enabled = False
  
End Sub

Private Sub cmdEditar_Click()
   Habilita Me
  
   suCmdEditar Me
   
   Me.MskData.SetFocus
   
End Sub


Private Sub CmdSair_Click()
   Unload Me
End Sub

Private Sub cmdUpdate_Click()
   
   gRs.Close
   If lIncluir Then
      gSql = "INSERT INTO tab_ldes (datalanc,ccusto,descricao,valor,operador,datatual) "
      gSql = gSql & "VALUES ( cdate('" & Me.MskData.Text & "'),"
      gSql = gSql & CboCCusto.ItemData(CboCCusto.ListIndex) & ",'"
      gSql = gSql & f_nulo(Me.TxtDescricao.Text, " ") & "',"
      gSql = gSql & Replace(IIf(Len(Me.TxtValor.Text) = 0, 0, Me.TxtValor.Text), ",", ".") & ",'"
      gSql = gSql & gOperador & "',Cdate('" & Date & "')) "
      ConDb.Execute gSql
      'lIncluir = False
   Else
      'gSql = "UPDATE tab_ldes SET datalanc = cdate('" & Me.MskData.Text & "'),"
      'gSql = gSql & "descricao = '" & Me.TxtDescricao.Text & "',"
      'gSql = gSql & "operador = '" & gOperador & "', datatual = cDate('" & Date & "')"
      'gSql = gSql & " WHERE codigo = " & Val(Me.lblcodigo.Caption)
      'ConDb.Execute gSql
      
   End If
                              
   Abre_Le_rst
   Carrega_Grid
   gRs.MoveLast
   Carrega_tela
   Desabilita Me
   
   suCmdUpdate Me
   cmdEditar.Enabled = False
   
End Sub


Private Sub Form_Activate()
   Abre_Le_rst
   
   Call Carrega_combo_Ccusto
   
   limpa_tela Me

   If gRs.BOF And gRs.EOF Then
      If MsgBox("Arquivo vazio. Incluir dados agora ?", vbYesNo, "Atenção ") = vbYes Then
         'gSql = "INSERT INTO tab_lDes (datalanc,ccusto,descricao,valor,"
         'gSql = gSql & "operador, datatual) "
         'gSql = gSql & " VALUES ( cdate('" & Date & "'),1,'"
         'gSql = gSql & f_nulo(Me.TxtDescricao.Text, " ") & "',"
         'gSql = gSql & f_nulo(Me.TxtValor.Text, 0) & ",'"
         'gSql = gSql & gOperador & "', cdate('" & Date & "') )"
         'ConDb.Execute gSql
         'gRs.Close
         'Abre_Le_rst
         'Me.MskData.Text = gRs!datalanc
         cmdEditar_Click
         lPrimeiro = True
         lIncluir = True
         Habilita Me
         MskData.SetFocus
      Else
         Desabilita Me
      End If

   Else
      gRs.MoveFirst
      Carrega_tela

      Desabilita Me
      Me.cmdEditar.Enabled = False
      lIncluir = False
      lPrimeiro = False
   End If
   Carrega_Grid

   lIncluir = True

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

Private Sub Abre_Le_rst()
   gSql = "select *,tab_ccusto.descricao as descCC FROM tab_ldes,tab_ccusto"
   gSql = gSql & " WHERE tab_ldes.ccusto = tab_ccusto.codigo"
   gSql = gSql & " ORDER BY datalanc desc"
   gRs.Open gSql, ConDb, adOpenKeyset, adLockOptimistic
    
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
    
    .Col = 0:   MskData.Text = .Text: .CellBackColor = vbYellow
    .Col = 1:   TxtValor.Text = .Text: .CellBackColor = vbYellow
    .Col = 2
    For i = 0 To Me.CboCCusto.ListCount - 1
       Me.CboCCusto.ListIndex = i
       If Me.CboCCusto.Text = .Text Then
          Me.CboCCusto.ListIndex = i
          Exit For
       End If
    Next
    .CellBackColor = vbYellow
    .Col = 3:   TxtDescricao.Text = .Text: .CellBackColor = vbYellow
    .Redraw = True
    .TopRow = .Row
    
  End With

End Sub

Private Sub MSFlexGrid1_SelChange()
   MSFlexGrid1_Click
End Sub

Private Sub Carrega_combo_Ccusto()
   
   gSql = "select codigo,Descricao as DescCC "
   gSql = gSql & "FROM tab_ccusto ORDER BY descricao "
   prsCCusto.Open gSql, ConDb, adOpenKeyset
   CboCCusto.Clear
   With prsCCusto
      '.MoveLast
      'nItem = .RecordCount
      .MoveFirst
      Do While Not .EOF
        CboCCusto.AddItem (prsCCusto!desccc)
        CboCCusto.ItemData(CboCCusto.NewIndex) = prsCCusto!codigo
        .MoveNext
      Loop
  End With

  prsCCusto.Close

End Sub



Private Sub MskData_Validate(Cancel As Boolean)
   If Not ChkData(MskData.Text) Then
      Cancel = True
   End If
End Sub
