VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmOutMo 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4845
   ClientLeft      =   795
   ClientTop       =   1290
   ClientWidth     =   8505
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4845
   ScaleWidth      =   8505
   ShowInTaskbar   =   0   'False
   Tag             =   "cadvend"
   Begin VB.TextBox TxttotItem 
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
      Enabled         =   0   'False
      Height          =   315
      Left            =   6270
      TabIndex        =   7
      Top             =   1050
      Width           =   1425
   End
   Begin VB.TextBox TxtPreco 
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
      Enabled         =   0   'False
      Height          =   300
      Left            =   3195
      TabIndex        =   6
      Top             =   1050
      Width           =   1365
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   1635
      Left            =   270
      TabIndex        =   19
      Top             =   1710
      Width           =   7875
      _ExtentX        =   13891
      _ExtentY        =   2884
      _Version        =   393216
      Cols            =   7
      FixedCols       =   0
      ScrollBars      =   2
      FormatString    =   "Tipo                         |E/S|  Data        | Produto                                          |Qtde.| Pço.Unit.| Total Item  "
   End
   Begin VB.TextBox TxtQtde 
      Alignment       =   1  'Right Justify
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   1
      EndProperty
      Enabled         =   0   'False
      Height          =   315
      Left            =   885
      TabIndex        =   5
      Top             =   1050
      Width           =   885
   End
   Begin VB.ComboBox CmbProduto 
      Height          =   315
      Left            =   3150
      TabIndex        =   4
      Text            =   "Combo1"
      Top             =   600
      Width           =   4545
   End
   Begin MSMask.MaskEdBox MskData 
      Height          =   315
      Left            =   885
      TabIndex        =   3
      Top             =   600
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   0   'False
      PromptChar      =   "_"
   End
   Begin VB.TextBox TxtE_S 
      Height          =   285
      Left            =   7410
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   165
      Width           =   285
   End
   Begin VB.ComboBox CmbTipo 
      Height          =   315
      Left            =   915
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   120
      Width           =   3555
   End
   Begin VB.Frame Frame1 
      Height          =   795
      Left            =   2040
      TabIndex        =   10
      Top             =   3720
      Width           =   4245
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Salvar"
         Enabled         =   0   'False
         Height          =   540
         Left            =   2190
         Picture         =   "frmOutMo.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   8
         Tag             =   "&Update"
         Top             =   180
         Width           =   615
      End
      Begin VB.CommandButton cmdEditar 
         Caption         =   "&Editar"
         Height          =   540
         Left            =   780
         Picture         =   "frmOutMo.frx":00FA
         Style           =   1  'Graphical
         TabIndex        =   15
         Tag             =   "&Refresh"
         Top             =   180
         Width           =   615
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Excluir"
         Height          =   540
         Left            =   1500
         Picture         =   "frmOutMo.frx":026C
         Style           =   1  'Graphical
         TabIndex        =   14
         Tag             =   "&Delete"
         Top             =   180
         Width           =   615
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Incluir"
         Height          =   540
         Left            =   120
         Picture         =   "frmOutMo.frx":03DE
         Style           =   1  'Graphical
         TabIndex        =   13
         Tag             =   "&Add"
         Top             =   180
         Width           =   615
      End
      Begin VB.CommandButton CmdSair 
         Caption         =   "&Sair"
         Height          =   540
         Left            =   3540
         Picture         =   "frmOutMo.frx":04C8
         Style           =   1  'Graphical
         TabIndex        =   12
         Tag             =   "&Update"
         Top             =   180
         Width           =   615
      End
      Begin VB.CommandButton cmddesfaz 
         Caption         =   "&Desfaz"
         Enabled         =   0   'False
         Height          =   540
         Left            =   2880
         Picture         =   "frmOutMo.frx":05C2
         Style           =   1  'Graphical
         TabIndex        =   11
         Tag             =   "&Update"
         Top             =   180
         Width           =   615
      End
   End
   Begin VB.Label Label5 
      Caption         =   "Total do Item"
      Height          =   195
      Left            =   5085
      TabIndex        =   21
      Top             =   1125
      Width           =   930
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Preço Unit."
      Height          =   195
      Left            =   2145
      TabIndex        =   20
      Top             =   1170
      Width           =   795
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Qtde.:"
      Height          =   195
      Left            =   345
      TabIndex        =   18
      Top             =   1140
      Width           =   435
   End
   Begin VB.Label Label1 
      Caption         =   "Produto:"
      Height          =   255
      Left            =   2385
      TabIndex        =   17
      Top             =   660
      Width           =   735
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "(E)ntrada/(S)aida"
      Height          =   195
      Left            =   6060
      TabIndex        =   16
      Top             =   195
      Width           =   1215
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Tipo"
      Height          =   195
      Index           =   0
      Left            =   345
      TabIndex        =   9
      Tag             =   "CODVEND:"
      Top             =   165
      Width           =   315
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Data:"
      Height          =   195
      Index           =   1
      Left            =   345
      TabIndex        =   0
      Tag             =   "NOME:"
      Top             =   645
      Width           =   390
   End
End
Attribute VB_Name = "frmOutMo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private lIncluir As Boolean
Private pRsFornec As New ADODB.Recordset
Private pRstipomov As New ADODB.Recordset
Private prsProduto As New ADODB.Recordset

Private Sub Abre_Le_rst()
  gSql = "select A.tipo,A.e_s,data,A.codprod,B.descricao as nomeprod, C.descricao as tipomov,qtde,precounit "
  gSql = gSql & " FROM tab_movestoque A, tab_produtos B,tipomov C "
  gSql = gSql & " WHERE A.codprod = B.codprod  AND"
  gSql = gSql & " Val(A.tipo) = C.tipo "
  gRs.Open gSql, ConDb, adOpenKeyset
  
End Sub

Private Sub CmbProduto_LostFocus()
   gSql = "select codprod,descricao,prevenda1 "
   gSql = gSql & "FROM tab_produtos WHERE codprod = '" & Format(CmbProduto.ItemData(CmbProduto.ListIndex), "000000") & "'"
   prsProduto.Open gSql, ConDb, adOpenKeyset
   Me.TxtPreco.Text = Format(prsProduto!prevenda1, "###,##0.00")
   prsProduto.Close
   
End Sub

Private Sub cmdAdd_Click()
   
   Habilita Me
   limpa_tela Me
   
   'Me.MskCnpj.SetFocus
   Me.cmdUpdate.Enabled = True
   Me.cmddesfaz.Enabled = True
   Me.cmdEditar.Enabled = False
   Me.cmdAdd.Enabled = False
   Me.CmdSair.Enabled = False
   Me.cmdDelete.Enabled = False
   CmbTipo.ListIndex = 0
   CmbProduto.ListIndex = 0
   lIncluir = True
End Sub

Private Sub cmdDelete_Click()

    'On Error GoTo ErroDelete
    
    'this may produce an error if you delete the last
    'record or the only record in the recordset
    ' Colocar aqui uma senha para apagar a movimentação
    If MsgBox("Deseja realmente apagar esta Movimentação ? ", vbYesNo, "Atenção") = vbYes Then
       gSql = "delete * from tab_movestoque where tipo = '" & Format(Me.CmbTipo.ItemData(CmbTipo.ListIndex), "00") & "'"
       gSql = gSql & " AND data = cdate('" & Me.MskData.Text & "')"
       gSql = gSql & " AND codprod = '" & Format(CmbProduto.ItemData(CmbProduto.ListIndex), "000000") & "'"
       
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
     MsgBox "Deu erro na exclusao da Movimentação " & Chr(13) & "Instrucao Sql = '" & _
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
 
   If lIncluir Then
      gSql = "INSERT INTO tab_movestoque (tipo,e_s,data,codprod,qtde,precounit,"
      gSql = gSql & "operador, datatual"
      gSql = gSql & ") "
      gSql = gSql & "VALUES ('" & Format(Me.CmbTipo.ItemData(CmbTipo.ListIndex), "00") & "','"
      gSql = gSql & Me.TxtE_S & "',cdate('"
      gSql = gSql & Me.MskData.Text & "'),'"
      gSql = gSql & Format(Me.CmbProduto.ItemData(CmbProduto.ListIndex), "000000") & "',"
      gSql = gSql & Me.TxtQtde.Text & ","
      gSql = gSql & Replace(Me.TxtPreco.Text, ",", ".")
      gSql = gSql & ",'" & gOperador & "',Cdate('" & Date & "') )"
      ConDb.Execute gSql
      lIncluir = False
   Else
      gSql = "UPDATE tab_movestoque SET qtde = " & Me.TxtQtde.Text & ","
      gSql = gSql & "E_S = '" & Me.TxtE_S & "',"
      gSql = gSql & "precounit = " & Replace(Me.TxtPreco.Text, ",", ".")
      gSql = gSql & " ,operador = '" & gOperador & "', datatual = cDate('" & Date & "')"
      gSql = gSql & " WHERE tipo = '" & Format(CmbTipo.ItemData(CmbTipo.ListIndex), "00") & "' "
      gSql = gSql & " AND data = cdate('" & Me.MskData.Text & "')"
      gSql = gSql & " AND codprod = '" & Format(CmbProduto.ItemData(CmbProduto.ListIndex), "000000") & "'"
      ConDb.Execute gSql
      
   End If
       
   Abre_Le_rst
   
   Carrega_Grid
   gRs.MoveFirst
   Carrega_tela
   Desabilita Me
   gRs.Close
   Me.cmdUpdate.Enabled = False
   Me.cmddesfaz.Enabled = False
   Me.cmdEditar.Enabled = True
   Me.cmdAdd.Enabled = True
   Me.CmdSair.Enabled = True
   Me.cmdDelete.Enabled = True
     
End Sub

Private Sub Form_Activate()
   Abre_Le_rst
   
   'Me.Lblcodgrupo.Caption = ""
   
   If gRs.BOF And gRs.EOF Then
      If MsgBox("Arquivo vazio. Incluir dados agora ?", vbYesNo, "Atenção ") = vbYes Then
         gSql = "INSERT INTO tab_movestoque (tipo, e_s,data,codvend,codprod,qtde,precounit,operador,datatual) "
         gSql = gSql & "VALUES ('01','E',cdate('" & Date & "')','01','" & Format(CmbProduto.ItemData(CmbProduto.ListIndex), "000000") & "'," & "1,1.00,"
         gSql = gSql & "'" & gOperador & "'," & Date & " ) "
         ConDb.Execute gSql
         gRs.Close
         Abre_Le_rst
         cmdEditar_Click
         lPrimeiro = True
      Else
         Desabilita Me
      End If
      
   Else
     ' gRs.MoveFirst
     ' Carrega_tela
    
      Desabilita Me
      lIncluir = False
      lPrimeiro = False
   End If
   
   Carrega_Combo_tipo
   Carrega_Combo_Produto
   
   Carrega_Grid
   gRs.MoveFirst
   Carrega_tela
   gRs.Close
   lIncluir = False

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'gRs.Close
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
    '.Row = 1
    .Col = 0:   .CellBackColor = vbYellow
    .Col = 3:   .CellBackColor = vbYellow
    .Col = 8:
    For i = 0 To CmbTipo.ListCount - 1
        If CmbTipo.ItemData(i) = .Text Then
           CmbTipo.ListIndex = i
           Exit For
        End If
    Next
    .Col = 1:  TxtE_S.Text = .Text: .CellBackColor = vbYellow
    .Col = 2:  MskData.Text = .Text: .CellBackColor = vbYellow
    
    .Col = 7
    For i = 0 To CmbProduto.ListCount - 1
        If CmbProduto.ItemData(i) = Val(.Text) Then
           CmbProduto.ListIndex = i
           Exit For
        End If
    Next
    .Col = 4:  TxtQtde.Text = .Text: .CellBackColor = vbYellow
    .Col = 5:  TxtPreco.Text = .Text: .CellBackColor = vbYellow
    .Col = 6:  TxttotItem.Text = .Text: .CellBackColor = vbYellow
    .TopRow = .Row
  End With

End Sub

Private Sub Carrega_tela()
   'Limpa as variaveis da tela se caso ficarem com dados da outra tela
   'limpa_tela Me
   'Carrega a tela com os dados do registro
   With MSFlexGrid1
      .Row = 1
      .Col = 0:  .CellBackColor = vbYellow
      .Col = 3:  .CellBackColor = vbYellow
      .Col = 8
      For i = 0 To CmbTipo.ListCount - 1
          If CmbTipo.ItemData(i) = Val(.Text) Then
             CmbTipo.ListIndex = i
             Exit For
          End If
      Next
      .Col = 1:  TxtE_S.Text = .Text: .CellBackColor = vbYellow
      .Col = 2:  MskData.Text = .Text: .CellBackColor = vbYellow
      .Col = 7
      For i = 0 To CmbProduto.ListCount - 1
      
          If CmbProduto.ItemData(i) = .Text Then
             CmbProduto.ListIndex = i
             Exit For
          End If
      Next
      .Col = 4:  TxtQtde.Text = Format(.Text, "000"): .CellBackColor = vbYellow
      .Col = 5:  TxtPreco.Text = Format(.Text, "###,##0.00"): .CellBackColor = vbYellow
      .Col = 6:  TxttotItem.Text = Format(.Text, "###,##0.00"): .CellBackColor = vbYellow
      .TopRow = .Row
   End With

End Sub

Private Sub Carrega_Grid()

'Teste do MsHFlexgrid1 - eh eh eh
  MSFlexGrid1.Row = 0
  
  With gRs
      '.MoveLast
      'nItem = .RecordCount
      '.MoveFirst
      MSFlexGrid1.Rows = 1
      MSFlexGrid1.Cols = 9
      Do While Not .EOF
         MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
         MSFlexGrid1.Row = MSFlexGrid1.Rows - 1
         MSFlexGrid1.Col = 0:  MSFlexGrid1.Text = f_nulo(!tipomov, "")
         MSFlexGrid1.Col = 1:  MSFlexGrid1.Text = f_nulo(!E_S, "")
         MSFlexGrid1.Col = 2:  MSFlexGrid1.Text = f_nulo(!Data, "")
         MSFlexGrid1.Col = 3:  MSFlexGrid1.Text = f_nulo(!nomeprod, "")
         MSFlexGrid1.Col = 4:  MSFlexGrid1.Text = Format(f_nulo(!qtde, 0), "##0")
         MSFlexGrid1.Col = 5:  MSFlexGrid1.Text = Format(f_nulo(!precounit, 0), "###,##0.00")
         MSFlexGrid1.Col = 6:  MSFlexGrid1.Text = Format(f_nulo(!qtde, 0) * f_nulo(!precounit, 0), "###,##0.00")
         MSFlexGrid1.Col = 7:  MSFlexGrid1.Visible = True: MSFlexGrid1.Text = f_nulo(!codprod, 0)
         MSFlexGrid1.Col = 8:  MSFlexGrid1.Visible = True: MSFlexGrid1.Text = f_nulo(!tipo, 0)
         .MoveNext
         
      Loop
      MSFlexGrid1.FixedRows = 1
          
  End With
  
End Sub

Private Sub Abre_Le_rst_tipomov()
   gSql = "select tipo,descricao "
   gSql = gSql & "FROM tipomov "
   pRstipomov.Open gSql, ConDb, adOpenKeyset
   
End Sub
Private Sub Carrega_Combo_tipo()

 Abre_Le_rst_tipomov
 
 CmbTipo.Clear
 With pRstipomov
      '.MoveLast
      'nItem = .RecordCount
      .MoveFirst
      Do While Not .EOF
        CmbTipo.AddItem (pRstipomov!descricao)
        CmbTipo.ItemData(CmbTipo.NewIndex) = pRstipomov!tipo
        .MoveNext
      Loop
  End With
  pRstipomov.Close
End Sub

Private Sub Abre_Le_rst_Produto()
   gSql = "select codprod,descricao "
   gSql = gSql & "FROM tab_produtos"
   prsProduto.Open gSql, ConDb, adOpenKeyset
   
End Sub
Private Sub Carrega_Combo_Produto()

 Abre_Le_rst_Produto
 
 CmbProduto.Clear
 With prsProduto
      '.MoveLast
      'nItem = .RecordCount
      .MoveFirst
      Do While Not .EOF
        CmbProduto.AddItem (prsProduto!descricao)
        CmbProduto.ItemData(CmbProduto.NewIndex) = prsProduto!codprod
        .MoveNext
      Loop
  End With
  prsProduto.Close
End Sub

Private Sub MskData_GotFocus()
   With MskData
      .SelStart = 0
      .SelLength = Len(.Text)
   End With

End Sub

Private Sub MskData_Validate(Cancel As Boolean)
   If Not IsDate(MskData.Text) Then
      MsgBox "Data inválida, favor corrigir", vbOKOnly, "Atenção " & gOperador
      Cancel = True
   Else
      Cancel = False
   End If
End Sub

Private Sub TxtE_S_GotFocus()
  With TxtE_S
      .SelStart = 0
      .SelLength = Len(.Text)
   End With

End Sub

Private Sub TxtE_S_LostFocus()
   TxtE_S.Text = UCase(TxtE_S.Text)
End Sub

Private Sub TxtE_S_Validate(Cancel As Boolean)
  If UCase(TxtE_S.Text) = "S" Or UCase(TxtE_S.Text) = "E" Then
     Cancel = False
  Else
     MsgBox "Digitar somente 'E' ou 'S'", vbOKOnly, "Atenção " & gOperador
     Cancel = True
  End If
  
End Sub

Private Sub TxtPreco_GotFocus()
   With TxtPreco
      .SelStart = 0
      .SelLength = Len(.Text)
   End With

End Sub

Private Sub TxtQtde_GotFocus()
With TxtQtde
      .SelStart = 0
      .SelLength = Len(.Text)
   End With
End Sub

Private Sub TxttotItem_GotFocus()
   TxttotItem.Text = Format(CDbl(TxtPreco.Text) * CDbl(TxtQtde.Text), "###,##0.00")
End Sub
