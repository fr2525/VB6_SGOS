VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form FrmApagar 
   Caption         =   "Contas a Pagar"
   ClientHeight    =   6225
   ClientLeft      =   1590
   ClientTop       =   1725
   ClientWidth     =   8010
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6225
   ScaleWidth      =   8010
   Begin MSFlexGridLib.MSFlexGrid VaSpread1 
      Height          =   2235
      Left            =   630
      TabIndex        =   25
      Top             =   2730
      Width           =   6825
      _ExtentX        =   12039
      _ExtentY        =   3942
      _Version        =   393216
      Cols            =   8
      FixedCols       =   0
      FormatString    =   $"FrmApagar.frx":0000
   End
   Begin MSMask.MaskEdBox MskVencto 
      Height          =   315
      Left            =   1530
      TabIndex        =   5
      Top             =   1005
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   0   'False
      MaxLength       =   10
      Format          =   "dd/mm/yyyy"
      Mask            =   "99/99/9999"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MskEmissao 
      Height          =   315
      Left            =   6000
      TabIndex        =   4
      Top             =   585
      Width           =   1020
      _ExtentX        =   1799
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   0   'False
      MaxLength       =   10
      Format          =   "dd/mm/yyyy"
      Mask            =   "99/99/9999"
      PromptChar      =   "_"
   End
   Begin VB.Frame FramePagto 
      Caption         =   "Dados de Pagamento"
      Enabled         =   0   'False
      Height          =   825
      Left            =   1410
      TabIndex        =   21
      Top             =   1740
      Width           =   5235
      Begin MSMask.MaskEdBox MskdtPagto 
         Height          =   300
         Left            =   945
         TabIndex        =   9
         Top             =   345
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "99/99/9999"
         PromptChar      =   "_"
      End
      Begin VB.TextBox TxtValorPago 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3270
         TabIndex        =   10
         Text            =   "0,00"
         Top             =   330
         Width           =   1575
      End
      Begin VB.Label Label8 
         Caption         =   "Valor"
         Height          =   225
         Left            =   2730
         TabIndex        =   23
         Top             =   375
         Width           =   480
      End
      Begin VB.Label Label7 
         Caption         =   "Data"
         Height          =   225
         Left            =   420
         TabIndex        =   22
         Top             =   360
         Width           =   555
      End
   End
   Begin VB.TextBox TxtValor 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   285
      Left            =   5355
      TabIndex        =   6
      Text            =   "0,00"
      Top             =   1005
      Width           =   1665
   End
   Begin VB.TextBox TxtNotafiscal 
      Enabled         =   0   'False
      Height          =   285
      Left            =   3930
      TabIndex        =   3
      Top             =   585
      Width           =   1155
   End
   Begin VB.ComboBox Cbofornecedor 
      Height          =   315
      Left            =   1530
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   150
      Width           =   5490
   End
   Begin VB.TextBox TxtDuplicata 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1545
      TabIndex        =   2
      Top             =   585
      Width           =   1155
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   1680
      TabIndex        =   0
      Top             =   5250
      Width           =   4905
      Begin VB.CommandButton CmdBaixar 
         Caption         =   "&Baixar"
         Height          =   540
         Left            =   2160
         Picture         =   "FrmApagar.frx":00A5
         Style           =   1  'Graphical
         TabIndex        =   24
         Tag             =   "&Refresh"
         Top             =   150
         Width           =   615
      End
      Begin VB.CommandButton cmddesfaz 
         Caption         =   "&Desfaz"
         Enabled         =   0   'False
         Height          =   540
         Left            =   3510
         Picture         =   "FrmApagar.frx":0CE7
         Style           =   1  'Graphical
         TabIndex        =   8
         Tag             =   "&Update"
         Top             =   150
         Width           =   615
      End
      Begin VB.CommandButton CmdSair 
         Caption         =   "&Sair"
         Height          =   540
         Left            =   4185
         Picture         =   "FrmApagar.frx":0DE1
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
         Picture         =   "FrmApagar.frx":0EDB
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
         Picture         =   "FrmApagar.frx":0FC5
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
         Picture         =   "FrmApagar.frx":1137
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
         Left            =   2835
         Picture         =   "FrmApagar.frx":12A9
         Style           =   1  'Graphical
         TabIndex        =   7
         Tag             =   "&Update"
         Top             =   150
         Width           =   615
      End
   End
   Begin VB.Label Label6 
      Caption         =   "Valor"
      Height          =   255
      Left            =   4785
      TabIndex        =   20
      Top             =   1035
      Width           =   675
   End
   Begin VB.Label Label5 
      Caption         =   "Nota Fiscal"
      Height          =   285
      Left            =   2925
      TabIndex        =   19
      Top             =   600
      Width           =   915
   End
   Begin VB.Label Label4 
      Caption         =   "Vencto."
      Height          =   285
      Left            =   585
      TabIndex        =   18
      Top             =   1065
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "Emissão"
      Height          =   225
      Left            =   5250
      TabIndex        =   17
      Top             =   615
      Width           =   660
   End
   Begin VB.Label Label2 
      Caption         =   "Fornecedor"
      Height          =   255
      Left            =   375
      TabIndex        =   16
      Top             =   180
      Width           =   915
   End
   Begin VB.Label Label1 
      Caption         =   "Duplicata"
      Height          =   345
      Left            =   540
      TabIndex        =   15
      Top             =   600
      Width           =   765
   End
End
Attribute VB_Name = "FrmApagar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private lIncluir As Boolean
Private lPrimeiro As Boolean
Private lbaixar As Boolean
Private prsFornece As New ADODB.Recordset

Private Sub Carrega_tela()
   'Limpa as variaveis da tela se caso ficarem com dados da outra tela
   'limpa_tela Me
   'Carrega a tela com os dados do registro
   With VaSpread1
      '.Row = 1
      '.Row = .ActiveRow
      .Col = 0: TxtDuplicata.Text = .Text
      .Col = 1: TxtNotafiscal.Text = .Text
      .Col = 2
      gSql = "select tab_fornece.codfor "
      gSql = gSql & " FROM tab_fornece "
      gSql = gSql & " Where tab_fornece.nome = '" & .Text & "'"
      prsFornece.Open gSql, ConDb, adOpenKeyset
      If Not prsFornece.EOF And Not prsFornece.BOF Then
         For i = 0 To Cbofornecedor.ListCount - 1
             If Cbofornecedor.ItemData(i) = prsFornece!codfor Then
                Cbofornecedor.ListIndex = i
                Exit For
             End If
         Next
      Else
         Cbofornecedor.ListIndex = -1
      End If
      prsFornece.Close
      .Col = 3:  Me.MskEmissao.Text = .Text
      .Col = 4:  Me.MskVencto.Text = .Text
      .Col = 5:  Me.TxtValor.Text = Format(.Text, "###,##0.00")
      .Col = 6:  Me.MskdtPagto.Text = f_nulo(.Text, "__/__/____")
      .Col = 7:  Me.TxtValorPago.Text = Format(.Text, "###,##0.00")
      .TopRow = .Row
   End With

End Sub

Private Sub Carrega_Grid()
   'vaSpread1.DataSource = gRs
   VaSpread1.Rows = gRs.RecordCount + 1
   gRs.MoveFirst
   For i = 1 To gRs.RecordCount
       VaSpread1.Row = i
       VaSpread1.Col = 0
       VaSpread1.Text = gRs!duplicata
       VaSpread1.Col = 1
       VaSpread1.Text = gRs!notafiscal
       VaSpread1.Col = 2
       VaSpread1.Text = gRs!nome
       VaSpread1.Col = 3
       VaSpread1.Text = Format(gRs!datamov, "dd/mm/yyyy")
       VaSpread1.Col = 4
       VaSpread1.Text = Format(gRs!vencto, "dd/mm/yyyy")
       VaSpread1.Col = 5
       VaSpread1.Text = Format(f_nulo(gRs!Valor, 0), "###,##0.00")
       VaSpread1.Col = 6
       VaSpread1.Text = Format(f_nulo(gRs!dtpagto, ""), "dd/mm/yyyy")
       VaSpread1.Col = 7
       VaSpread1.Text = Format(f_nulo(gRs!valorpago, 0), "###,##0.00")
       gRs.MoveNext
   Next
   VaSpread1.Row = 1
   VaSpread1.Col = 1
   
End Sub
Private Sub Abre_Le_rst()
   gSql = "SELECT duplicata, notafiscal,B.Nome,"
   gSql = gSql & "datamov,vencto,valor,dtpagto,valorpago "
   gSql = gSql & "FROM tab_apagar A, tab_fornece B "
   gSql = gSql & "Where a.codfor = b.codfor"
   gRs.Open gSql, ConDb, adOpenKeyset
   
End Sub

Private Sub cmdAdd_Click()
   
   lIncluir = True
   limpa_tela Me
   
   Me.TxtValorPago.Enabled = False
   Me.MskdtPagto.Enabled = False
   
   Me.TxtDuplicata.SetFocus
   
   Habilita Me
   FramePagto.Enabled = False
   suCmdAdd Me  'Habilita e desabilita botoes

End Sub

Private Sub CmdBaixar_Click()
   If lbaixar = False Then
      If IsDate(MskdtPagto.Text) Then
         MsgBox "Não Pode baixar Título já baixado", vbOKOnly, "Atenção " & gOperador
         Exit Sub
      End If
      
      Desabilita Me
      Me.FramePagto.Enabled = True
      Me.MskdtPagto.Text = CDate(Date)
      Me.TxtValorPago.Text = TxtValor
      Me.cmdUpdate.Enabled = False
      suCmdAdd Me  'Habilita e desabilita botoes
      Me.cmdUpdate.Enabled = False
      
      Me.CmdBaixar.Caption = "Confirma"
      'Me.MskdtPagto.SetFocus
      lbaixar = True
   Else
      lbaixar = False
      Me.CmdBaixar.Caption = "Baixar"
      lIncluir = False
      Call cmdUpdate_Click
   End If
        
End Sub

Private Sub cmdDelete_Click()
    'this may produce an error if you delete the last
    'record or the only record in the recordset
    If MsgBox("Deseja realmente apagar este registro ? ", vbYesNo, "Atenção") = vbYes Then
       gSql = "delete * from tab_apagar where duplicata = '" & TxtDuplicata.Text & "' AND "
       gSql = gSql & " codfor = " & Cbofornecedor.ItemData(Cbofornecedor.ListIndex)
       gSql = gSql & " AND notafiscal = '" & TxtNotafiscal.Text & "'"
       gSql = gSql & " AND datamov = cdate('" & MskEmissao.Text & "')"
       gSql = gSql & " AND vencto = cdate('" & MskVencto.Text & "')"
       ConDb.Execute gSql
       Abre_Le_rst
       Carrega_Grid
       gRs.MoveFirst
       Carrega_tela
       Desabilita Me
       gRs.Close
    End If
End Sub

Private Sub cmddesfaz_Click()
  lIncluir = False
  lbaixar = False
  Me.CmdBaixar.Caption = "Baixar"
  'Carrega_tela
  Desabilita Me
  Me.FramePagto.Enabled = False
  
  suCmdDesfaz Me
  
  Call VaSpread1_Click
  
End Sub

Private Sub cmdEditar_Click()
   Habilita Me
   
   suCmdEditar Me
   FramePagto.Enabled = False
   Me.TxtDuplicata.SetFocus
   
End Sub

Private Sub CmdSair_Click()
   Unload Me
End Sub

Private Sub cmdUpdate_Click()
 
   If lIncluir Then
      suInsert
      lIncluir = False
   Else
      If lbaixar = False Then
         gSql = "UPDATE tab_apagar SET"
         gSql = gSql & " codfor = " & Cbofornecedor.ItemData(Cbofornecedor.ListIndex)
         gSql = gSql & ",duplicata = '" & Me.TxtDuplicata.Text & "'"
         gSql = gSql & ",datamov = "
         If Me.MskEmissao.Text <> "" Then
            gSql = gSql & "'" & CDate(Me.MskEmissao.Text) & "'"
         Else
            gSql = gSql & "NULL"
         End If
         gSql = gSql & ",vencto = "
         If Me.MskVencto.Text <> "" Then
            gSql = gSql & "'" & CDate(Me.MskVencto.Text) & "'"
         Else
            gSql = gSql & "NULL"
         End If
         gSql = gSql & ",valor = " & Replace(CDbl(f_nulo(Me.TxtValor.Text, 0)), ",", ".")
         gSql = gSql & ",notafiscal = '" & Me.TxtNotafiscal.Text & "'"
         gSql = gSql & ",dtpagto = "
         If Me.MskdtPagto.Text = "__/__/____" Then
            gSql = gSql & "NULL"
         Else
            gSql = gSql & "'" & CDate(Me.MskdtPagto.Text) & "'"
         End If

         gSql = gSql & ",valorpago = " & Replace(CDbl(f_nulo(Me.TxtValorPago.Text, 0)), ",", ".")
         gSql = gSql & ",operador = '" & gOperador & "'"
         gSql = gSql & ",datatual = Cdate('" & Date & "')"
         gSql = gSql & "  where duplicata = '" & TxtDuplicata.Text & "' AND "
         gSql = gSql & " codfor = " & Cbofornecedor.ItemData(Cbofornecedor.ListIndex)
         gSql = gSql & " AND notafiscal = '" & TxtNotafiscal.Text & "'"
         ConDb.Execute gSql
      Else
         gSql = "UPDATE tab_apagar SET"
         gSql = gSql & " dtpagto = "
         If Me.MskdtPagto.Text <> "" Then
            gSql = gSql & "'" & CDate(Me.MskdtPagto.Text) & "'"
         Else
            MsgBox "Não digitou a data de pagamento ", vbOKOnly, "Atenção " & gOperador
            Exit Sub
            'gSql = gSql & "NULL"
            
         End If
         gSql = gSql & ",valorpago = " & Replace(CDbl(f_nulo(Me.TxtValorPago.Text, 0)), ",", ".")
         gSql = gSql & ",operador = '" & gOperador & "'"
         gSql = gSql & ",datatual = Cdate('" & Date & "')"
         gSql = gSql & "  where duplicata = '" & TxtDuplicata.Text & "' AND "
         gSql = gSql & " codfor = " & Cbofornecedor.ItemData(Cbofornecedor.ListIndex)
         gSql = gSql & " AND notafiscal = '" & TxtNotafiscal.Text & "'"
         ConDb.Execute gSql
         lbaixar = False
         FramePagto.Enabled = False
      End If
      
   End If
                              
   'Deixa os textbox desabilitados
   'Me.MskFixo.Enabled = False
   'Me.MskFixo.Text = ""
   
   Abre_Le_rst
   Carrega_Grid
   gRs.MoveLast
   Carrega_tela
   Desabilita Me
   suCmdUpdate Me
   gRs.Close
   
End Sub

Private Sub Form_Activate()
   
   Call Abre_Le_rst
   
   Call Carrega_combo_fornece
      
   If gRs.BOF And gRs.EOF Then
      If MsgBox("Arquivo vazio. Incluir dados agora ?", vbYesNo, "Atenção ") = vbYes Then
         '**--> função para dar o INSERT --->
         Cbofornecedor.ListIndex = 0
         suInsert
         Abre_Le_rst
         'Me.LblCodclie.Caption = gRs!codcli
         cmdEditar_Click
         lPrimeiro = True
         Exit Sub
      Else
         Desabilita Me
      End If
      
   Else
      'gRs.MoveFirst
      Carrega_Grid
      VaSpread1.Row = 1
      Carrega_tela
      Desabilita Me
      lIncluir = False
      lPrimeiro = False
   End If
   
   gRs.Close
   
   lIncluir = False

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
       SendKeys "{TAB}"
       KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()
  Call Centra(Me)
  lbaixar = False
   
End Sub
Private Sub suInsert()

    gSql = "INSERT INTO tab_apagar (codfor,duplicata,datamov,vencto,valor,"
    gSql = gSql & "notafiscal,dtpagto,valorpago,"
    gSql = gSql & "operador,datatual) "
    gSql = gSql & "VALUES (" & Cbofornecedor.ItemData(Cbofornecedor.ListIndex) & ",'"
    gSql = gSql & Me.TxtDuplicata.Text & "',"
    If Me.MskEmissao.Text <> "" Then
       gSql = gSql & "CDate('" & Me.MskEmissao.Text & "'),"
    Else
       gSql = gSql & "NULL,"
    End If
    If Me.MskVencto.Text <> "" Then
       gSql = gSql & "CDate('" & Me.MskVencto.Text & "'),"
    Else
       gSql = gSql & "NULL,"
    End If
    gSql = gSql & Replace(CDbl(f_nulo(Me.TxtValor.Text, 0)), ",", ".") & ",'"
    gSql = gSql & Me.TxtNotafiscal.Text & "',"
    
    If Me.MskdtPagto.Text = "__/__/____" Or Me.MskdtPagto.Text = "" Then
       gSql = gSql & "NULL,"
    Else
       gSql = gSql & "cdate('" & Me.MskdtPagto.Text & "'),"
    End If
    gSql = gSql & Replace(CDbl(f_nulo(Me.TxtValorPago.Text, 0)), ",", ".") & ",'"
    gSql = gSql & gOperador & "',Cdate('" & Date & "'))"
    ConDb.Execute gSql
    
End Sub

Private Sub MskdtPagto_GotFocus()
   With MskdtPagto
      .SelStart = 0
      .SelLength = Len(.Text)
   End With

End Sub

Private Sub MskdtPagto_Validate(Cancel As Boolean)
   If Not IsDate(MskdtPagto) Then
      MsgBox "Data Inválida", vbOKOnly, "Atenção " & gOperador
      Cancel = True
   End If

End Sub

Private Sub MskEmissao_GotFocus()
   With MskEmissao
      .SelStart = 0
      .SelLength = Len(.Text)
   End With


End Sub

Private Sub MskEmissao_Validate(Cancel As Boolean)
   If Not IsDate(MskEmissao) Then
      MsgBox "Data Inválida", vbOKOnly, "Atenção " & gOperador
      Cancel = True
   End If

End Sub

Private Sub MskVencto_GotFocus()
   With MskVencto
      .SelStart = 0
      .SelLength = Len(.Text)
   End With

End Sub

Private Sub MskVencto_Validate(Cancel As Boolean)
   If Not IsDate(MskVencto) Then
      MsgBox "Data Inválida", vbOKOnly, "Atenção " & gOperador
      Cancel = True
   End If

End Sub

Private Sub TxtDuplicata_GotFocus()
   With TxtDuplicata
      .SelStart = 0
      .SelLength = Len(.Text)
   End With

End Sub

Private Sub TxtNotafiscal_GotFocus()
   With TxtNotafiscal
      .SelStart = 0
      .SelLength = Len(.Text)
   End With
End Sub

Private Sub TxtValor_GotFocus()
   With TxtValor
      .SelStart = 0
      .SelLength = Len(.Text)
   End With
End Sub

Private Sub TxtValor_LostFocus()
   TxtValor.Text = Format(TxtValor.Text, "###,###,##0.00")
End Sub

Private Sub TxtValorPago_GotFocus()
   With TxtValorPago
      .SelStart = 0
      .SelLength = Len(.Text)
   End With
End Sub

Private Sub TxtValorPago_LostFocus()
   TxtValorPago.Text = Format(TxtValorPago.Text, "###,###,##0.00")
End Sub


Private Sub Carrega_combo_fornece()
   
   gSql = "select codfor,Nome "
   gSql = gSql & "FROM tab_fornece "
   prsFornece.Open gSql, ConDb, adOpenKeyset
   Cbofornecedor.Clear
   With prsFornece
      '.MoveLast
      'nItem = .RecordCount
      .MoveFirst
      Do While Not .EOF
        Cbofornecedor.AddItem (prsFornece!nome)
        Cbofornecedor.ItemData(Cbofornecedor.NewIndex) = prsFornece!codfor
        .MoveNext
      Loop
  End With

  prsFornece.Close

End Sub

Private Sub VaSpread1_Click()
   PintaGrid VaSpread1
   Carrega_tela
End Sub
