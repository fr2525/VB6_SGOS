VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form FrmUnid 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tipos de Unidade "
   ClientHeight    =   4305
   ClientLeft      =   1650
   ClientTop       =   2025
   ClientWidth     =   7080
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4305
   ScaleWidth      =   7080
   ShowInTaskbar   =   0   'False
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   1800
      Left            =   435
      TabIndex        =   12
      Top             =   1155
      Width           =   6120
      _ExtentX        =   10795
      _ExtentY        =   3175
      _Version        =   393216
      Cols            =   3
      FixedCols       =   0
      FormatString    =   $"FrmUnid.frx":0000
   End
   Begin VB.TextBox TxtQtde 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Left            =   3120
      MaxLength       =   9
      TabIndex        =   3
      Top             =   510
      Width           =   1020
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   1410
      TabIndex        =   4
      Top             =   3300
      Width           =   4245
      Begin VB.CommandButton cmddesfaz 
         Caption         =   "&Desfaz"
         Enabled         =   0   'False
         Height          =   540
         Left            =   2865
         Picture         =   "FrmUnid.frx":0060
         Style           =   1  'Graphical
         TabIndex        =   10
         Tag             =   "&Update"
         Top             =   150
         Width           =   615
      End
      Begin VB.CommandButton CmdSair 
         Caption         =   "&Sair"
         Height          =   540
         Left            =   3555
         Picture         =   "FrmUnid.frx":015A
         Style           =   1  'Graphical
         TabIndex        =   9
         Tag             =   "&Update"
         Top             =   135
         Width           =   615
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Incluir"
         Height          =   540
         Left            =   105
         Picture         =   "FrmUnid.frx":0254
         Style           =   1  'Graphical
         TabIndex        =   8
         Tag             =   "&Add"
         Top             =   150
         Width           =   615
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Excluir"
         Height          =   540
         Left            =   1485
         Picture         =   "FrmUnid.frx":033E
         Style           =   1  'Graphical
         TabIndex        =   7
         Tag             =   "&Delete"
         Top             =   150
         Width           =   615
      End
      Begin VB.CommandButton cmdEditar 
         Caption         =   "&Editar"
         Height          =   540
         Left            =   795
         Picture         =   "FrmUnid.frx":04B0
         Style           =   1  'Graphical
         TabIndex        =   6
         Tag             =   "&Refresh"
         Top             =   150
         Width           =   615
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Salvar"
         Enabled         =   0   'False
         Height          =   540
         Left            =   2190
         Picture         =   "FrmUnid.frx":0622
         Style           =   1  'Graphical
         TabIndex        =   5
         Tag             =   "&Update"
         Top             =   150
         Width           =   615
      End
   End
   Begin VB.TextBox TxtDescricao 
      Height          =   285
      Left            =   3120
      TabIndex        =   2
      Top             =   120
      Width           =   3270
   End
   Begin VB.Label Label1 
      Caption         =   "Descrição"
      Height          =   255
      Left            =   2250
      TabIndex        =   13
      Top             =   165
      Width           =   795
   End
   Begin VB.Label LblEntrada 
      AutoSize        =   -1  'True
      Caption         =   "Qtdes.por Unidade:"
      Height          =   195
      Left            =   1620
      TabIndex        =   11
      Top             =   540
      Width           =   1380
   End
   Begin VB.Label Lblcodunid 
      Caption         =   "codunid"
      ForeColor       =   &H000000FF&
      Height          =   270
      Left            =   1320
      TabIndex        =   1
      Top             =   165
      Width           =   525
   End
   Begin VB.Label lblTipo 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Código:"
      Height          =   195
      Index           =   0
      Left            =   645
      TabIndex        =   0
      Tag             =   "CODVEND:"
      Top             =   165
      Width           =   540
   End
End
Attribute VB_Name = "FrmUnid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private lIncluir As Boolean
Private lPrimeiro As Boolean

Private Sub Abre_Le_rst()
   gSql = "select * FROM tab_uni"
   gRs.Open gSql, ConDb, adOpenKeyset
   
End Sub

Private Sub Carrega_Grid()
'Teste do MsFlexgrid1
  
  MSFlexGrid1.Row = 0
  
  MSFlexGrid1.Col = 0
  MSFlexGrid1.Text = "Codigo"
  MSFlexGrid1.ColWidth(0) = 600
  MSFlexGrid1.Col = 1
  MSFlexGrid1.Text = "Descrição"
  MSFlexGrid1.ColWidth(1) = 4000
  MSFlexGrid1.Col = 2
  MSFlexGrid1.Text = "Qtdes. por Unidade"
  MSFlexGrid1.ColWidth(2) = 1200
  
  MSFlexGrid1.Row = 0
    
  With gRs
      '.MoveLast
      'nItem = .RecordCount
      '.MoveFirst
      MSFlexGrid1.Rows = 1
      Do While Not .EOF
         MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
         MSFlexGrid1.Row = MSFlexGrid1.Rows - 1
         MSFlexGrid1.Col = 0: MSFlexGrid1.Text = f_nulo(!uni_cod, "")
         MSFlexGrid1.Col = 1: MSFlexGrid1.Text = f_nulo(!uni_desc, "")
         MSFlexGrid1.Col = 2: MSFlexGrid1.Text = f_nulo(!uni_qtd, "")
         .MoveNext
       Loop
       MSFlexGrid1.FixedRows = 1
          
  End With
  
  End Sub

Private Sub Carrega_tela()
   'Limpa as variaveis da tela se caso ficarem com dados da outra tela
   limpa_tela Me
   'Carrega a tela com os dados do registro
   Me.Lblcodunid.Caption = gRs("uni_cod")
   Me.TxtDescricao.Text = gRs("uni_desc")
   Me.TxtQtde.Text = gRs("uni_qtd")
   
End Sub
Private Sub cmdAdd_Click()
   
   lIncluir = True
   limpa_tela Me
   
   Me.Lblcodunid.Caption = ""
   Me.TxtDescricao.SetFocus
   Me.cmdUpdate.Enabled = True
   Me.cmddesfaz.Enabled = True
   Me.cmdEditar.Enabled = False
   Me.cmdAdd.Enabled = False
   Me.CmdSair.Enabled = False
   Me.cmdDelete.Enabled = False
   
End Sub

Private Sub cmdDelete_Click()

    On Error GoTo ErroDelete
    
    'this may produce an error if you delete the last
    'record or the only record in the recordset
    If MsgBox("Deseja realmente apagar esta Unidade ? ", vbYesNo, "Atenção") = vbYes Then
        gSql = "delete * from tipovend where uni_cod = " & Val(Me.Lblcodunid.Caption)
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
     MsgBox "Deu erro na exclusao do Tipo de Unidade " & Chr(13) & "Instrucao Sql = '" & _
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
   Me.TxtDescricao.SetFocus
   
End Sub

Private Sub CmdSair_Click()
   Unload Me
End Sub

Private Sub cmdUpdate_Click()
   gRs.Close
   If lIncluir Then
      gSql = "INSERT INTO tab_uni (uni_cod,uni_desc,Uni_qtd,operador,datatual) "
      gSql = gSql & "VALUES ('" & Me.Lblcodunid.Caption & "','"
      gSql = gSql & Me.TxtDescricao.Text & "','" & Me.TxtQtde.Text & "','"
      gSql = gSql & "'" & gOperador & "',Cdate('" & Date & "'))"
      ConDb.Execute gSql
      lIncluir = False
      
   Else
      gSql = "UPDATE tab_uni SET uni_des = '" & Me.TxtDescricao.Text & "',"
      gSql = gSql & "uni_qtd = '" & Me.TxtQtde.Text & "',"
      gSql = gSql & " operador = '" & gOperador & "', datatual = Cdate('" & Date & "')"
      gSql = gSql & " WHERE uni_cod = " & Val(Lblcodunid.Caption)
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
    
  Me.Lblcodunid.Caption = ""
  If gRs.BOF And gRs.EOF Then
      If MsgBox("Arquivo vazio. Incluir dados agora ?", vbYesNo, "Atenção ") = vbYes Then
         gSql = "INSERT INTO tab_uni (uni_cod,uni_desc,uni_qtd,operador,datatual) "
         gSql = gSql & "VALUES ('01','" & Me.TxtDescricao.Text & "','"
         gSql = gSql & Me.TxtQtde.Text & "',"
         gSql = gSql & "'" & gOperador & "'," & Date & " )"
         ConDb.Execute gSql
         gRs.Close
         Abre_Le_rst
         Me.Lblcodunid.Caption = gRs!uni_cod
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
    Screen.MousePointer = vbDefault
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
    .Redraw = True
    
    .Row = oldrow
    
    .Col = 0:   Lblcodunid.Caption = .Text: .CellBackColor = vbYellow
    .Col = 1:   TxtDescricao.Text = .Text: .CellBackColor = vbYellow
    .Col = 2:   TxtQtde.Text = .Text: .CellBackColor = vbYellow
    .TopRow = .Row
   
End With

End Sub

