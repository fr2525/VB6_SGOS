VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmgrupos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Grupos"
   ClientHeight    =   4770
   ClientLeft      =   3555
   ClientTop       =   1935
   ClientWidth     =   5475
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4770
   ScaleWidth      =   5475
   ShowInTaskbar   =   0   'False
   Tag             =   "cadvend"
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   2595
      Left            =   1140
      TabIndex        =   9
      Top             =   660
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   4577
      _Version        =   393216
      Rows            =   5
      Cols            =   3
      FixedCols       =   0
      ScrollBars      =   2
      FormatString    =   "Código|Descrição                                      |"
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   630
      TabIndex        =   2
      Top             =   3540
      Width           =   4245
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Salvar"
         Enabled         =   0   'False
         Height          =   540
         Left            =   2190
         Picture         =   "frmgrupos.frx":0000
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
         Picture         =   "frmgrupos.frx":00FA
         Style           =   1  'Graphical
         TabIndex        =   7
         Tag             =   "&Refresh"
         Top             =   150
         Width           =   615
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Excluir"
         Height          =   540
         Left            =   1485
         Picture         =   "frmgrupos.frx":026C
         Style           =   1  'Graphical
         TabIndex        =   6
         Tag             =   "&Delete"
         Top             =   150
         Width           =   615
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Incluir"
         Height          =   540
         Left            =   105
         Picture         =   "frmgrupos.frx":03DE
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
         Picture         =   "frmgrupos.frx":04C8
         Style           =   1  'Graphical
         TabIndex        =   4
         Tag             =   "&Update"
         Top             =   150
         Width           =   615
      End
      Begin VB.CommandButton cmddesfaz 
         Caption         =   "&Desfaz"
         Enabled         =   0   'False
         Height          =   540
         Left            =   2865
         Picture         =   "frmgrupos.frx":05C2
         Style           =   1  'Graphical
         TabIndex        =   3
         Tag             =   "&Update"
         Top             =   150
         Width           =   615
      End
   End
   Begin VB.TextBox txtDescricao 
      Height          =   285
      Left            =   1845
      TabIndex        =   1
      Top             =   135
      Width           =   3390
   End
   Begin VB.Label Lblcodgrupo 
      Caption         =   "codgrupo"
      ForeColor       =   &H000000C0&
      Height          =   285
      Left            =   990
      TabIndex        =   10
      Top             =   135
      Width           =   660
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Griupo:"
      Height          =   195
      Index           =   0
      Left            =   180
      TabIndex        =   0
      Tag             =   "CODVEND:"
      Top             =   150
      Width           =   510
   End
End
Attribute VB_Name = "frmgrupos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private lIncluir As Boolean
Private lPrimeiro As Boolean

Private Sub Abre_Le_rst()
   gSql = "select * FROM tab_grupos"
   gRs.Open gSql, ConDb, adOpenKeyset

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
            
         MSFlexGrid1.Col = 0: MSFlexGrid1.Text = f_nulo(!codgrupo, "")
         MSFlexGrid1.Col = 1: MSFlexGrid1.Text = f_nulo(!descricao, "")
         
         .MoveNext
         
       Loop
       MSFlexGrid1.FixedRows = 1
          
  End With
  
End Sub

Private Sub Carrega_tela()
   'Limpa as variaveis da tela se caso ficarem com dados da outra tela
   limpa_tela Me
   'Carrega a tela com os dados do registro
   Me.Lblcodgrupo = gRs("codgrupo")
   Me.TxtDescricao.Text = gRs("descricao")
     
End Sub

Private Sub cmdAdd_Click()
   
   lIncluir = True
   
   limpa_tela Me
   
   Me.Lblcodgrupo.Caption = ""
         
   Me.cmdUpdate.Enabled = True
   Me.cmddesfaz.Enabled = True
   Me.cmdEditar.Enabled = False
   Me.cmdAdd.Enabled = False
   Me.CmdSair.Enabled = False
   Me.cmdDelete.Enabled = False
   Me.TxtDescricao.SetFocus
   lIncluir = True
   
End Sub

Private Sub cmdDelete_Click()

    On Error GoTo ErroDelete
    
    'this may produce an error if you delete the last
    'record or the only record in the recordset
    If MsgBox("Deseja realmente apagar este Grupo ? ", vbYesNo, "Atenção") = vbYes Then
       gSql = "delete * from tab_grupos where codgrupo = " & Val(Me.Lblcodgrupo.Caption)
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
     MsgBox "Deu erro na exclusao do Grupo " & Chr(13) & "Instrucao Sql = '" & _
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
    
   'Me.TxtDescricao.Enabled = True
   'Me.TxtDescricao.SetFocus
   
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
      gSql = "INSERT INTO tab_grupos (descricao,operador,datatual) "
      gSql = gSql & "VALUES ('" & Me.TxtDescricao.Text & "',"
      gSql = gSql & "'" & gOperador & "',Cdate('" & Date & "') ) "
      ConDb.Execute gSql
      lIncluir = False
   Else
      gSql = "UPDATE tab_grupos SET descricao = '" & Me.TxtDescricao.Text & "'"
      gSql = gSql & " ,operador = '" & gOperador & "', datatual = Cdate('" & Date & "')"
      gSql = gSql & " WHERE codgrupo = " & Val(Me.Lblcodgrupo.Caption)
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
   
   Me.Lblcodgrupo.Caption = ""
   If gRs.BOF And gRs.EOF Then
      If MsgBox("Arquivo vazio. Incluir dados agora ?", vbYesNo, "Atenção ") = vbYes Then
         gSql = "INSERT INTO tab_grupos (descricao,operador,datatual) "
         gSql = gSql & "VALUES ('" & Me.TxtDescricao.Text & "',"
         gSql = gSql & "'" & gOperador & "'," & Date & " ) "
         ConDb.Execute gSql
         gRs.Close
         Abre_Le_rst
         Me.Lblcodgrupo.Caption = gRs!codgrupo
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
    
    .Col = 0:   Lblcodgrupo.Caption = .Text: .CellBackColor = vbYellow
    .Col = 1:   TxtDescricao.Text = .Text: .CellBackColor = vbYellow
    .TopRow = .Row
    
  End With

End Sub

Private Sub TxtDescricao_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Or KeyAscii = 9 Then KeyAscii = 0
End Sub
