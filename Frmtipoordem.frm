VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmtipoordem 
   Caption         =   "Tipos de Ordem"
   ClientHeight    =   3960
   ClientLeft      =   3630
   ClientTop       =   1890
   ClientWidth     =   5715
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   3960
   ScaleWidth      =   5715
   Begin VB.TextBox TxtE_S 
      Height          =   315
      Left            =   4860
      TabIndex        =   8
      Text            =   "Text2"
      Top             =   540
      Width           =   315
   End
   Begin VB.TextBox TxtDescricao 
      Height          =   315
      Left            =   1200
      TabIndex        =   7
      Top             =   540
      Width           =   3030
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   870
      TabIndex        =   0
      Top             =   3000
      Width           =   4245
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Salvar"
         Enabled         =   0   'False
         Height          =   540
         Left            =   2190
         Picture         =   "frmtipoordem.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   6
         Tag             =   "&Update"
         Top             =   150
         Width           =   615
      End
      Begin VB.CommandButton cmdEditar 
         Caption         =   "&Editar"
         Height          =   540
         Left            =   795
         Picture         =   "frmtipoordem.frx":00FA
         Style           =   1  'Graphical
         TabIndex        =   5
         Tag             =   "&Refresh"
         Top             =   150
         Width           =   615
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Excluir"
         Height          =   540
         Left            =   1485
         Picture         =   "frmtipoordem.frx":026C
         Style           =   1  'Graphical
         TabIndex        =   4
         Tag             =   "&Delete"
         Top             =   150
         Width           =   615
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Incluir"
         Height          =   540
         Left            =   105
         Picture         =   "frmtipoordem.frx":03DE
         Style           =   1  'Graphical
         TabIndex        =   3
         Tag             =   "&Add"
         Top             =   150
         Width           =   615
      End
      Begin VB.CommandButton CmdSair 
         Caption         =   "&Sair"
         Height          =   540
         Left            =   3555
         Picture         =   "frmtipoordem.frx":04C8
         Style           =   1  'Graphical
         TabIndex        =   2
         Tag             =   "&Update"
         Top             =   150
         Width           =   615
      End
      Begin VB.CommandButton cmddesfaz 
         Caption         =   "&Desfaz"
         Enabled         =   0   'False
         Height          =   540
         Left            =   2865
         Picture         =   "frmtipoordem.frx":05C2
         Style           =   1  'Graphical
         TabIndex        =   1
         Tag             =   "&Update"
         Top             =   150
         Width           =   615
      End
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   1725
      Left            =   780
      TabIndex        =   9
      Top             =   1050
      Width           =   4305
      _ExtentX        =   7594
      _ExtentY        =   3043
      _Version        =   393216
      Rows            =   5
      Cols            =   3
      FixedCols       =   0
      ScrollBars      =   2
      FormatString    =   "Tipo|Descrição                                                       |E/S"
   End
   Begin VB.Label Lbltipomov 
      Caption         =   "Label3"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   1230
      TabIndex        =   13
      Top             =   210
      Width           =   465
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "E/S"
      Height          =   195
      Left            =   4440
      TabIndex        =   12
      Top             =   600
      Width           =   285
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Descrição:"
      Height          =   195
      Left            =   240
      TabIndex        =   11
      Top             =   630
      Width           =   765
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Tipo de Mov.:"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   10
      Tag             =   "CODVEND:"
      Top             =   210
      Width           =   990
   End
End
Attribute VB_Name = "frmtipoordem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private lIncluir As Boolean
Private lPrimeiro As Boolean

Private Sub Abre_Le_rst()
   gSql = "select * FROM tipomov"
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
            
         MSFlexGrid1.Col = 0: MSFlexGrid1.Text = f_nulo(!tipo, "")
         MSFlexGrid1.Col = 1: MSFlexGrid1.Text = f_nulo(!descricao, "")
         MSFlexGrid1.Col = 2: MSFlexGrid1.Text = f_nulo(!E_S, "")
         
         .MoveNext
         
       Loop
       MSFlexGrid1.FixedRows = 1
          
  End With
  
  End Sub

Private Sub Carrega_tela()
   
   'Carrega a tela com os dados do registro
   Me.Lbltipomov.Caption = gRs("tipo")
   Me.Txtdescricao.Text = gRs("descricao")
   Me.TxtE_S.Text = gRs("E_S")
   
End Sub

Private Sub cmdAdd_Click()
   lIncluir = True
   limpa_tela Me
   
   Me.Lbltipomov.Caption = ""
   Me.Txtdescricao.SetFocus
   Me.cmdUpdate.Enabled = True
   Me.cmddesfaz.Enabled = True
   Me.cmdEditar.Enabled = False
   Me.cmdAdd.Enabled = False
   Me.CmdSair.Enabled = False
   Me.cmdDelete.Enabled = False
   Me.Txtdescricao.SetFocus
End Sub

Private Sub cmdDelete_Click()

    On Error GoTo ErroDelete
    
    'this may produce an error if you delete the last
    'record or the only record in the recordset
    If MsgBox("Deseja realmente apagar este Tipo de Movimentação? ", vbYesNo, "Atenção") = vbYes Then
        gSql = "delete * from tipomov where tipo = " & Val(Me.Lbltipomov.Caption)
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
     MsgBox "Deu erro na exclusao do Tipo de Movimentacao " & Chr(13) & "Instrucao Sql = '" & _
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
   Me.Txtdescricao.SetFocus
End Sub

Private Sub CmdSair_Click()
   Unload Me
End Sub

Private Sub cmdUpdate_Click()
   gRs.Close
   If lIncluir Then
      gSql = "INSERT INTO tipomov (descricao,e_s,operador,datatual) "
      gSql = gSql & "VALUES ('" & Me.Txtdescricao.Text & "','" & Me.TxtE_S.Text & "','"
      gSql = gSql & gOperador & "',Cdate('" & Date & "') "
      ConDb.Execute gSql
      lIncluir = False
      
   Else
      gSql = "UPDATE tipomov SET descricao = '" & Me.Txtdescricao.Text
      gSql = gSql & "', e_s = '" & Me.TxtE_S.Text & "',"
      gSql = gSql & " operador = '" & gOperador & "', datatual = Cdate('" & Date & "')"
      gSql = gSql & " WHERE tipo = " & Val(Me.Lbltipomov.Caption)
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
   
   Me.Lbltipomov.Caption = ""
   If gRs.BOF And gRs.EOF Then
      If MsgBox("Arquivo vazio. Incluir dados agora ?", vbYesNo, "Atenção ") = vbYes Then
         gSql = "INSERT INTO tipomov (descricao,e_s,operador,datatual) "
         gSql = gSql & "VALUES ('" & Me.Txtdescricao.Text & "','" & Me.TxtE_S.Text & "','"
         gSql = gSql & gOperador & "'," & Date & " ) "
         ConDb.Execute gSql
         gRs.Close
         Abre_Le_rst
         Me.Lbltipomov.Caption = gRs!tipo
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
    
    .Col = 0:   Lbltipomov.Caption = .Text: .CellBackColor = vbYellow
    .Col = 1:   Txtdescricao.Text = .Text: .CellBackColor = vbYellow
    .Col = 2:   TxtE_S.Text = .Text: .CellBackColor = vbYellow
    .TopRow = .Row
   
End With


End Sub
