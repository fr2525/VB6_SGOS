VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmcadmarcas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Marcas"
   ClientHeight    =   4785
   ClientLeft      =   4755
   ClientTop       =   4080
   ClientWidth     =   6345
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4785
   ScaleWidth      =   6345
   ShowInTaskbar   =   0   'False
   Tag             =   "cadvend"
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   1575
      Left            =   480
      TabIndex        =   11
      Top             =   1440
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   2778
      _Version        =   393216
      Rows            =   6
      Cols            =   3
      FixedCols       =   0
      HighLight       =   2
      ScrollBars      =   2
      SelectionMode   =   1
      FormatString    =   "Codigo  | Descri��o                                                                             "
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   840
      TabIndex        =   4
      Top             =   3480
      Width           =   4455
      Begin VB.CommandButton cmddesfaz 
         Caption         =   "&Desfaz"
         Enabled         =   0   'False
         Height          =   540
         Left            =   3000
         Picture         =   "frmcadmarcas.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   10
         Tag             =   "&Update"
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton CmdSair 
         Caption         =   "&Sair"
         Height          =   540
         Left            =   3720
         Picture         =   "frmcadmarcas.frx":00FA
         Style           =   1  'Graphical
         TabIndex        =   9
         Tag             =   "&Update"
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Incluir"
         Height          =   540
         Left            =   120
         Picture         =   "frmcadmarcas.frx":01F4
         Style           =   1  'Graphical
         TabIndex        =   8
         Tag             =   "&Add"
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Excluir"
         Height          =   540
         Left            =   1560
         Picture         =   "frmcadmarcas.frx":02DE
         Style           =   1  'Graphical
         TabIndex        =   7
         Tag             =   "&Delete"
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton cmdEditar 
         Caption         =   "&Editar"
         Height          =   540
         Left            =   840
         Picture         =   "frmcadmarcas.frx":0450
         Style           =   1  'Graphical
         TabIndex        =   6
         Tag             =   "&Refresh"
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Salvar"
         Enabled         =   0   'False
         Height          =   540
         Left            =   2280
         Picture         =   "frmcadmarcas.frx":05C2
         Style           =   1  'Graphical
         TabIndex        =   5
         Tag             =   "&Update"
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.TextBox TxtDescricao 
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   1440
      TabIndex        =   0
      Top             =   420
      Width           =   4005
   End
   Begin VB.Label LblCodigo 
      Caption         =   "codigo"
      ForeColor       =   &H000000FF&
      Height          =   270
      Left            =   1425
      TabIndex        =   3
      Top             =   120
      Width           =   975
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Descri��o:"
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   2
      Tag             =   "NOME:"
      Top             =   420
      Width           =   855
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "C�digo:"
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   1
      Tag             =   "CODVEND:"
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "frmcadmarcas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private lIncluir As Boolean
Private lPrimeiro As Boolean
Private vRegAtual As Variant
Private nItem As Integer

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
         
         MSFlexGrid1.Col = 0: MSFlexGrid1.text = f_nulo(!codigo, "n/c")
         MSFlexGrid1.Col = 1: MSFlexGrid1.text = f_nulo(!nome, "n/c")
         MSFlexGrid1.Col = 2: MSFlexGrid1.text = f_nulo(!simbolo, "n/c")
         .MoveNext
      Loop
      MSFlexGrid1.FixedRows = 1
          
  End With
 
End Sub

Private Sub Carrega_tela()
   'Limpa as variaveis da tela se caso ficarem com dados da outra tela
   limpa_tela Me
   'Carrega a tela com os dados do registro
   Me.LblCodigo = gRs("codigo")
   Me.TxtDescricao.text = gRs("Nome")
   Me.TxtSimbolo.text = gRs("Simbolo")
   
End Sub
Private Sub cmdAdd_Click()
   
   limpa_tela Me
   
   Me.TxtSimbolo.Enabled = True
   Me.TxtSimbolo.text = ""
   Me.LblCodigo.Caption = ""
   Me.TxtDescricao.SetFocus
   suCmdAdd Me
   
End Sub

Private Sub cmdDelete_Click()
    'this may produce an error if you delete the last
    'record or the only record in the recordset
    If MsgBox("Deseja realmente apagar esta Moeda ? ", vbYesNo, "Aten��o") = vbYes Then
       gSql = "DELETE FROM Cadmoe WHERE cadmoe.codigo = " _
                & Me.LblCodigo.Caption & " AND cadmoe.nome = '" & Me.TxtDescricao.text & "'"
       ConDb.Execute gSql
       gRs.Close
       Abre_Le_rst
       gRs.MoveFirst
       Carrega_tela
       Desabilita Me
    End If
End Sub

Private Sub cmddesfaz_Click()
  lIncluir = False
  Desabilita Me
  MSFlexGrid1_Click
     
  suCmdDesfaz Me
  
End Sub

Private Sub cmdEditar_Click()
   Habilita Me
   suCmdEditar Me
   Me.TxtDescricao.SetFocus
   
End Sub

Private Sub CmdSair_Click()
   Unload Me
End Sub

Private Sub cmdUpdate_Click()
   gRs.Close
   If lIncluir Then
      gSql = "INSERT INTO cadmoe (nome,simbolo,operador,datatual) " & _
                          "VALUES ( '" & Me.TxtDescricao.text & "','" & _
                                         Me.TxtSimbolo.text & "','" & _
                                         gOperador & "','" & _
                                         Now & "')"
      ConDb.Execute gSql
      lIncluir = False
   Else
      gSql = "UPDATE cadmoe SET nome = '" & Me.TxtDescricao.text & "'," & _
                                " simbolo = '" & Me.TxtSimbolo.text & _
                                "', operador = '" & gOperador & _
                                "', datatual = '" & Now & "'" & _
                                " WHERE cadmoe.codigo = " & Me.LblCodigo.Caption
      ConDb.Execute gSql

   End If
   
   Abre_Le_rst
   Carrega_Grid
   gRs.MoveLast
   Carrega_tela
   Desabilita Me
   
   suCmdUpdate Me
    
End Sub

Private Sub Form_Activate()
   Abre_Le_rst
   limpa_tela Me
   
   Me.LblCodigo.Caption = ""
   
   If gRs.BOF And gRs.EOF Then
      If MsgBox("Arquivo vazio. Incluir dados agora ?", vbYesNo, "Aten��o ") = vbYes Then
         gSql = "INSERT INTO cadmoe (nome,simbolo,operador,datatual) " & _
                          "VALUES ( '" & Me.TxtDescricao.text & "','" & _
                                         Me.TxtSimbolo.text & "','" & _
                                         gOperador & "','" & _
                                         Now & "')"
         ConDb.Execute gSql
         gRs.Close
         Abre_Le_rst
         Me.LblCodigo.Caption = gRs!codigo
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
  If KeyCode = vbKeyReturn Then Sendkeys "{TAB}"
End Sub

Private Sub Form_Load()
  'Centraliza a tela no video
   Me.Move (Screen.Width - Me.Width) / 2, _
           (Screen.Height - Me.Height) / 2
   
  End Sub
Private Sub Abre_Le_rst()
   gSql = "select * from cadmoe"
   gRs.Open gSql, ConDb, adOpenKeyset, adLockOptimistic
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
       .Col = 0: .CellBackColor = vbWhite
       .Col = 1: .CellBackColor = vbWhite
       If .Row = .Rows - 1 Then
          Exit Do
       End If
    Loop
  
    .Refresh
    .Row = oldrow
    
    .Col = 0:   LblCodigo.Caption = .text: .CellBackColor = vbYellow
    .Col = 1:   TxtDescricao.text = .text: .CellBackColor = vbYellow
    .Col = 2:   TxtSimbolo.text = .text: .CellBackColor = vbYellow
     .Redraw = True
  End With

End Sub

Private Sub TxtDescricao_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Or KeyAscii = 9 Then KeyAscii = 0
End Sub

Private Sub TxtSimbolo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Or KeyAscii = 9 Then KeyAscii = 0
End Sub
