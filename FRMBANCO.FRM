VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmBanco 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tabela de Bancos"
   ClientHeight    =   4455
   ClientLeft      =   2565
   ClientTop       =   2835
   ClientWidth     =   5895
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4455
   ScaleWidth      =   5895
   ShowInTaskbar   =   0   'False
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   1575
      Left            =   240
      TabIndex        =   11
      Top             =   1200
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   2778
      _Version        =   393216
      Rows            =   6
      FixedCols       =   0
      FocusRect       =   2
      HighLight       =   2
      ScrollBars      =   2
      SelectionMode   =   1
      FormatString    =   "Código| Nome                                                                                                          "
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   600
      TabIndex        =   4
      Top             =   3000
      Width           =   4455
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Salvar"
         Enabled         =   0   'False
         Height          =   540
         Left            =   2280
         Picture         =   "frmbanco.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   10
         Tag             =   "&Update"
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton cmdEditar 
         Caption         =   "&Editar"
         Height          =   540
         Left            =   840
         Picture         =   "frmbanco.frx":00FA
         Style           =   1  'Graphical
         TabIndex        =   9
         Tag             =   "&Refresh"
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Excluir"
         Height          =   540
         Left            =   1560
         Picture         =   "frmbanco.frx":026C
         Style           =   1  'Graphical
         TabIndex        =   8
         Tag             =   "&Delete"
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Incluir"
         Height          =   540
         Left            =   120
         Picture         =   "frmbanco.frx":03DE
         Style           =   1  'Graphical
         TabIndex        =   7
         Tag             =   "&Add"
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton CmdSair 
         Caption         =   "&Sair"
         Height          =   540
         Left            =   3720
         Picture         =   "frmbanco.frx":04C8
         Style           =   1  'Graphical
         TabIndex        =   6
         Tag             =   "&Update"
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton cmddesfaz 
         Caption         =   "&Desfaz"
         Enabled         =   0   'False
         Height          =   540
         Left            =   3000
         Picture         =   "frmbanco.frx":05C2
         Style           =   1  'Graphical
         TabIndex        =   5
         Tag             =   "&Update"
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.TextBox TxtCodbco 
      Height          =   285
      Left            =   1080
      MaxLength       =   3
      TabIndex        =   0
      Top             =   120
      Width           =   495
   End
   Begin VB.TextBox TxtNomeBco 
      Height          =   285
      Left            =   1080
      MaxLength       =   30
      TabIndex        =   2
      Top             =   480
      Width           =   4170
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Código:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Tag             =   "CODVEND:"
      Top             =   120
      Width           =   735
   End
   Begin VB.Label LblNomebco 
      Alignment       =   1  'Right Justify
      Caption         =   "Nome:"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Tag             =   "NOME:"
      Top             =   480
      Width           =   675
   End
End
Attribute VB_Name = "FrmBanco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private rst As New ADODB.Recordset
Private cSql As String
Private lIncluir As Boolean
Private lPrimeiro As Boolean
Private vRegAtual As Variant
Private nItem As Integer

Private Sub Carrega_Grid()
  
 'Teste do MsHFlexgrid1 - eh eh eh
  With MSFlexGrid1
  '    .BackColorFixed = vbCyan
  '    .Clear
  '    .Rows = 2
  '    .Cols = 3
  '    .Row = 0
  '    .FixedCols = 0
  '    .FixedRows = 1
  '    '.Height = 5625: .Width = 7600:
  '    .Row = 0
  '    .Col = 0: .ColWidth(0) = 800: .Text = "Código": .ColAlignment(0) = flexAlignLeftCenter
  '    .Col = 1: .ColWidth(1) = 5700: .Text = "Banco": .ColAlignment(1) = flexAlignLeftCenter
        
  '     .SelectionMode = flexSelectionByRow
 
  End With
  MSFlexGrid1.Row = 0
  
  With rst
      .MoveLast
      nItem = .RecordCount
      .MoveFirst
      MSFlexGrid1.Rows = 1
      If .AbsolutePosition <> -1 Then
         Do While Not .EOF
            MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
            MSFlexGrid1.Row = MSFlexGrid1.Rows - 1
            MSFlexGrid1.Col = 0: MSFlexGrid1.Text = IIf(IsNull(!codbco), " ", !codbco)
            MSFlexGrid1.Col = 1: MSFlexGrid1.Text = IIf(IsNull(!nomedobanco), " ", !nomedobanco)
            .MoveNext
            
          Loop
          MSFlexGrid1.FixedRows = 1
      End If
          
  End With
  
  
End Sub

Private Sub Carrega_tela()
   limpa_tela Me
   Me.TxtCodbco.Text = rst!codbco
   If Not IsNull(rst!nomedobanco) Then Me.TxtNomeBco.Text = rst!nomedobanco
     
End Sub

Private Sub cmdAdd_Click()
   limpa_tela Me
   Me.TxtCodbco.SetFocus
   Me.cmdUpdate.Enabled = True
   Me.cmddesfaz.Enabled = True
   Me.cmdEditar.Enabled = False
   Me.cmdAdd.Enabled = False
   Me.CmdSair.Enabled = False
   Me.cmdDelete.Enabled = False
   lIncluir = True

End Sub

Private Sub cmdDelete_Click()
    'this may produce an error if you delete the last
    'record or the only record in the recordset
    If MsgBox("Deseja realmente apagar este Banco ? ", vbYesNo, "Atenção") = vbYes Then
        rst.Close
        cSql = "DELETE FROM Cadbco WHERE cadbco.codbco = '" & Me.TxtCodbco.Text & "'"
        cnn.Execute cSql
        On Error GoTo ErroDelete
        Abre_Le_rst
        rst.MoveFirst
              
        Carrega_tela
        Desabilita Me
        Carrega_Grid
     End If
     Exit Sub
     
ErroDelete:
     MsgBox "Deu erro na Exclusao do Banco " & Chr(13) & "cSql = '" & _
            cSql & "'  "
End Sub

Private Sub cmddesfaz_Click()
  lIncluir = False
  Desabilita Me
    
  Me.cmdUpdate.Enabled = False
  Me.cmddesfaz.Enabled = False
  Me.cmdEditar.Enabled = True
  Me.cmdAdd.Enabled = True
  Me.CmdSair.Enabled = True
  Me.cmdDelete.Enabled = True
  
End Sub

Private Sub cmdEditar_Click()
   Habilita Me
   Me.TxtCodbco.SetFocus
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
      'rst.AddNew
      rst.Close
      cSql = "INSERT INTO cadbco (codbco,nomedobanco,operador,datatual) " & _
                          "VALUES ( '" & Me.TxtCodbco.Text & "','" & _
                                         Me.TxtNomeBco.Text & "','" & _
                                         gOperador & "','" & _
                                         Now & "')"
      cnn.Execute cSql
                          
      lIncluir = False
   Else
      rst.Close
      cSql = "UPDATE cadbco SET codbco = '" & Me.TxtCodbco.Text & "'," & _
                                " Nomedobanco = '" & Me.TxtNomeBco.Text & _
                                "', operador = '" & gOperador & _
                                "', datatual = '" & Now & "'" & _
                                " WHERE codbco = '" & Me.TxtCodbco.Text & "'"
      cnn.Execute cSql
      
      'If Not lPrimeiro Then
      '   rst.Edit
      'Else
         lPrimeiro = False
      'End If
   End If
   
   Abre_Le_rst
   
   'rst!codbco = Me.TxtCodbco.Text
   'rst!nomedobanco = Me.TxtNomeBco.Text
   'rst!Operador = gOperador
   'rst!Datatual = Now
   'rst.Update
   'rst.MoveLast
   
   'Abre_Le_rst
      
   Carrega_tela
   'Deixa os textbox desabilitados
   Desabilita Me
   
   Me.cmdUpdate.Enabled = False
   Me.cmddesfaz.Enabled = False
   Me.cmdEditar.Enabled = True
   Me.cmdAdd.Enabled = True
   Me.CmdSair.Enabled = True
   Me.cmdDelete.Enabled = True
       
   Carrega_Grid
 
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub Form_Load()
  
   'Abrindo o Recordset
   Abre_Le_rst
          
   'Centraliza a tela no video
   Me.Move (Screen.Width - Me.Width) / 2, _
           (Screen.Height - Me.Height) / 2
         
  If rst.BOF And rst.EOF Then
      If MsgBox("Arquivo vazio. Incluir dados agora ?", vbYesNo, "Atenção ") = vbYes Then
         'rst.AddNew
         With rst
           .AddNew
           !nomedobanco = ""
           .Update
         End With
         cmdEditar_Click
         lPrimeiro = True
      Else
         Desabilita Me
      End If
      
   Else
      rst.MoveFirst
      Carrega_tela
    
      Desabilita Me
      lIncluir = False
      lPrimeiro = False
   End If
   Carrega_Grid
  
  End Sub

Private Sub Abre_Le_rst()
   cSql = "select * from cadbco"
   rst.Open cSql, cnn, adOpenKeyset, adLockOptimistic, adCmdText
End Sub
   
      
Private Sub Form_Unload(Cancel As Integer)
   Screen.MousePointer = vbDefault
   rst.Close
End Sub

Private Sub Grid1_Click()

 End With

End Sub

Private Sub vaSpr1_Click(ByVal Col As Long, ByVal Row As Long)
    vaSpr1.Row = Row
    vaSpr1.Col = 1:   TxtCodbco.Text = vaSpr1.Text
    vaSpr1.Col = 2:   TxtNomeBco.Text = vaSpr1.Text
'    rst.MoveFirst
'    Do While Not rst.EOF
'       If rst!codbco = TxtCodbco.Text Then
'          Exit Do
'       End If
'       rst.MoveNext
'    Loop
    
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
    
    .Col = 0:   TxtCodbco.Text = .Text: .CellBackColor = vbYellow
    .Col = 1:   TxtNomeBco.Text = .Text: .CellBackColor = vbYellow
    .Redraw = True
  End With

End Sub
