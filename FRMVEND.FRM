VERSION 5.00
Begin VB.Form frmvend 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Balconistas"
   ClientHeight    =   4815
   ClientLeft      =   4755
   ClientTop       =   4080
   ClientWidth     =   5955
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4815
   ScaleWidth      =   5955
   ShowInTaskbar   =   0   'False
   Tag             =   "cadvend"
   Begin VB.PictureBox vaSpr1 
      Height          =   1695
      Left            =   240
      ScaleHeight     =   1635
      ScaleWidth      =   5235
      TabIndex        =   19
      Top             =   1680
      Width           =   5295
   End
   Begin VB.TextBox TxtFixo 
      Alignment       =   1  'Right Justify
      Height          =   300
      Left            =   1305
      TabIndex        =   5
      Top             =   720
      Width           =   1470
   End
   Begin VB.TextBox txtNivel 
      Height          =   330
      Left            =   4275
      TabIndex        =   11
      Top             =   1095
      Width           =   375
   End
   Begin VB.TextBox TxtSenha 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1320
      PasswordChar    =   "*"
      TabIndex        =   9
      Top             =   1095
      Width           =   1470
   End
   Begin VB.TextBox TxtComissao 
      Alignment       =   1  'Right Justify
      Height          =   300
      Left            =   4290
      TabIndex        =   7
      Top             =   720
      Width           =   1020
   End
   Begin VB.TextBox TxtNome 
      Height          =   285
      Left            =   1305
      TabIndex        =   3
      Top             =   375
      Width           =   4005
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   720
      TabIndex        =   12
      Top             =   3600
      Width           =   4695
      Begin VB.CommandButton cmddesfaz 
         Caption         =   "&Desfaz"
         Enabled         =   0   'False
         Height          =   540
         Left            =   3105
         Picture         =   "frmvend.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   18
         Tag             =   "&Update"
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton CmdSair 
         Caption         =   "&Sair"
         Height          =   540
         Left            =   3795
         Picture         =   "frmvend.frx":00FA
         Style           =   1  'Graphical
         TabIndex        =   17
         Tag             =   "&Update"
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Incluir"
         Height          =   540
         Left            =   345
         Picture         =   "frmvend.frx":01F4
         Style           =   1  'Graphical
         TabIndex        =   16
         Tag             =   "&Add"
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Excluir"
         Height          =   540
         Left            =   1725
         Picture         =   "frmvend.frx":02DE
         Style           =   1  'Graphical
         TabIndex        =   15
         Tag             =   "&Delete"
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton cmdEditar 
         Caption         =   "&Editar"
         Height          =   540
         Left            =   1035
         Picture         =   "frmvend.frx":0450
         Style           =   1  'Graphical
         TabIndex        =   14
         Tag             =   "&Refresh"
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Salvar"
         Enabled         =   0   'False
         Height          =   540
         Left            =   2415
         Picture         =   "frmvend.frx":05C2
         Style           =   1  'Graphical
         TabIndex        =   13
         Tag             =   "&Update"
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Label LblCodvend 
      Caption         =   "codvend"
      ForeColor       =   &H000000FF&
      Height          =   270
      Left            =   1305
      TabIndex        =   1
      Top             =   75
      Width           =   975
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Nível:"
      Height          =   255
      Index           =   5
      Left            =   3240
      TabIndex        =   10
      Tag             =   "NIVEL:"
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Senha:"
      Height          =   255
      Index           =   4
      Left            =   240
      TabIndex        =   8
      Tag             =   "SENHA:"
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Comissão:"
      Height          =   255
      Index           =   3
      Left            =   3120
      TabIndex        =   6
      Tag             =   "COMISSAO:"
      Top             =   720
      Width           =   975
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Fixo:"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   4
      Tag             =   "SALARIO:"
      Top             =   705
      Width           =   975
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Nome:"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Tag             =   "NOME:"
      Top             =   375
      Width           =   975
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Código:"
      Height          =   255
      Index           =   0
      Left            =   150
      TabIndex        =   0
      Tag             =   "CODVEND:"
      Top             =   75
      Width           =   975
   End
End
Attribute VB_Name = "frmvend"
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

  With rst
      .MoveLast
      nItem = .RecordCount
      .MoveFirst
      
      vaSpr1.MaxRows = 0: vaSpr1.MaxRows = nItem
      
      nItem = 0
      
      While Not .EOF
         nItem = nItem + 1
        
         vaSpr1.Row = nItem
         vaSpr1.Col = 1: vaSpr1.Text = fuNulo(!codvend, "")
         vaSpr1.Col = 2: vaSpr1.Text = fuNulo(!nome, "")
         vaSpr1.Col = 3: vaSpr1.Text = fuNulo(!Salario, "")
         vaSpr1.Col = 4: vaSpr1.Text = fuNulo(!comissao, "")
         vaSpr1.Col = 5: vaSpr1.Text = fuNulo(!senha, "")
         vaSpr1.Col = 6: vaSpr1.Text = fuNulo(!Nivel, "")
         .MoveNext
     Wend
     vaSpr1.MaxRows = nItem
         
  End With
End Sub

Private Sub Carrega_tela()
   'Limpa as variaveis da tela se caso ficarem com dados da outra tela
   limpa_tela Me
   'Carrega a tela com os dados do registro
   Me.LblCodvend = rst("codvend")
   Me.TxtNome.Text = rst("nome")
   Me.TxtFixo.Text = Format(rst("salario"), "###,##0.00")
   Me.TxtComissao.Text = Format(rst("comissao"), "#0.00")
   Me.TxtSenha.Text = rst("senha")
   Me.txtNivel.Text = rst("nivel")
   
End Sub
Private Sub cmdAdd_Click()
   
   lIncluir = True
   limpa_tela Me
   
   frmvend.TxtFixo.Enabled = True
   frmvend.TxtFixo.Text = ""
   frmvend.LblCodvend.Caption = ""
   frmvend.TxtNome.SetFocus
   frmvend.cmdUpdate.Enabled = True
   frmvend.cmddesfaz.Enabled = True
   frmvend.cmdEditar.Enabled = False
   frmvend.cmdAdd.Enabled = False
   frmvend.CmdSair.Enabled = False
   frmvend.cmdDelete.Enabled = False
   
End Sub


Private Sub cmdDelete_Click()
    'this may produce an error if you delete the last
    'record or the only record in the recordset
    If MsgBox("Deseja realmente apagar este balconista ? ", vbYesNo, "Atenção") = vbYes Then
        rst.Close
        cSql = "DELETE FROM cadvend WHERE cadvend.codvend = " _
                & Me.LblCodvend.Caption & " AND cadvend.Nome = '" & Me.TxtNome.Text & "'"
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
     MsgBox "Deu erro na exclusao do Balconista" & Chr(13) & "Instrucao Sql = '" & _
            cSql & "'  "
End Sub


Private Sub cmddesfaz_Click()
  lIncluir = False
  Desabilita Me
  Me.TxtFixo.Enabled = False
   
  Me.cmdUpdate.Enabled = False
  Me.cmddesfaz.Enabled = False
  Me.cmdEditar.Enabled = True
  Me.cmdAdd.Enabled = True
  Me.CmdSair.Enabled = True
  Me.cmdDelete.Enabled = True
 
End Sub

Private Sub cmdEditar_Click()
   Carrega_tela
   
   Me.TxtFixo.Enabled = True
   
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
      rst.Close
      cSql = "INSERT INTO cadvend (Nome,Salario,Comissao,Senha,Nivel,operador,datatual) " & _
                          "VALUES ( '" & Me.TxtNome.Text & "','" & _
                                         Me.TxtComissao.Text & "','" & _
                                         Me.TxtSenha.Text & "','" & _
                                         Me.txtNivel.Text & "','" & _
                                         gOperador & "','" & _
                                         Now & "')"
      cnn.Execute cSql
                          
      lIncluir = False
   Else
      rst.Close
      cSql = "UPDATE cadvend SET nome = '" & Me.TxtNome.Text & "'," & _
                                " Salario = '" & Me.TxtFixo.Text & "'," & _
                                " Comissao = '" & Me.TxtComissao.Text & "'," & _
                                " Senha = '" & Me.TxtSenha.Text & "'," & _
                                " Nivel = '" & Me.Nivel.Text & "'," & _
                                " operador = '" & gOperador & "'," & _
                                " datatual = '" & Now & "'" & _
                                " WHERE cadvend.codvend = " & CLng(Me.LblCodvend.Caption)
      cnn.Execute cSql
   
      lPrimeiro = False
   End If
     
   Abre_Le_rst
   
   Carrega_tela
   
   Desabilita Me
   
   Me.cmdUpdate.Enabled = False
   Me.cmddesfaz.Enabled = False
   Me.cmdEditar.Enabled = True
   Me.cmdAdd.Enabled = True
   frmvend.CmdSair.Enabled = True
   Me.cmdDelete.Enabled = True
  
   Carrega_Grid
   
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
     
End Sub

Private Sub Form_Load()

   'Abre e le o recordset
   Abre_Le_rst
   
   'Centraliza a tela no video
   frmvend.Move (Screen.Width - frmvend.Width) / 2, _
           (Screen.Height - frmvend.Height) / 2
   
   limpa_tela Me
   
   frmvend.LblCodvend.Caption = ""
   
     If rst.BOF And rst.EOF Then
      If MsgBox("Arquivo vazio. Incluir dados agora ?", vbYesNo, "Atenção ") = vbYes Then
         'rst.AddNew
         With rst
           .AddNew
           !nome = ""
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
   cSql = "select * from cadvend"
   rst.Open cSql, cnn, adOpenKeyset, adLockOptimistic, adCmdText
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Screen.MousePointer = vbDefault
    rst.Close
    
End Sub


