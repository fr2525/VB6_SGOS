VERSION 5.00
Begin VB.Form frmcadvend 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Balconistas"
   ClientHeight    =   3090
   ClientLeft      =   4755
   ClientTop       =   4080
   ClientWidth     =   5955
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   5955
   ShowInTaskbar   =   0   'False
   Tag             =   "cadvend"
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
      Height          =   1455
      Left            =   720
      TabIndex        =   12
      Top             =   1440
      Width           =   4695
      Begin VB.CommandButton cmddesfaz 
         Caption         =   "&Desfaz"
         Enabled         =   0   'False
         Height          =   540
         Left            =   3105
         Picture         =   "frmcadvend.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   22
         Tag             =   "&Update"
         Top             =   720
         Width           =   615
      End
      Begin VB.CommandButton CmdSair 
         Caption         =   "&Sair"
         Height          =   540
         Left            =   3795
         Picture         =   "frmcadvend.frx":00FA
         Style           =   1  'Graphical
         TabIndex        =   21
         Tag             =   "&Update"
         Top             =   720
         Width           =   615
      End
      Begin VB.CommandButton CmdUltimo 
         Height          =   300
         Left            =   2880
         Picture         =   "frmcadvend.frx":01F4
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton CmdProximo 
         Height          =   300
         Left            =   2520
         Picture         =   "frmcadvend.frx":02EE
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmdAnterior 
         Height          =   300
         Left            =   2160
         Picture         =   "frmcadvend.frx":03E8
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmdPrimeiro 
         Height          =   300
         Left            =   1800
         Picture         =   "frmcadvend.frx":04E2
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Incluir"
         Height          =   540
         Left            =   345
         Picture         =   "frmcadvend.frx":05DC
         Style           =   1  'Graphical
         TabIndex        =   16
         Tag             =   "&Add"
         Top             =   720
         Width           =   615
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Excluir"
         Height          =   540
         Left            =   1725
         Picture         =   "frmcadvend.frx":06C6
         Style           =   1  'Graphical
         TabIndex        =   15
         Tag             =   "&Delete"
         Top             =   720
         Width           =   615
      End
      Begin VB.CommandButton cmdEditar 
         Caption         =   "&Editar"
         Height          =   540
         Left            =   1035
         Picture         =   "frmcadvend.frx":0838
         Style           =   1  'Graphical
         TabIndex        =   14
         Tag             =   "&Refresh"
         Top             =   720
         Width           =   615
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Salvar"
         Enabled         =   0   'False
         Height          =   540
         Left            =   2415
         Picture         =   "frmcadvend.frx":09AA
         Style           =   1  'Graphical
         TabIndex        =   13
         Tag             =   "&Update"
         Top             =   720
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
Attribute VB_Name = "frmcadvend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private pRsVend As Recordset
Private Sub Desabilita()
'Deixa os textbox desabilitados
   Dim i
   For i = 1 To Me.Controls.Count - 1
       If TypeOf Me.Controls(i) Is TextBox Then
          Me.Controls(i).Enabled = False
       
       End If
       If TypeOf Me.Controls(i) Is MaskEdBox Then
          Me.Controls(i).Enabled = False
         
       End If
   Next i
   Me.TxtFixo.Enabled = False
      
End Sub

Private Sub limpa_tela()
 Dim i
   For i = 1 To Me.Controls.Count - 1
       If TypeOf Me.Controls(i) Is TextBox Then
          Me.Controls(i).Enabled = True
          Me.Controls(i).Text = ""
       End If
       If TypeOf Me.Controls(i) Is MaskEdBox Then
          Me.Controls(i).Enabled = True
          Me.Controls(i).Text = ""
       End If
   Next i
     Me.TxtFixo.Enabled = True
End Sub
Private Sub Carrega_tela()
   'Limpa as variaveis da tela se caso ficarem com dados da outra tela
   limpa_tela
   'Carrega a tela com os dados do registro
   frmcadvend.LblCodvend = pRsVend("codvend")
   frmcadvend.TxtNome.Text = pRsVend("nome")
   frmcadvend.TxtFixo.Text = Format(pRsVend("salario"), "###,##0.00")
   frmcadvend.TxtComissao.Text = Format(pRsVend("comissao"), "#0.00")
   frmcadvend.TxtSenha.Text = pRsVend("senha")
   frmcadvend.txtNivel.Text = pRsVend("nivel")
   
End Sub
Private Sub cmdAdd_Click()
   
   limpa_tela
   
   frmcadvend.TxtFixo.Enabled = True
   frmcadvend.TxtFixo.Text = ""
   frmcadvend.LblCodvend.Caption = ""
   frmcadvend.TxtNome.SetFocus
   
   frmcadvend.cmdUpdate.Enabled = True
   frmcadvend.cmddesfaz.Enabled = True
   frmcadvend.cmdEditar.Enabled = False
   frmcadvend.cmdAdd.Enabled = False
   frmcadvend.CmdSair.Enabled = False
   frmcadvend.cmdDelete.Enabled = False
   frmcadvend.cmdAnterior.Enabled = False
   frmcadvend.cmdPrimeiro.Enabled = False
   frmcadvend.CmdProximo.Enabled = False
   frmcadvend.CmdUltimo.Enabled = False
  
End Sub


Private Sub cmdAnterior_Click()
  pRsVend.MovePrevious
  If pRsVend.BOF Then
       MsgBox "Primeiro registro", vbOKOnly, "Atenção"
    Else
       Carrega_tela
       Desabilita
    End If
End Sub

Private Sub cmdDelete_Click()
    'this may produce an error if you delete the last
    'record or the only record in the recordset
    If MsgBox("Deseja realmente apagar este balconista ? ", vbYesNo, "Atenção") = vbYes Then
        With pRsVend
            .Delete
            .MoveNext
            If .EOF Then .MoveLast
        End With
        Carrega_tela
        Desabilita
     End If
End Sub


Private Sub cmddesfaz_Click()
  
  Carrega_tela
  Desabilita
  Me.TxtFixo.Enabled = False
   
  Me.cmdUpdate.Enabled = False
  Me.cmddesfaz.Enabled = False
  Me.cmdEditar.Enabled = True
  Me.cmdAdd.Enabled = True
  Me.CmdSair.Enabled = True
  Me.cmdDelete.Enabled = True
  Me.cmdAnterior.Enabled = True
  Me.cmdPrimeiro.Enabled = True
  Me.CmdProximo.Enabled = True
  Me.CmdUltimo.Enabled = True

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
   Me.cmdAnterior.Enabled = False
   Me.cmdPrimeiro.Enabled = False
   Me.CmdProximo.Enabled = False
   Me.CmdUltimo.Enabled = False

End Sub

Private Sub cmdPrimeiro_Click()
    pRsVend.MoveFirst
    Carrega_tela
    Desabilita
End Sub

Private Sub CmdProximo_Click()
   pRsVend.MoveNext
   If pRsVend.EOF Then
       MsgBox "Ultimo registro", vbOKOnly, "Atenção"
   Else
      Carrega_tela
      Desabilita
   End If
End Sub

Private Sub CmdSair_Click()
   Unload Me
End Sub

Private Sub CmdUltimo_Click()
   pRsVend.MoveLast
   Carrega_tela
   Desabilita
End Sub

Private Sub cmdUpdate_Click()
   pRsVend.AddNew
   'pRsVend("codvend") = frmcadvend.LblCodvend
   pRsVend("nome") = frmcadvend.TxtNome.Text
   pRsVend("salario") = Val(frmcadvend.TxtFixo.Text)
   pRsVend("comissao") = Val(frmcadvend.TxtComissao.Text)
   pRsVend("senha") = frmcadvend.TxtSenha.Text
   pRsVend("nivel") = frmcadvend.txtNivel.Text
   pRsVend.Update
     
   'Deixa os textbox desabilitados
   frmcadvend.TxtFixo.Enabled = False
   frmcadvend.TxtFixo.Text = ""
   pRsVend.MoveLast
   
   Carrega_tela
   
   Desabilita
   
   frmcadvend.cmdUpdate.Enabled = False
   frmcadvend.cmddesfaz.Enabled = False
   frmcadvend.cmdEditar.Enabled = True
   frmcadvend.cmdAdd.Enabled = True
   frmcadvend.CmdSair.Enabled = True
   frmcadvend.cmdDelete.Enabled = True
   frmcadvend.cmdAnterior.Enabled = True
   frmcadvend.cmdPrimeiro.Enabled = True
   frmcadvend.CmdProximo.Enabled = True
   frmcadvend.CmdUltimo.Enabled = True
     
End Sub



Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
     
End Sub

Private Sub Form_Load()
   'Centraliza a tela no video
   frmcadvend.Move (Screen.Width - frmcadvend.Width) / 2, _
           (Screen.Height - frmcadvend.Height) / 2
   
   limpa_tela
   
   frmcadvend.LblCodvend.Caption = ""
   gSql = "select * from cadvend"
   Set pRsVend = gDb.OpenRecordset(gSql)
   pRsVend.MoveFirst
   
   Carrega_tela
   
   Desabilita
   
   'frmcadvend.TxtFixo.Enabled = False
   'frmcadvend.TxtFixo.Text = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Screen.MousePointer = vbDefault
End Sub



