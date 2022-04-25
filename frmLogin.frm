VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   2250
   ClientLeft      =   2670
   ClientTop       =   2970
   ClientWidth     =   3810
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2250
   ScaleWidth      =   3810
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "Login"
   Begin VB.TextBox TxtCotacao 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   315
      Left            =   1320
      MaxLength       =   16
      TabIndex        =   3
      Text            =   "0,00"
      Top             =   870
      Visible         =   0   'False
      Width           =   1755
   End
   Begin VB.ComboBox CboUsuario 
      Height          =   315
      Left            =   1320
      TabIndex        =   1
      ToolTipText     =   "Escolha o usuario do sistema"
      Top             =   120
      Width           =   1815
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancela"
      Height          =   360
      Left            =   2175
      TabIndex        =   5
      Tag             =   "Cancel"
      Top             =   1590
      Width           =   1140
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   360
      Left            =   570
      TabIndex        =   4
      Tag             =   "OK"
      Top             =   1560
      Width           =   1140
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1320
      MaxLength       =   8
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   525
      Width           =   1785
   End
   Begin VB.Label LblTroco 
      Caption         =   "&Troco:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Tag             =   "&Password:"
      Top             =   930
      Visible         =   0   'False
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Senha:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   248
      Index           =   1
      Left            =   105
      TabIndex        =   0
      Tag             =   "&Password:"
      Top             =   540
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Usuário:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   248
      Index           =   0
      Left            =   105
      TabIndex        =   6
      Tag             =   "&User Name:"
      Top             =   150
      Width           =   1080
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpbuffer As String, nSize As Long) As Long

Public OK As Boolean
Private plTem_troco As Boolean

Private Sub CboUsuario_Click()
  '  If CboUsuario.ListIndex = -1 Then
  '     txtPassword = ""
  '     Exit Sub
  '  End If
  '  gRs.FindFirst "codvend = " & CboUsuario.ItemData(CboUsuario.ListIndex)
  '  If gRs.NoMatch Then
  '     Exit Sub
  '  End If
  '  txtPassword.Enabled = True
  '  txtPassword.SetFocus
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
 'Esse código permite a mudança de quadro de
  'texto através do Enter
  If KeyAscii = 13 Then
    'Se o tipo do controle ativo for TextBox
    If TypeOf Screen.ActiveControl Is TextBox Then
      'Simula o pressionamento da tecla TAB
      Sendkeys "{tab}"
      'A linha a seguir evita ouvir um bip
      KeyAscii = 0
    End If
  End If
End Sub

Private Sub Form_Load()
    plTem_troco = False

    gSql = "SELECT * from tab_operador ORDER BY nome"
         
    gRs.Open gSql, ConDb, adOpenKeyset
          
    With CboUsuario
         .Clear
         .AddItem "Master"
         Do While Not gRs.EOF
            .AddItem gRs("nome")
            .ItemData(.NewIndex) = gRs("id")
            gRs.MoveNext
         Loop
    End With
    gRs.Close
    'gSql = "select * from tab_operador"
    'Set gRs = ConDb.Execute(gSql)
    
'     gSql = "SELECT Max(hoje) as dia_de_hoje from caixa"
          
'     gRs.Open gSql, ConDb, adOpenKeyset
     
'     If gRs!dia_de_hoje <> Date Then
'        Me.LblTroco.Visible = True
'        Me.TxtCotacao.Visible = True
'        Me.TxtCotacao.Enabled = True
'        plTem_troco = True
'     End If
'     gRs.Close
     
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdCancel_Click()
    OK = False
    Me.Hide
End Sub

Private Sub CmdOk_Click()
    'To Do - create test for correct password
    'check for correct password
    If txtPassword.text = "oyster" Then
       'operador já está setado -> gOperador = "Master"
       OK = True
       Me.Hide
'       If plTem_troco Then
'          gSql = "INSERT into caixa (hoje,troco) VALUES ( CDate('" & Date & "'), "
'          gSql = gSql & Replace(IIf(Len(Me.TxtCotacao.Text) = 0, 0, Me.TxtCotacao.Text), ",", ".") & " ) "
'         ConDb.Execute gSql
'       End If
       Exit Sub
   End If

   If CboUsuario.ListIndex = -1 Then
      MsgBox "Informe o Usuário", vbCritical, gRs!nome
      CboUsuario.SetFocus
      Exit Sub
   End If
   'guSUARIO = CboUsuario.ItemData(CboUsuario.ListIndex)
   gSql = "select nome,id,senha,nivel from tab_operador"
   gSql = gSql & " WHERE nome = '" & CboUsuario.text & "'"
   gSql = gSql & " AND id = " & CboUsuario.ItemData(CboUsuario.ListIndex)
    
   gRs.Open gSql, ConDb, adOpenKeyset
   
   If gRs.BOF And gRs.EOF Then
      MsgBox "Operador Não encontrado...", vbCritical, CboUsuario.text
      CboUsuario.SetFocus
      gRs.Close
      Exit Sub
   End If
   If Not UCase$(gRs!senha) = fuEncript(UCase$(Trim(txtPassword.text)), "oyster") Then
      MsgBox "Senha Inválida...", vbCritical, gRs("nome")
      txtPassword.SetFocus
      gRs.Close
      Exit Sub
   End If
   gOperador = gRs!nome
   gnCodOperador = gRs!id
   gNivel = gRs!nivel
   gRs.Close
   OK = True
   Me.Hide
'   If plTem_troco Then
'       gSql = "INSERT into caixa (hoje,troco) VALUES ( CDate('" & Date & "'), "
'      gSql = gSql & Replace(IIf(Len(Me.TxtCotacao.Text) = 0, 0, Me.TxtCotacao.Text), ",", ".") & " ) "
'      ConDb.Execute gSql
'   End If
 End Sub

