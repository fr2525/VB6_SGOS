VERSION 5.00
Begin VB.Form frmLoja 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "cadloja"
   ClientHeight    =   5835
   ClientLeft      =   1095
   ClientTop       =   330
   ClientWidth     =   6765
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5835
   ScaleWidth      =   6765
   Begin VB.CommandButton CmdSair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   615
      Left            =   3720
      MaskColor       =   &H00808080&
      Picture         =   "frmLoja.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   4920
      Width           =   735
   End
   Begin VB.CommandButton CmdAtualiza 
      Caption         =   "Atuali&zar"
      Default         =   -1  'True
      Height          =   615
      Left            =   2040
      MaskColor       =   &H00808080&
      Picture         =   "frmLoja.frx":00FA
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   4920
      Width           =   735
   End
   Begin VB.Frame Frame1 
      Caption         =   "Mensagens"
      Height          =   1215
      Left            =   1080
      TabIndex        =   26
      Top             =   3360
      Width           =   4335
      Begin VB.TextBox txtFields 
         DataField       =   "MENSAGEM2"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   14
         Left            =   480
         TabIndex        =   28
         Top             =   720
         Width           =   3375
      End
      Begin VB.TextBox txtFields 
         DataField       =   "MENSAGEM1"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   12
         Left            =   480
         TabIndex        =   27
         Top             =   360
         Width           =   3375
      End
   End
   Begin VB.TextBox txtFields 
      DataField       =   "DIVCUPOM"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   13
      Left            =   4440
      TabIndex        =   24
      Top             =   2880
      Width           =   375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "SENHA"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   11
      Left            =   1440
      TabIndex        =   22
      Top             =   2880
      Width           =   375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "CELULAR"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   10
      Left            =   4440
      TabIndex        =   20
      Top             =   2520
      Width           =   1695
   End
   Begin VB.TextBox txtFields 
      DataField       =   "TELEFONE"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   9
      Left            =   1440
      TabIndex        =   18
      Top             =   2520
      Width           =   1575
   End
   Begin VB.TextBox txtFields 
      DataField       =   "INSC_EST"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   8
      Left            =   4440
      TabIndex        =   16
      Top             =   2160
      Width           =   1695
   End
   Begin VB.TextBox txtFields 
      DataField       =   "CGC"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   7
      Left            =   1440
      TabIndex        =   14
      Top             =   2160
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "CEP"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   6
      Left            =   4560
      TabIndex        =   12
      Top             =   1800
      Width           =   1575
   End
   Begin VB.TextBox txtFields 
      DataField       =   "ESTADO"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   5
      Left            =   1440
      TabIndex        =   10
      Top             =   1785
      Width           =   375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "CIDADE"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   4
      Left            =   1440
      TabIndex        =   8
      Top             =   1455
      Width           =   4695
   End
   Begin VB.TextBox txtFields 
      DataField       =   "BAIRRO"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   3
      Left            =   1440
      TabIndex        =   6
      Top             =   1140
      Width           =   4695
   End
   Begin VB.TextBox txtFields 
      DataField       =   "ENDERECO"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   2
      Left            =   1440
      TabIndex        =   4
      Top             =   840
      Width           =   4695
   End
   Begin VB.TextBox txtFields 
      DataField       =   "NOME"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   1
      Left            =   1440
      TabIndex        =   2
      Top             =   495
      Width           =   4695
   End
   Begin VB.Label LblLoja 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   1560
      TabIndex        =   25
      Top             =   120
      Width           =   495
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Cupom?:"
      Height          =   255
      Index           =   13
      Left            =   3120
      TabIndex        =   23
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Senha? (S/N):"
      Height          =   255
      Index           =   11
      Left            =   120
      TabIndex        =   21
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Celular:"
      Height          =   255
      Index           =   10
      Left            =   3120
      TabIndex        =   19
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Telefone:"
      Height          =   255
      Index           =   9
      Left            =   120
      TabIndex        =   17
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Insc.Est:"
      Height          =   255
      Index           =   8
      Left            =   3120
      TabIndex        =   15
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "CGC:"
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   13
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "CEP:"
      Height          =   255
      Index           =   6
      Left            =   3840
      TabIndex        =   11
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Estado:"
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   9
      Top             =   1785
      Width           =   1215
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Cidade:"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   7
      Top             =   1455
      Width           =   1215
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Bairro:"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   5
      Top             =   1140
      Width           =   1215
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Endereço:"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   3
      Top             =   825
      Width           =   1215
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Nome:"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label lblLabels 
      Caption         =   "Loja:"
      Height          =   255
      Index           =   0
      Left            =   960
      TabIndex        =   0
      Top             =   120
      Width           =   375
   End
End
Attribute VB_Name = "frmLoja"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()

End Sub

Private Sub CmdSair_Click()
  Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault
End Sub

Private Sub datPrimaryRS_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)
  'This is where you would put error handling code
  'If you want to ignore errors, comment out the next line
  'If you want to trap them, add code here to handle them
  MsgBox "Data error event hit err:" & Description
End Sub

Private Sub datPrimaryRS_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  'This will display the current record position for this recordset
  datPrimaryRS.Caption = "Record: " & CStr(datPrimaryRS.Recordset.AbsolutePosition)
End Sub

Private Sub datPrimaryRS_WillChangeRecord(ByVal adReason As ADODB.EventReasonEnum, ByVal cRecords As Long, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  'This is where you put validation code
  'This event gets called when the following actions occur
  Dim bCancel As Boolean

  Select Case adReason
  Case adRsnAddNew
  Case adRsnClose
  Case adRsnDelete
  Case adRsnFirstChange
  Case adRsnMove
  Case adRsnRequery
  Case adRsnResynch
  Case adRsnUndoAddNew
  Case adRsnUndoDelete
  Case adRsnUndoUpdate
  Case adRsnUpdate
  End Select

  If bCancel Then adStatus = adStatusCancel
End Sub

Private Sub cmdRefresh_Click()
  'This is only needed for multi user apps
  On Error GoTo RefreshErr
  datPrimaryRS.Refresh
  Exit Sub
RefreshErr:
  MsgBox Err.Description
End Sub

Private Sub cmdUpdate_Click()
  On Error GoTo UpdateErr

  datPrimaryRS.Recordset.UpdateBatch adAffectAll
  Exit Sub
UpdateErr:
  MsgBox Err.Description
End Sub

Private Sub cmdClose_Click()
  Unload Me
End Sub

