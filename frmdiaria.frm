VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmdiaria 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cotação"
   ClientHeight    =   3270
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3270
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtData 
      Height          =   375
      Left            =   1320
      TabIndex        =   0
      Text            =   "  /  /  "
      Top             =   480
      Width           =   1335
   End
   Begin MSMask.MaskEdBox MskValor 
      Height          =   375
      Left            =   1320
      TabIndex        =   2
      Top             =   1440
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   13
      Mask            =   "##,###,###.##"
      PromptChar      =   "_"
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Height          =   495
      Left            =   2400
      Picture         =   "frmdiaria.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Cancelar"
      Top             =   2400
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Height          =   495
      Left            =   1560
      Picture         =   "frmdiaria.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Gravar"
      Top             =   2400
      Width           =   615
   End
   Begin VB.ComboBox CboMoedas 
      Height          =   315
      Left            =   1320
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   960
      Width           =   2895
   End
   Begin VB.Label Label3 
      Caption         =   "Moeda"
      Height          =   255
      Left            =   360
      TabIndex        =   7
      Top             =   960
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "Cotação"
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Data"
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   480
      Width           =   735
   End
End
Attribute VB_Name = "frmdiaria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private pRsMoedas As Recordset
Private pRsCotacao As Recordset
Private gSql As String

Public OK As Boolean

Private Sub Command1_Click()
  OK = True
  Me.Hide
     
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub Form_Load()
   'Centraliza a tela no video
   Me.Move (Screen.Width - Me.Width) / 2, _
           (Screen.Height - Me.Height) / 2
 
   gSql = "Select nome,codigo from cadmoe"
   Set pRsMoedas = gDb.OpenRecordset(gSql)
   With CboMoedas
       .Clear
       Do While Not pRsMoedas.EOF
          .AddItem pRsMoedas!nome
          .ItemData(.NewIndex) = pRsMoedas!codigo
          pRsMoedas.MoveNext
       Loop
   End With

OK = False

End Sub


Private Sub TxtData_GotFocus()
    TxtData.SelStart = 0
    TxtData.SelLength = Len(TxtData.Text)
End Sub

Private Sub TxtData_LostFocus()
    If Not ValidaData(TxtData.Text) Then
       If MsgBox(" Data Inválida. " + Chr(13) + " Tentar Novamente ?", vbYesNo, "Atenção") = vbYes Then
          TxtData.Text = " "
          TxtData.SetFocus
       Else
          End
       End If
    End If
    
       
       
       
       
       
       
       
       
       
       
    
End Sub
