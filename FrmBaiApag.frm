VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form FrmBaiApag 
   Caption         =   "Baixa de Contas a pagar"
   ClientHeight    =   5595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7260
   LinkTopic       =   "Form1"
   ScaleHeight     =   5595
   ScaleWidth      =   7260
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdSair 
      Caption         =   "&Sair"
      Height          =   540
      Left            =   5730
      Picture         =   "FrmBaiApag.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   18
      Tag             =   "&Update"
      Top             =   4740
      Width           =   615
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Ok"
      Height          =   540
      Left            =   5070
      Picture         =   "FrmBaiApag.frx":00FA
      Style           =   1  'Graphical
      TabIndex        =   17
      Tag             =   "&Update"
      Top             =   4740
      Width           =   615
   End
   Begin VB.Frame Frame1 
      Caption         =   "Dados do Pagto."
      Height          =   1290
      Left            =   4335
      TabIndex        =   12
      Top             =   3195
      Width           =   2700
      Begin VB.TextBox Text7 
         Height          =   300
         Left            =   1170
         TabIndex        =   16
         Text            =   "Text6"
         Top             =   705
         Width           =   1320
      End
      Begin VB.TextBox Text6 
         Height          =   300
         Left            =   1170
         TabIndex        =   15
         Text            =   "Text6"
         Top             =   270
         Width           =   1320
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Valor Pago:"
         Height          =   195
         Left            =   210
         TabIndex        =   14
         Top             =   735
         Width           =   825
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Data:"
         Height          =   195
         Left            =   630
         TabIndex        =   13
         Top             =   330
         Width           =   390
      End
   End
   Begin VB.TextBox Text5 
      Height          =   300
      Left            =   5460
      TabIndex        =   11
      Text            =   "Text1"
      Top             =   2640
      Width           =   1335
   End
   Begin VB.TextBox Text4 
      Height          =   300
      Left            =   5460
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   2250
      Width           =   1335
   End
   Begin VB.TextBox Text3 
      Height          =   300
      Left            =   5460
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   1845
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Height          =   300
      Left            =   5460
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   1440
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   300
      Left            =   5460
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   1020
      Width           =   1335
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   4005
      Left            =   210
      TabIndex        =   1
      Top             =   960
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   7064
      _Version        =   393216
      Cols            =   4
      FixedCols       =   0
      FormatString    =   "Duplicata  |  Emissão  |  Vencto.  |  Valor              "
   End
   Begin VB.ComboBox CboFornece 
      Height          =   315
      Left            =   210
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   315
      Width           =   3945
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Nota fiscal:"
      Height          =   195
      Left            =   4515
      TabIndex        =   10
      Top             =   2685
      Width           =   795
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Valor:"
      Height          =   195
      Left            =   4875
      TabIndex        =   5
      Top             =   2280
      Width           =   405
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Vencto:"
      Height          =   195
      Left            =   4725
      TabIndex        =   4
      Top             =   1890
      Width           =   555
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Emissão:"
      Height          =   195
      Left            =   4650
      TabIndex        =   3
      Top             =   1485
      Width           =   630
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Duplicata:"
      Height          =   195
      Left            =   4560
      TabIndex        =   2
      Top             =   1050
      Width           =   720
   End
End
Attribute VB_Name = "FrmBaiApag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private lIncluir As Boolean
Private lPrimeiro As Boolean
Private lbaixar As Boolean
Private prsAPagar As New ADODB.Recordset

Private Sub Form_Load()

 'Centraliza a tela no video
   Me.Move (Screen.Width - Me.Width) / 2, _
           (Screen.Height - Me.Height) / 2
   
   lbaixar = False

End Sub

