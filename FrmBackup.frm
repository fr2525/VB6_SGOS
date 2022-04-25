VERSION 5.00
Begin VB.Form FrmBackup 
   Caption         =   "Copia de Segurança"
   ClientHeight    =   3360
   ClientLeft      =   2790
   ClientTop       =   2070
   ClientWidth     =   3855
   LinkTopic       =   "Form1"
   ScaleHeight     =   3360
   ScaleWidth      =   3855
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   2100
      TabIndex        =   3
      Top             =   525
      Width           =   1470
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Iniciar"
      Height          =   540
      Left            =   2205
      Picture         =   "FrmBackup.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Tag             =   "&Update"
      Top             =   2160
      Width           =   615
   End
   Begin VB.CommandButton CmdSair 
      Caption         =   "&Sair"
      Height          =   540
      Left            =   2865
      Picture         =   "FrmBackup.frx":00FA
      Style           =   1  'Graphical
      TabIndex        =   1
      Tag             =   "&Update"
      Top             =   2160
      Width           =   615
   End
   Begin VB.Image Image1 
      Height          =   3225
      Left            =   45
      Picture         =   "FrmBackup.frx":01F4
      Top             =   45
      Width           =   1725
   End
   Begin VB.Label lblDestino 
      Caption         =   "Unidade de Destino:"
      Height          =   315
      Left            =   2070
      TabIndex        =   0
      Top             =   180
      Width           =   1680
   End
End
Attribute VB_Name = "FrmBackup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdSair_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   'Centraliza a tela no video
   Me.Move (Screen.Width - Me.Width) / 2, _
           (Screen.Height - Me.Height) / 2

End Sub
