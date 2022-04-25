VERSION 5.00
Begin VB.Form FrmProcessa 
   Caption         =   "Form1"
   ClientHeight    =   3075
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7320
   LinkTopic       =   "Form1"
   ScaleHeight     =   3075
   ScaleWidth      =   7320
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Sair"
      Height          =   705
      Left            =   5850
      TabIndex        =   5
      Top             =   2115
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Processar"
      Height          =   705
      Left            =   4500
      TabIndex        =   4
      Top             =   2115
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Height          =   345
      Left            =   1890
      TabIndex        =   3
      Text            =   "Text2"
      Top             =   1425
      Width           =   5085
   End
   Begin VB.TextBox Text1 
      Height          =   345
      Left            =   1890
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   915
      Width           =   5085
   End
   Begin VB.Label Label3 
      Caption         =   "Adequação do Banco de dados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   360
      TabIndex        =   6
      Top             =   285
      Width           =   6585
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   315
      Left            =   180
      TabIndex        =   2
      Top             =   1395
      Width           =   1485
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   345
      Left            =   180
      TabIndex        =   0
      Top             =   915
      Width           =   1515
   End
End
Attribute VB_Name = "FrmProcessa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
