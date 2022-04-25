VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3855
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7980
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   7980
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Frame fraMainFrame 
      Height          =   3780
      Left            =   45
      TabIndex        =   0
      Top             =   -15
      Width           =   7860
      Begin VB.PictureBox picLogo 
         AutoSize        =   -1  'True
         Height          =   540
         Left            =   120
         Picture         =   "frmSplash.frx":0000
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   1
         Top             =   240
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Sistema Licenciado para:"
         Height          =   195
         Left            =   1755
         TabIndex        =   9
         Top             =   240
         Width           =   1785
      End
      Begin VB.Label lblproduto 
         AutoSize        =   -1  'True
         Caption         =   "Sistema de Gerenciamento Comercial"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1755
         TabIndex        =   8
         Tag             =   "Product"
         Top             =   1200
         Width           =   5220
      End
      Begin VB.Label lblCompanyProduct 
         AutoSize        =   -1  'True
         Caption         =   "Vibe Informatica"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1710
         TabIndex        =   7
         Tag             =   "CompanyProduct"
         Top             =   480
         Width           =   4005
      End
      Begin VB.Label lblPlatform 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Windows 95/98/NT/2000/XP"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1755
         TabIndex        =   6
         Tag             =   "Platform"
         Top             =   2280
         Width           =   4095
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Version 1.0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1755
         TabIndex        =   5
         Tag             =   "Version"
         Top             =   1920
         Width           =   1275
      End
      Begin VB.Label lblWarning 
         Caption         =   "Warning"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   2
         Tag             =   "Warning"
         Top             =   3240
         Visible         =   0   'False
         Width           =   6855
      End
      Begin VB.Label lblCompany 
         AutoSize        =   -1  'True
         Caption         =   "Info Sistemas Ltda. - Fone: (13) 3289-5374 "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   1755
         TabIndex        =   4
         Tag             =   "Company"
         Top             =   2895
         Width           =   3150
      End
      Begin VB.Label lblCopyright 
         Caption         =   "Copyright"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1755
         TabIndex        =   3
         Tag             =   "Copyright"
         Top             =   2640
         Width           =   2415
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
'    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
'    lblProductName.Caption = App.Title
lblCompanyProduct.Caption = "Mercadão do Gesso" 'gCCliente
lblproduto.Caption = "S.M.G."    'gCSistema"
End Sub



