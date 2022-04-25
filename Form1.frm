VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{B1C46850-3E6A-11D2-8FEB-00104B9E07A7}#3.0#0"; "ssdw3ao.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3255
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6870
   LinkTopic       =   "Form1"
   ScaleHeight     =   3255
   ScaleWidth      =   6870
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   270
      Top             =   2865
      Visible         =   0   'False
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   2
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "FILE NAME=H:\Aplics\VB6.0\SMG\dbsmg.dsn"
      OLEDBString     =   ""
      OLEDBFile       =   "H:\Aplics\VB6.0\SMG\dbsmg.dsn"
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   "oyster"
      RecordSource    =   "select * from tab_movestoque"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin SSDataWidgets_A_OLEDB.SSOleDBCommand SSOleDBCommand1 
      Height          =   675
      Left            =   6090
      TabIndex        =   2
      Top             =   1830
      Width           =   765
      _Version        =   196612
      _ExtentX        =   1349
      _ExtentY        =   1191
      _StockProps     =   78
      Caption         =   "SSOleDBCommand1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid SSOleDBGrid1 
      Bindings        =   "Form1.frx":0000
      Height          =   2175
      Left            =   180
      TabIndex        =   1
      Top             =   420
      Width           =   5715
      _Version        =   196616
      AllowAddNew     =   -1  'True
      AllowDelete     =   -1  'True
      RowHeight       =   423
      Columns(0).Width=   3200
      Columns(0).DataType=   8
      Columns(0).FieldLen=   4096
      _ExtentX        =   10081
      _ExtentY        =   3836
      _StockProps     =   79
      Caption         =   "SSOleDBGrid1"
      BeginProperty PageFooterFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty PageHeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin SSDataWidgets_A_OLEDB.SSOleDBData SSOleDBData1 
      Bindings        =   "Form1.frx":0015
      Height          =   390
      Left            =   750
      TabIndex        =   0
      Top             =   30
      Width           =   4680
      _Version        =   196613
      _ExtentX        =   8255
      _ExtentY        =   688
      _StockProps     =   79
      Caption         =   "SSOleDBData1"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub SSOleDBCommand1_Click()
 Unload Me
End Sub
