VERSION 5.00
Begin VB.Form FrmMedicao 
   Caption         =   "Medição"
   ClientHeight    =   2460
   ClientLeft      =   6315
   ClientTop       =   5250
   ClientWidth     =   3735
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   2460
   ScaleWidth      =   3735
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   705
      TabIndex        =   4
      Top             =   1380
      Width           =   2235
      Begin VB.CommandButton BtSalvar 
         Caption         =   "&Salvar"
         Height          =   540
         Left            =   135
         Picture         =   "FrmMedicao.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   7
         Tag             =   "&Update"
         Top             =   210
         Width           =   615
      End
      Begin VB.CommandButton Btsair 
         Caption         =   "&Sair"
         Height          =   540
         Left            =   1470
         Picture         =   "FrmMedicao.frx":00FA
         Style           =   1  'Graphical
         TabIndex        =   6
         Tag             =   "&Update"
         Top             =   210
         Width           =   615
      End
      Begin VB.CommandButton Btcorrige 
         Caption         =   "&Corrige"
         Height          =   540
         Left            =   795
         Picture         =   "FrmMedicao.frx":01F4
         Style           =   1  'Graphical
         TabIndex        =   5
         Tag             =   "&Update"
         Top             =   210
         Width           =   615
      End
   End
   Begin VB.TextBox TxtDta_medicao 
      Height          =   315
      Left            =   1905
      TabIndex        =   3
      Top             =   735
      Width           =   1305
   End
   Begin VB.TextBox TxtMedicao 
      Height          =   360
      Left            =   1890
      TabIndex        =   1
      Top             =   165
      Width           =   1320
   End
   Begin VB.Label Label2 
      Caption         =   "Data da Medição"
      Height          =   300
      Left            =   240
      TabIndex        =   2
      Top             =   780
      Width           =   1515
   End
   Begin VB.Label Label1 
      Caption         =   "Numero da Medição"
      Height          =   390
      Left            =   180
      TabIndex        =   0
      Top             =   240
      Width           =   1575
   End
End
Attribute VB_Name = "FrmMedicao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Btsair_Click()
   Unload Me
End Sub

Private Sub BtSalvar_Click()
   gSql = "INSERT INTO tab_med (cod_obra,medicao,dta_medicao, operador, datatual) "
   gSql = gSql & " VALUES( " & FrmPlanilha.CmbObra.ItemData(FrmPlanilha.CmbObra.ListIndex) & ",'" & TxtMedicao.Text & "',cdate('" & TxtDta_medicao.Text & "'),'" & gOperador & "',Cdate('" & Date & "') )"
   ConDb.Execute gSql
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
     
End Sub

Private Sub TxtDta_medicao_Validate(Cancel As Boolean)
If Not IsDate(TxtDta_medicao.Text) Then
   MsgBox " Data Invalida !! ", vbCritical, " Erro na Data "
   Cancel = True
End If

End Sub
