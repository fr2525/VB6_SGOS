VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form FrmDatas1 
   Caption         =   "Entre com as datas"
   ClientHeight    =   2385
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3600
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   2385
   ScaleWidth      =   3600
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdSair 
      Height          =   570
      Left            =   2550
      Picture         =   "FrmDatas1.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Sai "
      Top             =   1530
      Width           =   675
   End
   Begin VB.CommandButton cmdimprimir 
      Height          =   570
      Left            =   1590
      Picture         =   "FrmDatas1.frx":00FA
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Imprime o orçamento"
      Top             =   1545
      Width           =   675
   End
   Begin MSMask.MaskEdBox MskAte 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd/MM/yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   3
      EndProperty
      Height          =   390
      Left            =   1605
      TabIndex        =   3
      Top             =   810
      Width           =   1650
      _ExtentX        =   2910
      _ExtentY        =   688
      _Version        =   393216
      MaxLength       =   10
      Mask            =   "99/99/9999"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MskDe 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd/MM/yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   3
      EndProperty
      Height          =   390
      Left            =   1605
      TabIndex        =   2
      Top             =   285
      Width           =   1650
      _ExtentX        =   2910
      _ExtentY        =   688
      _Version        =   393216
      MaxLength       =   10
      Mask            =   "99/99/9999"
      PromptChar      =   "_"
   End
   Begin VB.Label Label2 
      Caption         =   "Data Final:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   270
      Left            =   345
      TabIndex        =   1
      Top             =   885
      Width           =   1620
   End
   Begin VB.Label Label1 
      Caption         =   "Data Inicial:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   255
      TabIndex        =   0
      Top             =   330
      Width           =   1605
   End
End
Attribute VB_Name = "FrmDatas1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private myRel As Variant
Private mySel As Variant
Private pNomerel As String
Private pcSelecao As String

Public Property Get PassRel() As Variant
   PassRel = myRel
End Property

Public Property Let PassRel(ByVal vNovoRel As Variant)
   myRel = vNovoRel
End Property

Public Property Get PassSel() As Variant
   PassSel = mySel
End Property

Public Property Let PassSel(ByVal vNovoSel As Variant)
   mySel = vNovoSel
End Property

Private Sub cmdimprimir_Click()
Dim Diai$, Mesi$, Anoi$
Dim Diaf$, Mesf$, Anof$

Dim cat As New ADOX.Catalog
Dim cmd As New ADODB.Command


Diai = Str(Day(MskDe.Text)): Mesi = Str(Month(MskDe.Text)): Anoi = Str(Year(MskDe.Text))
Diaf = Str(Day(MskAte.Text)): Mesf = Str(Month(MskAte.Text)): Anof = Str(Year(MskAte.Text))
   
'On Error GoTo trata_erro
   
'abre o catalogo
 cat.ActiveConnection = "Provider=Microsoft.Jet.OLEDB.4.0;" _
& "Data Source=" & ConDb.DefaultDatabase & ".mdb"

'agora vamos criar a consulta

gSql = "SELECT A.codprod, (select descricao from tab_produtos where a.codprod = tab_Produtos.codprod) AS descricao, Sum(A.QTDEp) AS qtdep, Sum(A.QtdeA) AS qtdea"
gSql = gSql & " FROM (Tab_ItemVenda AS A INNER JOIN Tab_produtos AS B ON A.codprod = B.codprod) INNER JOIN Tab_Vendas ON A.nsu = Tab_Vendas.nsu"
gSql = gSql & " WHERE Tab_Vendas.DTA_VENDA >= CDate(#"
gSql = gSql & Format(Mesi, "00") & "/" & Format(Diai, "00") & "/" & Format(Anoi, "0000") & "#)"
gSql = gSql & "AND Tab_Vendas.DTA_VENDA <= CDate(#"
gSql = gSql & Format(Mesf, "00") & "/" & Format(Diaf, "00") & "/" & Format(Anof, "0000") & "#)"
gSql = gSql & " GROUP BY A.codprod"
gSql = gSql & " ORDER BY A.codprod"

cmd.CommandText = gSql

cat.Views.Append "Vendidos", cmd
   
'On Error Resume Next

   FrmCompras.CrRelcomp.DataFiles(0) = ConDb.DefaultDatabase & ".mdb"

   FrmCompras.CrRelcomp.Formulas(0) = "intervalo = " & "'" _
                      & " Vendas no Periodo de " & MskDe.Text _
                      & " a " & MskAte.Text & "'"
     
   FrmCompras.CrRelcomp.Destination = 0 'Vídeo
   FrmCompras.CrRelcomp.WindowState = crptMaximized
   FrmCompras.CrRelcomp.WindowTitle = "Visualização das Vendas do Período"
   FrmCompras.CrRelcomp.SelectionFormula = ""
   FrmCompras.CrRelcomp.Formulas(1) = "nomeloja = '" & gNome & "'"
   FrmCompras.CrRelcomp.SortFields(0) = "-{Vendidos.qtdep}"
   FrmCompras.CrRelcomp.ReportFileName = gPathRel & PassRel
   FrmCompras.CrRelcomp.Action = 1
     
cat.Views.Delete ("Vendidos")
     
Set cat = Nothing

Exit Sub

trata_erro:
MsgBox Err.Description, vbCritical, "Consulta já existe..."
     
End Sub

Private Sub CmdSair_Click()
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyReturn Then
      SendKeys "{TAB}"
      KeyAscii = 0
   End If
   If KeyCode = vbKeyEscape Then Unload Me

End Sub

Private Sub Form_Load()
    Centra Me
    pNomerel = myRel
    pSelecao = mySel
End Sub

Private Sub MskAte_Validate(Cancel As Boolean)
   If Not IsDate(MskAte) Then
       MsgBox "Favor entrar com uma data válida", vbOKOnly, "Atenção Operador"
       Cancel = True
    End If
    If MskAte < MskDe Then
       MsgBox "Favor entrar com uma data válida", vbOKOnly, "Atenção Operador"
       Cancel = True
    End If
      
End Sub

Private Sub MskDe_Validate(Cancel As Boolean)
    If Not IsDate(MskDe) Then
       MsgBox "Favor entrar com uma data válida", vbOKOnly, "Atenção Operador"
       Cancel = True
    End If
    
End Sub
