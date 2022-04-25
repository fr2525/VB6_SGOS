VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form Frmpesqprod 
   Caption         =   "Pesquisa de Produtos "
   ClientHeight    =   5910
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8010
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5910
   ScaleWidth      =   8010
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Txtescolhido 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   360
      Left            =   1530
      TabIndex        =   3
      Text            =   "123456"
      Top             =   5310
      Width           =   885
   End
   Begin VB.CommandButton CmdOk 
      Caption         =   "Ok"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   6900
      TabIndex        =   1
      Top             =   5235
      Width           =   735
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexprod 
      Height          =   4755
      Left            =   285
      TabIndex        =   0
      Top             =   330
      Width           =   7365
      _ExtentX        =   12991
      _ExtentY        =   8387
      _Version        =   393216
      Rows            =   3
      Cols            =   3
      FixedCols       =   0
      AllowBigSelection=   0   'False
      SelectionMode   =   1
      FormatString    =   "Referencia   |<Descrição                                   |<Fornecedor"
   End
   Begin VB.Label LblNomeProd 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   2655
      TabIndex        =   4
      Top             =   5355
      Width           =   720
   End
   Begin VB.Label LblEscolhido 
      Caption         =   "Escolhido"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   225
      TabIndex        =   2
      Top             =   5340
      Width           =   1200
   End
End
Attribute VB_Name = "Frmpesqprod"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private pRsProduto As New ADODB.Recordset
Public pcCodprod As String

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub CmdOk_Click()
   MSFlexprod.Col = 0
   If FrmVendas.Visible = True Then
      FrmVendas.Txtreferencia = MSFlexprod.Text
     'FrmVendas.TxtReferencia.SetFocus
   Else
      FrmEntradas.Txtreferencia = MSFlexprod.Text
  '    FrmEntradas.TxtReferencia.SetFocus
   End If
   'pRsProduto.Close
   Unload Me
End Sub



Private Sub Form_Load()

   gSql = "select codprod,descricao,nome from tab_produtos,tab_fornece where "
   gSql = gSql & "tab_produtos.ativo = 'S' and tab_produtos.codfor = tab_fornece.codfor ORDER by descricao"
   pRsProduto.Open gSql, ConDb, adOpenKeyset
    
   'Set MSFlexprod.SelectionMode = flexSelectionByRow
    
   With MSFlexprod
   

        .Clear
        .Cols = 3
        .Rows = 1
        .Row = 0
        .Col = 0
        .Text = "Referencia"
        .Col = 1
        .ColWidth(1) = 4330
        .Text = "Descricao                    "
        .Col = 2
        .ColWidth(2) = 3000
        .Text = "Fornecedor"
        Do While Not pRsProduto.EOF
           .AddItem pRsProduto!codprod & vbTab & pRsProduto!descricao & _
                                         vbTab & pRsProduto!nome
           pRsProduto.MoveNext
        Loop
        .Row = 1
        .Col = 0
        Txtescolhido.Text = .TextMatrix(1, 0)
        LblNomeProd.Caption = .TextMatrix(1, 1)
   End With
   pRsProduto.Close
   If Not IsNumeric(gCodProd) Then
      For i = 1 To MSFlexprod.Rows - 1
          If UCase(gCodProd) = Mid(UCase(MSFlexprod.TextMatrix(i, 1)), 1, Len(gCodProd)) Then
             MSFlexprod.Row = i
             MSFlexprod_Click
             Exit For
          End If
      Next
   'TxtDescEscolhido = MSFlexprod.Text
   'MSFlexprod.Col = 0
  End If
End Sub

Private Sub MSFlexprod_Click()
   Dim oldrow As Long
   Dim lcColGrid As Double
  
   With MSFlexprod
      oldrow = .Row
      If .Row = 0 Then
         lcColGrid = .Col
         .Col = lcColGrid
         .Sort = flexSortStringAscending
         Exit Sub
      End If
      Txtescolhido.Text = .TextMatrix(.Row, 0)
      LblNomeProd.Caption = .TextMatrix(.Row, 1)
      .Row = 0
      .Redraw = False
      Do While True
         .Row = .Row + 1
         For i = 0 To .Cols - 1
             .Col = i: .CellBackColor = vbWhite
         Next
         If .Row = .Rows - 1 Then
            Exit Do
         End If
      Loop
       .Redraw = True
       .Row = oldrow
       .Col = 0:   .CellBackColor = vbYellow
       .Col = 1:   .CellBackColor = vbYellow
       .Col = 2:   .CellBackColor = vbYellow
       .TopRow = .Row
   End With
  
End Sub

Private Sub MSFlexprod_GotFocus()
   MSFlexprod_Click
End Sub

Private Sub MSFlexprod_KeyDown(KeyCode As Integer, Shift As Integer)
Dim tecla

tecla = KeyCode

   If tecla = vbKeyReturn Then
      MSFlexprod.Col = 0
      Txtescolhido.Text = MSFlexprod.Text
      CmdOk.SetFocus
   End If
         
If tecla = vbKeyUp Or tecla = vbKeyDown Then
   MSFlexprod_Click
End If

End Sub

