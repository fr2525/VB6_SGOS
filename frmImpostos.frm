VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmImpostos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cad. Impostos"
   ClientHeight    =   5205
   ClientLeft      =   4545
   ClientTop       =   2445
   ClientWidth     =   7245
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5205
   ScaleWidth      =   7245
   ShowInTaskbar   =   0   'False
   Tag             =   "cadvend"
   Begin MSMask.MaskEdBox MskValor 
      Height          =   270
      Left            =   4500
      TabIndex        =   3
      Top             =   960
      Width           =   1965
      _ExtentX        =   3466
      _ExtentY        =   476
      _Version        =   393216
      Format          =   "R$#,##0.00;(R$#,##0.00)"
      PromptChar      =   "_"
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   2085
      Left            =   585
      TabIndex        =   4
      Top             =   1500
      Width           =   5820
      _ExtentX        =   10266
      _ExtentY        =   3678
      _Version        =   393216
      Rows            =   5
      Cols            =   3
      FixedCols       =   0
      FormatString    =   "Descrição                                                                          |Valor                        "
   End
   Begin VB.Frame Frame1 
      Height          =   795
      Left            =   1560
      TabIndex        =   1
      Top             =   3990
      Width           =   4245
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Salvar"
         Enabled         =   0   'False
         Height          =   540
         Left            =   2220
         Picture         =   "frmImpostos.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   7
         Tag             =   "&Update"
         Top             =   180
         Width           =   615
      End
      Begin VB.CommandButton cmdEditar 
         Caption         =   "&Editar"
         Height          =   540
         Left            =   780
         Picture         =   "frmImpostos.frx":00FA
         Style           =   1  'Graphical
         TabIndex        =   12
         Tag             =   "&Refresh"
         Top             =   180
         Width           =   615
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Excluir"
         Height          =   540
         Left            =   1500
         Picture         =   "frmImpostos.frx":026C
         Style           =   1  'Graphical
         TabIndex        =   11
         Tag             =   "&Delete"
         Top             =   180
         Width           =   615
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Incluir"
         Height          =   540
         Left            =   120
         Picture         =   "frmImpostos.frx":03DE
         Style           =   1  'Graphical
         TabIndex        =   10
         Tag             =   "&Add"
         Top             =   180
         Width           =   615
      End
      Begin VB.CommandButton CmdSair 
         Caption         =   "&Sair"
         Height          =   540
         Left            =   3540
         Picture         =   "frmImpostos.frx":04C8
         Style           =   1  'Graphical
         TabIndex        =   9
         Tag             =   "&Update"
         Top             =   180
         Width           =   615
      End
      Begin VB.CommandButton cmddesfaz 
         Caption         =   "&Desfaz"
         Enabled         =   0   'False
         Height          =   540
         Left            =   2880
         Picture         =   "frmImpostos.frx":05C2
         Style           =   1  'Graphical
         TabIndex        =   8
         Tag             =   "&Update"
         Top             =   180
         Width           =   615
      End
   End
   Begin VB.TextBox TxtDescricao 
      Height          =   285
      Left            =   1260
      TabIndex        =   2
      Top             =   510
      Width           =   5175
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Valor:"
      Height          =   195
      Left            =   3855
      TabIndex        =   13
      Top             =   990
      Width           =   405
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Código:"
      Height          =   195
      Index           =   0
      Left            =   600
      TabIndex        =   0
      Tag             =   "CODVEND:"
      Top             =   120
      Width           =   540
   End
   Begin VB.Label LblCod_imposto 
      Caption         =   "cod_imposto"
      ForeColor       =   &H000000FF&
      Height          =   210
      Left            =   1290
      TabIndex        =   6
      Top             =   135
      Width           =   975
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Descrição"
      Height          =   195
      Index           =   1
      Left            =   375
      TabIndex        =   5
      Tag             =   "NOME:"
      Top             =   585
      Width           =   720
   End
End
Attribute VB_Name = "frmImpostos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private lIncluir As Boolean

Private Sub Abre_Le_rst()
  gSql = "select * FROM tab_impostos"
  gRs.Open gSql, ConDb, adOpenKeyset
  
End Sub

Private Sub cmdAdd_Click()
   
   limpa_tela Me
   
   Me.LblCod_imposto.Caption = ""
   Me.TxtDescricao.SetFocus
   Me.cmdUpdate.Enabled = True
   Me.cmddesfaz.Enabled = True
   Me.cmdEditar.Enabled = False
   Me.cmdAdd.Enabled = False
   Me.CmdSair.Enabled = False
   Me.cmdDelete.Enabled = False
   
   lIncluir = True
End Sub

Private Sub cmdDelete_Click()

    On Error GoTo ErroDelete
    
    'this may produce an error if you delete the last
    'record or the only record in the recordset
    If MsgBox("Deseja realmente apagar este Imposto ? ", vbYesNo, "Atenção") = vbYes Then
       gSql = "delete * from tab_impostos where cod_imposto = " & Me.LblCod_imposto.Caption
       ConDb.Execute gSql
       gRs.Close
       Abre_Le_rst
       Carrega_Grid
       gRs.MoveFirst
       Carrega_tela
       Desabilita Me
     End If
     Exit Sub
     
ErroDelete:
     MsgBox "Deu erro na exclusao do Imposto " & Chr(13) & "Instrucao Sql = '" & _
            cSql & "'  "
End Sub


Private Sub cmddesfaz_Click()
  
  lIncluir = False
  
  ' Carrega_tela
  Desabilita Me
  MSFlexGrid1_Click
  Me.cmdUpdate.Enabled = False
  Me.cmddesfaz.Enabled = False
  Me.cmdEditar.Enabled = True
  Me.cmdAdd.Enabled = True
  Me.CmdSair.Enabled = True
  Me.cmdDelete.Enabled = True

End Sub

Private Sub cmdEditar_Click()
   ' Carrega_tela
   Habilita Me
   Me.TxtDescricao.SetFocus
   Me.cmdUpdate.Enabled = True
   Me.cmddesfaz.Enabled = True
   Me.cmdEditar.Enabled = False
   Me.cmdAdd.Enabled = False
   Me.CmdSair.Enabled = False
   Me.cmdDelete.Enabled = False
   
End Sub

Private Sub CmdSair_Click()
   Unload Me
End Sub

Private Sub cmdUpdate_Click()
 
   gRs.Close
   If lIncluir Then
      gSql = "INSERT INTO tab_impostos (descricao,valor,"
      gSql = gSql & "operador, datatual"
      gSql = gSql & ") "
      gSql = gSql & "VALUES ('" & Me.TxtDescricao.Text & "',"
      gSql = gSql & Val(Me.MskValor.Text) & ",'"
      gSql = gSql & gOperador & "',Cdate('" & Date & "' ))"
      ConDb.Execute gSql
      lIncluir = False
   Else
      gSql = "UPDATE tab_impostos SET descricao = '" & Me.TxtDescricao.Text & "',"
      gSql = gSql & "valor = " & Val(Me.MskValor.Text) & ","
      gSql = gSql & " operador = '" & gOperador & "', datatual = Cdate('" & Date & "'))"
      gSql = gSql & " WHERE cod_imposto = " & Me.LblCod_imposto.Caption
      ConDb.Execute gSql
      
   End If
       
   Abre_Le_rst
   
   Carrega_Grid
   gRs.MoveFirst
   Carrega_tela
   Desabilita Me
      
   Me.cmdUpdate.Enabled = False
   Me.cmddesfaz.Enabled = False
   Me.cmdEditar.Enabled = True
   Me.cmdAdd.Enabled = True
   Me.CmdSair.Enabled = True
   Me.cmdDelete.Enabled = True
     
End Sub

Private Sub Form_Activate()

    Abre_Le_rst
   
   
   Me.LblCod_imposto.Caption = ""
   If gRs.BOF And gRs.EOF Then
      If MsgBox("Arquivo vazio. Incluir dados agora ?", vbYesNo, "Atenção ") = vbYes Then
         gSql = "INSERT INTO tab_impostos (descricao,valor,"
         gSql = gSql & "operador, datatual"
         gSql = gSql & ") "
         gSql = gSql & "VALUES ('" & f_nulo(Me.TxtDescricao.Text, " ") & "',"
         gSql = gSql & Val(Me.MskValor.Text) & ",'"
         gSql = gSql & gOperador & "'," & Date & " )"
         ConDb.Execute gSql
         gRs.Close
         Abre_Le_rst
         Me.LblCod_imposto.Caption = gRs!cod_Imposto
         cmdEditar_Click
         lPrimeiro = True
      Else
         Desabilita Me
      End If
      
   Else
      gRs.MoveFirst
      Carrega_tela
    
      Desabilita Me
      lIncluir = False
      lPrimeiro = False
   End If
   Carrega_Grid
   
   lIncluir = False
   
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub Form_Load()
   'Centraliza a tela no video
   Me.Move (Screen.Width - Me.Width) / 2, _
           (Screen.Height - Me.Height) / 2
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
    gRs.Close
    Screen.MousePointer = vbDefault
End Sub

Private Sub MSFlexGrid1_Click()
 Dim oldrow As Long
  
  oldrow = MSFlexGrid1.Row
  
  MSFlexGrid1.Row = 0
  
  With MSFlexGrid1
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
    
    '.Col = 0:   LblCod_imposto.Caption = .Text: .CellBackColor = vbYellow
    .Col = 0:   TxtDescricao.Text = .Text: .CellBackColor = vbYellow
    .Col = 1:   MskValor.Text = Format(.Text, "R$#,##0.00;(R$#,##0.00)"): .CellBackColor = vbYellow
    .Col = 2:   LblCod_imposto.Caption = .Text: .CellBackColor = vbYellow
    .TopRow = .Row
    
    
End With


End Sub


Private Sub Carrega_tela()
   'Limpa as variaveis da tela se caso ficarem com dados da outra tela
   limpa_tela Me
   'Carrega a tela com os dados do registro
   Me.LblCod_imposto.Caption = gRs("cod_imposto")
   Me.TxtDescricao.Text = "" & gRs("descricao")
   Me.MskValor.Text = "" & Format(gRs!Valor, "R$#,##0.00;(R$#,##0.00)")
   
   
End Sub

Private Sub Carrega_Grid()

'Teste do MsHFlexgrid1 - eh eh eh
  MSFlexGrid1.Row = 0
  
  With gRs
      '.MoveLast
      'nItem = .RecordCount
      '.MoveFirst
      MSFlexGrid1.Rows = 1
      MSFlexGrid1.ColWidth(2) = 10
      Do While Not .EOF
         MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
         MSFlexGrid1.Row = MSFlexGrid1.Rows - 1
            
         'MSFlexGrid1.Col = 0:  MSFlexGrid1.Text = f_nulo(!cod_imposto, "")
         MSFlexGrid1.Col = 0:  MSFlexGrid1.Text = f_nulo(!descricao, "")
         MSFlexGrid1.Col = 1:  MSFlexGrid1.Text = Format(!Valor, "R$#,##0.00;(R$#,##0.00)")
         MSFlexGrid1.Col = 2:  MSFlexGrid1.Text = !cod_Imposto
         .MoveNext
         
       Loop
       MSFlexGrid1.FixedRows = 1
          
  End With
  
End Sub

Private Sub MskValor_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Or KeyAscii = 9 Then KeyAscii = 0
End Sub

Private Sub TxtDescricao_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Or KeyAscii = 9 Then KeyAscii = 0
End Sub
