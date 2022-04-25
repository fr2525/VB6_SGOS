VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmBalco 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Operadores"
   ClientHeight    =   3450
   ClientLeft      =   3315
   ClientTop       =   2460
   ClientWidth     =   6315
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3450
   ScaleWidth      =   6315
   ShowInTaskbar   =   0   'False
   Tag             =   "cadvend"
   Begin VB.TextBox TxtNivel 
      Height          =   315
      Left            =   5670
      TabIndex        =   9
      Top             =   420
      Width           =   345
   End
   Begin VB.TextBox txtComissao 
      Height          =   315
      Left            =   2910
      MaxLength       =   4
      TabIndex        =   7
      Top             =   420
      Width           =   495
   End
   Begin MSMask.MaskEdBox MskFixo 
      Height          =   315
      Left            =   960
      TabIndex        =   6
      Top             =   420
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   556
      _Version        =   393216
      Format          =   "R$#,##0.00;(R$#,##0.00)"
      PromptChar      =   "_"
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   1305
      Left            =   615
      TabIndex        =   18
      Top             =   960
      Width           =   5085
      _ExtentX        =   8969
      _ExtentY        =   2302
      _Version        =   393216
      Rows            =   5
      Cols            =   6
      FixedCols       =   0
      SelectionMode   =   1
      FormatString    =   "Código|Nome                                     | Fixo          | Comissão  | Senha   |Nivel"
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   870
      TabIndex        =   1
      Top             =   2460
      Width           =   4245
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Salvar"
         Enabled         =   0   'False
         Height          =   540
         Left            =   2160
         Picture         =   "frmbalco.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   12
         Tag             =   "&Update"
         Top             =   150
         Width           =   615
      End
      Begin VB.CommandButton cmdEditar 
         Caption         =   "&Editar"
         Height          =   540
         Left            =   795
         Picture         =   "frmbalco.frx":00FA
         Style           =   1  'Graphical
         TabIndex        =   17
         Tag             =   "&Refresh"
         Top             =   150
         Width           =   615
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Excluir"
         Height          =   540
         Left            =   1485
         Picture         =   "frmbalco.frx":026C
         Style           =   1  'Graphical
         TabIndex        =   16
         Tag             =   "&Delete"
         Top             =   150
         Width           =   615
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Incluir"
         Height          =   540
         Left            =   105
         Picture         =   "frmbalco.frx":03DE
         Style           =   1  'Graphical
         TabIndex        =   15
         Tag             =   "&Add"
         Top             =   150
         Width           =   615
      End
      Begin VB.CommandButton CmdSair 
         Caption         =   "&Sair"
         Height          =   540
         Left            =   3555
         Picture         =   "frmbalco.frx":04C8
         Style           =   1  'Graphical
         TabIndex        =   13
         Tag             =   "&Update"
         Top             =   150
         Width           =   615
      End
      Begin VB.CommandButton cmddesfaz 
         Caption         =   "&Desfaz"
         Enabled         =   0   'False
         Height          =   540
         Left            =   2865
         Picture         =   "frmbalco.frx":05C2
         Style           =   1  'Graphical
         TabIndex        =   14
         Tag             =   "&Update"
         Top             =   150
         Width           =   615
      End
   End
   Begin VB.TextBox TxtSenha 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   4170
      MaxLength       =   8
      PasswordChar    =   "*"
      TabIndex        =   8
      Top             =   420
      Width           =   810
   End
   Begin VB.TextBox TxtNome 
      Height          =   285
      Left            =   2490
      MaxLength       =   50
      TabIndex        =   4
      Top             =   90
      Width           =   3525
   End
   Begin VB.Label LblNivel 
      AutoSize        =   -1  'True
      Caption         =   "Nível"
      Height          =   195
      Left            =   5190
      TabIndex        =   19
      Top             =   480
      Width           =   390
   End
   Begin VB.Label LblCodvend 
      Caption         =   "codvend"
      ForeColor       =   &H000000FF&
      Height          =   270
      Left            =   1020
      TabIndex        =   2
      Top             =   90
      Width           =   645
   End
   Begin VB.Label lblsenha 
      Alignment       =   1  'Right Justify
      Caption         =   "Senha:"
      Height          =   255
      Index           =   4
      Left            =   3540
      TabIndex        =   11
      Tag             =   "SENHA:"
      Top             =   450
      Width           =   525
   End
   Begin VB.Label lblcomissao 
      Alignment       =   1  'Right Justify
      Caption         =   "Comissão:"
      Height          =   255
      Index           =   3
      Left            =   2100
      TabIndex        =   10
      Tag             =   "COMISSAO:"
      Top             =   450
      Width           =   735
   End
   Begin VB.Label lblfixo 
      Alignment       =   1  'Right Justify
      Caption         =   "Fixo:"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   5
      Tag             =   "SALARIO:"
      Top             =   420
      Width           =   735
   End
   Begin VB.Label lblNome 
      Alignment       =   1  'Right Justify
      Caption         =   "Nome:"
      Height          =   255
      Index           =   1
      Left            =   1800
      TabIndex        =   3
      Tag             =   "NOME:"
      Top             =   90
      Width           =   525
   End
   Begin VB.Label lblcodigo 
      Alignment       =   1  'Right Justify
      Caption         =   "Código:"
      Height          =   255
      Index           =   0
      Left            =   150
      TabIndex        =   0
      Tag             =   "CODVEND:"
      Top             =   75
      Width           =   765
   End
End
Attribute VB_Name = "frmBalco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private lIncluir As Boolean
Private lPrimeiro As Boolean


Private Sub Carrega_tela()
   'Limpa as variaveis da tela se caso ficarem com dados da outra tela
   limpa_tela Me
   'Carrega a tela com os dados do registro
   Me.LblCodvend = gRs!id
   Me.TxtNome.text = gRs!nome
   Me.MskFixo.text = Format(gRs!salario, "R$#,##0.00;(R$#,##0.00)")
   Me.txtComissao.text = Format(gRs!comissao / 100, "##0%")
   Me.TxtSenha.text = fuEncript(gRs!senha, "oyster")
   Me.TxtNivel.text = gRs("nivel")
   
End Sub
Private Sub Carrega_Grid()

'Teste do MsHFlexgrid1 - eh eh eh
  MSFlexGrid1.Row = 0
  
  If gRs.BOF And gRs.EOF Then
     Exit Sub
  End If
  
  With gRs
      '.MoveLast
      'nItem = .RecordCount
      '.MoveFirst
      MSFlexGrid1.Rows = 1
      Do While Not .EOF
         MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
         MSFlexGrid1.Row = MSFlexGrid1.Rows - 1
            
         MSFlexGrid1.Col = 0: MSFlexGrid1.text = f_nulo(!id, "")
         MSFlexGrid1.Col = 1: MSFlexGrid1.text = f_nulo(!nome, "")
         MSFlexGrid1.Col = 2: MSFlexGrid1.text = Format(f_nulo(!salario, 0), "0.00")
         MSFlexGrid1.Col = 3: MSFlexGrid1.text = Format(f_nulo(!comissao / 100, 0), "##0%")
         
         MSFlexGrid1.Col = 4: MSFlexGrid1.ColWidth(4) = 0
                              MSFlexGrid1.text = Format(fuEncript(f_nulo(!senha, ""), "oyster"), "******")
         
         MSFlexGrid1.Col = 5: MSFlexGrid1.text = f_nulo(!nivel, "")
         .MoveNext
         
       Loop
       MSFlexGrid1.FixedRows = 1
          
  End With
  
  End Sub

Private Sub cmdAdd_Click()
   
   lIncluir = True
   limpa_tela Me
   
   Me.MskFixo.Enabled = True
   Me.MskFixo.text = ""
   Me.LblCodvend.Caption = ""
   Me.TxtNome.SetFocus
   suCmdAdd Me  'Habilita e desabilita botoes
   
End Sub

Private Sub cmdDelete_Click()
    'this may produce an error if you delete the last
    'record or the only record in the recordset
    If MsgBox("Deseja realmente apagar este balconista ? ", vbYesNo, "Atenção") = vbYes Then
       gSql = "delete * from tab_operador where id = " & Val(Me.LblCodvend.Caption)
       ConDb.Execute gSql
       gRs.Close
       Abre_Le_rst
       Carrega_Grid
       gRs.MoveFirst
       Carrega_tela
       Desabilita Me
     End If
End Sub

Private Sub cmddesfaz_Click()
  lIncluir = False
  'Carrega_tela
  Desabilita Me
  'Me.TxtFixo.Enabled = False
    
  MSFlexGrid1_Click
  
  suCmdDesfaz Me
  
End Sub

Private Sub cmdEditar_Click()
   Habilita Me
   
   suCmdEditar Me
   Me.TxtNome.SetFocus
   
End Sub


Private Sub CmdSair_Click()
   Unload Me
End Sub

Private Sub cmdUpdate_Click()
   
   gRs.Close
   If lIncluir Then
      gSql = "INSERT INTO tab_operador (Nome,senha,salario,comissao,nivel,operador,datatual) "
      gSql = gSql & "VALUES ('" & Me.TxtNome.text & "','" & fuEncript(Me.TxtSenha.text, "oyster") & "',"
      gSql = gSql & Val(Me.MskFixo.text) & "," & Val(Me.txtComissao.text)
      gSql = gSql & "," & Val(Me.TxtNivel.text)
      gSql = gSql & ",'" & gOperador & "','" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "')"
      ConDb.Execute gSql
      lIncluir = False
   Else
      gSql = "UPDATE tab_operador SET Nome = '" & Me.TxtNome.text & "', senha = '"
      gSql = gSql & fuEncript(Me.TxtSenha.text, "oyster") & "', salario = " & Val(Me.MskFixo.text) & ","
      gSql = gSql & "comissao = " & Val(Me.txtComissao.text)
      gSql = gSql & ",nivel = " & Val(Me.TxtNivel.text)
      gSql = gSql & " ,operador = '" & gOperador & "', datatual = '" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "'"
      gSql = gSql & " WHERE id = " & Val(Me.LblCodvend.Caption)
      ConDb.Execute gSql
      
   End If
                              
   'Deixa os textbox desabilitados
   Me.MskFixo.Enabled = False
   Me.MskFixo.text = ""
   
   Abre_Le_rst
   Carrega_Grid
   gRs.MoveLast
   Carrega_tela
   Desabilita Me
   
   suCmdUpdate Me
    
End Sub


Private Sub Form_Activate()
   Abre_Le_rst
   
   limpa_tela Me

   Me.LblCodvend.Caption = ""

   If gRs.BOF And gRs.EOF Then
      If MsgBox("Arquivo vazio. Incluir dados agora ?", vbYesNo, "Atenção ") = vbYes Then
         gSql = "INSERT INTO tab_operador (Nome,senha,nivel,salario,"
         gSql = gSql & "comissao,operador, datatual) "
         gSql = gSql & " VALUES ( '" & f_nulo(Me.TxtNome.text, " ") & "','"
         gSql = gSql & f_nulo(Me.TxtSenha.text, " ") & "',"
         gSql = gSql & Val(Me.TxtNivel.text) & ","
         gSql = gSql & Val(Me.MskFixo.text) & ","
         gSql = gSql & Val(Me.txtComissao.text) & ",'"
         gSql = gSql & gOperador & "'," & Date & ")"
         ConDb.Execute gSql
         gRs.Close
         Abre_Le_rst
         Me.LblCodvend.Caption = gRs!id
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
  If KeyCode = vbKeyReturn Then Sendkeys "{TAB}"
     
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

Private Sub Abre_Le_rst()
   gSql = "select * FROM tab_operador"
   gRs.Open gSql, ConDb, adOpenKeyset, adLockOptimistic
    
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
  
    .Refresh
    .Row = oldrow
    
    .Col = 0:   LblCodvend.Caption = .text: .CellBackColor = vbYellow
    .Col = 1:   TxtNome.text = .text: .CellBackColor = vbYellow
    .Col = 2:   MskFixo.text = .text: .CellBackColor = vbYellow
    .Col = 3:   txtComissao.text = .text: .CellBackColor = vbYellow
    .Col = 4:   TxtSenha.text = .text: .CellBackColor = vbYellow
    .Col = 5:   TxtNivel.text = .text: .CellBackColor = vbYellow
    .Redraw = True
    .TopRow = .Row
    
  End With

End Sub


Private Sub MSFlexGrid1_SelChange()
   MSFlexGrid1_Click
End Sub

Private Sub MskFixo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Or KeyAscii = 9 Then KeyAscii = 0
End Sub

Private Sub txtComissao_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Or KeyAscii = 9 Then KeyAscii = 0
End Sub

Private Sub TxtNivel_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Or KeyAscii = 9 Then KeyAscii = 0
End Sub

Private Sub TxtNome_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Or KeyAscii = 9 Then KeyAscii = 0
End Sub

Private Sub TxtSenha_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Or KeyAscii = 9 Then KeyAscii = 0
End Sub
