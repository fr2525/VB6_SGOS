VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmCCusto 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3630
   ClientLeft      =   3315
   ClientTop       =   2460
   ClientWidth     =   6315
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3630
   ScaleWidth      =   6315
   ShowInTaskbar   =   0   'False
   Tag             =   "cadvend"
   Begin VB.TextBox TxtLocal 
      Height          =   300
      Left            =   975
      TabIndex        =   13
      Top             =   765
      Width           =   3525
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   1305
      Left            =   390
      TabIndex        =   12
      Top             =   1215
      Width           =   5370
      _ExtentX        =   9472
      _ExtentY        =   2302
      _Version        =   393216
      Rows            =   5
      Cols            =   6
      FixedCols       =   0
      BackColorSel    =   16777215
      SelectionMode   =   1
      FormatString    =   "Código|Descrição                                        | Local                            "
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   900
      TabIndex        =   1
      Top             =   2640
      Width           =   4245
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Salvar"
         Enabled         =   0   'False
         Height          =   540
         Left            =   2190
         Picture         =   "frmCCusto.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   6
         Tag             =   "&Update"
         Top             =   150
         Width           =   615
      End
      Begin VB.CommandButton cmdEditar 
         Caption         =   "&Editar"
         Height          =   540
         Left            =   795
         Picture         =   "frmCCusto.frx":00FA
         Style           =   1  'Graphical
         TabIndex        =   11
         Tag             =   "&Refresh"
         Top             =   150
         Width           =   615
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Excluir"
         Height          =   540
         Left            =   1485
         Picture         =   "frmCCusto.frx":026C
         Style           =   1  'Graphical
         TabIndex        =   10
         Tag             =   "&Delete"
         Top             =   150
         Width           =   615
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Incluir"
         Height          =   540
         Left            =   105
         Picture         =   "frmCCusto.frx":03DE
         Style           =   1  'Graphical
         TabIndex        =   9
         Tag             =   "&Add"
         Top             =   150
         Width           =   615
      End
      Begin VB.CommandButton CmdSair 
         Caption         =   "&Sair"
         Height          =   540
         Left            =   3555
         Picture         =   "frmCCusto.frx":04C8
         Style           =   1  'Graphical
         TabIndex        =   7
         Tag             =   "&Update"
         Top             =   150
         Width           =   615
      End
      Begin VB.CommandButton cmddesfaz 
         Caption         =   "&Desfaz"
         Enabled         =   0   'False
         Height          =   540
         Left            =   2865
         Picture         =   "frmCCusto.frx":05C2
         Style           =   1  'Graphical
         TabIndex        =   8
         Tag             =   "&Update"
         Top             =   150
         Width           =   615
      End
   End
   Begin VB.TextBox TxtDescricao 
      Height          =   285
      Left            =   990
      MaxLength       =   50
      TabIndex        =   4
      Top             =   390
      Width           =   3525
   End
   Begin VB.Label LblCodigo 
      Caption         =   "codigo"
      ForeColor       =   &H000000FF&
      Height          =   270
      Left            =   1020
      TabIndex        =   2
      Top             =   90
      Width           =   645
   End
   Begin VB.Label lblfixo 
      Alignment       =   1  'Right Justify
      Caption         =   "Local:"
      Height          =   255
      Index           =   2
      Left            =   105
      TabIndex        =   5
      Tag             =   "SALARIO:"
      Top             =   780
      Width           =   735
   End
   Begin VB.Label lblDescricao 
      Alignment       =   1  'Right Justify
      Caption         =   "Nome:"
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   3
      Tag             =   "NOME:"
      Top             =   405
      Width           =   525
   End
   Begin VB.Label Labelcodigo 
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
Attribute VB_Name = "frmCCusto"
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
   Me.lblcodigo = gRs!codigo
   Me.TxtDescricao.Text = gRs!descricao
   Me.TxtLocal.Text = f_nulo(gRs!Local, "")
   
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
            
         MSFlexGrid1.Col = 0: MSFlexGrid1.Text = f_nulo(!codigo, "")
         MSFlexGrid1.Col = 1: MSFlexGrid1.Text = f_nulo(!descricao, "")
         MSFlexGrid1.Col = 2: MSFlexGrid1.Text = f_nulo(!Local, "")
         .MoveNext
         
       Loop
       MSFlexGrid1.FixedRows = 1
          
  End With
  
  End Sub

Private Sub cmdAdd_Click()
   
   lIncluir = True
   limpa_tela Me
   
   Me.lblcodigo.Caption = ""
   Me.TxtDescricao.SetFocus
   suCmdAdd Me  'Habilita e desabilita botoes
   
End Sub

Private Sub cmdDelete_Click()
    'this may produce an error if you delete the last
    'record or the only record in the recordset
    If MsgBox("Deseja realmente apagar este Centro de Custo ? ", vbYesNo, "Atenção") = vbYes Then
       gSql = "delete * from tab_ccusto where codigo = " & Val(Me.lblcodigo.Caption)
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
   Me.TxtDescricao.SetFocus
   
End Sub


Private Sub CmdSair_Click()
   Unload Me
End Sub

Private Sub cmdUpdate_Click()
   
   gRs.Close
   If lIncluir Then
      gSql = "INSERT INTO tab_ccusto (descricao,Local,operador,datatual) "
      gSql = gSql & "VALUES ('" & Me.TxtDescricao.Text & "','"
      gSql = gSql & Me.TxtLocal.Text & "','"
      gSql = gSql & gOperador & "',Cdate('" & Date & "')) "
      ConDb.Execute gSql
      lIncluir = False
   Else
      gSql = "UPDATE tab_ccusto SET descricao = '" & Me.TxtDescricao.Text & "',"
      gSql = gSql & "Local = '" & Me.TxtLocal.Text & "',"
      gSql = gSql & "operador = '" & gOperador & "', datatual = cDate('" & Date & "')"
      gSql = gSql & " WHERE codigo = " & Val(Me.lblcodigo.Caption)
      ConDb.Execute gSql
      
   End If
                              
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

   Me.lblcodigo.Caption = ""

   If gRs.BOF And gRs.EOF Then
      If MsgBox("Arquivo vazio. Incluir dados agora ?", vbYesNo, "Atenção ") = vbYes Then
         gSql = "INSERT INTO tab_ccusto (descricao,Local,"
         gSql = gSql & "operador, datatual) "
         gSql = gSql & " VALUES ( '" & f_nulo(Me.TxtDescricao.Text, " ") & "','"
         gSql = gSql & f_nulo(Me.TxtLocal.Text, " ") & "','"
         gSql = gSql & gOperador & "', cdate('" & Date & "') )"
         ConDb.Execute gSql
         gRs.Close
         Abre_Le_rst
         Me.lblcodigo.Caption = gRs!codigo
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

Private Sub Abre_Le_rst()
   gSql = "select * FROM tab_ccusto order by descricao"
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
    
    .Col = 0:   lblcodigo.Caption = .Text: .CellBackColor = vbYellow
    .Col = 1:   TxtDescricao.Text = .Text: .CellBackColor = vbYellow
    .Col = 2:   TxtLocal.Text = .Text: .CellBackColor = vbYellow
    .Redraw = True
    .TopRow = .Row
    
  End With

End Sub

Private Sub MSFlexGrid1_SelChange()
   MSFlexGrid1_Click
End Sub


