VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form Frmtipovend 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tipos de Venda"
   ClientHeight    =   5340
   ClientLeft      =   1650
   ClientTop       =   2025
   ClientWidth     =   7080
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5340
   ScaleWidth      =   7080
   ShowInTaskbar   =   0   'False
   Tag             =   "cadvend"
   Begin VB.TextBox TxtEspecial 
      Height          =   285
      Left            =   1440
      TabIndex        =   6
      Top             =   990
      Width           =   4215
   End
   Begin VB.TextBox TxtParcelas 
      Height          =   285
      Left            =   4185
      TabIndex        =   5
      Top             =   600
      Width           =   495
   End
   Begin VB.TextBox TxtDias 
      Height          =   285
      Left            =   2520
      TabIndex        =   4
      Top             =   600
      Width           =   615
   End
   Begin VB.TextBox TxtEntrada 
      Height          =   285
      Left            =   1440
      MaxLength       =   1
      TabIndex        =   3
      Top             =   600
      Width           =   375
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   1380
      TabIndex        =   7
      Top             =   4260
      Width           =   4245
      Begin VB.CommandButton cmddesfaz 
         Caption         =   "&Desfaz"
         Enabled         =   0   'False
         Height          =   540
         Left            =   2865
         Picture         =   "Frmtipovend.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   13
         Tag             =   "&Update"
         Top             =   150
         Width           =   615
      End
      Begin VB.CommandButton CmdSair 
         Caption         =   "&Sair"
         Height          =   540
         Left            =   3555
         Picture         =   "Frmtipovend.frx":00FA
         Style           =   1  'Graphical
         TabIndex        =   12
         Tag             =   "&Update"
         Top             =   150
         Width           =   615
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Incluir"
         Height          =   540
         Left            =   105
         Picture         =   "Frmtipovend.frx":01F4
         Style           =   1  'Graphical
         TabIndex        =   11
         Tag             =   "&Add"
         Top             =   150
         Width           =   615
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Excluir"
         Height          =   540
         Left            =   1485
         Picture         =   "Frmtipovend.frx":02DE
         Style           =   1  'Graphical
         TabIndex        =   10
         Tag             =   "&Delete"
         Top             =   150
         Width           =   615
      End
      Begin VB.CommandButton cmdEditar 
         Caption         =   "&Editar"
         Height          =   540
         Left            =   795
         Picture         =   "Frmtipovend.frx":0450
         Style           =   1  'Graphical
         TabIndex        =   9
         Tag             =   "&Refresh"
         Top             =   150
         Width           =   615
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Salvar"
         Enabled         =   0   'False
         Height          =   540
         Left            =   2190
         Picture         =   "Frmtipovend.frx":05C2
         Style           =   1  'Graphical
         TabIndex        =   8
         Tag             =   "&Update"
         Top             =   150
         Width           =   615
      End
   End
   Begin VB.TextBox TxtDescricao 
      Height          =   285
      Left            =   2370
      TabIndex        =   2
      Top             =   120
      Width           =   3270
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   2535
      Left            =   360
      TabIndex        =   14
      Top             =   1440
      Width           =   6435
      _ExtentX        =   11351
      _ExtentY        =   4471
      _Version        =   393216
      Rows            =   5
      Cols            =   6
      FixedCols       =   0
      FormatString    =   "Código|Descrição                                                       | Entrada|Dias|Parcelas|Especial                     "
   End
   Begin VB.Label Label1 
      Caption         =   "Especial:"
      Height          =   255
      Left            =   435
      TabIndex        =   18
      Top             =   990
      Width           =   855
   End
   Begin VB.Label LblParcelas 
      AutoSize        =   -1  'True
      Caption         =   "Parcelas"
      Height          =   195
      Left            =   3420
      TabIndex        =   17
      Top             =   630
      Width           =   615
   End
   Begin VB.Label LblDias 
      AutoSize        =   -1  'True
      Caption         =   "Dias"
      Height          =   195
      Left            =   2040
      TabIndex        =   16
      Top             =   615
      Width           =   315
   End
   Begin VB.Label LblEntrada 
      AutoSize        =   -1  'True
      Caption         =   "Entrada ?"
      Height          =   195
      Left            =   480
      TabIndex        =   15
      Top             =   600
      Width           =   690
   End
   Begin VB.Label Lbltipovend 
      Caption         =   "codvend"
      ForeColor       =   &H000000FF&
      Height          =   270
      Left            =   1740
      TabIndex        =   1
      Top             =   150
      Width           =   525
   End
   Begin VB.Label lblTipo 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Tipo de Venda;"
      Height          =   195
      Index           =   0
      Left            =   450
      TabIndex        =   0
      Tag             =   "CODVEND:"
      Top             =   150
      Width           =   1095
   End
End
Attribute VB_Name = "Frmtipovend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private lIncluir As Boolean
Private lPrimeiro As Boolean

Private Sub Abre_Le_rst()
   gSql = "select * FROM tipovend"
   gRs.Open gSql, ConDb, adOpenKeyset
   
End Sub

Private Sub Carrega_Grid()
'Teste do MsFlexgrid1
  MSFlexGrid1.Row = 0
  
  With gRs
      '.MoveLast
      'nItem = .RecordCount
      '.MoveFirst
      MSFlexGrid1.Rows = 1
      Do While Not .EOF
         MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
         MSFlexGrid1.Row = MSFlexGrid1.Rows - 1
         MSFlexGrid1.Col = 0: MSFlexGrid1.Text = f_nulo(!código, "")
         MSFlexGrid1.Col = 1: MSFlexGrid1.Text = f_nulo(!descricao, "")
         MSFlexGrid1.Col = 2: MSFlexGrid1.Text = f_nulo(!entrada, "")
         MSFlexGrid1.Col = 3: MSFlexGrid1.Text = f_nulo(!dias, "")
         MSFlexGrid1.Col = 4: MSFlexGrid1.Text = f_nulo(!parcelas, "")
         MSFlexGrid1.Col = 5: MSFlexGrid1.Text = f_nulo(!Especial, "")
         .MoveNext
         
       Loop
       MSFlexGrid1.FixedRows = 1
          
  End With
  
  End Sub

Private Sub Carrega_tela()
   'Limpa as variaveis da tela se caso ficarem com dados da outra tela
   limpa_tela Me
   'Carrega a tela com os dados do registro
   Me.Lbltipovend.Caption = gRs("código")
   Me.TxtDescricao.Text = gRs("descricao")
   Me.TxtDias.Text = gRs("dias")
   Me.TxtEntrada = gRs!entrada
   Me.TxtParcelas = gRs!parcelas
   Me.TxtEspecial = gRs!Especial
   
End Sub
Private Sub cmdAdd_Click()
   
   lIncluir = True
   limpa_tela Me
   
   Me.Lbltipovend.Caption = ""
   Me.TxtDescricao.SetFocus
   Me.cmdUpdate.Enabled = True
   Me.cmddesfaz.Enabled = True
   Me.cmdEditar.Enabled = False
   Me.cmdAdd.Enabled = False
   Me.CmdSair.Enabled = False
   Me.cmdDelete.Enabled = False
   
End Sub

Private Sub cmdDelete_Click()

    On Error GoTo ErroDelete
    
    'this may produce an error if you delete the last
    'record or the only record in the recordset
    If MsgBox("Deseja realmente apagar este Tipo de Venda ? ", vbYesNo, "Atenção") = vbYes Then
        gSql = "delete * from tipovend where código = " & Val(Me.Lbltipovend.Caption)
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
     MsgBox "Deu erro na exclusao do Tipo de Venda " & Chr(13) & "Instrucao Sql = '" & _
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
   
   Habilita Me
        
   Me.cmdUpdate.Enabled = True
   Me.cmddesfaz.Enabled = True
   Me.cmdEditar.Enabled = False
   Me.cmdAdd.Enabled = False
   Me.CmdSair.Enabled = False
   Me.cmdDelete.Enabled = False
   Me.TxtDescricao.SetFocus
   
End Sub

Private Sub CmdSair_Click()
   Unload Me
End Sub

Private Sub cmdUpdate_Click()
   gRs.Close
   If lIncluir Then
      gSql = "INSERT INTO tipovend (descricao,entrada,dias,parcelas,especial,operador,datatual) "
      gSql = gSql & "VALUES ('" & Me.TxtDescricao.Text & "','"
      gSql = gSql & Me.TxtEntrada.Text & "','" & Me.TxtDias.Text & "','"
      gSql = gSql & Me.TxtParcelas.Text & "','"
      gSql = gSql & Me.TxtEspecial.Text & "',"
      gSql = gSql & "'" & gOperador & "',Cdate('" & Date & "'))"
      ConDb.Execute gSql
      lIncluir = False
      
   Else
      gSql = "UPDATE tipovend SET descricao = '" & Me.TxtDescricao.Text & "',"
      gSql = gSql & "entrada = '" & IIf(Me.TxtEntrada.Text = "S", "S", "N") & "',"
      gSql = gSql & "dias = '" & Val(Me.TxtDias.Text) & "',"
      gSql = gSql & "parcelas = '" & Val(Me.TxtParcelas.Text) & "',"
      gSql = gSql & "especial = '" & Me.TxtEspecial.Text & "',"
      gSql = gSql & " operador = '" & gOperador & "', datatual = Cdate('" & Date & "')"
      gSql = gSql & " WHERE código = " & Val(Lbltipovend.Caption)
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
  limpa_tela Me
   
  Me.Lbltipovend.Caption = ""
  If gRs.BOF And gRs.EOF Then
      If MsgBox("Arquivo vazio. Incluir dados agora ?", vbYesNo, "Atenção ") = vbYes Then
         gSql = "INSERT INTO tipovend (descricao,entrada,dias,parcelas,especial,operador,datatual) "
         gSql = gSql & "VALUES ('" & Me.TxtDescricao.Text & "','"
         gSql = gSql & Me.TxtEntrada.Text & "','" & Me.TxtDias.Text & "','"
         gSql = gSql & Me.TxtParcelas.Text & "','"
         gSql = gSql & Me.TxtEspecial.Text & "',"
         gSql = gSql & "'" & gOperador & "'," & Date & " )"
         ConDb.Execute gSql
         gRs.Close
         Abre_Le_rst
         Me.Lbltipovend.Caption = gRs!código
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
    Screen.MousePointer = vbDefault
    gRs.Close
    
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
    
    .Col = 0:   Lbltipovend.Caption = .Text: .CellBackColor = vbYellow
    .Col = 1:   TxtDescricao.Text = .Text: .CellBackColor = vbYellow
    .Col = 2:   TxtEntrada.Text = .Text: .CellBackColor = vbYellow
    .Col = 3:   TxtDias.Text = .Text: .CellBackColor = vbYellow
    .Col = 4:   TxtParcelas.Text = .Text: .CellBackColor = vbYellow
    .Col = 5:   TxtEspecial.Text = .Text: .CellBackColor = vbYellow
    .TopRow = .Row
   
End With

End Sub

Private Sub TxtEntrada_Validate(Cancel As Boolean)
    TxtEntrada.Text = UCase(TxtEntrada.Text)
    If UCase(TxtEntrada.Text) <> "S" And UCase(TxtEntrada.Text) <> "N" Then
       MsgBox "Digite somente 'S' ou 'N' por favor", vbOKOnly, "Atenção: " + gOperador
       Cancel = True
    End If
    
End Sub
