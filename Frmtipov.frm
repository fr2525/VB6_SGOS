VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "ss32x25.ocx"
Begin VB.Form Frmtipovend 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tipos de Venda"
   ClientHeight    =   4935
   ClientLeft      =   3600
   ClientTop       =   3015
   ClientWidth     =   7110
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   7110
   ShowInTaskbar   =   0   'False
   Tag             =   "cadvend"
   Begin FPSpread.vaSpread vaSpr1 
      Height          =   1665
      Left            =   255
      TabIndex        =   19
      Top             =   1830
      Width           =   6405
      _Version        =   131077
      _ExtentX        =   11298
      _ExtentY        =   2937
      _StockProps     =   64
      DisplayRowHeaders=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   6
      OperationMode   =   2
      SelectBlockOptions=   0
      SpreadDesigner  =   "Frmtipov.frx":0000
      UserResize      =   1
   End
   Begin VB.TextBox TxtEspecial 
      Height          =   285
      Left            =   1155
      TabIndex        =   18
      Top             =   1320
      Width           =   5490
   End
   Begin VB.TextBox TxtParcelas 
      Height          =   285
      Left            =   6195
      TabIndex        =   4
      Top             =   915
      Width           =   450
   End
   Begin VB.TextBox Txtdias 
      Height          =   285
      Left            =   3645
      TabIndex        =   3
      Top             =   915
      Width           =   480
   End
   Begin VB.TextBox TxtEntrada 
      Height          =   285
      Left            =   1155
      TabIndex        =   2
      Top             =   915
      Width           =   225
   End
   Begin VB.TextBox TxtDescricao 
      Height          =   285
      Left            =   1155
      TabIndex        =   1
      Top             =   435
      Width           =   5490
   End
   Begin VB.Frame Frame1 
      Height          =   885
      Left            =   1290
      TabIndex        =   7
      Top             =   3735
      Width           =   4695
      Begin VB.CommandButton cmddesfaz 
         Caption         =   "&Desfaz"
         Enabled         =   0   'False
         Height          =   540
         Left            =   3105
         Picture         =   "Frmtipov.frx":1819
         Style           =   1  'Graphical
         TabIndex        =   13
         Tag             =   "&Update"
         Top             =   210
         Width           =   615
      End
      Begin VB.CommandButton CmdSair 
         Caption         =   "&Sair"
         Height          =   540
         Left            =   3795
         Picture         =   "Frmtipov.frx":1913
         Style           =   1  'Graphical
         TabIndex        =   12
         Tag             =   "&Update"
         Top             =   210
         Width           =   615
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Incluir"
         Height          =   540
         Left            =   345
         Picture         =   "Frmtipov.frx":1A0D
         Style           =   1  'Graphical
         TabIndex        =   11
         Tag             =   "&Add"
         Top             =   210
         Width           =   615
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Excluir"
         Height          =   540
         Left            =   1725
         Picture         =   "Frmtipov.frx":1AF7
         Style           =   1  'Graphical
         TabIndex        =   10
         Tag             =   "&Delete"
         Top             =   210
         Width           =   615
      End
      Begin VB.CommandButton cmdEditar 
         Caption         =   "&Editar"
         Height          =   540
         Left            =   1035
         Picture         =   "Frmtipov.frx":1C69
         Style           =   1  'Graphical
         TabIndex        =   9
         Tag             =   "&Refresh"
         Top             =   210
         Width           =   615
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Salvar"
         Enabled         =   0   'False
         Height          =   540
         Left            =   2415
         Picture         =   "Frmtipov.frx":1DDB
         Style           =   1  'Graphical
         TabIndex        =   8
         Tag             =   "&Update"
         Top             =   225
         Width           =   615
      End
   End
   Begin VB.Label LblEspecial 
      Alignment       =   1  'Right Justify
      Caption         =   "Especial:"
      Height          =   255
      Index           =   0
      Left            =   180
      TabIndex        =   17
      Tag             =   "NOME:"
      Top             =   1335
      Width           =   795
   End
   Begin VB.Label LblParcelas 
      Caption         =   "Parcelas"
      Height          =   195
      Left            =   5280
      TabIndex        =   16
      Top             =   945
      Width           =   645
   End
   Begin VB.Label Lbldias 
      Caption         =   "Dias entre parcelas:"
      Height          =   240
      Left            =   1815
      TabIndex        =   15
      Top             =   930
      Width           =   1590
   End
   Begin VB.Label LblEntrada 
      Caption         =   "Entrada?:"
      Height          =   210
      Left            =   285
      TabIndex        =   14
      Top             =   930
      Width           =   690
   End
   Begin VB.Label Lbltipovend 
      Caption         =   "tipovenda"
      ForeColor       =   &H000000FF&
      Height          =   270
      Left            =   1200
      TabIndex        =   5
      Top             =   60
      Width           =   975
   End
   Begin VB.Label LblDescricao 
      Alignment       =   1  'Right Justify
      Caption         =   "Descrição:"
      Height          =   255
      Index           =   1
      Left            =   60
      TabIndex        =   6
      Tag             =   "NOME:"
      Top             =   420
      Width           =   915
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Código:"
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Tag             =   "CODVEND:"
      Top             =   75
      Width           =   975
   End
End
Attribute VB_Name = "Frmtipovend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private rst As New ADODB.Recordset
Private cSql As String
Private lIncluir As Boolean
Private lPrimeiro As Boolean
Private vRegAtual As Variant
Private nItem As Integer
Private Sub Carrega_Grid()

  With rst
      .MoveLast
      nItem = .RecordCount
      .MoveFirst
      
      vaSpr1.MaxRows = 0: vaSpr1.MaxRows = nItem
      
      nItem = 0
      
      While Not .EOF
         nItem = nItem + 1
        
         vaSpr1.Row = nItem
         vaSpr1.Col = 1: vaSpr1.Text = fuNulo(!tipovenda, "")
         vaSpr1.Col = 2: vaSpr1.Text = fuNulo(!descricao, "")
         vaSpr1.Col = 3: vaSpr1.Text = fuNulo(!entrada, "")
         vaSpr1.Col = 4: vaSpr1.Text = fuNulo(!dias, "")
         vaSpr1.Col = 5: vaSpr1.Text = fuNulo(!parcelas, "")
         vaSpr1.Col = 6: vaSpr1.Text = fuNulo(!especial, "")
         .MoveNext
     Wend
     vaSpr1.MaxRows = nItem
         
  End With
End Sub

Private Sub Carrega_tela()
   limpa_tela Me
   Frmtipovend.Lbltipovend.Caption = rst("tipovenda")
   If Not IsNull(rst("descricao")) Then Frmtipovend.TxtDescricao.Text = rst("descricao")
   If Not IsNull(rst("entrada")) Then Frmtipovend.TxtEntrada.Text = rst("entrada")
   If Not IsNull(rst("dias")) Then Frmtipovend.Txtdias.Text = rst("dias")
   If Not IsNull(rst("parcelas")) Then Frmtipovend.TxtParcelas.Text = rst("parcelas")
   If Not IsNull(rst("especial")) Then Frmtipovend.TxtEspecial.Text = rst("especial")
 
   
End Sub
Private Sub cmdAdd_Click()

   Frmtipovend.Lbltipovend.Caption = ""
   limpa_tela Me
   Frmtipovend.TxtDescricao.SetFocus
   Me.cmdUpdate.Enabled = True
   Me.cmddesfaz.Enabled = True
   Me.cmdEditar.Enabled = False
   Me.cmdAdd.Enabled = False
   Me.CmdSair.Enabled = False
   Me.cmdDelete.Enabled = False
   
   lIncluir = True
End Sub

Private Sub cmdDelete_Click()
    'this may produce an error if you delete the last
    'record or the only record in the recordset
    If MsgBox("Deseja realmente apagar este Tipo de venda ? ", vbYesNo, "Atenção") = vbYes Then
        rst.Close
        cSql = "DELETE FROM tipovend WHERE tipovend.tipovenda = " _
                & Me.Lbltipovend.Caption & " AND tipovend.descricao = '" & Me.TxtDescricao.Text & "'"
        cnn.Execute cSql
        On Error GoTo ErroDelete
        Abre_Le_rst
        rst.MoveFirst
              
        Carrega_tela
        Desabilita Me
        Carrega_Grid
     End If
     Exit Sub
     
ErroDelete:
     MsgBox "Deu erro na exclusao do Tipo de Venda" & Chr(13) & "Instrucao Sql = '" & _
            cSql & "'  "
End Sub

Private Sub cmddesfaz_Click()
  
  lIncluir = False
  Desabilita Me
   
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
 
End Sub

Private Sub CmdSair_Click()
   Unload Me
End Sub

Private Sub cmdUpdate_Click()
   If lIncluir Then
      rst.Close
      cSql = "INSERT INTO tipovend (descricao,entrada,dias,parcelas,especial,operador,datatual) " & _
                          "VALUES ( '" & Me.TxtDescricao.Text & "','" & _
                                         Me.TxtEntrada.Text & "','" & _
                                         Me.Txtdias.Text & "','" & _
                                         Me.TxtParcelas.Text & "','" & _
                                         Me.TxtEspecial.Text & "','" & _
                                         gOperador & "','" & _
                                         Now & "')"
      cnn.Execute cSql
                          
      lIncluir = False
   Else
      rst.Close
      cSql = "UPDATE tipovend SET descricao = '" & Me.TxtDescricao.Text & "'," & _
                                " Entrada = '" & Me.TxtEntrada.Text & "'," & _
                                " Dias = '" & Me.Txtdias.Text & "'," & _
                                " Parcelas = '" & Me.TxtParcelas.Text & "'," & _
                                " Especial = '" & Me.TxtEspecial.Text & "'," & _
                                " operador = '" & gOperador & "'," & _
                                " datatual = '" & Now & "'" & _
                                " WHERE tipovend.tipovenda = " & CLng(Me.Lbltipovend.Caption)
      cnn.Execute cSql
   
      lPrimeiro = False
   End If
   
   Abre_Le_rst
   
   Carrega_tela
   
   Desabilita Me
   
   Me.cmdUpdate.Enabled = False
   Me.cmddesfaz.Enabled = False
   Me.cmdEditar.Enabled = True
   Me.cmdAdd.Enabled = True
   Me.CmdSair.Enabled = True
   Me.cmdDelete.Enabled = True
   
   Carrega_Grid
           
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
     
End Sub

Private Sub Form_Load()
  
    'Abrindo o Recordset
    Abre_Le_rst
    
   'Centraliza a tela no video
   Me.Move (Screen.Width - Me.Width) / 2, _
           (Screen.Height - Me.Height) / 2
 
   Me.Lbltipovend.Caption = ""
   
    If rst.BOF And rst.EOF Then
      If MsgBox("Arquivo vazio. Incluir dados agora ?", vbYesNo, "Atenção ") = vbYes Then
         'rst.AddNew
         With rst
           .AddNew
           !descricao = ""
           .Update
         End With
         cmdEditar_Click
         lPrimeiro = True
      Else
         Desabilita Me
      End If
   
  Else
      rst.MoveFirst
      Carrega_tela
    
      Desabilita Me
      lIncluir = False
      lPrimeiro = False
   End If
   Carrega_Grid
   
End Sub
Private Sub Abre_Le_rst()
   cSql = "select * from tipovend"
   rst.Open cSql, cnn, adOpenKeyset, adLockOptimistic, adCmdText
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Screen.MousePointer = vbDefault
    rst.Close
   
End Sub



