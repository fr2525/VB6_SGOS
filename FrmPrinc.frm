VERSION 5.00
Begin VB.Form FrmMain 
   Caption         =   "Sistema de Gerenciamento de Loja"
   ClientHeight    =   9150
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   14550
   LinkTopic       =   "MdiForm1"
   ScaleHeight     =   9150
   ScaleWidth      =   14550
   WindowState     =   2  'Maximized
   Begin VB.Image Image1 
      Height          =   5115
      Left            =   480
      Picture         =   "FrmPrinc.frx":0000
      Stretch         =   -1  'True
      Top             =   240
      Width           =   14160
   End
   Begin VB.Menu Mn1 
      Caption         =   "&Arquivos"
      Begin VB.Menu Mn11 
         Caption         =   "&Lojas"
      End
      Begin VB.Menu Mn12 
         Caption         =   "&Clientes"
      End
      Begin VB.Menu Mn13 
         Caption         =   "&Fornecedores"
      End
      Begin VB.Menu mn14 
         Caption         =   "&Balconistas"
      End
      Begin VB.Menu Mn15 
         Caption         =   "&Grupos"
      End
      Begin VB.Menu Mn16 
         Caption         =   "&Produtos"
      End
      Begin VB.Menu mn17 
         Caption         =   "&tipos de Movimentação"
      End
      Begin VB.Menu Mn18 
         Caption         =   "Tipos de &Venda"
      End
   End
   Begin VB.Menu Mn2 
      Caption         =   "&Movimentações"
      Begin VB.Menu Mn21 
         Caption         =   "&Vendas"
      End
      Begin VB.Menu Mntraco1 
         Caption         =   "-"
      End
      Begin VB.Menu Mn22 
         Caption         =   "&Compras"
      End
      Begin VB.Menu Mntraco2 
         Caption         =   "-"
      End
      Begin VB.Menu Mn23 
         Caption         =   "Cance&lar Vendas"
      End
      Begin VB.Menu Mn24 
         Caption         =   "Cancelar co&mpras"
      End
      Begin VB.Menu Mn25 
         Caption         =   "&Trocas"
      End
      Begin VB.Menu Mn26 
         Caption         =   "&Outras Movimentações"
      End
      Begin VB.Menu Mn27 
         Caption         =   "&Despesas"
      End
   End
   Begin VB.Menu Mn3 
      Caption         =   "Ctas. a &Pagar"
   End
   Begin VB.Menu MN4 
      Caption         =   "Ctas a &Receber"
   End
   Begin VB.Menu Mn5 
      Caption         =   "Re&latórios"
      Begin VB.Menu Mn51 
         Caption         =   "E&tiquetas de Produtos"
      End
      Begin VB.Menu Mn52 
         Caption         =   "Produtos com Estoque &zero"
      End
      Begin VB.Menu Mn53 
         Caption         =   "&Lista de Preços"
      End
      Begin VB.Menu Mn54 
         Caption         =   "Posição &Física"
      End
      Begin VB.Menu Mn55 
         Caption         =   "Relat.p/ &Inventário "
      End
      Begin VB.Menu traco3 
         Caption         =   "-"
      End
      Begin VB.Menu Mn56 
         Caption         =   "&Vendas"
      End
      Begin VB.Menu Mn57 
         Caption         =   "M&ovimentações por Período"
      End
      Begin VB.Menu mn58 
         Caption         =   "Produtos Mais V&endidos"
      End
      Begin VB.Menu Mn59 
         Caption         =   "Produtos me&nos vendidos"
      End
   End
   Begin VB.Menu Mn6 
      Caption         =   "&Utilitários"
      Begin VB.Menu Mn61 
         Caption         =   "&Reparar Banco de dados"
      End
      Begin VB.Menu Mn62 
         Caption         =   "Cópia de Segurança (Backup)"
      End
   End
   Begin VB.Menu Mn7 
      Caption         =   "A&juda"
      Begin VB.Menu Mn71 
         Caption         =   "So&bre"
      End
   End
   Begin VB.Menu Mn8 
      Caption         =   "&Sair"
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Mn11_Click()
  frmLoja.Show
End Sub

Private Sub Mn12_Click()
   frmClientes.Show
End Sub

Private Sub Mn13_Click()
   frmFornecedores.Show
End Sub

Private Sub mn14_Click()
   frmBalconistas.Show
End Sub

Private Sub Mn15_Click()
   frmGrupos.Show
End Sub

Private Sub Mn16_Click()
   frmProdutos.Show
End Sub

Private Sub mn17_Click()
   frmtipomov.Show
End Sub

Private Sub Mn18_Click()
   frmtipovend.Show
End Sub

Private Sub Mn21_Click()
   frmBalcao.Show
End Sub

Private Sub Mn71_Click()
  frmsobre.Show
End Sub

Private Sub Mn8_Click()
   Unload Me
   End
End Sub
