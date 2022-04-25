VERSION 5.00
Begin VB.UserControl CtlPct 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.PictureBox PctFoto 
      Height          =   3510
      Left            =   60
      ScaleHeight     =   3450
      ScaleWidth      =   4635
      TabIndex        =   0
      Top             =   45
      Width           =   4695
   End
End
Attribute VB_Name = "CtlPct"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private Sub PctFoto_Click()
     FrmCadsoc.Visible = False
     FrmTatoo.Show
     FrmTatoo.PctFoto = FrmCadsoc.PctFoto
     FrmTatoo.Command2.Visible = False
     FrmTatoo.TxtData.Enabled = False
     FrmTatoo.MskValor.Enabled = False
End Sub
