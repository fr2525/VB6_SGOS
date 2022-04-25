Attribute VB_Name = "Module1"
Option Explicit

'--> Para fazer a pesquisa incremental no Combo
#If Win32 Then
    Declare Function SendMessage Lib "User32" Alias "SendMessageA" _
        (ByVal hWnd As Long, ByVal wMsg As Long, _
         ByVal wParam As Long, lParam As Any) As Long
#Else
    Declare Function SendMessage Lib "User" _
        (ByVal hWnd As Integer, ByVal wMsg As Integer, _
         ByVal wParam As Integer, lParam As Any) As Long
#End If

Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" _
         (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, _
          ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" _
         (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, _
          ByVal lpFileName As String) As Long

Private Declare Function GetModuleFileName Lib "kernel32" _
         Alias "GetModuleFileNameA" _
         (ByVal hModule As Long, _
         ByVal lpFileName As String, _
         ByVal nSize As Long) As Long

Global ConDb As New ADODB.Connection
Global gCmd As New Command

'Global gWk As Workspace
'Global gDb As Database
Global gRs As ADODB.Recordset

'Global gRsTemp As Recordset
'*------- Variavel indicando se é desenvolvimento ou produção
Global gDesenv As Boolean

Global gsDatabase As String
Global gsConnect As String
Global gsRecordsource As String
Global gOperador As String
Global gnCodOperador As Integer

Global gSql As String
Global gNivel As Integer

Global gNome As String
Global gSenha As String
Global gImpresso As String
Global gAtualizado As String
Global gCupom As String
Global gMensagem1 As String
Global gMensagem2 As String
Global gPalavra As String
'
Global gCodObra As Double
Global gUnidade As String
Global gNumMedicao As Double
Global gDataMedicao As Date
Global gCodProd As String
Global gnSequencia As Integer
Global gnCodcli As Double
Global gnTotPed As Double
Global gnAPrazo As Boolean
Global gNomeRelato As String
Global gSelecao As String
Global gPathRel As String

'
'Global imp_12cp As String
'Global imp_10cp As String
'Global imp_acom As String
'Global imp_dcom As String
'Global imp_aneg As String
'Global imp_dneg As String
'Global imp_acex As String
'Global imp_dcex As String
'Global imp_aexp As String
'Global imp_dexp As String
'Global imp_asub As String
'Global imp_dsub As String
'Global imp_norm As String
''
'imp_12cp = vbKeyEscape + "M"
'imp_10cp = Chr(27) & "P"
'imp_acom = Chr(15)
'imp_dcom = Chr(18)
'imp_aneg = "G"
'imp_dneg = "H"
'imp_acex = "W"
'imp_dcex = "W" + Chr(0)
'imp_aexp = "W"
'imp_dexp = "W" + Chr(0)
'imp_asub = "-"
'imp_dsub = "-" + Chr(0)
'imp_norm = imp_10cp + imp_dcom + imp_dneg + imp_dcex + imp_dexp + imp_dsub

Public fMainForm As FrmVendas

Sub Main()
    
    'Variaveis usadas para ver se é versão compilada ou de desenvolvimento
    Dim Strfilename As String
    Dim lngCount As Long
    Dim ResX As Long, ResY As Long

    Dim PicX As Long, Picy As Long, Res As Long
    
    gDesenv = False
    
    Screen.MousePointer = vbHourglass
    
    If App.PrevInstance Then
       MsgBox ("Não pode carregar o programa novamente."), vbExclamation, "A aplicação já está aberta"
       'Unload Me
       End
    End If
 
    ' -> VB WorkShop 2000
    
    Strfilename = String(255, 0)
    lngCount = GetModuleFileName(App.hInstance, Strfilename, 255)   'Pega o nome do programa
    Strfilename = Left(Strfilename, lngCount)    'Retira os espaços do fim do nome
   
    If UCase(Left(Right(Strfilename, 7), 3)) = "VB5" Then   'Verifica se é o VB5.EXE
       gDesenv = True
    Else
       If UCase(Left(Right(Strfilename, 7), 3)) = "VB6" Then   'Verifica se é o VB6.EXE
          gDesenv = True
        Else
          gDesenv = False
          'MsgBox "Versão Compilada - Executável"
    
       End If
    End If
    
    'Verifica a resolução de video
    PicX = Screen.TwipsPerPixelX
    Picy = Screen.TwipsPerPixelY
    ResX = Screen.Width \ PicX
    ResY = Screen.Height \ Picy
    Res = ResX * ResY
    If Res < 480000 Then 'Se a Resolução menor 800 x 600 então
       MsgBox "Resolução de vídeo incompatível. " & vbCr & vbLf & " Altere para no minimo 800 X 600 e reexecute o programa ", vbOKOnly, "Atenção Operador"
       End
    End If

    gPathRel = ReadINI("Geral", "Path_rel", App.Path & "\sgl.ini")
    
    Set gRs = New ADODB.Recordset
    gRs.CursorType = adOpenForwardOnly
    gRs.CursorLocation = adUseServer
    gRs.LockType = adLockReadOnly
   
    Dim fLogin As New frmLogin
    '
    'Set ConDb = New ADODB.Connection
    '***************************************************
    'Pode ser assim -->
    'ConDb.Provider = "Microsoft.Jet.OLEDB.3.51"
    'ConDb.ConnectionString = APP.PATH & "\dados\DBsgl.mdb"
    'ConDb.Open
    '***************************************************
    'Ou assim -->
    'ConDb.Open "Provider = Microsoft.Jet.OLEDB.3.51;Data Source = " & App.Path & "\dados\DBsgl.mdb"
    '
    'FileCopy App.Path & "\dados\dbsmg.mdb", App.Path & "\dados\dbcopia.mdb"
    '
    ConDb.Open "FILEDSN=" & App.Path & "\dbsgl.dsn;UID=admin;PWD=oyster;"
    'ConDb.Open "FILEDSN=C:\ARQUIVOS DE PROGRAMAS\SMG\dbsmg.dsn;UID=admin;PWD=oyster;"
    gOperador = "Master"
    gnCodOperador = 99
    
    '
    '**************************************************
    ' Ou ainda conexao com o Interbase
    '**************************************************
    'conexão com o IBProvider
    'adoConn.Open "provider=LCPI.IBProvider;data source=C:\teste\Employee.gdb;ctype=win1251;user 'id=sysdba;password=masterkey"
    'conexão com o SIBProvider
    ' adoConn.Open "provider=sibprovider;data source=c:\teste\employee.gdb", "sysdba", "masterkey"
    'conexão com o IbOleDb Provider
    'adoConn.Open "Provider=IbOleDb;Data Source=c:\teste\employee.gdb", "sysdba", "masterkey"
    
    '**************************************************
    ' Ou ainda conexao com o SQL SERVER
    '**************************************************
    'ConDb.Provider = "SQLOLEDB"
    'ConDb.Properties("Data Source").Value = "CENUDBDSNV01"
    'ConDb.Properties("integrated security").Value = "SSPI"
    'ConDb.Properties("Initial Catalog").Value = "dbs_poc_opcredito"
    '
    gSql = "select * FROM tab_lojas"
    gRs.Open gSql, ConDb, adOpenKeyset, adLockOptimistic
    '
    gNome = gRs!nome
    gSenha = gRs!senha
    gImpresso = f_nulo(gRs!Impresso, "")
    'gAtualizado = gRs!Atualizad
    gCupom = gRs!divcupom
    gMensagem1 = f_nulo(gRs!Mensagem1, "")
    gMensagem2 = f_nulo(gRs!Mensagem1, "")
    gPalavra = f_nulo(gRs!Palavra, "")
    gRs.Close
    '
    If gSenha Then
       fLogin.Show vbModal
       If Not fLogin.OK Then
          'Login Failed so exit app
          End
       End If
       Unload fLogin
    End If
    '
    frmSplash.lblCompanyProduct.Caption = gNome
    frmSplash.Show
    frmSplash.Refresh
    Screen.MousePointer = vbNormal
    Set fMainForm = New FrmVendas
    Load fMainForm
'    If gNivel > 2 Then
'       fMainForm.MnArquivos.Visible = False
'       fMainForm.mnCancvenda.Visible = False
'       fMainForm.Mnfechacli.Visible = False
'       fMainForm.MnCompras.Visible = False
'       fMainForm.mnoutrasmov.Visible = False
'       fMainForm.mnuApagar.Visible = False
'       fMainForm.mnRelato.Visible = False
'       fMainForm.mnutilitarios.Visible = False
'    End If
    
    Unload frmSplash
    'fmainform.Picture = app.path & "\fotos\imagem.bmp"
    fMainForm.Caption = gNome
    fMainForm.Show
End Sub


