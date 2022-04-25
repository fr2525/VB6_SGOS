Attribute VB_Name = "Module2"
Option Explicit

'--> Para fazer a pesquisa incremental no Combo
#If Win32 Then
    Declare Function SendMessage Lib "User32" Alias "SendMessageA" _
        (ByVal hwnd As Long, ByVal wMsg As Long, _
         ByVal wParam As Long, lParam As Any) As Long
#Else
    Declare Function SendMessage Lib "User" _
        (ByVal hwnd As Integer, ByVal wMsg As Integer, _
         ByVal wParam As Integer, lParam As Any) As Long
#End If

Declare Function GetPrivateProfileString Lib "Kernel32" Alias "GetPrivateProfileStringA" _
         (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, _
          ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Declare Function WritePrivateProfileString Lib "Kernel32" Alias "WritePrivateProfileStringA" _
         (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, _
          ByVal lpFileName As String) As Long

Private Declare Function GetModuleFileName Lib "Kernel32" _
         Alias "GetModuleFileNameA" _
         (ByVal hModule As Long, _
         ByVal lpFileName As String, _
         ByVal nSize As Long) As Long

Public Declare Function ShellExecute Lib "shell32.dll" _
        Alias "ShellExecuteA" _
        (ByVal hwnd As Long, _
        ByVal lpOperation As String, _
        ByVal lpFile As String, _
        ByVal lpParameters As String, _
        ByVal lpDirectory As String, _
        ByVal nShowCmd As Long) As Long

Public Declare Function WinExec Lib "Kernel32" _
        (ByVal lpCmdLine As String, _
        ByVal nCmdShow As Long) As Long

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

Public fMainForm As frmMain

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
       
    If UCase(Left(Right(Strfilename, 7), 3)) = "VB6" Then   'Verifica se é o VB6.EXE
        gDesenv = True
    Else
        gDesenv = False
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

    gPathRel = ReadINI("Geral", "Path_rel", App.Path & "\sgOS.ini")
    
    Set gRs = New ADODB.Recordset
    gRs.CursorType = adOpenForwardOnly
    gRs.CursorLocation = adUseServer
    gRs.LockType = adLockReadOnly
   
    Dim fLogin As New frmLogin
    
    'String para SQLite ------------------
    '*****************************************************
    '*
    'ConDb.Open "DRIVER=SQLite3 ODBC Driver;Database=" & App.Path & "\dados\DbClube.db;LongNames=0;Timeout=1000;NoTXN=0;SyncPragma=NORMAL;StepAPI=0;"
    '*
    '*****************************************************
    '
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
    'ConDb.Open "FILEDSN=" & App.Path & "\dbsgOS.dsn;UID=root;PWD=oyster;"
    Call sConectaBanco
    'ConDb.Open "Driver={MySQL ODBC 5.3 ANSI Driver};Server=localhost:3306;Database=DbSGOS;uid=root;pwd=oyster"
    gOperador = "Master"
    gnCodOperador = 99
    gNivel = 1
    gNome = "Reinert Informática"
    
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
    If gRs.BOF And gRs.EOF Then
        If MsgBox("Necessário cadastrar o cliente, deseja fazer isso agora?", vbYesNo, "Sistema de ordens de serviço") = vbYes Then
            gRs.Close
            ConDb.Close
            frmlojas.Show vbModal
            Exit Sub
        Else
            End
        End If
    End If
    gNome = gRs!nome
    gSenha = gRs!senha
    
    'gAtualizado = gRs!Atualizad
    gCupom = gRs!divcupom
    gMensagem1 = f_nulo(gRs!Mensagem1, "")
    gMensagem2 = f_nulo(gRs!Mensagem1, "")
    
    gRs.Close
   
    '
    If gSenha = "S" Then
       fLogin.Show vbModal
       If Not fLogin.OK Then
          'Login Failed so exit app
          End
       End If
       Unload fLogin
    End If
    '
    'frmSplash.lblCompanyProduct.Caption = gNome
    'frmSplash.Show
    'frmSplash.Refresh
    Screen.MousePointer = vbNormal
    Set fMainForm = New frmMain
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
    
    'Unload frmSplash
    'fmainform.Picture = app.path & "\fotos\imagem.bmp"
    fMainForm.Caption = gNome
    fMainForm.Show
End Sub
'---------------------------------------------
'**** Exemplo de como chamar programa externo ----
'Private Sub Bar_Menu_ButtonClick(ByVal Button As MSComctlLib.Button)
'
'    Dim ret%
'    Dim param$
'    Dim Pos%
'    Dim permitir As Boolean
'
'    param$ = Fu_Formata_Parametros_Chamada_Executavel
'
'    On Error Resume Next
'
'    If Usuario$ <> "ADMNET" And Usuario$ <> "ADMCIT" And Usuario$ <> "BERK01" And Usuario$ <> "ADMSEGUR" And TipoUsuario$ <> "MASTER" Or (InStr(UCase(empresa$), "AXA") <> 0 And (TipoUsuario$ = "MASTER CIT" Or TipoUsuario$ = "MASTER CITNET")) Then
'
'        permitir = False
'        Select Case Button.Key
'
'            Case "Mnu_CITTAB"
'
'                If UCase(TipoUsuario$) <> "SUBSCRITOR" Or UCase(empresa) <> "SULAMERICA" Then
'                    If Fu_Verfica_Permissao("TTFTB001") = True Then
'                        ret% = Shell(DirArq$ & "\CITTAB.EXE " & param$, 1)
'                    End If
'                Else
'                    MsgBox "Módulo não disponível para este perfil de usuario.", 64, "CIT - Controle Integrado de Transportes"
'                End If
'
'            Case "Mnu_CITINT"
'                If Fu_Verfica_Permissao("TTFCI001") = True Then
'                    ret% = Shell(DirArq$ & "\CITINT.EXE " & param$, 1)
'                End If
'
'            Case "Mnu_CITNAC"
'
'                If Fu_Verfica_Permissao("TTFCN001") = True Then
'                    ret% = Shell(DirArq$ & "\CITNAC.EXE " & param$, 1)
'
'                End If
'
'            Case "Mnu_CITAJT"
'                If Fu_Verfica_Permissao("TTFCN001") = True Or Fu_Verfica_Permissao("TTFCI001") = True Then
'                    If InStr(UCase(empresa$), "CHUBB") <> 0 Then
'                        param$ = DirArq$ & "@" & empresa$ & "@" & DirTxt$ & "@" & PrMinProp$ & "@" & VencFat$ & "@" & CalcEnc$ & "@" & CancFat$ & "@" & GeraInterface$ & "@" & LiberaCadEmissao$ & "@" & LiberaApolEmissao$ & "@" & CossegCorret$ & "@" & OBSApoliceAverbacao$ & "@" & CodSeguradora$ & "@" & NomeEmissaoEspecial$ & "@" & Usuario$ & "@" & FilialUsuario% & "@" & TipoUsuario$ & "@" & NomeODBC$ & "@" & UserDB$ & "@" & SenhaDB$ & "@" & Acesso$ & "@" & Versao$ & "@"
'                    End If
'                    permitir = True
'                End If
'
'                If InStr(UCase(empresa$), "XL") <> 0 Then permitir = True
'
'                If permitir = True Then
'                    ret% = Shell(DirArq$ & "\CITAJT.EXE " & param$, 1)
'                Else
'                    MsgBox "Módulo não disponível para este perfil de usuario.", 64, "CIT - Controle Integrado de Transportes"
'                End If
'
'            Case "Mnu_CITEXP" '
'
'                If UCase(TipoUsuario$) <> "SUBSCRITOR" Or UCase(empresa) <> "SULAMERICA" Then
'                    If Fu_Verfica_Permissao("TTFEX101") = True Then
'                        ret% = Shell(DirArq$ & "\CITEXP.EXE " & param$, 1)
'                    End If
'                Else
'                    MsgBox "Módulo não disponível para este perfil de usuario.", 64, "CIT - Controle Integrado de Transportes"
'                End If
'
'            Case "Mnu_CITIMP"
'                If UCase(TipoUsuario$) <> "SUBSCRITOR" Or UCase(empresa) <> "SULAMERICA" Then
'                    If Fu_Verfica_Permissao("TTFIM101") = True Then
'                        ret% = Shell(DirArq$ & "\CITIMP.EXE " & param$, 1)
'                    End If
'                Else
'                    MsgBox "Módulo não disponível para este perfil de usuario.", 64, "CIT - Controle Integrado de Transportes"
'                End If
'            Case "Mnu_CITCVR"
'                '06/06/2018 Itamar Alterado para ACE e Chubb
'                If UCase(TipoUsuario$) <> "SUBSCRITOR" Or UCase(empresa) <> "SULAMERICA" Then
'                    If Fu_Verfica_Permissao("TTFCV001") = True Then
'                        permitir = True
'                    End If
'                End If
'
'                '06/06/2018 Itamar Para subscritor não permite
'                If InStr(UCase(empresa$), "CHUBB") <> 0 Or InStr(UCase(empresa$), "ACE") <> 0 Then
'                    If TipoUsuario$ = "SUBSCRITOR" Then
'                        permitir = False
'                    End If
'                End If
'
'                If permitir = True Then
'                    ret% = Shell(DirArq$ & "\CITCVR.EXE " & param$, 1)
'                Else
'                    MsgBox "Módulo não disponível para este perfil de usuario.", 64, "CIT - Controle Integrado de Transportes"
'                End If
'            Case "Mnu_CITSTR"
'                If UCase(TipoUsuario$) <> "SUBSCRITOR" Or UCase(empresa) <> "SULAMERICA" Then
'                    If Fu_Verfica_Permissao("TTFST001") = True Then
'                        permitir = True
'                    End If
'                End If
'
'                If permitir = True Then
'                    ret% = Shell(DirArq$ & "\CITSTR.EXE " & param$, 1)
'                Else
'                    MsgBox "Módulo não disponível para este perfil de usuario.", 64, "CIT - Controle Integrado de Transportes"
'                End If
'
'            '10/01/2018 - Dyogo - Alterada a Liberação do Modulo da SulAmerica para AXA
'            Case "Mnu_CITLIB"
'                If InStr(UCase(empresa$), "HDI") <> 0 Or InStr(UCase(empresa$), "HANNOVER") <> 0 Or InStr(UCase(empresa$), "GLOBAL") <> 0 Or _
'                   InStr(UCase(empresa$), "TOKIO") <> 0 Or InStr(UCase(empresa$), "ACE") <> 0 Or _
'                   InStr(UCase(empresa$), "LIBERTY") <> 0 Or InStr(UCase(empresa$), "FAIRFAX") <> 0 _
'                   Or InStr(UCase(empresa$), "SIMETRIAS") <> 0 Or InStr(UCase(empresa$), "MITSUI") <> 0 _
'                   Or InStr(UCase(empresa$), "BERKLEY") <> 0 Or InStr(UCase(empresa$), "MAPFRE") <> 0 _
'                   Or InStr(UCase(empresa$), "AIG") <> 0 Or InStr(UCase(empresa$), "PORTO") <> 0 _
'                   Or InStr(UCase(empresa$), "SWISSRE") <> 0 Or InStr(UCase(empresa$), "XL") <> 0 _
'                   Or InStr(UCase(empresa$), "ARGO") <> 0 Or InStr(UCase(empresa$), "GENERALI") <> 0 Or InStr(UCase(empresa$), "CHUBB") <> 0 _
'                   Or InStr(UCase(empresa$), "BRADESCO") <> 0 Or InStr(UCase(empresa$), "STARR") <> 0 Or InStr(UCase(empresa$), "EZZE") <> 0 Or InStr(UCase(empresa$), "AUSTRAL") <> 0 Or InStr(UCase(empresa$), "ALBATROZ") <> 0 Then
'
'                    If InStr(UCase(empresa$), "HDI") <> 0 Or InStr(UCase(empresa$), "GLOBAL") <> 0 Then
'                        If TipoUsuario$ <> "GESTOR" And TipoUsuario$ <> "PRODUTO" Then
'                            MsgBox "Usuário com acesso negado a este módulo.", vbCritical, "CIT - Controle Integrado se Transportes"
'                        Else
'                            ret% = Shell(DirArq$ & "\CITLIB.EXE " & param$, 1)
'                        End If
'                    Else
'                        If TipoUsuario$ <> "ADMINISTRADOR" Then
'                            ret% = Shell(DirArq$ & "\CITLIB.EXE " & param$, 1)
'                        Else
'                            MsgBox "Usuário com acesso negado a este módulo.", vbCritical, "CIT - Controle Integrado se Transportes"
'                        End If
'                    End If
'                Else
'                    MsgBox "Módulo não disponível para a Seguradora.", 64, "CIT - Controle Integrado de Transportes"
'                End If
'
'            Case "Mnu_CITITF"
'                If UCase(TipoUsuario$) <> "COMPLIANCE" And UCase(TipoUsuario$) <> "ADMINISTRADOR" Then
'                    If UCase(TipoUsuario$) <> "SUBSCRITOR" Or UCase(empresa) <> "SULAMERICA" Then
'                        If InStr(UCase(empresa$), "HDI") <> 0 Or InStr(UCase(empresa$), "HANNOVER") <> 0 Or InStr(UCase(empresa$), "GLOBAL") <> 0 Or _
'                           InStr(UCase(empresa$), "TOKIO") <> 0 Or InStr(UCase(empresa$), "ACE") <> 0 Or _
'                           InStr(UCase(empresa$), "MAPFRE") <> 0 Or InStr(UCase(empresa$), "ZURICH") <> 0 Or _
'                           InStr(UCase(empresa$), "CHUBB") <> 0 Or InStr(UCase(empresa$), "MITSUI") <> 0 Or _
'                           InStr(UCase(empresa$), "FAIRFAX") <> 0 Or InStr(UCase(empresa$), "MARITIMA") <> 0 Or InStr(UCase(empresa$), "SULAMERICA") <> 0 Or _
'                           InStr(UCase(empresa$), "SIMETRIAS") <> 0 Or InStr(UCase(empresa$), "BERKLEY") <> 0 Or _
'                           InStr(UCase(empresa$), "AIG") <> 0 Or InStr(UCase(empresa$), "ARGO") <> 0 Or InStr(UCase(empresa$), "PORTO") <> 0 Or _
'                           InStr(UCase(empresa$), "JMALUCELLI") <> 0 Or InStr(UCase(empresa$), "STARR") <> 0 Or _
'                           InStr(UCase(empresa$), "QBE") <> 0 Or InStr(UCase(empresa$), "LIBERTY") <> 0 Or InStr(UCase(empresa$), "EZZE") <> 0 Or InStr(UCase(empresa$), "AUSTRAL") <> 0 Or InStr(UCase(empresa$), "ALBATROZ") <> 0 Then
'
'                           If InStr(UCase(empresa$), "SULAMERICA") <> 0 And UCase(TipoUsuario) = "CONSULTA GERAL" Then
'                                ret = 0
'                                MsgBox "Módulo de Interface indisponível para seu perfil de usuário.", 64, "CIT - Controle Integrado de Transportes"
'                           Else
'                                '06/06/2018 Itamar
'                                If InStr(UCase(empresa$), "CHUBB") <> 0 Or InStr(UCase(empresa$), "ACE") <> 0 Or InStr(UCase(empresa$), "ALBATROZ") <> 0 Then
'                                    If TipoUsuario$ = "SUBSCRITOR" Then
'                                        MsgBox "Módulo não disponível para a Seguradora.", 64, "CIT - Controle Integrado de Transportes"
'                                    Else
'                                        ret% = Shell(DirArq$ & "\CITITF.EXE " & param$, 1)
'                                    End If
'                                Else
'                                    ret% = Shell(DirArq$ & "\CITITF.EXE " & param$, 1)
'                                End If
'                           End If
'
'                        Else
'                            MsgBox "Módulo não disponível para a Seguradora.", 64, "CIT - Controle Integrado de Transportes"
'                        End If
'                    End If
'                Else
'                    MsgBox "Usuário não possui permissão para esse módulo.", 64, "CIT - Controle Integrado de Transportes"
'                End If
'
'            Case "Mnu_CITGES"
'                If UCase(TipoUsuario$) <> "SUBSCRITOR" Or UCase(empresa) <> "SULAMERICA" Then
'                    If InStr(UCase(empresa$), "TOKIO") <> 0 Or InStr(UCase(empresa$), "CHUBB") <> 0 Or _
'                       InStr(UCase(empresa$), "ACE") <> 0 Or InStr(UCase(empresa$), "SIMETRIAS") <> 0 Or _
'                       InStr(UCase(empresa$), "FAIRFAX") <> 0 Or InStr(UCase(empresa$), "BERKLEY") <> 0 Or _
'                       InStr(UCase(empresa$), "ZURICH") <> 0 Or InStr(UCase(empresa$), "MAPFRE") <> 0 Or _
'                       InStr(UCase(empresa$), "SULAMERICA") <> 0 Or InStr(UCase(empresa$), "XL") <> 0 Or _
'                       InStr(UCase(empresa$), "ARGO") <> 0 Or InStr(UCase(empresa$), "LIBERTY") <> 0 Or _
'                       InStr(UCase(empresa$), "BRADESCO") <> 0 Or InStr(UCase(empresa$), "GLOBAL") <> 0 Or _
'                       InStr(UCase(empresa$), "GENERALI") <> 0 Or InStr(UCase(empresa$), "SWISSRE") <> 0 Or _
'                       InStr(UCase(empresa$), "UBF") <> 0 Or InStr(UCase(empresa$), "STARR") <> 0 Or _
'                       InStr(UCase(empresa$), "ALIANÇA") <> 0 Or InStr(UCase(empresa$), "AXA") <> 0 Or _
'                       InStr(UCase(empresa$), "AIG") <> 0 Or InStr(UCase(empresa$), "EZZE") <> 0 Or InStr(UCase(empresa$), "AUSTRAL") <> 0 Or InStr(UCase(empresa$), "ALBATROZ") <> 0 Then
'                       '23/04/2018 Itamar Liberado para o Subscritor da ACE e CHUBB
'                      If InStr(UCase(empresa$), "ACE") <> 0 Or InStr(UCase(empresa$), "CHUBB") <> 0 And _
'                         (TipoUsuario$ = "SUBSCRITOR" Or TipoUsuario$ = "GESTOR" Or TipoUsuario$ = "PRODUTO" Or TipoUsuario = "FILIAL" Or TipoUsuario = "FILIAL/CONSULTA" Or TipoUsuario$ = "CONSULTA GERAL" Or TipoUsuario$ = "EMISSAO" Or TipoUsuario$ = "EMISSÃO GESTOR" Or TipoUsuario$ = "INFORMAÇÕES GERAIS") Then
'                            ret% = Shell(DirArq$ & "\CITGES.EXE " & param$, 1)
'
'                        ElseIf InStr(UCase(empresa$), "ARGO") <> 0 And (TipoUsuario$ = "EMISSAO") Then
'
'                            ret% = Shell(DirArq$ & "\CITGES.EXE " & param$, 1)
'
'                        ElseIf TipoUsuario$ = "GESTOR" Or TipoUsuario$ = "PRODUTO" Or TipoUsuario = "FILIAL" Or TipoUsuario = "FILIAL/CONSULTA" Then
'
'                            ret% = Shell(DirArq$ & "\CITGES.EXE " & param$, 1)
'
'                        Else
'                            MsgBox "Módulo não disponível para usuário do tipo " & TipoUsuario$ & ".", 64, "CIT - Controle Integrado de Transportes"
'                        End If
'                    Else
'                        MsgBox "Módulo não disponível para a Seguradora.", 64, "CIT - Controle Integrado de Transportes"
'                    End If
'                Else
'                    MsgBox "Módulo não disponível para este perfil de usuario.", 64, "CIT - Controle Integrado de Transportes"
'                End If
'
'            Case "Mnu_CITAVU"
'                If InStr(UCase(empresa$), "MITSUI") <> 0 Or InStr(UCase(empresa$), "TOKIO") <> 0 Or InStr(UCase(empresa$), "MAPFRE") <> 0 Or InStr(UCase(empresa$), "SULAMERICA") <> 0 Then
'                    If TipoUsuario$ = "GESTOR" Or TipoUsuario$ = "PRODUTO" Then
'                        permitir = True
'                    End If
'                End If
'
'                If permitir = True Then
'                    ret% = Shell(DirArq$ & "\CITAVU.EXE " & param$, 1)
'                Else
'                    MsgBox "Módulo não disponível para a Seguradora.", 64, "CIT - Controle Integrado de Transportes"
'                End If
'
'            Case "Mnu_CITUSU"
'                '27/04/2018 Itamar
'                If InStr(UCase(empresa$), "CHUBB") <> 0 Or InStr(UCase(empresa$), "ACE") <> 0 Then
'                    If TipoUsuario$ = "GESTOR" Then
'                        ret% = Shell(DirArq$ & "\CITUSU.EXE " & param$, 1)
'                    Else
'                        MsgBox "Módulo de Segurança disponível apenas para senha do Administrador do Sistema.", 64, "CIT - Controle Integrado de Transportes"
'                    End If
'                'sandra venturim
'                ElseIf Ind_grp_admin% = True And TipoUsuario$ = "ADMINISTRADOR" Then
'                    ret% = Shell(DirArq$ & "\CITUSU.EXE " & param$, 1)
'                '07/06/2018 - Andreia Ferraraccio - jira 4094
'                ElseIf InStr(UCase(empresa$), "AXA") <> 0 And (TipoUsuario$ = "MASTER CIT" Or TipoUsuario$ = "MASTER CITNET" Or TipoUsuario$ = "ADMCIT") Then
'                    ret% = Shell(DirArq$ & "\CITUSU.EXE " & param$, 1)
'                ElseIf InStr(UCase(empresa$), "FAIRFAX") <> 0 And TipoUsuario$ = "ADMINISTRADOR" Then
'                    ret% = Shell(DirArq$ & "\CITUSU.EXE " & param$, 1)
'                Else
'                    MsgBox "Módulo de Segurança disponível apenas para senha do Administrador do Sistema.", 64, "CIT - Controle Integrado de Transportes"
'                End If
'
'            Case "Mnu_IMPRE"
'                '08/02/2018 - Dyogo - Liberação do Modulo para ACE e CHUBB
'                '02/03/2018 - Dyogo - Bloqueio do Modulo para AXA
'                '19/09/2018 - Itamar Liberado para a Liberty
'                If InStr(UCase(empresa$), "CHUBB") <> 0 Or InStr(UCase(empresa$), "ACE") <> 0 Or InStr(UCase(empresa$), "XL") <> 0 Or InStr(UCase(empresa$), "LIBERTY") <> 0 Or InStr(UCase(empresa$), "SWISSRE") Then
'                    If TipoUsuario$ = "GESTOR" Or TipoUsuario$ = "PRODUTO" Then
'                        ret% = Shell(DirArq$ & "\TTPRE001.EXE " & param$, 1)
'                    Else
'                        MsgBox "Módulo não disponível para este perfil de usuario.", 64, "CIT - Controle Integrado de Transportes"
'                    End If
'                Else
'                    MsgBox "Módulo não disponível para a Seguradora.", 64, "CIT - Controle Integrado de Transportes"
'                End If
'
'            Case "Mnu_Sair"
'                If InStr(UCase(empresa$), "AIG") <> 0 Then  'CIT-3241-[16mai2017]
'                    If MsgBox("Todos os Módulos do CIT serão encerrados, as operações em andamento serão interrompidas, gerando dados incompletos  e/ou  corrompidos na base de dados. " & Chr(13) & "Deseja encerrar o CIT ? ", vbYesNo, "CIT - Controle Integrado de Transportes ") <> 6 Then
'                        Exit Sub
'                    End If
'                Else
'                    If MsgBox("Deseja encerrar o CIT ? ", vbYesNo, "CIT - Controle Integrado de Transportes ") <> 6 Then
'                        Exit Sub
'                    End If
'                End If
'                Unload Me
'                End
'        End Select
'    Else
'        Select Case Button.Key
'
'            Case "Mnu_CITUSU"
'                ret% = Shell(DirArq$ & "\CITUSU.EXE " & param$, 1)
'
'            Case "Mnu_Sair"
'                Unload Me
'                End
'
'            Case Else
'                MsgBox "Usuário do Administrador inválido para a opção selecionada.", 16, "CIT - Controle Integrado de Transportes"
'
'        End Select
'    End If
'
'    On Error GoTo 0
'
'End Sub


