Attribute VB_Name = "Module2"
Option Explicit

Private Sub Carrega_Grid()

'Teste do MsHFlexgrid1 - eh eh eh
  MSFlexGrid1.Row = 0
  
  With rst
      .MoveLast
      nItem = .RecordCount
      .MoveFirst
      MSFlexGrid1.Rows = 1
      If .AbsolutePosition <> -1 Then
         Do While Not .EOF
            MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
            MSFlexGrid1.Row = MSFlexGrid1.Rows - 1
            
            MSFlexGrid1.Col = 0: MSFlexGrid1.Text = f_nulo(!codigo, "")
            MSFlexGrid1.Col = 1: MSFlexGrid1.Text = f_nulo(!nome, "")
            MSFlexGrid1.Col = 2: MSFlexGrid1.Text = f_nulo(!endereco, "")
            MSFlexGrid1.Col = 3: MSFlexGrid1.Text = f_nulo(!bairro, "")
            MSFlexGrid1.Col = 4: MSFlexGrid1.Text = f_nulo(!Cidade, "")
            MSFlexGrid1.Col = 5: MSFlexGrid1.Text = f_nulo(!estado, "")
            MSFlexGrid1.Col = 6: MSFlexGrid1.Text = f_nulo(!cep, "")
            MSFlexGrid1.Col = 7: MSFlexGrid1.Text = f_nulo(!cgccpf, "")
            MSFlexGrid1.Col = 8: MSFlexGrid1.Text = f_nulo(!rg, "")
            MSFlexGrid1.Col = 9: MSFlexGrid1.Text = f_nulo(!Telefone, "")
            MSFlexGrid1.Col = 10: MSFlexGrid1.Text = f_nulo(!celular, "")
            MSFlexGrid1.Col = 11: MSFlexGrid1.Text = f_nulo(!diaAniver, "")
            MSFlexGrid1.Col = 12: MSFlexGrid1.Text = f_nulo(!MesAniver, "")
            MSFlexGrid1.Col = 13: MSFlexGrid1.Text = f_nulo(!AnoAniver, "")
            MSFlexGrid1.Col = 14: MSFlexGrid1.Text = f_nulo(!Ultcompra, "01/01/1901")
            MSFlexGrid1.Col = 15: MSFlexGrid1.Text = f_nulo(!email, "")
            
            .MoveNext
            
          Loop
          MSFlexGrid1.FixedRows = 1
      End If
          
  End With
   

  End Sub
Private Sub Carrega_tela()
   'Limpa as variaveis da tela se caso ficarem com dados da outra tela
   limpa_tela Me
   'Carrega a tela com os dados do registro
   FrmClientes.LblCodclie = rst!codigo
   If Not IsNull(rst!nome) Then Me.TxtNome.Text = rst!nome
   If Not IsNull(rst!endereco) Then Me.TxtEndereco.Text = rst!endereco
   If Not IsNull(rst!bairro) Then Me.TxtBairro.Text = rst!bairro
   If Not IsNull(rst!Cidade) Then Me.TxtCidade.Text = rst!Cidade
   If Not IsNull(rst!estado) Then Me.TxtUf.Text = rst!estado
   If Not IsNull(rst!cep) Then Me.TxtCep.Text = rst!cep
   If Not IsNull(rst!cgccpf) Then Me.Txtcgc_cpf.Text = rst!cgccpf
   If Not IsNull(rst!rg) Then Me.TxtRG.Text = rst!rg
   If Not IsNull(rst!Telefone) Then Me.TxtTelefone.Text = rst!Telefone
   If Not IsNull(rst!celular) Then Me.TxtCelular.Text = rst!celular
   If Not IsNull(rst!diaAniver) Then Me.TxtDiaAniver.Text = rst!diaAniver
   If Not IsNull(rst!MesAniver) Then Me.TxtMesAniver.Text = rst!MesAniver
   If Not IsNull(rst!AnoAniver) Then Me.TxtAnoAniver.Text = rst!AnoAniver
   If Not IsNull(rst!Ultcompra) Then Me.TxtUltimacompra.Text = Format(rst!Ultcompra, "dd/mm/YYYY")
   If Not IsNull(rst!email) Then Me.TxtEmail.Text = rst!email
   
End Sub

Private Sub cmdAdd_Click()

   Me.LblCodclie.Caption = ""
   limpa_tela Me
   Me.TxtNome.SetFocus
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
    If MsgBox("Deseja realmente apagar este Cliente ? ", vbYesNo, "Atenção") = vbYes Then
        rst.Close
        cSql = "DELETE FROM Cadclie WHERE cadclie.codigo = " _
                & Me.LblCodclie.Caption & " AND cadclie.nome = '" & Me.TxtNome.Text & "'"
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
     MsgBox "Deu erro na exclusao do Cliente" & Chr(13) & "Instrucao Sql = '" & _
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
   Me.TxtNome.SetFocus
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
   If Not f_ValidaData(TxtUltimacompra.Text) Then
       Exit Sub
   End If
   
   If lIncluir Then
      rst.Close
      cSql = "INSERT INTO cadclie (nome,endereco,bairro,cidade,estado,cep,cgccpf,rg,telefone,celular,diaaniver,mesaniver,anoaniver,ultcompra,email,operador,datatual) " & _
                          "VALUES ( '" & Me.TxtNome.Text & "','" & _
                                         Me.TxtEndereco.Text & "','" & _
                                         Me.TxtBairro.Text & "','" & _
                                         Me.TxtCidade.Text & "','" & _
                                         Me.TxtUf.Text & "','" & _
                                         Me.TxtCep.Text & "','" & _
                                         Me.Txtcgc_cpf.Text & "','" & _
                                         Me.TxtRG.Text & "','" & _
                                         Me.TxtTelefone.Text & "','" & _
                                         Me.TxtCelular.Text & "','" & _
                                         Me.TxtDiaAniver.Text & "','" & _
                                         Me.TxtMesAniver.Text & "','" & _
                                         Me.TxtAnoAniver.Text & "','" & _
                                         Me.TxtUltimacompra.Text & "','" & _
                                         Me.TxtEmail.Text & "','" & _
                                         gOperador & "','" & _
                                         Now & "')"
      cnn.Execute cSql
                          
      lIncluir = False
   Else
      rst.Close
      cSql = "UPDATE cadclie SET nome = '" & Me.TxtNome.Text & "'," & _
                                " Endereco = '" & Me.TxtEndereco.Text & "'," & _
                                " Bairro = '" & Me.TxtBairro.Text & "'," & _
                                " Cidade = '" & Me.TxtCidade.Text & "'," & _
                                " Estado = '" & Me.TxtUf.Text & "'," & _
                                " CEP = '" & Me.TxtCep.Text & "'," & _
                                " CGCCPF = '" & Me.Txtcgc_cpf.Text & "'," & _
                                " RG = '" & Me.TxtRG.Text & "'," & _
                                " Telefone = '" & Me.TxtTelefone.Text & "'," & _
                                " Celular = '" & Me.TxtCelular.Text & "'," & _
                                " DiaAniver = '" & Me.TxtDiaAniver.Text & "'," & _
                                " Mesaniver = '" & Me.TxtMesAniver.Text & "'," & _
                                " Anoaniver = '" & Me.TxtAnoAniver.Text & "'," & _
                                " ultcompra = '" & Me.TxtUltimacompra.Text & "'," & _
                                " Email = '" & Me.TxtEmail.Text & "'," & _
                                " operador = '" & gOperador & "'," & _
                                " datatual = '" & Now & "'" & _
                                " WHERE cadclie.codigo = " & CLng(Me.LblCodclie.Caption)
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
    
   Abre_Le_rst
  
   'Centraliza a tela no video
   Me.Move (Screen.Width - Me.Width) / 2, _
           (Screen.Height - Me.Height) / 2
 
   Me.LblCodclie.Caption = ""

   If rst.BOF And rst.EOF Then
      If MsgBox("Arquivo vazio. Incluir dados agora ?", vbYesNo, "Atenção ") = vbYes Then
         'rst.AddNew
         With rst
           .AddNew
           !nome = ""
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
   Set qd = db.QueryDefs("MostraTodosClientes")
   Set rst = qd.OpenRecordset
   'cSql = "select * from cadclie"
   'rst.Open cSql, cnn, adOpenKeyset, adLockOptimistic, adCmdText
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Screen.MousePointer = vbDefault
    rst.Close
       
End Sub


Private Sub MSFlexGrid1_Click()
  
  Dim oldrow As Long
  oldrow = MSFlexGrid1.Row
  MSFlexGrid1.Row = 0
  With MSFlexGrid1
    .Redraw = False
    Do While True
       .Row = .Row + 1
       For i = 0 To 15
           .Col = i: .CellBackColor = vbWhite
       Next
       If .Row = .Rows - 1 Then
          Exit Do
       End If
    Loop
  
    .Refresh
    .Row = oldrow
    
    .Col = 0:   LblCodclie.Caption = .Text: .CellBackColor = vbYellow
    .Col = 1:   TxtNome.Text = .Text: .CellBackColor = vbYellow
    .Col = 2:   TxtEndereco.Text = .Text: .CellBackColor = vbYellow
    .Col = 3:   TxtBairro.Text = .Text: .CellBackColor = vbYellow
    .Col = 4:   TxtCidade.Text = .Text: .CellBackColor = vbYellow
    .Col = 5:   TxtUf.Text = .Text: .CellBackColor = vbYellow
    .Col = 6:   TxtCep.Text = .Text: .CellBackColor = vbYellow
    .Col = 7:   Txtcgc_cpf.Text = .Text: .CellBackColor = vbYellow
    .Col = 8:   TxtRG.Text = .Text: .CellBackColor = vbYellow
    .Col = 9:   TxtTelefone.Text = .Text: .CellBackColor = vbYellow
    .Col = 10:   TxtCelular.Text = .Text: .CellBackColor = vbYellow
    .Col = 11:   TxtDiaAniver.Text = Left(.Text, 2): .CellBackColor = vbYellow
    .Col = 12:   TxtMesAniver.Text = Mid(.Text, 4, 2): .CellBackColor = vbYellow
    .Col = 13:   TxtAnoAniver.Text = Right(.Text, 2): .CellBackColor = vbYellow
    .Col = 14:   TxtUltimacompra.Text = .Text: .CellBackColor = vbYellow
    .Col = 15:   TxtEmail.Text = .Text: .CellBackColor = vbYellow
  
    .Redraw = True
    
  End With


  
'    rst.MoveFirst
'    Do While Not rst.EOF
'       If rst!codigo = LblCodigo.Caption Then
'          Exit Do
'       End If
'       rst.MoveNext
'    Loop
End Sub

Private Sub vaSpr1_Click(ByVal Col As Long, ByVal Row As Long)
    Dim borda
    
    vaSpr1.Row = Row
    borda = vaSpr1.BorderStyle
    
    'suMarcaSpr vaSpr1, vaSpr1.Row
    
    vaSpr1.BorderStyle = 1
    vaSpr1.Col = 1:   LblCodclie.Caption = vaSpr1.Text
    vaSpr1.Col = 2:   TxtNome.Text = vaSpr1.Text
    vaSpr1.Col = 3:   TxtEndereco.Text = vaSpr1.Text
    vaSpr1.Col = 4:   TxtBairro.Text = vaSpr1.Text
    vaSpr1.Col = 5:   TxtCidade.Text = vaSpr1.Text
    vaSpr1.Col = 6:   TxtUf.Text = vaSpr1.Text
    vaSpr1.Col = 7:   TxtCep.Text = vaSpr1.Text
    vaSpr1.Col = 8:   Txtcgc_cpf.Text = vaSpr1.Text
    vaSpr1.Col = 9:   TxtRG.Text = vaSpr1.Text
    vaSpr1.Col = 10:   TxtTelefone.Text = vaSpr1.Text
    vaSpr1.Col = 11:   TxtCelular.Text = vaSpr1.Text
    vaSpr1.Col = 12:   TxtDiaAniver.Text = Left(vaSpr1.Text, 2)
    vaSpr1.Col = 13:   TxtMesAniver.Text = Mid(vaSpr1.Text, 4, 2)
    vaSpr1.Col = 14:   TxtAnoAniver.Text = Right(vaSpr1.Text, 2)
    vaSpr1.Col = 15:   TxtUltimacompra.Text = vaSpr1.Text
    vaSpr1.Col = 16:   TxtEmail.Text = vaSpr1.Text
  
'    rst.MoveFirst
'    Do While Not rst.EOF
'       If rst!codigo = LblCodigo.Caption Then
'          Exit Do
'       End If
'       rst.MoveNext
'    Loop
 
End Sub


