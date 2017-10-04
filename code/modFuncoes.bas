Attribute VB_Name = "modFuncoes"
Option Explicit

'   ________________________________________________________________________________
'   \  ____________________________________________________________________________ \
'    \ \         ____    ____   __      __      ____     ____      ____   __       \ \
'     \ \       / ___\  / ___\ /\ \    /\_\    / __ \  /\___ \    / ___\ /\ \       \ \
'      \ \     /\ \__/ /\ \__/ \ \ \   \/\ \  /\ \_\ \ \/___\ \  /\ \__/ \ \ \       \ \
'       \ \    \ \  __\\ \  _\  \ \ \   \ \ \ \ \  __/   /\_ \ \ \ \  __\ \ \ \       \ \
'        \ \    \ \ \_/ \ \ \/   \ \ \   \ \ \ \ \ \/    \/_\ \ \ \ \ \_/  \ \ \       \ \
'         \ \    \ \ \   \ \ \___ \ \ \___\ \ \ \ \ \       _\_\ \ \ \ \    \ \ \___    \ \
'          \ \    \ \_\   \ \____\ \ \____\\ \_\ \ \_\     /\_____\ \ \_\    \ \____\    \ \
'           \ \    \/_/    \/____/  \/____/ \/_/  \/_/     \/_____/  \/_/     \/____/     \ \
'            \ \                                                                           \ \
'             \ \___________________________________________________________________________\ \
'              \_Felip3FL______________________________________________________________________\
'

Public ativaFormulario As Boolean

Private alturaPadraoFrame As Integer
Private tamanhoPadraoFrame As Integer

Public wAnimaSaida As Integer
Public wAnimaEntrda As Integer

Public Declare Sub Sleep Lib "kernel32" _
        (ByVal dwMilliseconds As Long)


Public Function pastaAtual() As String
    pastaAtual = CurDir & "\"
    'pastaAtual = "C:\Users\felipel ima\google drive\Documents\Desenv\Visual Basic 6\FL\ProjetoChurros\"
End Function

Public Function abrirConexaoADO(ByRef AdoVar, ByVal nomeServidor As String, _
                                 ByVal nomeBanco As String) As Boolean

    On Error GoTo ConexaoErro

    AdoVar.Provider = "SQLOLEDB"
    AdoVar.Properties("Data Source").Value = nomeServidor
    AdoVar.Properties("Initial Catalog").Value = nomeBanco
    AdoVar.Properties("User ID").Value = "felipelima"
    AdoVar.Properties("Password").Value = "felipe"
    
    AdoVar.Open
    
    abrirConexaoADO = True
    
    Exit Function
    
ConexaoErro:
    abrirConexaoADO = False
    
    MsgBox "Erro na Conexão ADO", vbCritical, "Erro 56"
    
End Function


Public Function comandoSQL(variavelConexao As ADODB.Recordset, comando As scriptSQL)

    Dim sql As String
    
    sql = comando.select & vbNewLine & _
          comando.from & vbNewLine & _
          comando.where & vbNewLine & _
          comando.orderBy
    
    variavelConexao.CursorLocation = adUseClient
    variavelConexao.Open sql, adoCNLoja, adOpenDynamic, adLockPessimistic
    
    comandoSQL = variavelConexao

End Function

Public Function insercaoSQL(comando As scriptSQL) As Boolean

    Dim sql As String
    
    sql = comando.delete & vbNewLine & _
          comando.insert & vbNewLine & _
          comando.update & vbNewLine & _
          comando.from & vbNewLine & _
          comando.where
    
    adoCNLoja.Execute sql
    
    insercaoSQL = True

End Function

Public Function updateSQL(comando As scriptSQL) As Integer

    Dim sql As String
    
    sql = comando.update & vbNewLine & _
          comando.where
    
    adoCNLoja.Execute sql
    
    updateSQL = 0

End Function

Public Sub selecionaBotao(botao, Index As Integer, variosSelecionado As Boolean)
    Dim i As Byte
    
    If variosSelecionado Then
        If botao(Index).BackColor = CORBOTAONORMAL Then
            botao(Index).BackColor = CORBOTAOPRESIONADO
        Else
            botao(Index).BackColor = CORBOTAONORMAL
        End If
    Else
        For i = 0 To botao.UBound
            botao(i).BackColor = CORBOTAONORMAL
        Next i
        
        botao(Index).BackColor = CORBOTAOPRESIONADO
    End If
    
End Sub

Public Sub criaNovoBotao(QTDE As Byte, vetorObjeto)
    
    Dim i As Byte
    
    If QTDE > vetorObjeto.UBound Then
    
        For i = vetorObjeto.UBound + 1 To QTDE
            Load vetorObjeto(i)
            vetorObjeto(i).Visible = True
        Next i
    
    ElseIf QTDE < vetorObjeto.UBound Then
        
        For i = QTDE + 1 To vetorObjeto.UBound
            Unload vetorObjeto(i)
        Next i
        
    End If
    
End Sub

Public Function formataPreco(preco As String) As String
    formataPreco = Format(Replace(preco, ".", ","), "####,###,#0.00")
End Function

Public Sub ajustaMenu(formulario As Form)

    formulario.Width = 5580
    formulario.Height = Screen.Height
    formulario.Top = 0
    formulario.left = Screen.Width
    formulario.borda1.X1 = 0
    formulario.borda1.X2 = 0
    formulario.borda1.Y1 = 0
    formulario.borda1.Y2 = Screen.Height
    
    AlwaysOnTop formulario, True
    
End Sub

Public Sub abilitaFormularioAnima(abilita As Boolean, formulario As Form, desabilitaForm As Boolean)
    
Dim valorAnimacao As Integer
    
    If abilita Then
    


    '    If (velocidadeAnimaMenu + wAnimaSaida) < 0 Then
    '        valorAnimacao = valorAnimacao + 12
    '    Else
    '        wAnimaSaida = wAnimaSaida - 12
    '    End If
    
        formulario.Visible = True
        
        valorAnimacao = (velocidadeAnimaMenu + wAnimaEntrda)
        'wAnimaEntrda = wAnimaEntrda - 12
        
        formulario.left = formulario.left - valorAnimacao
        
        If (formulario.left + formulario.Width) - valorAnimacao < frmControle.Width Then
            formulario.left = (frmControle.Width - formulario.Width)
            formulario.tmrAnimacao.Enabled = False
            formulario.SetFocus
            wAnimaEntrda = 0
        End If
    
    Else
    
        'Dim valorAnimacao As Integer
        
        valorAnimacao = (velocidadeAnimaMenu + wAnimaSaida)
        formulario.left = formulario.left + valorAnimacao
        
        If formulario.left >= frmControle.Width Or _
           (formulario.left + formulario.Width) < 0 Then
           
            formulario.tmrAnimacao.Enabled = False
            formulario.Visible = False
            wAnimaSaida = 0
            If desabilitaForm Then Unload formulario
            
        End If
        
        'abilitaFormularioAnima = True
    
    End If
    
End Sub

Public Sub expandirMenu(frame As frame, formulario As Form)

    If alturaPadraoFrame = 0 Then
        alturaPadraoFrame = frame.Top
        tamanhoPadraoFrame = frame.Height
        frame.Top = 0
        frame.Height = frmControle.Height - frame.Top
        frame.ZOrder 0
    Else
        frame.Top = alturaPadraoFrame
        frame.Height = tamanhoPadraoFrame
        alturaPadraoFrame = 0
    End If
    
End Sub

Public Function contadorCaracteres(ByVal Texto As String) As Integer
    
    Dim vetorCaractereSubs(1) As String
    Dim caractere As Variant
    
    vetorCaractereSubs(0) = "I"
    vetorCaractereSubs(1) = ","
    'vetorCaractereSubs(2) = "O"
    'vetorCaractereSubs(3) = "T"
    
    For Each caractere In vetorCaractereSubs
        Do While Texto Like "*" & caractere & "*"
            Texto = left$(Texto, (InStr(Texto, caractere) - 1)) _
            & "" _
            & Right$(Texto, (Len(Texto) - (InStr(Texto, caractere))))
            contadorCaracteres = contadorCaracteres + 1
        Loop
    Next
    

    
    
End Function

Public Sub limpaCaractere(ByRef campoTexto As TextBox)
    If campoTexto.Text = Empty Then Exit Sub
    campoTexto.Text = Mid(campoTexto.Text, 1, Len(campoTexto.Text) - 1)
End Sub

Public Function digitoNumerico(KeyAscii As Integer) As Boolean
    digitoNumerico = False
    If KeyAscii >= 48 And KeyAscii <= 57 Then
        digitoNumerico = True
    End If
End Function

Public Function campoNumerico(KeyAscii As Integer) As Integer
    If digitoNumerico(KeyAscii) Or digitoPadrao(KeyAscii) Then
        campoNumerico = KeyAscii
    End If
End Function

Public Function digitoPadrao(KeyAscii As Integer) As Boolean
    digitoPadrao = False
    If KeyAscii = 8 Or KeyAscii = 13 Or KeyAscii = 27 Then
        digitoPadrao = True
    End If
End Function

Public Function adicionaCPF(CPF As String) As Boolean
    If ValidaCPF(CPF) Then
        frmControle.adicionaCPF (CPF)
        adicionaCPF = True
    End If
End Function

Public Sub campoValido(Label As Label, valido As Boolean)
    If valido Then
        Label.Caption = Label.ToolTipText
        Label.ForeColor = CORFONTETITULO
    Else
        Label.Caption = Label.ToolTipText & " inválido!"
        Label.ForeColor = vbRed
    End If
End Sub

Function ValidaCPF(CPF As String) As Boolean
'
    Dim soma As Integer
    Dim Resto As Integer
    Dim i As Integer
    
    'Valida argumento
    If Len(CPF) <> 11 Then
        ValidaCPF = False
        Exit Function
    End If

        
    
    soma = 0
    For i = 1 To 9
        soma = soma + Val(Mid$(CPF, i, 1)) * (11 - i)
    Next i
    Resto = 11 - (soma - (Int(soma / 11) * 11))
    If Resto = 10 Or Resto = 11 Then Resto = 0
    If Resto <> Val(Mid$(CPF, 10, 1)) Then
        ValidaCPF = False
        Exit Function
    End If
        
    soma = 0
    For i = 1 To 10
        soma = soma + Val(Mid$(CPF, i, 1)) * (12 - i)
    Next i
    Resto = 11 - (soma - (Int(soma / 11) * 11))
    If Resto = 10 Or Resto = 11 Then Resto = 0
    If Resto <> Val(Mid$(CPF, 11, 1)) Then
        ValidaCPF = False
        Exit Function
    End If
    
    ValidaCPF = True

End Function

Public Sub centroFormulario(componente)
    componente.left = (Screen.Width / 2) - (componente.Width / 2)
End Sub

Public Sub centroFormularioHeight(componente)
    componente.Top = (Screen.Height - componente.Height) / 2
End Sub

Public Function Replace(Texto As String, caracter As String, caracterParaSubstituir As String) As String
    
    Do While Texto Like "*" & caracter & "*"
        Texto = left$(Texto, (InStr(Texto, caracter) - 1)) _
        & caracterParaSubstituir _
        & Right$(Texto, (Len(Texto) - (InStr(Texto, caracter))))
    Loop
    
    Replace = Texto
    
End Function

Public Function formataPrecoBD(preco As String) As String
    formataPrecoBD = Format(preco, "0.00")
    formataPrecoBD = Replace(formataPrecoBD, ",", ".")
End Function

Public Function MSGBotaoCarregando(formulario As Form, botao, Mensagem As String)
    botao.Caption = vbNewLine & Space(6) & Mensagem & " . . ."
    botao.ForeColor = CORFONTETITULO
    formulario.Enabled = False
    formulario.Refresh
End Function

Public Function MSGBotaoNormal(formulario As Form, botao, Mensagem As String)
    botao.Caption = vbNewLine & Mensagem
    botao.ForeColor = CORFONTEBOTAO
    formulario.Enabled = True
    formulario.Refresh
End Function

Public Sub montaTecladoNumerico(teclado, corTeclado As String)

    Dim i  As Byte

    For i = 0 To 9
        teclado(i).Caption = vbNewLine & vbNewLine & i
        teclado(i).BackColor = corTeclado
        teclado(i).ForeColor = CORFONTEBOTAO
    Next i
    

    If teclado.UBound > 9 Then
        If teclado.UBound = 11 Then
            teclado(11).Caption = vbNewLine & vbNewLine & ","
            teclado(11).BackColor = corTeclado
            teclado(11).FontBold = True
            teclado(11).ForeColor = CORFONTEBOTAO
        End If
        'teclado(10).Caption = vbNewLine & vbNewLine & "OK"
        'teclado(10).BackColor = corTeclado
        'teclado(10).ForeColor = CORFONTEBOTAO
        
    End If
    
End Sub

Public Sub botaoApagaVisivel(botao, Texto As String)
    If Len(Texto) > 0 Then
        botao.Visible = True
    Else
        botao.Visible = False
    End If
End Sub

Public Sub entradaCaractereVirtual(campo As TextBox, ByVal Index As Integer, botaoLimpar)
    campo.SetFocus
    campo.SelStart = Len(campo.Text)
    campo.SelText = Index
    botaoApagaVisivel botaoLimpar, campo.Text
End Sub

Public Sub selecionaTexto(campoTexto As TextBox)
    campoTexto.SelStart = 1
    campoTexto.SelLength = Len(campoTexto.Text)
End Sub

Public Sub sairSistema()
    If vendaAberta Then
        If frmLogin.validaLogin(glbUsuarioCodigo, True, True, _
        "Se você sair do sistema, irá cancelar essa venda", True) Then
                CupomCancela
                End
        End If
    Else
        End
    End If
End Sub

 Public Function pegarNumeroPedido()
 
    Dim sql As scriptSQL
    Dim RSUsuario As New ADODB.Recordset
    
    sql.select = "SELECT max(ITV_NotaFiscal) + 1 as pedido "
    sql.from = "FROM ItensVenda"
'
    comandoSQL RSUsuario, sql
'
    pegarNumeroPedido = RSUsuario("pedido")
 
 End Function


Sub Esperar(ByVal Tempo As Long)
    
    Dim StartTime As Long
    StartTime = Timer
    Do While Timer < StartTime + Tempo
        DoEvents
    Loop

End Sub
