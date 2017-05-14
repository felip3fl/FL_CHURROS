Attribute VB_Name = "Module1"
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

Public Sub centroFormularioWidth(componente)
    componente.Left = (Screen.Width / 2) - (componente.Width / 2)
End Sub

Public Sub centroFormularioHeight(componente)
    componente.Top = (Screen.Height - componente.Height) / 2
End Sub

Public Sub centroFormulario(componente)
    centroFormularioWidth componente
    centroFormularioHeight componente
End Sub

Public Function pastaAtual() As String
    pastaAtual = CurDir & "\"
End Function

Public Sub sairSistema()
    End
End Sub

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



Public Sub abilitaFormularioAnima(abilita As Boolean, formulario As Frame, desabilitaForm As Boolean)
    
Dim valorAnimacao As Integer
frmProdutoCadastro.tmrAnimacao.Enabled = True
    
    If abilita Then
    
        formulario.Visible = True
        
        'valorAnimacao = (300)
        
        formulario.Left = formulario.Left - 300
        
        If (formulario.Left + formulario.Width) < (frmProdutoCadastro.Width + 200) Then
            formulario.Left = (frmProdutoCadastro.Width - formulario.Width)
            frmProdutoCadastro.tmrAnimacao.Enabled = False
            'formulario.SetFocus
            'wAnimaEntrda = 0
        End If
    
    Else
        
        'valorAnimacao = (velocidadeAnimaMenu + wAnimaSaida)
        formulario.Left = formulario.Left + 200
        
        If formulario.Left >= frmProdutoCadastro.Width Or _
           (formulario.Left + formulario.Width) < 0 Then
           
            frmProdutoCadastro.tmrAnimacao.Enabled = False
            'frmProdutoCadastro.Visible = False
            'wAnimaSaida = 0
            'If desabilitaForm Then Unload formulario
            
        End If
    
    End If
    
End Sub

Public Function campoNumerico(KeyAscii As Integer) As Integer
    If digitoNumerico(KeyAscii) Or digitoPadrao(KeyAscii) Then
        campoNumerico = KeyAscii
    End If
End Function

Public Function digitoNumerico(KeyAscii As Integer) As Boolean
    digitoNumerico = False
    If KeyAscii >= 48 And KeyAscii <= 57 Then
        digitoNumerico = True
    End If
End Function

Public Function digitoPadrao(KeyAscii As Integer) As Boolean
    digitoPadrao = False
    If KeyAscii = 8 Or KeyAscii = 13 Or KeyAscii = 27 Then
        digitoPadrao = True
    End If
End Function

Public Sub botaoApagaVisivel(botao, Texto As String)
    If Len(Texto) > 0 Then
        botao.Visible = True
    Else
        botao.Visible = False
    End If
End Sub
