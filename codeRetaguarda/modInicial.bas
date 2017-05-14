Attribute VB_Name = "Module2"
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

Sub main()

    verificaAppExecucao
    If Not obterInformacaoBancosINI Then sairSistema
    If Not abrirConexaoBD(adoCNLoja, glbServidorLocal, glbBancoLocal) Then sairSistema
    
    frmProdutoCadastro.Show
    
End Sub

Public Sub verificaAppExecucao()
    If App.PrevInstance Then
       MsgBox "Aplicativo executando", vbInformation
       End
    End If
End Sub

Private Function obterInformacaoBancosINI()

    Dim conexaoBDINI As String
    Dim adoCNAccess As New ADODB.Connection
    Dim rdoConexaoINI As New ADODB.Recordset
    Dim sql As String

    conexaoBDINI = "Driver={Microsoft Access Driver (*.mdb)};" & _
                   "Dbq=" & pastaAtual & "BDini.mdb;" & _
                   "Uid=Admin; Pwd=felipe"

    adoCNAccess.Open conexaoBDINI
    
    sql = "Select * from ConexaoSistema"
          
    rdoConexaoINI.CursorLocation = adUseClient
    rdoConexaoINI.Open sql, adoCNAccess, adOpenForwardOnly, adLockPessimistic
    
        If rdoConexaoINI.RecordCount = 0 Then
            MsgBox "Não há informação de conexão no INI! ", vbCritical, "Conexão com Banco de Dados"
        ElseIf rdoConexaoINI.RecordCount = 1 Then
            
            glbBancoLocal = rdoConexaoINI("GLB_BancoLocal")
            glbServidorLocal = rdoConexaoINI("GLB_ServidorLocal")
            glbCodigoLoja = rdoConexaoINI("GLB_CodigoLoja")
            
        End If
    
    adoCNAccess.Close
    obterInformacaoBancosINI = True
    
    Exit Function
    
ConexaoErro:

    MsgBox "Erro ao abrir conexão com o INI! ", vbCritical, "Conexão INI"
    End
End Function

Public Function abrirConexaoBD(variavelBanco, servidor As String, Banco As String) As Boolean

    Dim ConexaoDLLAdo As New DMACD.conexaoADO

    If ConexaoDLLAdo.abrirConexaoADO(variavelBanco, servidor, Banco) Then
        abrirConexaoBD = True
    Else
        MsgBox "Erro ao abrir conexão com o Banco de Dados!", vbCritical, "Conexão Banco de Dados"
    End If
    
End Function

