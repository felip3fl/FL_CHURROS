Attribute VB_Name = "modVariaveis"
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

Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Public Type POINTAPI
    X As Long
    Y As Long
End Type

Public Posicao As POINTAPI

Public vendaAberta As Boolean
Public precoEspecial As Byte

Public glbUsuarioNome As String
Public glbUsuarioCodigo As String


'''''''''''''''''''''''''''''''''
''' INFORMACAO BD '''''''''''''''
'''''''''''''''''''''''''''''''''

Global adoCNLoja As New ADODB.Connection
Global glbBancoLocal As String
Global glbServidorLocal As String
Global glbCodigoLoja As String

Type scriptSQL
    insert As String
    delete As String
    update As String
    select As String
    from As String
    where As String
    orderBy As String
End Type

'''''''''''''''''''''''''''''''''
''' CORES PADRAO ''''''''''''''''
'''''''''''''''''''''''''''''''''

Public Const CORBOTAONORMAL = &H4047B
Public Const CORBOTAOPRESIONADO = &H7171BF
Public Const CORFONTE = &HEBEBF5
Public Const CORFONTEBOTAO = &HD8D8FE
Public Const CORFONTETITULO = &H7171BF
Public Const CORFUNDOMENU = &H8C&
'&H007171BF&


'''''''''''''''''''''''''''''''''
''' TAMANHO BOTAO '''''''''''''''
'''''''''''''''''''''''''''''''''

Public Const BOTAOLARGURA = 3400
Public Const BOTAOALTURA = 1700

Public Const BOTAOPOSICAOTOPO = 200
Public Const BOTAOPOSICAOESQUERDA = 200

Public Const BOTAONUMLINHAS = 4

Public Const MARGEMCOLUNA = BOTAOALTURA + 260
Public Const MARGEMLINHA = BOTAOLARGURA + 260

Public Const BOTAOBORDASELEC = 81
Public Const BOTAOLARGURASELEC = BOTAOLARGURA + (BOTAOBORDASELEC * 2)
Public Const BOTAOALTURASELEC = BOTAOALTURA + (BOTAOBORDASELEC * 2)

Public Const BOTAOTEXTOPOSICAOTOPO = BOTAOALTURA - 570
Public Const BOTAOTEXTOPOSICAOESQUERDA = 180

Public Const MENUPOSICAOTOPO = BOTAOALTURA + 70
Public Const MENUPOSICAOESQUERDA = 360

Public Const MARGEMMENUTOPO = 200
Public Const MARGEMMENUESQUERDA = 300

'''''''''''''''''''''''''''''''''
''' TEXTO PADRAO' '''''''''''''''
'''''''''''''''''''''''''''''''''

Public Const MSGCARREGANDO = "      Carregando . . ."


'''''''''''''''''''''''''''''''''
''' PASTAS ''''''''''''''''''''''
'''''''''''''''''''''''''''''''''

Public Const ENDPASTAIMG = "data\"
Public Const ENDPASTATEMP = "temp\"


'''''''''''''''''''''''''''''''''
''' PASTAS ''''''''''''''''''''''
'''''''''''''''''''''''''''''''''

Public Const nomeIMGPadrao = "imgpadrao"

'''''''''''''''''''''''''''''''''
''' PASTAS ''''''''''''''''''''''
'''''''''''''''''''''''''''''''''

Public Const velocidadeAnimaMenu = 300
