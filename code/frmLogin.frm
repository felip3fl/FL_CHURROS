VERSION 5.00
Begin VB.Form frmLogin 
   AutoRedraw      =   -1  'True
   BackColor       =   &H000000FF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   10455
   ClientLeft      =   -450
   ClientTop       =   1365
   ClientWidth     =   19995
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   10455
   ScaleWidth      =   19995
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmrAnima 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   13140
      Top             =   3075
   End
   Begin VB.Timer timerSenha 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   13140
      Top             =   2205
   End
   Begin VB.Frame frmSenha 
      Appearance      =   0  'Flat
      BackColor       =   &H000000C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   9030
      Left            =   7395
      TabIndex        =   2
      Top             =   120
      Width           =   5595
      Begin VB.Frame frmCaracteSenha 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   780
         Left            =   300
         TabIndex        =   4
         Top             =   1980
         Width           =   4935
         Begin VB.Image imgSenha 
            Appearance      =   0  'Flat
            Height          =   735
            Index           =   0
            Left            =   0
            Picture         =   "frmLogin.frx":0000
            Top             =   555
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.Image imgSenha 
            Appearance      =   0  'Flat
            Height          =   735
            Index           =   1
            Left            =   1400
            Picture         =   "frmLogin.frx":09D6
            Top             =   555
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.Image imgSenha 
            Appearance      =   0  'Flat
            Height          =   735
            Index           =   2
            Left            =   2800
            Picture         =   "frmLogin.frx":13AC
            Top             =   555
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.Image imgSenha 
            Appearance      =   0  'Flat
            Height          =   735
            Index           =   3
            Left            =   4200
            Picture         =   "frmLogin.frx":1D82
            Top             =   555
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.Image imgSenha 
            Appearance      =   0  'Flat
            Height          =   735
            Index           =   4
            Left            =   0
            Picture         =   "frmLogin.frx":2758
            Top             =   15
            Width           =   735
         End
         Begin VB.Image imgSenha 
            Appearance      =   0  'Flat
            Height          =   735
            Index           =   5
            Left            =   1400
            Picture         =   "frmLogin.frx":32AE
            Top             =   15
            Width           =   735
         End
         Begin VB.Image imgSenha 
            Appearance      =   0  'Flat
            Height          =   735
            Index           =   6
            Left            =   2800
            Picture         =   "frmLogin.frx":3E04
            Top             =   15
            Width           =   735
         End
         Begin VB.Image imgSenha 
            Appearance      =   0  'Flat
            Height          =   735
            Index           =   7
            Left            =   4200
            Picture         =   "frmLogin.frx":495A
            Top             =   15
            Width           =   735
         End
      End
      Begin VB.Frame frmTeclado 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   5175
         Left            =   780
         TabIndex        =   3
         Top             =   3375
         Width           =   4005
         Begin VB.Label cmdTecladoNum 
            Alignment       =   2  'Center
            BackColor       =   &H0000008C&
            Caption         =   "1"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   1200
            Index           =   0
            Left            =   1410
            TabIndex        =   15
            ToolTipText     =   "0"
            Top             =   3900
            Width           =   1200
         End
         Begin VB.Label cmdTecladoNum 
            Alignment       =   2  'Center
            BackColor       =   &H0000008C&
            Caption         =   "1"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   1200
            Index           =   9
            Left            =   2685
            TabIndex        =   14
            ToolTipText     =   "9"
            Top             =   2625
            Width           =   1200
         End
         Begin VB.Label cmdTecladoNum 
            Alignment       =   2  'Center
            BackColor       =   &H0000008C&
            Caption         =   "1"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   1200
            Index           =   8
            Left            =   1410
            TabIndex        =   13
            ToolTipText     =   "8"
            Top             =   2625
            Width           =   1200
         End
         Begin VB.Label cmdTecladoNum 
            Alignment       =   2  'Center
            BackColor       =   &H0000008C&
            Caption         =   "1"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   1200
            Index           =   7
            Left            =   135
            TabIndex        =   12
            ToolTipText     =   "7"
            Top             =   2625
            Width           =   1200
         End
         Begin VB.Label cmdTecladoNum 
            Alignment       =   2  'Center
            BackColor       =   &H0000008C&
            Caption         =   "1"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   1200
            Index           =   6
            Left            =   2685
            TabIndex        =   11
            ToolTipText     =   "6"
            Top             =   1350
            Width           =   1200
         End
         Begin VB.Label cmdTecladoNum 
            Alignment       =   2  'Center
            BackColor       =   &H0000008C&
            Caption         =   "1"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   1200
            Index           =   5
            Left            =   1410
            TabIndex        =   10
            ToolTipText     =   "5"
            Top             =   1350
            Width           =   1200
         End
         Begin VB.Label cmdTecladoNum 
            Alignment       =   2  'Center
            BackColor       =   &H0000008C&
            Caption         =   "1"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   1200
            Index           =   4
            Left            =   135
            TabIndex        =   9
            ToolTipText     =   "4"
            Top             =   1350
            Width           =   1200
         End
         Begin VB.Label cmdTecladoNum 
            Alignment       =   2  'Center
            BackColor       =   &H0000008C&
            Caption         =   "1"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   1200
            Index           =   3
            Left            =   2685
            TabIndex        =   8
            ToolTipText     =   "3"
            Top             =   75
            Width           =   1200
         End
         Begin VB.Label cmdTecladoNum 
            Alignment       =   2  'Center
            BackColor       =   &H0000008C&
            Caption         =   "1"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   1200
            Index           =   2
            Left            =   1410
            TabIndex        =   7
            ToolTipText     =   "2"
            Top             =   75
            Width           =   1200
         End
         Begin VB.Label cmdTecladoNum 
            Alignment       =   2  'Center
            BackColor       =   &H0000008C&
            Caption         =   "1"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   1200
            Index           =   1
            Left            =   135
            TabIndex        =   6
            ToolTipText     =   "1"
            Top             =   90
            Width           =   1200
         End
      End
      Begin VB.Label lblMensagem 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "mensagem"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   795
         Left            =   45
         TabIndex        =   16
         Top             =   90
         Width           =   5505
      End
      Begin VB.Label lblMSGLogin 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Entre com a senha:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   405
         Left            =   45
         TabIndex        =   5
         Top             =   930
         Width           =   5520
      End
   End
   Begin VB.Frame frmUsuario 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   9075
      Left            =   390
      TabIndex        =   0
      Top             =   885
      Visible         =   0   'False
      Width           =   6000
      Begin VB.Image cmdUsuario 
         Appearance      =   0  'Flat
         Height          =   1500
         Index           =   0
         Left            =   0
         Picture         =   "frmLogin.frx":54B0
         Top             =   0
         Width           =   1500
      End
      Begin VB.Label lblNomeUsuario 
         BackStyle       =   0  'Transparent
         Caption         =   "lblNomeUsuario"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   0
         Left            =   2850
         TabIndex        =   1
         Top             =   2190
         Width           =   2145
      End
   End
   Begin VB.Image imgSombra 
      Height          =   915
      Left            =   14280
      Picture         =   "frmLogin.frx":698C
      Top             =   9330
      Width           =   30000
   End
   Begin VB.Image imgDivisao 
      Height          =   10500
      Left            =   6945
      Picture         =   "frmLogin.frx":9133
      Top             =   0
      Width           =   45
   End
   Begin VB.Image cmdSair 
      Appearance      =   0  'Flat
      Height          =   1200
      Left            =   18990
      Picture         =   "frmLogin.frx":987C
      ToolTipText     =   "Sair do Sistema"
      Top             =   165
      Width           =   1200
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

Dim anima As Integer

Private loginValido As Boolean
Private ativaFormulario As Boolean

Private animacaoEntrada As Boolean
Private animacaoSaida As Boolean

Private RSUsuario As New ADODB.Recordset
Private senha As String

Private Sub cmdSair_Click()

    If glbUsuarioNome = Empty Then sairSistema
    animacaoSaida = True
    ativaFormulario = False
    tmrAnima.Enabled = True
    loginValido = False

End Sub

Private Sub cmdTecladoNum_DblClick(Index As Integer)
    cmdTecladoNum_MouseDown Index, 0, 0, 0, 0
End Sub

Private Sub cmdTecladoNum_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgSenha(Len(senha)).Visible = True
    senha = senha & Index
    
    If Len(senha) = 4 Then
        frmSenha.Enabled = False
        
        If senha = RSUsuario("senha") Then
            loginValido = True
            glbUsuarioNome = RSUsuario("nome")
            glbUsuarioCodigo = RSUsuario("codigo")
            
            lblMSGLogin.Caption = "Sucesso"
            frmLogin.Refresh
            frmLogin.Enabled = False
            
            Sleep 200
    
            ativaFormulario = False
            tmrAnima.Enabled = True
        Else
            senha = ""
            lblMSGLogin.Caption = "Senha Incorreta!"
            timerSenha.Enabled = True
        End If
        
    End If
End Sub

Private Sub cmdUsuario_Click(Index As Integer)
    Dim i As Byte
    
    For i = 0 To cmdUsuario.UBound
        cmdUsuario(i).Picture = LoadPicture(endIMGUsuario(CStr(i), False))
        lblNomeUsuario(i).FontBold = False
        lblNomeUsuario(i).left = (cmdUsuario(i).left + cmdUsuario(i).Height)
    Next i
    
    RSUsuario.MoveFirst
    RSUsuario.Move Index
    
    cmdUsuario(Index).Picture = LoadPicture(endIMGUsuario(RSUsuario("codigo"), True))
    lblNomeUsuario(Index).FontBold = True
    lblNomeUsuario(Index).left = (cmdUsuario(Index).left + cmdUsuario(Index).Height) + 160
    
    lblMSGLogin.Caption = "Entre com a senha:"
    
    limpaCaractereSenha

End Sub

Private Sub Command1_Click()
    loginValido = False
    Unload Me
End Sub

Private Sub Form_Activate()
    loginValido = False
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    If KeyAscii = 27 Then
        cmdSair_Click
    End If
    
    If KeyAscii >= 48 And KeyAscii <= 57 Then
        cmdTecladoNum_MouseDown KeyAscii - 48, 0, 0, 0, 0
    End If
    
End Sub

Private Sub Form_Load()
    ajustaMenuComponentes
    'carregaUsuariosBD glbUsuarioCodigo
    'ajustaMenuComponentes
End Sub

Private Sub ajustaMenuComponentes()
    
    Dim i As Byte
    
        
    'lblMensagem.top = (lblMSGLogin.top - 600)
    
    frmLogin.Top = -(Screen.Height)
    frmLogin.left = 0
    frmLogin.Height = Screen.Height
    frmLogin.Width = Screen.Width
    
    imgSombra.Top = (Screen.Height - imgSombra.Height)
    imgSombra.left = 0
    imgSombra.Height = Screen.Width
    
    frmLogin.BackColor = CORFUNDOMENU

    frmTeclado.Visible = True
    frmTeclado.BackColor = CORFUNDOMENU
    
    centroFormulario frmSenha
    centroFormularioHeight frmSenha
    frmSenha.left = frmSenha.left + 4000
    frmSenha.BackColor = CORFUNDOMENU
    
    frmCaracteSenha.Visible = True
    frmCaracteSenha.BackColor = CORFUNDOMENU
    For i = 0 To imgSenha.UBound
        imgSenha(i).Top = 15
    Next i
    
    lblMSGLogin.Visible = True
    
    frmUsuario.Visible = True
    frmUsuario.BackColor = CORFUNDOMENU
    centroFormulario frmUsuario
    frmUsuario.left = frmUsuario.left - 3000
    
    montaTecladoNumerico cmdTecladoNum, CORFUNDOMENU
    
    cmdSair.left = Screen.Width - cmdSair.Width
    cmdSair.Top = 0
    
    centroFormulario imgDivisao
    centroFormularioHeight imgDivisao
    
End Sub

Private Sub carregaUsuariosBD(usuarioCodigo As String)

    Dim sql As scriptSQL
    
    sql.select = "select top 10 rtrim(US_Codigo) as codigo," & _
                 vbNewLine & "Upper(rtrim(US_Nome)) as nome, " & _
                 vbNewLine & "US_TipoUsuario as tipoUsuario, " & _
                 vbNewLine & "US_Senha as senha"
    sql.from = "from usuario"
    sql.orderBy = "order by US_Codigo"
    
    If usuarioCodigo <> Empty Then sql.where = "WHERE US_Codigo = " & usuarioCodigo
    
    comandoSQL RSUsuario, sql
    
    If RSUsuario.RecordCount > 0 Then
        carregaPosicaoBotaoUsuario RSUsuario.RecordCount - 1
        carregaNomeUsuario RSUsuario.RecordCount - 1
        cmdUsuario_Click (0)
    Else
    
        cmdUsuario(0).Visible = False
        lblNomeUsuario(0).AutoSize = True
        lblNomeUsuario(0).Visible = True
        lblNomeUsuario(0).Caption = "ERRO! Falha ao tenta abrir o sistema" & vbNewLine & _
                                    "Entre em contato com o suporte"
        lblNomeUsuario(0).left = 0
        centroFormularioHeight lblNomeUsuario(0)
        lblNomeUsuario(0).Top = lblNomeUsuario(0).Top - frmUsuario.Top
        
        lblMensagem.Caption = Empty
        lblMSGLogin.Caption = "Nenhum usuario selecionado"
        frmSenha.Enabled = False
    
    End If
    
End Sub

Private Sub carregaNomeUsuario(qtdeBotao As Byte)

    Dim i As Byte
    
    For i = 0 To qtdeBotao
        lblNomeUsuario(i).Caption = RSUsuario("nome")
        lblNomeUsuario(i).ToolTipText = RSUsuario("nome")
        cmdUsuario(i).ToolTipText = RSUsuario("nome")
        lblNomeUsuario(i).Visible = True
        RSUsuario.MoveNext
    Next i
    
End Sub

Private Sub carregaPosicaoBotaoUsuario(qtdeBotao As Byte)

    Dim i As Byte

    criaNovoBotao qtdeBotao, cmdUsuario
    criaNovoBotao qtdeBotao, lblNomeUsuario
    
    For i = 0 To qtdeBotao
        cmdUsuario(i).Top = cmdUsuario(i).Top + ((cmdUsuario(i).Height + 15) * i)
        cmdUsuario(i).left = 600
        
        lblNomeUsuario(i).Top = (cmdUsuario(i).Top + (cmdUsuario(i).Height / 2)) - 120
        'lblNomeUsuario(i).left = (cmdUsuario(i).left + cmdUsuario(i).Height) + 150
    Next i
    
    frmUsuario.Height = (cmdUsuario(qtdeBotao).Top + cmdUsuario(qtdeBotao).Width)
    centroFormularioHeight frmUsuario
    
End Sub

Private Function endIMGUsuario(codigo As String, ativo As Boolean) As String

    Dim Arquivo As String
    Dim enderecoArquivo As String
    Dim tipoIMG As String * 1
    
    If ativo Then
        tipoIMG = "p"
    Else
        tipoIMG = "n"
    End If
    
    enderecoArquivo = pastaAtual & ENDPASTATEMP & "usu\usu" & codigo & tipoIMG
    Arquivo = Dir(enderecoArquivo, vbDirectory)
    
    If Arquivo = Empty Then
        endIMGUsuario = pastaAtual & ENDPASTAIMG & "usu" & tipoIMG
    Else
        endIMGUsuario = enderecoArquivo
    End If

    
End Function


Private Sub Form_Unload(Cancel As Integer)
    RSUsuario.Close
End Sub

Private Sub limpaCaractereSenha()
   Dim i As Byte
   
    For i = 0 To 3
        imgSenha(i).Visible = False
    Next
    
    senha = Empty
    
End Sub

Private Sub timerSenha_Timer()
    limpaCaractereSenha
    lblMSGLogin.Caption = "Entre com a senha:"
    frmSenha.Enabled = True
    timerSenha.Enabled = False
End Sub

Private Sub tmrAnima_Timer()

    If ativaFormulario Then
        
        If (frmLogin.Top + 100 + anima) > 0 Or animacaoEntrada = False Then
            frmLogin.Top = 0
            anima = 0
            tmrAnima.Enabled = False
        Else
            frmLogin.Top = (frmLogin.Top + anima) + 100
            anima = anima + 30
        End If

    Else
    
        If (frmLogin.Top + frmLogin.Height) < 0 Or animacaoSaida = False Then
            tmrAnima.Enabled = False
            anima = 0
            Unload Me
        Else
            frmLogin.Top = (frmLogin.Top - anima) - 100
            anima = anima + 30
        End If
        
    End If
End Sub

Public Function validaLogin(codigoUsuario As String, animaEntrada As Boolean, _
                            animaSaida As Boolean, Mensagem, aguardaResponta As Boolean) As Boolean

    If Len(Mensagem) < 20 Then Mensagem = vbNewLine & Mensagem
    lblMensagem.Caption = Mensagem

    carregaUsuariosBD codigoUsuario
    
    animacaoEntrada = animaEntrada
    animacaoSaida = animaSaida
    ativaFormulario = True
    tmrAnima.Enabled = True
    tmrAnima_Timer
    
    If aguardaResponta Then
        frmLogin.Show 1
    Else
        frmLogin.Show 0
    End If
    
    validaLogin = loginValido
    
End Function

Private Sub propriedadeInicial()
    frmSenha.Enabled = True
    limpaCaractereSenha
End Sub

Public Function aberturaCaixa()

    If glbUsuarioNome = Empty Then
        
        If caixaAberto(glbCodigoLoja) = False Then
            If frmLogin.validaLogin("", False, True, "Abertura de Caixa", True) = False Then End
            abrirCaixa glbUsuarioCodigo, glbCodigoLoja
        Else
            Call frmLogin.validaLogin(glbUsuarioCodigo, False, True, "Bem-Vindo de volta", False)
        End If
        
    End If
    
End Function

Private Sub abrirCaixa(codigoUsuario As String, loja As String)

    Dim sql As scriptSQL
    
    sql.insert = "insert controlesistema ("
    sql.insert = sql.insert & vbNewLine & "CS_Loja, "
    sql.insert = sql.insert & vbNewLine & "CS_Usuario, "
    sql.insert = sql.insert & vbNewLine & "CS_Situacao, "
    sql.insert = sql.insert & vbNewLine & "CS_DataInicial)"
    
    sql.insert = sql.insert & vbNewLine & "values ("
    sql.insert = sql.insert & vbNewLine & "" & loja & ","
    sql.insert = sql.insert & vbNewLine & "" & codigoUsuario & ","
    sql.insert = sql.insert & vbNewLine & "'" & "A" & "',"
    sql.insert = sql.insert & vbNewLine & "" & "GETDATE()" & ")"
    
    Call insercaoSQL(sql)
    
End Sub

Private Function caixaAberto(loja As String) As Boolean

    Dim sql As scriptSQL
    Dim RSCaixa As New ADODB.Recordset
    
    sql.select = "select top 1 cs_usuario as usuarioCodigo"
    sql.from = "from controlesistema"
    sql.where = "where CS_Situacao = 'A'" & _
                 vbNewLine & "and CS_Loja = " & loja
    
    comandoSQL RSCaixa, sql
    
    If RSCaixa.RecordCount > 0 Then
        glbUsuarioCodigo = RSCaixa("usuarioCodigo")
        caixaAberto = True
    End If
    
    RSCaixa.Close
    
End Function
