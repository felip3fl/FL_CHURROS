VERSION 5.00
Begin VB.Form frmCPFNota 
   BackColor       =   &H0004047B&
   BorderStyle     =   0  'None
   Caption         =   "CPFNota"
   ClientHeight    =   10155
   ClientLeft      =   11235
   ClientTop       =   705
   ClientWidth     =   5565
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   10155
   ScaleWidth      =   5565
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frameCPF 
      Appearance      =   0  'Flat
      BackColor       =   &H0000008C&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7080
      Left            =   295
      TabIndex        =   1
      Top             =   800
      Width           =   5000
      Begin VB.Frame frameTeclado 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   5175
         Left            =   480
         TabIndex        =   6
         Top             =   1530
         Width           =   4020
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
            TabIndex        =   16
            ToolTipText     =   "1"
            Top             =   90
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
            TabIndex        =   15
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
            Index           =   3
            Left            =   2685
            TabIndex        =   14
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
            Index           =   4
            Left            =   135
            TabIndex        =   13
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
            Index           =   5
            Left            =   1410
            TabIndex        =   12
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
            Index           =   7
            Left            =   135
            TabIndex        =   10
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
            Index           =   8
            Left            =   1410
            TabIndex        =   9
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
            Index           =   9
            Left            =   2685
            TabIndex        =   8
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
            Index           =   0
            Left            =   1410
            TabIndex        =   7
            ToolTipText     =   "0"
            Top             =   3900
            Width           =   1200
         End
      End
      Begin VB.Frame frmCodigoBarras 
         Appearance      =   0  'Flat
         BackColor       =   &H007171BF&
         BorderStyle     =   0  'None
         Caption         =   "Frame4"
         ForeColor       =   &H80000008&
         Height          =   630
         Left            =   0
         TabIndex        =   3
         Top             =   660
         Width           =   5000
         Begin VB.TextBox txtCPF 
            Appearance      =   0  'Flat
            BackColor       =   &H007171BF&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00D8D8FE&
            Height          =   360
            Left            =   135
            MaxLength       =   11
            TabIndex        =   4
            Text            =   "txtCPF"
            Top             =   135
            Width           =   3150
         End
         Begin VB.Image cmdLimparTXT 
            Height          =   465
            Left            =   4275
            Picture         =   "frmCPFNota.frx":0000
            Top             =   90
            Visible         =   0   'False
            Width           =   660
         End
      End
      Begin VB.Line Line4 
         BorderColor     =   &H003333A4&
         X1              =   -15
         X2              =   6000
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Label lblMSGCPF 
         BackStyle       =   0  'Transparent
         Caption         =   "CPF"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H007171BF&
         Height          =   330
         Left            =   0
         TabIndex        =   2
         ToolTipText     =   "CPF"
         Top             =   200
         Width           =   3075
      End
   End
   Begin VB.Timer tmrAnimacao 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   5025
      Top             =   330
   End
   Begin VB.Label cmdAdiciona 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Adiciona"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00EBEBF5&
      Height          =   1000
      Left            =   0
      TabIndex        =   5
      Top             =   8000
      Width           =   3435
   End
   Begin VB.Label lblMSGTitulo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nota Fiscal Paulista"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H007171BF&
      Height          =   435
      Left            =   390
      TabIndex        =   0
      Top             =   195
      Width           =   3405
   End
   Begin VB.Line borda1 
      BorderColor     =   &H00040460&
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   30120
   End
End
Attribute VB_Name = "frmCPFNota"
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

Dim MSGBotaoAciciona As String

Private Sub cmdAdiciona_Click()

    MSGBotaoCarregando Me, cmdAdiciona, "VALIDANDO"

    If adicionaCPF(txtCPF.Text) Then
        Form_Deactivate
        vendaAberta = True
        frmControle.statusVendaAberta
    Else
        limpaCampos
        campoValido lblMSGCPF, False
    End If
    
End Sub

Private Sub cmdLimparTXT_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    limpaCaractere txtCPF
    txtCPF.SetFocus
End Sub

Private Sub limpaCampos()

    MSGBotaoNormal Me, cmdAdiciona, "Adiciona"
    campoValido lblMSGCPF, True
    
    txtCPF.Text = Empty
    
End Sub

Private Sub cmdTecladoNum_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    entradaCaractereVirtual txtCPF, Index, cmdLimparTXT
End Sub

Private Sub Form_Activate()
    ajustaMenuComponentes
    limpaCampos
    ativaFormulario = True
    tmrAnimacao.Enabled = True
    txtCPF.SetFocus
End Sub

Private Sub Form_Deactivate()

    ativaFormulario = False
    tmrAnimacao.Enabled = True
    
End Sub

Private Sub Form_Load()
    ajustaMenu Me
    ajustaMenuComponentes
End Sub


Private Sub frmQuantidade_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub tmrAnimacao_Timer()
    'If ativaFormulario = True Then
        'abilitaFormularioAnima Me'
    'Else
        abilitaFormularioAnima ativaFormulario, Me, True
    'End If
End Sub

Private Sub ajustaMenuComponentes()

    cmdAdiciona.left = 0
    cmdAdiciona.Top = Me.Height - cmdAdiciona.Height
    cmdAdiciona.Width = Me.Width
    
    lblMSGTitulo.left = MARGEMMENUESQUERDA
    lblMSGTitulo.Top = MARGEMMENUTOPO
    
    frameCPF.BackColor = Me.BackColor
    frameCPF.left = MARGEMMENUESQUERDA
    
    frameTeclado.BackColor = Me.BackColor
    
    montaTecladoNumerico cmdTecladoNum, Me.BackColor
    
End Sub

Private Sub txtCPF_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        cmdAdiciona_Click
    End If
    
    KeyAscii = campoNumerico(KeyAscii)
    
End Sub

Private Sub txtCPF_KeyUp(KeyCode As Integer, Shift As Integer)
    botaoApagaVisivel cmdLimparTXT, txtCPF
End Sub
