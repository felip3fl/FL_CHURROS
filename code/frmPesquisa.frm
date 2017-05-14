VERSION 5.00
Begin VB.Form frmPesquisa 
   BackColor       =   &H0004047B&
   BorderStyle     =   0  'None
   Caption         =   "Pesquisa"
   ClientHeight    =   10155
   ClientLeft      =   2910
   ClientTop       =   705
   ClientWidth     =   5565
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   10155
   ScaleWidth      =   5565
   ShowInTaskbar   =   0   'False
   Begin VB.Frame framePesquisa 
      Appearance      =   0  'Flat
      BackColor       =   &H0000008C&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6765
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
         Left            =   465
         TabIndex        =   6
         Top             =   1470
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
         Top             =   645
         Width           =   5000
         Begin VB.TextBox txtPesquisa 
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
            MaxLength       =   6
            TabIndex        =   4
            Text            =   "txtPesquisa"
            Top             =   135
            Width           =   3150
         End
         Begin VB.Image cmdLimparTXT 
            Height          =   465
            Left            =   4275
            Picture         =   "frmPesquisa.frx":0000
            Top             =   90
            Visible         =   0   'False
            Width           =   660
         End
      End
      Begin VB.Line Line4 
         BorderColor     =   &H003333A4&
         X1              =   -30
         X2              =   5865
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Label lblMSGCodigo 
         BackStyle       =   0  'Transparent
         Caption         =   "Código"
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
         ToolTipText     =   "Código"
         Top             =   195
         Width           =   3075
      End
   End
   Begin VB.Timer tmrAnimacao 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   5055
      Top             =   330
   End
   Begin VB.Label cmdAdiciona 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Pesquisa"
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
      Caption         =   "Pesquisa"
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
      Width           =   1575
   End
   Begin VB.Line borda1 
      BorderColor     =   &H00040460&
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   30120
   End
End
Attribute VB_Name = "frmPesquisa"
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

Dim caractereInicial As String

Private Sub cmdAdiciona_Click()
    txtPesquisa_KeyPress 13
End Sub

Private Sub cmdLimparTXT_Click()
    limpaCaractere txtPesquisa
End Sub

Private Sub limpaCampos()
    MSGBotaoNormal Me, cmdAdiciona, "Pesquisa"
    txtPesquisa.Text = Empty
    campoValido lblMSGCodigo, True
End Sub

Private Sub cmdTecladoNum_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    entradaCaractereVirtual txtPesquisa, Index, cmdLimparTXT
End Sub

Private Sub Form_Activate()
    limpaCampos
    ativaFormulario = True
    tmrAnimacao.Enabled = True
    txtPesquisa.Text = caractereInicial
    txtPesquisa.SelStart = Len(txtPesquisa.Text)
    txtPesquisa.SetFocus
End Sub

Private Sub Form_Deactivate()

    caractereInicial = ""
    ativaFormulario = False
    tmrAnimacao.Enabled = True
    
End Sub

Private Sub Form_Load()
    ajustaMenu Me
    ajustaMenuComponentes
End Sub


Private Sub tmrAnimacao_Timer()
    'If ativaFormulario = True Then
       ' abilitaFormularioAnima Me
    'Else
        abilitaFormularioAnima ativaFormulario, Me, False
    'End If
End Sub

Private Sub ajustaMenuComponentes()

    cmdAdiciona.left = 0
    cmdAdiciona.Top = Me.Height - cmdAdiciona.Height
    cmdAdiciona.Width = Me.Width
    
    lblMSGTitulo.left = MARGEMMENUESQUERDA
    lblMSGTitulo.Top = MARGEMMENUTOPO
    
    framePesquisa.BackColor = Me.BackColor
    framePesquisa.left = MARGEMMENUESQUERDA
    
    frameTeclado.BackColor = Me.BackColor
    
    montaTecladoNumerico cmdTecladoNum, Me.BackColor
    
End Sub

Private Sub txtPesquisa_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
         frmControle.carregaFamilia (montaWhereSQL(txtPesquisa.Text))
         txtPesquisa.Text = Empty
         Form_Deactivate
    End If
    KeyAscii = campoNumerico(KeyAscii)
End Sub

Public Function pesquisar(primeiroCaractere As String) As String
    
    frmPesquisa.Show
    caractereInicial = primeiroCaractere
    pesquisar = montaWhereSQL(txtPesquisa.Text)
    
End Function

Public Function montaWhereSQL(codigoPesquisa As String) As String

    Dim sql As scriptSQL

    sql.select = "select pro_familia"
    sql.from = "from ProdutoLoja, FamiliaProduto"
    sql.where = "where PRO_Codigo like '%" & codigoPesquisa & "%' and PRO_Familia = FAP_CodigoFamilia group by pro_familia "
    
    montaWhereSQL = sql.select & " " & _
                sql.from & " " & _
                sql.where
                
End Function

Private Sub txtPesquisa_KeyUp(KeyCode As Integer, Shift As Integer)
    botaoApagaVisivel cmdLimparTXT, txtPesquisa
End Sub
