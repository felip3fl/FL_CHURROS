VERSION 5.00
Object = "{D76D7130-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7d.ocx"
Begin VB.Form frmFuncoes 
   BackColor       =   &H0004047B&
   BorderStyle     =   0  'None
   Caption         =   "frmFuncoes"
   ClientHeight    =   10080
   ClientLeft      =   13515
   ClientTop       =   1080
   ClientWidth     =   5655
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10080
   ScaleWidth      =   5655
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmrAnimacao 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   4005
      Top             =   450
   End
   Begin VB.Frame frmFuncoes 
      Appearance      =   0  'Flat
      BackColor       =   &H0000008C&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7305
      Left            =   200
      TabIndex        =   0
      Top             =   800
      Width           =   5000
      Begin VB.PictureBox cmdLeituraZ 
         Appearance      =   0  'Flat
         BackColor       =   &H0004047B&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   900
         Left            =   3540
         Picture         =   "frmFuncoes.frx":0000
         ScaleHeight     =   900
         ScaleWidth      =   900
         TabIndex        =   13
         Top             =   780
         Width           =   900
      End
      Begin VB.PictureBox cmdLeituraX 
         Appearance      =   0  'Flat
         BackColor       =   &H0004047B&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   900
         Left            =   2010
         Picture         =   "frmFuncoes.frx":126C
         ScaleHeight     =   900
         ScaleWidth      =   900
         TabIndex        =   7
         Top             =   780
         Width           =   900
      End
      Begin VB.PictureBox cmdVendas 
         Appearance      =   0  'Flat
         BackColor       =   &H0004047B&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   900
         Left            =   3075
         Picture         =   "frmFuncoes.frx":24D4
         ScaleHeight     =   900
         ScaleWidth      =   900
         TabIndex        =   6
         Top             =   4530
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.PictureBox cmdCancelaNF 
         Appearance      =   0  'Flat
         BackColor       =   &H0004047B&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   900
         Left            =   465
         Picture         =   "frmFuncoes.frx":3674
         ScaleHeight     =   900
         ScaleWidth      =   900
         TabIndex        =   5
         Top             =   780
         Width           =   900
      End
      Begin VB.PictureBox cmdFechaCaixa 
         Appearance      =   0  'Flat
         BackColor       =   &H0004047B&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   900
         Left            =   465
         Picture         =   "frmFuncoes.frx":456D
         ScaleHeight     =   900
         ScaleWidth      =   900
         TabIndex        =   4
         Top             =   2430
         Width           =   900
      End
      Begin VB.PictureBox cmdSair 
         Appearance      =   0  'Flat
         BackColor       =   &H0004047B&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   900
         Left            =   2010
         Picture         =   "frmFuncoes.frx":53FB
         ScaleHeight     =   900
         ScaleLeft       =   10
         ScaleMode       =   0  'User
         ScaleTop        =   3
         ScaleWidth      =   900
         TabIndex        =   3
         Top             =   2430
         Width           =   900
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Leitura Z"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00D8D8FE&
         Height          =   315
         Left            =   3540
         TabIndex        =   14
         Top             =   1725
         Width           =   900
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Sair"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00D8D8FE&
         Height          =   315
         Left            =   2010
         TabIndex        =   12
         Top             =   3390
         Width           =   900
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Caixa"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00D8D8FE&
         Height          =   315
         Left            =   300
         TabIndex        =   11
         Top             =   3390
         Width           =   1290
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Cancela CF"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00D8D8FE&
         Height          =   315
         Left            =   360
         TabIndex        =   10
         Top             =   1725
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Leitura X"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00D8D8FE&
         Height          =   315
         Left            =   2010
         TabIndex        =   9
         Top             =   1725
         Width           =   900
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Vendas"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00D8D8FE&
         Height          =   315
         Left            =   3150
         TabIndex        =   8
         Top             =   5715
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Tarefas"
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
         TabIndex        =   1
         Top             =   195
         Width           =   4000
      End
      Begin VB.Line Line2 
         BorderColor     =   &H003333A4&
         X1              =   -60
         X2              =   5840
         Y1              =   0
         Y2              =   0
      End
   End
   Begin VSFlex7DAOCtl.VSFlexGrid grdPreco 
      Height          =   1140
      Left            =   765
      TabIndex        =   15
      Top             =   8505
      Visible         =   0   'False
      Width           =   3840
      _cx             =   6773
      _cy             =   2011
      _ConvInfo       =   1
      Appearance      =   0
      BorderStyle     =   0
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   263291
      ForeColor       =   7434687
      BackColorFixed  =   263291
      ForeColorFixed  =   7434687
      BackColorSel    =   263291
      ForeColorSel    =   7434687
      BackColorBkg    =   263291
      BackColorAlternate=   263291
      GridColor       =   263291
      GridColorFixed  =   263291
      TreeColor       =   263291
      FloodColor      =   263291
      SheetBorder     =   263291
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   3
      Cols            =   2
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmFuncoes.frx":64F1
      ScrollTrack     =   0   'False
      ScrollBars      =   2
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   -1  'True
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   263291
      ForeColorFrozen =   7434687
      WallPaperAlignment=   9
   End
   Begin VB.Label lblMSGTitulo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Funções do Sistema"
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
      Left            =   200
      TabIndex        =   2
      Top             =   225
      Width           =   3510
   End
   Begin VB.Line borda1 
      BorderColor     =   &H00040460&
      X1              =   30
      X2              =   30
      Y1              =   45
      Y2              =   30165
   End
End
Attribute VB_Name = "frmFuncoes"
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

Dim formDeactivate As Boolean

Public Sub cmdCancelaNF_Click()

    If frmLogin.validaLogin(glbUsuarioCodigo, True, True, _
    "Cancelar último cupom", True) Then
        Form_Deactivate
        CupomCancela
    End If
    
End Sub

Private Sub cmdFechaCaixa_Click()
    
    Dim msgLeituraZ As String
    
    msgLeituraZ = "Fechamento do Caixa"

    If frmLogin.validaLogin(glbUsuarioCodigo, True, False, msgLeituraZ, True) Then
        fecharCaixa glbCodigoLoja, glbUsuarioCodigo
        cmdSair_Click
    End If
    
End Sub

Private Sub fecharCaixa(loja As String, codigoUsuario As String)

    Dim sql As scriptSQL
                         
    sql.update = "update controlesistema  set "
    sql.update = sql.update & vbNewLine & "CS_Situacao = 'F'"
    sql.update = sql.update & vbNewLine & ",CS_dataFinal = GETDATE()"
    
    sql.where = vbNewLine & "where"
    sql.where = sql.where & vbNewLine & "CS_Situacao = 'A'"
    sql.where = sql.where & vbNewLine & "and CS_USUARIO = " & codigoUsuario
    sql.where = sql.where & vbNewLine & "and CS_loja = " & loja
    
    Call insercaoSQL(sql)
    
End Sub

Private Sub cmdLeituraX_Click()

    Screen.MousePointer = 11

    CupomLeituraX
    
    Screen.MousePointer = 0
    Form_Deactivate
    
End Sub

Private Sub cmdLeituraZ_Click()

    Dim msgLeituraZ As String
    
    msgLeituraZ = "ATENÇÃO! A leitura Z irá encerrar todas as operações de venda do dia!"

    If frmLogin.validaLogin(glbUsuarioCodigo, True, False, msgLeituraZ, True) Then
        CupomLeituraZ
        cmdSair_Click
    End If
    
End Sub

Public Sub cmdSair_Click()
    sairSistema
End Sub

Private Sub cmdVendas_Click()
    Form_Deactivate
End Sub

Private Sub Form_Activate()
    ajustaMenu Me
    formDeactivate = True
    ativaFormulario = True
    tmrAnimacao.Enabled = True
End Sub

Private Sub Form_Deactivate()
    If formDeactivate Then
        ativaFormulario = False
        tmrAnimacao.Enabled = True
    End If
End Sub

Private Sub Form_Load()
    frmFuncoes.left = MARGEMMENUESQUERDA
    frmFuncoes.BackColor = Me.BackColor
    lblMSGTitulo.left = MARGEMMENUESQUERDA
End Sub

Private Sub Picture1_Click()

End Sub

Private Sub tmrAnimacao_Timer()
    'If ativaFormulario = True Then
        'abilitaFormularioAnima Me
    'Else
        abilitaFormularioAnima ativaFormulario, Me, False
    'End If
End Sub
