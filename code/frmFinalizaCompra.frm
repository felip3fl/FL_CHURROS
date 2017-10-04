VERSION 5.00
Object = "{D76D7130-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7d.ocx"
Begin VB.Form frmFinalizaCompra 
   BackColor       =   &H0004047B&
   BorderStyle     =   0  'None
   Caption         =   "frmFinalizaCompra"
   ClientHeight    =   10290
   ClientLeft      =   2985
   ClientTop       =   765
   ClientWidth     =   14460
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   ScaleHeight     =   10290
   ScaleWidth      =   14460
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frmItens 
      Appearance      =   0  'Flat
      BackColor       =   &H0000008C&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5595
      Left            =   405
      TabIndex        =   6
      Top             =   2640
      Width           =   5000
      Begin VSFlex7DAOCtl.VSFlexGrid grdItens 
         Height          =   4245
         Left            =   0
         TabIndex        =   8
         Top             =   720
         Width           =   4995
         _cx             =   8811
         _cy             =   7488
         _ConvInfo       =   1
         Appearance      =   0
         BorderStyle     =   0
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
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
         Rows            =   2
         Cols            =   5
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmFinalizaCompra.frx":0000
         ScrollTrack     =   0   'False
         ScrollBars      =   2
         ScrollTips      =   0   'False
         MergeCells      =   1
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
      Begin VB.Label lblExpandirTamanho 
         BackStyle       =   0  'Transparent
         Caption         =   "Itens (Quantidade: 1)"
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
         TabIndex        =   7
         Top             =   200
         Width           =   4000
      End
      Begin VB.Line Line3 
         BorderColor     =   &H003333A4&
         X1              =   -195
         X2              =   5060
         Y1              =   0
         Y2              =   0
      End
   End
   Begin VB.Frame frmModalidade 
      Appearance      =   0  'Flat
      BackColor       =   &H0000008C&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   8535
      Left            =   7965
      TabIndex        =   2
      Top             =   750
      Width           =   5000
      Begin VB.Frame frmTeclado 
         Appearance      =   0  'Flat
         BackColor       =   &H0004047B&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   5175
         Left            =   495
         TabIndex        =   12
         Top             =   3200
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
            Index           =   11
            Left            =   2685
            TabIndex        =   23
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
            Index           =   1
            Left            =   135
            TabIndex        =   22
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
            TabIndex        =   21
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
            TabIndex        =   20
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
            TabIndex        =   19
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
            TabIndex        =   18
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
            TabIndex        =   17
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
            TabIndex        =   16
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
            TabIndex        =   15
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
            TabIndex        =   14
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
            TabIndex        =   13
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
         Top             =   2205
         Width           =   5000
         Begin VB.TextBox txtDinheiro 
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
            Picture         =   "frmFinalizaCompra.frx":00AF
            Top             =   90
            Visible         =   0   'False
            Width           =   660
         End
      End
      Begin VB.Image cmdPagamento 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   795
         Index           =   4
         Left            =   4125
         Picture         =   "frmFinalizaCompra.frx":0914
         Stretch         =   -1  'True
         ToolTipText     =   "TEF"
         Top             =   735
         Width           =   795
      End
      Begin VB.Label lblModalidade 
         BackStyle       =   0  'Transparent
         Caption         =   "Dinheiro:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H007171BF&
         Height          =   330
         Left            =   0
         TabIndex        =   10
         Top             =   1740
         Width           =   4995
      End
      Begin VB.Image cmdPagamento 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   795
         Index           =   0
         Left            =   75
         Picture         =   "frmFinalizaCompra.frx":1DBF
         Stretch         =   -1  'True
         ToolTipText     =   "Dinheiro"
         Top             =   735
         Width           =   795
      End
      Begin VB.Image cmdPagamento 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   795
         Index           =   3
         Left            =   3120
         Picture         =   "frmFinalizaCompra.frx":33B9
         Stretch         =   -1  'True
         ToolTipText     =   "Cartão"
         Top             =   735
         Width           =   795
      End
      Begin VB.Image cmdPagamento 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   795
         Index           =   1
         Left            =   1110
         Picture         =   "frmFinalizaCompra.frx":4B85
         Stretch         =   -1  'True
         ToolTipText     =   "Cheque"
         Top             =   735
         Width           =   795
      End
      Begin VB.Image cmdPagamento 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   795
         Index           =   2
         Left            =   2115
         Picture         =   "frmFinalizaCompra.frx":5AD8
         Stretch         =   -1  'True
         ToolTipText     =   "Vale Alimentação"
         Top             =   735
         Width           =   795
      End
      Begin VB.Label lblMSGCPF 
         BackStyle       =   0  'Transparent
         Caption         =   "Modalidade"
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
         TabIndex        =   5
         ToolTipText     =   "CPF"
         Top             =   200
         Width           =   5000
      End
      Begin VB.Line Line4 
         BorderColor     =   &H003333A4&
         X1              =   -45
         X2              =   5050
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Label lblSelecionado 
         Appearance      =   0  'Flat
         BackColor       =   &H00B9B9FB&
         ForeColor       =   &H80000008&
         Height          =   945
         Left            =   0
         TabIndex        =   9
         Top             =   660
         Width           =   945
      End
   End
   Begin VB.Timer tmrAnimacao 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   5850
      Top             =   630
   End
   Begin VSFlex7DAOCtl.VSFlexGrid grdPreco 
      Height          =   1140
      Left            =   480
      TabIndex        =   11
      Top             =   975
      Width           =   3720
      _cx             =   6562
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
      FormatString    =   $"frmFinalizaCompra.frx":7274
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
   Begin VB.Label lblTotal 
      BackStyle       =   0  'Transparent
      Caption         =   "Total"
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
      Left            =   2010
      TabIndex        =   26
      Top             =   195
      Width           =   1200
   End
   Begin VB.Label cmdProximo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Próximo"
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
      Height          =   1005
      Left            =   4545
      TabIndex        =   25
      Top             =   8640
      Width           =   3435
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
      Height          =   1005
      Left            =   945
      TabIndex        =   24
      Top             =   9420
      Width           =   3435
   End
   Begin VB.Label cmdFinalizar 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Finaliza"
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
      Height          =   1005
      Left            =   720
      TabIndex        =   1
      Top             =   8565
      Width           =   3435
   End
   Begin VB.Label lblMSGTotal 
      BackStyle       =   0  'Transparent
      Caption         =   "Total"
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
      Left            =   300
      TabIndex        =   0
      Top             =   195
      Width           =   1200
   End
   Begin VB.Line borda1 
      BorderColor     =   &H00040460&
      X1              =   30
      X2              =   30
      Y1              =   90
      Y2              =   30210
   End
End
Attribute VB_Name = "frmFinalizaCompra"
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


Private RSItens As New ADODB.Recordset
Private RSCodigoPagamento As New ADODB.Recordset

Private Sub exibirTotal(Valor As String)
    lblTotal.Caption = formataPreco(Valor)
End Sub

Private Sub limpaVariavel()

    cmdFinalizar.Visible = False
    MSGBotaoNormal Me, cmdFinalizar, "Finaliza"
    
    cmdAdiciona.Visible = False
    MSGBotaoNormal Me, cmdAdiciona, "Adiciona"
    
    cmdProximo.Visible = True
    MSGBotaoNormal Me, cmdProximo, "Próximo"
    
    grdPreco.Visible = False
    grdPreco.Row = 1
    
    frmModalidade.Visible = False
    frmItens.Visible = True
    
    txtDinheiro.Text = Empty
    botaoApagaVisivel cmdLimparTXT, txtDinheiro.Text
    
    grdPreco.TextMatrix(0, 1) = 0
    grdPreco.TextMatrix(1, 1) = 0
    grdPreco.TextMatrix(2, 1) = 0
    
    lblModalidade.BackColor = CORFONTETITULO
    
    Me.left = Screen.Width
    
End Sub

Private Sub carregaCodigoPagamentoBD()
    
    Dim sql As scriptSQL
    
    If RSCodigoPagamento.State <> 0 Then RSCodigoPagamento.Close
    
    sql.select = "select rtrim(cp_codigo) as codigoPagamento"
    sql.from = "from codigoPagamento"
    sql.orderBy = "order by codigoPagamento "
    
    comandoSQL RSCodigoPagamento, sql
    
    If RSCodigoPagamento.RecordCount <= 0 Then
        MsgBox "Erro! Não foi encontrado códigos de pagamento", vbCritical, "Finaliza Compra"
        Form_Deactivate
    End If
    
End Sub

Private Sub cmdAdiciona_Click()
    txtDinheiro_KeyPress 13
End Sub

Private Sub cmdFinalizar_Click()
    
    MSGBotaoCarregando Me, cmdFinalizar, "FINALIZANDO"
    Esperar 1

    CupomIniciaFechamento
    cupomTerminaFechamento ""
    vendaAberta = False
    frmControle.limpaVariavel True
    Form_Deactivate
    
End Sub

Private Sub cmdLimparTXT_Click()
    limpaCaractere txtDinheiro
End Sub

Private Sub cmdPagamento_Click(Index As Integer)

    selecionaMenu cmdPagamento(Index), True
    RSCodigoPagamento.MoveFirst
    RSCodigoPagamento.Move Index
    
    lblModalidade.Caption = cmdPagamento(Index).ToolTipText
    lblModalidade.ToolTipText = cmdPagamento(Index).ToolTipText
    lblModalidade.ForeColor = CORFONTETITULO
    atualizaPagamento 0, Index
    
    If frmModalidade.Enabled And frmModalidade.Visible Then txtDinheiro.SetFocus
    
End Sub

Private Sub cmdProximo_Click()

    MSGBotaoCarregando Me, cmdProximo, "RESETANDO PAGAMENTO"
    Esperar 0.5
    desfazerPagamentoBD wNumeroCupom, "CF", glbCodigoLoja
    
    frmModalidade.Visible = True
    grdPreco.Visible = True
    frmItens.Visible = False
    
    cmdProximo.Visible = False
    cmdFinalizar.Visible = False
    cmdAdiciona.Visible = True
    Me.Enabled = True
    cmdPagamento_Click (0)
    
End Sub

Private Sub cmdTecladoNum_Click(Index As Integer)
    
    campoValido lblModalidade, True
    If Index = 11 Then
        txtDinheiro.Text = txtDinheiro.Text & ","
    ElseIf Index = 10 Then
        txtDinheiro_KeyPress 13
    Else
        txtDinheiro.Text = txtDinheiro.Text & Index
    End If
    
End Sub

Private Sub Form_Activate()

    limpaVariavel
    exibirTotal frmControle.lblPrecoTotal

    ativaFormulario = True
    tmrAnimacao.Enabled = True
    
    Dim i As Byte
    
    grdItens.Row = 0
    For i = 0 To grdItens.Cols - 1
        grdItens.Col = i
        grdItens.CellFontBold = True
    Next i
    
    grdItens.SetFocus
    
End Sub

Private Sub Form_Deactivate()
    'If RSCodigoPagamento.State <> 0 Then RSCodigoPagamento.Close
    ativaFormulario = False
    tmrAnimacao.Enabled = True
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdFinalizar_Click
    End If
End Sub

Private Sub Form_Load()
    ajustaMenu Me
    ajustaMenuComponentes
    carregaCodigoPagamentoBD
End Sub

Private Sub tmrAnimacao_Timer()
    abilitaFormularioAnima ativaFormulario, Me, True
End Sub


Public Sub carregaTamanho(pedido As String)

    Dim sql As scriptSQL
    Dim RSItens As New ADODB.Recordset
    
    sql.select = "select LOWER(rtrim(pro_descricaoMenu)) as DescricaoMenu," & vbNewLine & _
                 "ITV_NotaFiscal as nota," & vbNewLine & _
                 "ITV_Item as item, " & vbNewLine & _
                 "rtrim(ITV_CodigoProduto) as codigo, " & vbNewLine & _
                 "ITV_Quantidade as qtde, " & vbNewLine & _
                 "ITV_PrecoUnitario as preco," & vbNewLine & _
                 "rtrim(PRO_descricao) as descricao"
    sql.from = "from itensvenda, ProdutoLoja"
    sql.where = "where ITV_NotaFiscal = '" & pedido & "' and ITV_CodigoProduto = PRO_Codigo"
    
    comandoSQL RSItens, sql
    
    If RSItens.RecordCount > 0 Then
        grdItens.Rows = 1
        Do While Not RSItens.EOF
        carregaInfomacaoItens RSItens("item"), RSItens("codigo"), Val(RSItens("qtde")), CDbl(RSItens("preco")), RSItens("descricao")
        RSItens.MoveNext
        Loop
    End If
    
    RSItens.Close
    
End Sub

Private Sub carregaInfomacaoItens(Item As String, codigo As String, QTDE As Integer, preco As Double, descricao As String)
    grdItens.AddItem ""
    grdItens.AddItem Item & Chr(9) & codigo & Chr(9) & QTDE & Chr(9) & formataPreco(Str(preco)) & Chr(9) & formataPreco(Str(QTDE * preco))
    grdItens.AddItem descricao & Chr(9) & descricao & Chr(9) & descricao & Chr(9) & descricao & Chr(9) & descricao
    grdItens.MergeRow(grdItens.Rows - 1) = True
End Sub

Private Sub ajustaMenuComponentes()
    
    Dim i As Byte
    
    For i = 0 To grdPreco.Rows - 1
        grdPreco.TextMatrix(i, 1) = 0
    Next i
    
    lblTotal.Top = MARGEMMENUTOPO
    lblMSGTotal.Top = MARGEMMENUTOPO
    lblMSGTotal.left = MARGEMMENUESQUERDA
    
    cmdFinalizar.left = 0
    cmdFinalizar.Top = (Me.Height - cmdFinalizar.Height)
    cmdFinalizar.Width = Me.Width
    
    cmdAdiciona.left = 0
    cmdAdiciona.Top = (Me.Height - cmdAdiciona.Height)
    cmdAdiciona.Width = Me.Width
    
    cmdProximo.left = 0
    cmdProximo.Top = (Me.Height - cmdProximo.Height)
    cmdProximo.Width = Me.Width
    
    frmItens.Top = (lblMSGTotal.Top + lblMSGTotal.Height) + MARGEMMENUTOPO
    frmItens.Height = (cmdFinalizar.Top - frmItens.Top)
    frmItens.left = MARGEMMENUESQUERDA
    frmItens.Visible = True
    frmItens.BackColor = &H4047B
    
    grdItens.left = 0
    grdItens.Height = (frmItens.Height - grdItens.Top)
    
    grdPreco.Top = (lblMSGTotal.Top + lblMSGTotal.Height) + MARGEMMENUTOPO
    grdPreco.left = MARGEMMENUESQUERDA
    grdPreco.Visible = False
    
    frmModalidade.Top = (grdPreco.Top + grdPreco.Height) + MARGEMMENUTOPO
    frmModalidade.left = MARGEMMENUESQUERDA
    frmModalidade.Height = (cmdFinalizar.Top - frmModalidade.Top)
    frmModalidade.Visible = False
    frmModalidade.BackColor = &H4047B
    
    txtDinheiro.Text = Empty
    
    montaTecladoNumerico cmdTecladoNum, Me.BackColor
    
    
    
'    frmModalidade.top = (frmModalidade.top + frmModalidade.Height)
'    frmModalidade.left = frmModalidade.left
'    frmModalidade.Visible = False
    
End Sub

Private Sub selecionaMenu(Menu, ativa As Boolean)
    If ativa Then
        lblSelecionado.ZOrder 1
        lblSelecionado.Top = Menu.Top - Val(BOTAOBORDASELEC)
        lblSelecionado.left = Menu.left - Val(BOTAOBORDASELEC)
        lblSelecionado.Visible = True
    Else
        lblSelecionado.Visible = False
    End If
End Sub

Private Sub txtDinheiro_KeyDown(KeyCode As Integer, Shift As Integer)
    botaoApagaVisivel cmdLimparTXT, txtDinheiro
End Sub

Private Sub txtDinheiro_KeyPress(KeyAscii As Integer)
    
    campoValido lblModalidade, True
    If KeyAscii = 13 Then
        If txtDinheiro.Text = Empty Then txtDinheiro.Text = 0
        If validaValor(CDbl(txtDinheiro.Text)) Then
            atualizaPagamento CDbl(txtDinheiro.Text), 1
            gravaPagamentoBD wNumeroCupom, "CF", CDbl(txtDinheiro.Text), glbCodigoLoja, glbUsuarioCodigo, 0, RSCodigoPagamento("codigoPagamento")
            txtDinheiro.Text = Empty
        Else
            campoValido lblModalidade, False
            txtDinheiro.Text = Empty
        End If
    End If
    
End Sub

Private Function validaValor(Valor As Double)
    If Valor > 0 Then validaValor = True
End Function

Private Function atualizaPagamentoBD(notaFiscal As String, _
                                serie As String, _
                                loja As String, _
                                FormaPagamento As String, _
                                Valor As Double) As Boolean

    Dim sql As scriptSQL
    Dim RSPagamento As New ADODB.Recordset
    
    atualizaPagamentoBD = False
    
    sql.select = "select count(PGT_NotaFiscal) as qtdePagamento"
    sql.from = "from pagamento"
    
    sql.where = "where"
    sql.where = sql.where & vbNewLine & "PGT_NotaFiscal = " & notaFiscal & ""
    sql.where = sql.where & vbNewLine & "and PGT_Serie = '" & serie & "'"
    sql.where = sql.where & vbNewLine & "and PGT_Loja = " & loja & ""
    sql.where = sql.where & vbNewLine & "and PGT_FormaPagamento = '" & FormaPagamento & "'"
    
    comandoSQL RSPagamento, sql
    
    If Val(RSPagamento("qtdePagamento")) > 0 Then
            
            sql.update = "update pagamento  set "
            sql.update = sql.update & vbNewLine & "PGT_Valor = PGT_Valor + " & formataPrecoBD(CStr(Valor))
                                                          
            Call updateSQL(sql)
            atualizaPagamentoBD = True
            
    End If
    
    RSPagamento.Close
    
End Function

Private Sub gravaPagamentoBD(notaFiscal As String, _
                             serie As String, _
                             Valor As Double, _
                             loja As String, _
                             usuario As String, _
                             caixa As String, _
                             FormaPagamento As String)
             
    If atualizaPagamentoBD(notaFiscal, serie, loja, FormaPagamento, Valor) = False Then
    
        Dim sql As scriptSQL
    
        sql.insert = "insert PAGAMENTO ("
        sql.insert = sql.insert & vbNewLine & "PGT_NotaFiscal, "
        sql.insert = sql.insert & vbNewLine & "PGT_Serie, "
        sql.insert = sql.insert & vbNewLine & "PGT_Loja, "
        sql.insert = sql.insert & vbNewLine & "PGT_Caixa, "
        sql.insert = sql.insert & vbNewLine & "PGT_Operador, "
        sql.insert = sql.insert & vbNewLine & "PGT_FormaPagamento, "
        sql.insert = sql.insert & vbNewLine & "PGT_Valor)"
        
        sql.insert = sql.insert & vbNewLine & "values ("
        sql.insert = sql.insert & vbNewLine & "" & notaFiscal & ","
        sql.insert = sql.insert & vbNewLine & "'" & serie & "',"
        sql.insert = sql.insert & vbNewLine & "'" & loja & "',"
        sql.insert = sql.insert & vbNewLine & "" & caixa & ","
        sql.insert = sql.insert & vbNewLine & "'" & usuario & "',"
        sql.insert = sql.insert & vbNewLine & "'" & FormaPagamento & "',"
        sql.insert = sql.insert & vbNewLine & "" & formataPrecoBD(CStr(Valor)) & ")"
        
        Call insercaoSQL(sql)
        
    End If

End Sub

Private Sub desfazerPagamentoBD(notaFiscal As String, _
                             serie As String, _
                             codigoLoja As String)
                             
    Dim sql As scriptSQL
    
    sql.delete = "delete pagamento"
    sql.where = "where PGT_NotaFiscal =  " & notaFiscal & ""
    sql.where = sql.where & vbNewLine & "and PGT_Serie = '" & serie & "'"
    sql.where = sql.where & vbNewLine & "and PGT_Loja = '" & codigoLoja & "'"
    

    Call insercaoSQL(sql)

End Sub

Private Sub atualizaPagamento(Valor As Double, indexPagamento)

    grdPreco.TextMatrix(0, 1) = CDbl(grdPreco.TextMatrix(0, 1)) + Valor
    grdPreco.TextMatrix(1, 1) = CDbl(lblTotal) - CDbl(grdPreco.TextMatrix(0, 1))
    
    grdPreco.Row = 2
    
    If CDbl(grdPreco.TextMatrix(1, 1)) <= 0 Then
        cmdFinalizar.Visible = True
        cmdAdiciona.Visible = False
        If CDbl(grdPreco.TextMatrix(1, 1)) < 0 Then
            grdPreco.TextMatrix(2, 1) = (CDbl(grdPreco.TextMatrix(1, 1))) * -1
            txtDinheiro.Text = CDbl(txtDinheiro) - CDbl(grdPreco.TextMatrix(2, 1))
            grdPreco.Col = 0
            grdPreco.CellFontBold = True
            grdPreco.Col = 1
            grdPreco.CellFontBold = True
            grdPreco.TextMatrix(1, 1) = 0
        Else
            grdPreco.Col = 0
            grdPreco.CellFontBold = False
            grdPreco.Col = 1
            grdPreco.CellFontBold = False
        End If
    End If
    
    grdPreco.TextMatrix(0, 1) = formataPreco(grdPreco.TextMatrix(0, 1))
    grdPreco.TextMatrix(1, 1) = formataPreco(grdPreco.TextMatrix(1, 1))
    grdPreco.TextMatrix(2, 1) = formataPreco(grdPreco.TextMatrix(2, 1))
    
End Sub

Private Sub txtDinheiro_KeyUp(KeyCode As Integer, Shift As Integer)
    botaoApagaVisivel cmdLimparTXT, txtDinheiro
End Sub
