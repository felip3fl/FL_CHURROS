VERSION 5.00
Object = "{D76D7130-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7d.ocx"
Begin VB.Form frmProdutoCadastro 
   BackColor       =   &H0000008C&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Produto"
   ClientHeight    =   8355
   ClientLeft      =   885
   ClientTop       =   1920
   ClientWidth     =   19110
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8355
   ScaleWidth      =   19110
   Begin VB.Frame frameFamilia 
      Appearance      =   0  'Flat
      BackColor       =   &H0004047B&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   8340
      Left            =   13140
      TabIndex        =   42
      Top             =   60
      Width           =   5670
      Begin VB.Timer tmrAnimacao 
         Enabled         =   0   'False
         Interval        =   10
         Left            =   4950
         Top             =   4365
      End
      Begin VB.Frame framePesquisa 
         Appearance      =   0  'Flat
         BackColor       =   &H0000008C&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   5865
         Left            =   300
         TabIndex        =   43
         Top             =   800
         Width           =   5000
         Begin VB.Frame frmCodigoBarras 
            Appearance      =   0  'Flat
            BackColor       =   &H007171BF&
            BorderStyle     =   0  'None
            Caption         =   "Frame4"
            ForeColor       =   &H80000008&
            Height          =   630
            Left            =   0
            TabIndex        =   44
            Top             =   645
            Width           =   5000
            Begin VB.TextBox txtPesquisaFamilia 
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
               TabIndex        =   45
               Text            =   "txtPesquisa"
               Top             =   135
               Width           =   3150
            End
            Begin VB.Image cmdLimparTXT 
               Height          =   465
               Left            =   4275
               Picture         =   "produtoCadastro.frx":0000
               Top             =   90
               Visible         =   0   'False
               Width           =   660
            End
         End
         Begin VSFlex7DAOCtl.VSFlexGrid grdFamilia 
            Height          =   4140
            Left            =   0
            TabIndex        =   47
            Top             =   1605
            Width           =   4995
            _cx             =   8811
            _cy             =   7302
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
            Cols            =   4
            FixedRows       =   0
            FixedCols       =   0
            RowHeightMin    =   5
            RowHeightMax    =   5
            ColWidthMin     =   5
            ColWidthMax     =   5
            ExtendLastCol   =   -1  'True
            FormatString    =   $"produtoCadastro.frx":0865
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
         Begin VB.Line Line3 
            BorderColor     =   &H003333A4&
            X1              =   -30
            X2              =   5865
            Y1              =   0
            Y2              =   0
         End
         Begin VB.Label lblMSGCodigo 
            BackStyle       =   0  'Transparent
            Caption         =   "Pesquisar Código"
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
            TabIndex        =   46
            ToolTipText     =   "Código"
            Top             =   195
            Width           =   3075
         End
      End
      Begin VB.Line borda1 
         BorderColor     =   &H00040460&
         X1              =   0
         X2              =   0
         Y1              =   -30
         Y2              =   30090
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Familia"
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
         TabIndex        =   49
         Top             =   195
         Width           =   1245
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
         Left            =   1185
         TabIndex        =   48
         Top             =   7755
         Visible         =   0   'False
         Width           =   3435
      End
   End
   Begin VB.Frame framePrincipal 
      Appearance      =   0  'Flat
      BackColor       =   &H00000080&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5715
      Left            =   330
      TabIndex        =   4
      Top             =   165
      Width           =   11475
      Begin VB.Frame Frame9 
         Appearance      =   0  'Flat
         BackColor       =   &H007171BF&
         BorderStyle     =   0  'None
         Caption         =   "Frame4"
         ForeColor       =   &H80000008&
         Height          =   525
         Left            =   2835
         TabIndex        =   39
         Top             =   3480
         Width           =   2650
         Begin VB.TextBox txtPrecoVenda2 
            Appearance      =   0  'Flat
            BackColor       =   &H007171BF&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00D8D8FE&
            Height          =   285
            Left            =   135
            MaxLength       =   11
            TabIndex        =   40
            Text            =   "txtCPF"
            Top             =   120
            Width           =   3150
         End
      End
      Begin VB.Frame txtBordaFamilia 
         Appearance      =   0  'Flat
         BackColor       =   &H007171BF&
         BorderStyle     =   0  'None
         Caption         =   "Frame4"
         ForeColor       =   &H80000008&
         Height          =   525
         Left            =   0
         TabIndex        =   37
         Top             =   4830
         Width           =   1635
         Begin VB.TextBox txtFamilia 
            Appearance      =   0  'Flat
            BackColor       =   &H007171BF&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00D8D8FE&
            Height          =   285
            Left            =   135
            Locked          =   -1  'True
            MaxLength       =   10
            TabIndex        =   38
            Top             =   120
            Width           =   3150
         End
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         BackColor       =   &H007171BF&
         BorderStyle     =   0  'None
         Caption         =   "Frame4"
         ForeColor       =   &H80000008&
         Height          =   525
         Left            =   1635
         TabIndex        =   23
         Top             =   1095
         Width           =   2235
         Begin VB.TextBox txtProdutoBarras 
            Appearance      =   0  'Flat
            BackColor       =   &H007171BF&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00D8D8FE&
            Height          =   285
            Left            =   135
            MaxLength       =   14
            TabIndex        =   24
            Text            =   "00000000000000"
            Top             =   120
            Width           =   3150
         End
      End
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         BackColor       =   &H007171BF&
         BorderStyle     =   0  'None
         Caption         =   "Frame4"
         ForeColor       =   &H80000008&
         Height          =   525
         Left            =   4020
         TabIndex        =   21
         Top             =   1095
         Width           =   7065
         Begin VB.TextBox txtDescricao 
            Appearance      =   0  'Flat
            BackColor       =   &H007171BF&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00D8D8FE&
            Height          =   285
            Left            =   135
            MaxLength       =   50
            TabIndex        =   22
            Text            =   "##################################################"
            Top             =   120
            Width           =   7440
         End
      End
      Begin VB.Frame Frame4 
         Appearance      =   0  'Flat
         BackColor       =   &H007171BF&
         BorderStyle     =   0  'None
         Caption         =   "Frame4"
         ForeColor       =   &H80000008&
         Height          =   525
         Left            =   0
         TabIndex        =   19
         Top             =   2130
         Width           =   2070
         Begin VB.TextBox txtNCM 
            Appearance      =   0  'Flat
            BackColor       =   &H007171BF&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00D8D8FE&
            Height          =   285
            Left            =   135
            MaxLength       =   11
            TabIndex        =   20
            Text            =   "00000000000"
            Top             =   120
            Width           =   3150
         End
      End
      Begin VB.Frame Frame5 
         Appearance      =   0  'Flat
         BackColor       =   &H007171BF&
         BorderStyle     =   0  'None
         Caption         =   "Frame4"
         ForeColor       =   &H80000008&
         Height          =   525
         Left            =   5685
         TabIndex        =   17
         Top             =   3480
         Width           =   1635
         Begin VB.TextBox txtICMS 
            Appearance      =   0  'Flat
            BackColor       =   &H007171BF&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00D8D8FE&
            Height          =   285
            Left            =   135
            MaxLength       =   10
            TabIndex        =   18
            Text            =   "0000000000"
            Top             =   120
            Width           =   3150
         End
      End
      Begin VB.Frame Frame6 
         Appearance      =   0  'Flat
         BackColor       =   &H007171BF&
         BorderStyle     =   0  'None
         Caption         =   "Frame4"
         ForeColor       =   &H80000008&
         Height          =   525
         Left            =   7530
         TabIndex        =   15
         Top             =   3480
         Width           =   1635
         Begin VB.TextBox txtIPI 
            Appearance      =   0  'Flat
            BackColor       =   &H007171BF&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00D8D8FE&
            Height          =   285
            Left            =   135
            MaxLength       =   10
            TabIndex        =   16
            Text            =   "0000000000"
            Top             =   120
            Width           =   3150
         End
      End
      Begin VB.Frame Frame7 
         Appearance      =   0  'Flat
         BackColor       =   &H007171BF&
         BorderStyle     =   0  'None
         Caption         =   "Frame4"
         ForeColor       =   &H80000008&
         Height          =   525
         Left            =   0
         TabIndex        =   13
         Top             =   3480
         Width           =   2650
         Begin VB.TextBox txtPrecoVenda1 
            Appearance      =   0  'Flat
            BackColor       =   &H007171BF&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00D8D8FE&
            Height          =   285
            Left            =   135
            MaxLength       =   11
            TabIndex        =   14
            Text            =   "txtCPF"
            Top             =   120
            Width           =   3150
         End
      End
      Begin VB.Frame Frame10 
         Appearance      =   0  'Flat
         BackColor       =   &H007171BF&
         BorderStyle     =   0  'None
         Caption         =   "Frame4"
         ForeColor       =   &H80000008&
         Height          =   525
         Left            =   0
         TabIndex        =   11
         Top             =   1095
         Width           =   1500
         Begin VB.TextBox txtCodigo 
            Appearance      =   0  'Flat
            BackColor       =   &H007171BF&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00D8D8FE&
            Height          =   285
            Left            =   135
            MaxLength       =   6
            TabIndex        =   12
            Text            =   "000000"
            Top             =   120
            Width           =   3150
         End
      End
      Begin VB.Frame Frame12 
         Appearance      =   0  'Flat
         BackColor       =   &H007171BF&
         BorderStyle     =   0  'None
         Caption         =   "Frame4"
         ForeColor       =   &H80000008&
         Height          =   525
         Left            =   2205
         TabIndex        =   9
         Top             =   2130
         Width           =   1680
         Begin VB.TextBox txtUnidade 
            Appearance      =   0  'Flat
            BackColor       =   &H007171BF&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00D8D8FE&
            Height          =   285
            Left            =   135
            MaxLength       =   3
            TabIndex        =   10
            Text            =   "###"
            Top             =   120
            Width           =   3150
         End
      End
      Begin VB.Frame Frame3 
         Appearance      =   0  'Flat
         BackColor       =   &H007171BF&
         BorderStyle     =   0  'None
         Caption         =   "Frame4"
         ForeColor       =   &H80000008&
         Height          =   525
         Left            =   4020
         TabIndex        =   7
         Top             =   2130
         Width           =   7065
         Begin VB.TextBox txtDescricaoMenu 
            Appearance      =   0  'Flat
            BackColor       =   &H007171BF&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00D8D8FE&
            Height          =   285
            Left            =   135
            MaxLength       =   50
            TabIndex        =   8
            Text            =   "##################################################"
            Top             =   120
            Width           =   7440
         End
      End
      Begin VB.Frame Frame13 
         Appearance      =   0  'Flat
         BackColor       =   &H007171BF&
         BorderStyle     =   0  'None
         Caption         =   "Frame4"
         ForeColor       =   &H80000008&
         Height          =   525
         Left            =   9390
         TabIndex        =   5
         Top             =   3480
         Width           =   1635
         Begin VB.TextBox txtCST 
            Appearance      =   0  'Flat
            BackColor       =   &H007171BF&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00D8D8FE&
            Height          =   285
            Left            =   135
            MaxLength       =   10
            TabIndex        =   6
            Text            =   "0000000000"
            Top             =   120
            Width           =   3150
         End
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Preço Venda 2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H007171BF&
         Height          =   285
         Left            =   2850
         TabIndex        =   41
         ToolTipText     =   "CPF"
         Top             =   3150
         Width           =   1665
      End
      Begin VB.Label lblMSGCPF 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Código"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H007171BF&
         Height          =   285
         Left            =   0
         TabIndex        =   36
         ToolTipText     =   "CPF"
         Top             =   765
         Width           =   840
      End
      Begin VB.Label lblMSGTitulo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cadastro Produto"
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
         Left            =   0
         TabIndex        =   35
         Top             =   0
         Width           =   3045
      End
      Begin VB.Line Line4 
         BorderColor     =   &H003333A4&
         X1              =   0
         X2              =   11065
         Y1              =   570
         Y2              =   570
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Produto Barras"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H007171BF&
         Height          =   285
         Left            =   1695
         TabIndex        =   34
         ToolTipText     =   "CPF"
         Top             =   765
         Width           =   1755
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Descrição"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H007171BF&
         Height          =   285
         Left            =   3990
         TabIndex        =   33
         ToolTipText     =   "CPF"
         Top             =   765
         Width           =   1155
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Descrição Menu"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H007171BF&
         Height          =   285
         Left            =   4020
         TabIndex        =   32
         ToolTipText     =   "CPF"
         Top             =   1800
         Width           =   1845
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NCM"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H007171BF&
         Height          =   285
         Left            =   0
         TabIndex        =   31
         ToolTipText     =   "CPF"
         Top             =   1800
         Width           =   555
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ICMS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H007171BF&
         Height          =   285
         Left            =   5685
         TabIndex        =   30
         ToolTipText     =   "CPF"
         Top             =   3150
         Width           =   600
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "IPI"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H007171BF&
         Height          =   285
         Left            =   7530
         TabIndex        =   29
         ToolTipText     =   "CPF"
         Top             =   3105
         Width           =   285
      End
      Begin VB.Label lblFamilia 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Familia"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H007171BF&
         Height          =   285
         Left            =   0
         TabIndex        =   28
         ToolTipText     =   "CPF"
         Top             =   4485
         Width           =   810
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Unidade"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H007171BF&
         Height          =   285
         Left            =   2190
         TabIndex        =   27
         ToolTipText     =   "CPF"
         Top             =   1800
         Width           =   960
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CST"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H007171BF&
         Height          =   285
         Left            =   9360
         TabIndex        =   26
         ToolTipText     =   "CPF"
         Top             =   3120
         Width           =   495
      End
      Begin VB.Line Line1 
         BorderColor     =   &H003333A4&
         X1              =   0
         X2              =   11100
         Y1              =   2955
         Y2              =   2955
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Preço Venda 1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H007171BF&
         Height          =   285
         Left            =   0
         TabIndex        =   25
         ToolTipText     =   "CPF"
         Top             =   3150
         Width           =   1665
      End
      Begin VB.Line Line2 
         BorderColor     =   &H003333A4&
         X1              =   0
         X2              =   11065
         Y1              =   4305
         Y2              =   4305
      End
   End
   Begin VB.Frame frameAcoes 
      Appearance      =   0  'Flat
      BackColor       =   &H003333A4&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      ForeColor       =   &H80000008&
      Height          =   1450
      Left            =   -2535
      TabIndex        =   0
      Top             =   6900
      Width           =   14100
      Begin VB.Frame frameGrava 
         Appearance      =   0  'Flat
         BackColor       =   &H003333A4&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1335
         Left            =   11235
         TabIndex        =   1
         Top             =   45
         Width           =   2505
         Begin VB.PictureBox cmdCarinho 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1095
            Left            =   765
            Picture         =   "produtoCadastro.frx":08F4
            ScaleHeight     =   1095
            ScaleWidth      =   1170
            TabIndex        =   2
            Top             =   -15
            Width           =   1170
         End
         Begin VB.Label lblGrupo 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Salva"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00EBEBF5&
            Height          =   630
            Index           =   0
            Left            =   765
            TabIndex        =   3
            Top             =   1125
            Width           =   1200
         End
         Begin VB.Image Image5 
            Height          =   1260
            Left            =   0
            Picture         =   "produtoCadastro.frx":1941
            Top             =   0
            Visible         =   0   'False
            Width           =   45
         End
      End
   End
End
Attribute VB_Name = "frmProdutoCadastro"
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

Dim ativaFormulario As Boolean

Private Sub Command1_Click()
    End
End Sub

Private Sub Command2_Click()
    End
End Sub

Private Sub cmdCarinho_Click()

    gravar txtCodigo.Text, txtProdutoBarras, txtDescricao, _
    txtDescricaoMenu.Text, txtNCM, txtICMS, txtCST, txtIPI, _
    txtPrecoVenda1, txtPrecoVenda2, txtUnidade, 0

End Sub

Private Sub Form_Load()
    
    propriedadeInicial
    limpaCampos
    'grdFamilia.h
    'grdFamilia.Row = 0
    'grdFamilia.Col = 0
    'grdFamilia.RowHeight = 100
    
End Sub

Private Sub propriedadeInicial()

    'frmProdutoCadastro.Height = 8820
    frmProdutoCadastro.Width = 12600

    framePrincipal.Top = 300
    framePrincipal.Left = 700
    framePrincipal.BackColor = Me.BackColor
    
'    frameFamilia.BackColor = Me.BackColor
    framePesquisa.BackColor = frameFamilia.BackColor
    
    frameAcoes.Width = frmProdutoCadastro.Width
    'frameAcoes.Top = (frmProdutoCadastro.Height - frameAcoes.Height)
    frameAcoes.Left = 0
    frameGrava.Left = (Me.Width - frameGrava.Width) - 300
    
    frameFamilia.Top = 0
    frameFamilia.Left = frmProdutoCadastro.Width
    
End Sub

Private Sub limpaCampos()
    txtCodigo.Text = Empty
    txtProdutoBarras.Text = Empty
    txtDescricao.Text = Empty
    txtNCM.Text = Empty
    txtUnidade.Text = Empty
    txtDescricaoMenu.Text = Empty
    txtPrecoVenda1.Text = Empty
    txtPrecoVenda2.Text = Empty
    txtICMS.Text = Empty
    txtIPI.Text = Empty
    txtCST.Text = Empty
    txtPesquisaFamilia.Text = Empty
End Sub



Private Sub gravar(codigo As String, codigobarras As String, _
                    descricao As String, descricaoMenu As String, _
                    NCM As String, ICMS As String, CTS As String, _
                    IPI As String, precoVenda1 As String, _
                    precoVenda2 As String, unidade As String, _
                    familia As String)

    Dim sql As scriptSQL

    sql.insert = "insert produtoLoja ("
    sql.insert = sql.insert & vbNewLine & "PRO_Codigo, "
    sql.insert = sql.insert & vbNewLine & "PRO_CodigoBarras, "
    sql.insert = sql.insert & vbNewLine & "PRO_Descricao, "
    sql.insert = sql.insert & vbNewLine & "PRO_DescricaoMenu, "
    sql.insert = sql.insert & vbNewLine & "PRO_NCM, "
    sql.insert = sql.insert & vbNewLine & "PRO_ICMS, "
    sql.insert = sql.insert & vbNewLine & "PRO_CST, "
    sql.insert = sql.insert & vbNewLine & "PRO_IPI, "
    sql.insert = sql.insert & vbNewLine & "PRO_PrecoVenda1, "
    sql.insert = sql.insert & vbNewLine & "PRO_PrecoVenda2, "
    sql.insert = sql.insert & vbNewLine & "PRO_Unidade, "
    sql.insert = sql.insert & vbNewLine & "PRO_Familia"
    sql.insert = sql.insert & vbNewLine & ")"
    
    sql.insert = sql.insert & vbNewLine & "values ("
    sql.insert = sql.insert & vbNewLine & "'" & codigo & "',"
    sql.insert = sql.insert & vbNewLine & "'" & codigobarras & "',"
    sql.insert = sql.insert & vbNewLine & "'" & descricao & "',"
    sql.insert = sql.insert & vbNewLine & "'" & descricaoMenu & "',"
    sql.insert = sql.insert & vbNewLine & "'" & NCM & "',"
    sql.insert = sql.insert & vbNewLine & "" & ICMS & ","
    sql.insert = sql.insert & vbNewLine & "'" & CTS & "',"
    sql.insert = sql.insert & vbNewLine & "" & IPI & ","
    sql.insert = sql.insert & vbNewLine & "" & precoVenda1 & ","
    sql.insert = sql.insert & vbNewLine & "" & precoVenda2 & ","
    sql.insert = sql.insert & vbNewLine & "'" & unidade & "',"
    sql.insert = sql.insert & vbNewLine & "" & 0 & ")"
    
    Call insercaoSQL(sql)
    
End Sub

Public Sub carregaFamilia(where As String)

    Dim sql As scriptSQL
    Dim RSItens As New ADODB.Recordset
    
    sql.select = "select LOWER(rtrim(FAP_Descricao)) as Descricao," & vbNewLine & _
                 "rtrim(FAP_CodigoFamilia) as CodigoFamilia," & vbNewLine & _
                 "rtrim(FAP_SubFamilia) as SubFamilia, " & vbNewLine & _
                 "FAP_ComplementoAdicional as ComplementoAdicional "
    sql.from = "from familiaProduto"
    sql.where = where
    
    comandoSQL RSItens, sql
    
    
    grdFamilia.Rows = 1
    grdFamilia.AddItem ""
        
    If RSItens.RecordCount > 0 Then
        Do While Not RSItens.EOF
            grdFamilia.AddItem RSItens("CodigoFamilia") & Chr(9) & _
            RSItens("descricao") & Chr(9) & _
            RSItens("SubFamilia") & Chr(9) & _
            RSItens("ComplementoAdicional")
            grdFamilia.AddItem ""
            RSItens.MoveNext
        Loop
    Else
        grdFamilia.AddItem "Código não encontrado" & Chr(9) & "Código não encontrado" & Chr(9) & "Código não encontrado" & Chr(9) & "Código não encontrado"
    End If
    
    RSItens.Close
    
End Sub


Private Sub Frame8_Click()

End Sub

Private Sub Frame8_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub frameAcoes_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub framePrincipal_Click()
    ativaFormulario = False
    tmrAnimacao.Enabled = True
End Sub

Private Sub grdFamilia_Click()
    If grdFamilia.Row > 1 Then
        txtFamilia.Text = grdFamilia.TextMatrix(grdFamilia.Row, 0)
        If txtFamilia.Text = Empty Then
            txtFamilia.Text = grdFamilia.TextMatrix(grdFamilia.Row - 1, 0)
        End If
        ativaFormulario = False
        tmrAnimacao.Enabled = True
    End If
End Sub

Private Sub Text8_Click()

End Sub

Private Sub Label10_Click()

End Sub

Private Sub lblFamilia_Click()
    txtBordaFamilia_Click
End Sub

Private Sub tmrAnimacao_Timer()
    abilitaFormularioAnima ativaFormulario, frameFamilia, True
End Sub

Private Sub txtBordaFamilia_Click()
    carregaFamilia ""
    txtPesquisaFamilia.Text = Empty
    ativaFormulario = True
    tmrAnimacao.Enabled = True
    txtPesquisaFamilia.SetFocus
End Sub

Private Sub txtFamilia_Click()
    txtBordaFamilia_Click
End Sub

Public Function montaWhereFamilia(codigoPesquisa As String) As String

    Dim sql As scriptSQL

    sql.where = "where FAP_CodigoFamilia like '%" & codigoPesquisa & "%' "
    
    montaWhereFamilia = sql.where

End Function

Private Sub txtPesquisaFamilia_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        carregaFamilia montaWhereFamilia(txtPesquisaFamilia.Text)
    End If
End Sub

Private Sub txtPesquisaFamilia_KeyUp(KeyCode As Integer, Shift As Integer)
    botaoApagaVisivel cmdLimparTXT, txtPesquisaFamilia
End Sub
