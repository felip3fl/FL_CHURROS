VERSION 5.00
Begin VB.Form frmControle 
   Appearance      =   0  'Flat
   BackColor       =   &H0000008C&
   BorderStyle     =   0  'None
   ClientHeight    =   11115
   ClientLeft      =   2535
   ClientTop       =   3255
   ClientWidth     =   16005
   Icon            =   "inicio.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11115
   ScaleWidth      =   16005
   ShowInTaskbar   =   0   'False
   Tag             =   "5"
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Frame frmCarregando 
      Appearance      =   0  'Flat
      BackColor       =   &H0004047B&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      ForeColor       =   &H80000008&
      Height          =   1080
      Left            =   5910
      TabIndex        =   21
      Top             =   4770
      Width           =   6330
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "   CARREGANDO . . ."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00EBEBF5&
         Height          =   345
         Left            =   15
         TabIndex        =   22
         Top             =   375
         Width           =   6300
      End
   End
   Begin VB.Frame frmGrupo 
      Appearance      =   0  'Flat
      BackColor       =   &H003333A4&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      ForeColor       =   &H80000008&
      Height          =   1450
      Left            =   0
      TabIndex        =   1
      Top             =   10065
      Width           =   30000
      Begin VB.PictureBox cmdGrupo 
         Appearance      =   0  'Flat
         BackColor       =   &H003333A4&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1200
         Index           =   0
         Left            =   795
         Picture         =   "inicio.frx":628A
         ScaleHeight     =   1200
         ScaleWidth      =   1200
         TabIndex        =   19
         Top             =   0
         Width           =   1200
      End
      Begin VB.Label lblGrupo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Cardapio"
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
         TabIndex        =   2
         Top             =   1185
         Width           =   1200
      End
      Begin VB.Image imgDivisaoGrupo 
         Height          =   1260
         Index           =   5
         Left            =   16620
         Picture         =   "inicio.frx":72FA
         Top             =   -60
         Width           =   45
      End
      Begin VB.Label lblGrupo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Pesquisa"
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
         Index           =   5
         Left            =   13300
         TabIndex        =   7
         Top             =   1160
         Width           =   1140
      End
      Begin VB.Label lblGrupo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Favoritos"
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
         Index           =   4
         Left            =   10800
         TabIndex        =   6
         Top             =   1160
         Width           =   1140
      End
      Begin VB.Label lblGrupo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Bebidas"
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
         Index           =   3
         Left            =   8300
         TabIndex        =   5
         Top             =   1160
         Width           =   1140
      End
      Begin VB.Label lblGrupo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Doces"
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
         Index           =   2
         Left            =   5800
         TabIndex        =   4
         Top             =   1160
         Width           =   1140
      End
      Begin VB.Label lblGrupo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Salgados"
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
         Index           =   1
         Left            =   3300
         TabIndex        =   3
         Top             =   1160
         Width           =   1140
      End
      Begin VB.Image imgDivisaoGrupo 
         Height          =   1260
         Index           =   4
         Left            =   12600
         Picture         =   "inicio.frx":7521
         Top             =   45
         Width           =   45
      End
      Begin VB.Image imgDivisaoGrupo 
         Height          =   1260
         Index           =   3
         Left            =   10100
         Picture         =   "inicio.frx":7748
         Top             =   45
         Width           =   45
      End
      Begin VB.Image imgDivisaoGrupo 
         Height          =   1260
         Index           =   2
         Left            =   7600
         Picture         =   "inicio.frx":796F
         Top             =   45
         Width           =   45
      End
      Begin VB.Image imgDivisaoGrupo 
         Height          =   1260
         Index           =   1
         Left            =   5100
         Picture         =   "inicio.frx":7B96
         Top             =   45
         Width           =   45
      End
      Begin VB.Image imgDivisaoGrupo 
         Height          =   1260
         Index           =   0
         Left            =   2600
         Picture         =   "inicio.frx":7DBD
         Top             =   45
         Width           =   45
      End
   End
   Begin VB.Frame frmTopo 
      Appearance      =   0  'Flat
      BackColor       =   &H003333A4&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      ForeColor       =   &H80000008&
      Height          =   1450
      Left            =   90
      TabIndex        =   0
      Top             =   0
      Width           =   30000
      Begin VB.Timer timerFoca 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   0
         Top             =   0
      End
      Begin VB.PictureBox cmdLogo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1410
         Left            =   -15
         Picture         =   "inicio.frx":7FE4
         ScaleHeight     =   1410
         ScaleWidth      =   2925
         TabIndex        =   18
         Top             =   -60
         Width           =   2925
      End
      Begin VB.Frame frmCPF 
         Appearance      =   0  'Flat
         BackColor       =   &H003333A4&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1335
         Left            =   7155
         TabIndex        =   15
         Top             =   45
         Width           =   4320
         Begin VB.Label lblMsgCPF 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Nota Fiscal Paulista"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00EBEBF5&
            Height          =   345
            Left            =   375
            TabIndex        =   17
            Top             =   300
            Width           =   3885
         End
         Begin VB.Label lblCPF 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Toque aqui para adicionar"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00EBEBF5&
            Height          =   330
            Left            =   375
            TabIndex        =   16
            Top             =   720
            Width           =   3885
         End
         Begin VB.Image Image7 
            Height          =   1260
            Left            =   0
            Picture         =   "inicio.frx":B27F
            Top             =   0
            Width           =   45
         End
      End
      Begin VB.Frame frmPagamento 
         Appearance      =   0  'Flat
         BackColor       =   &H003333A4&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1335
         Left            =   11805
         TabIndex        =   12
         Top             =   45
         Width           =   3435
         Begin VB.PictureBox cmdCarinho 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1185
            Left            =   1950
            Picture         =   "inicio.frx":B4A6
            ScaleHeight     =   1185
            ScaleWidth      =   1170
            TabIndex        =   20
            Top             =   45
            Width           =   1170
         End
         Begin VB.Label lblPrecoTotal 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "51,63"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   21.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00EBEBF5&
            Height          =   495
            Left            =   195
            TabIndex        =   14
            Top             =   180
            Width           =   1605
         End
         Begin VB.Label lblQTDETotal 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "6"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00EBEBF5&
            Height          =   405
            Left            =   195
            TabIndex        =   13
            Top             =   735
            Width           =   1605
         End
         Begin VB.Image Image5 
            Height          =   1260
            Left            =   0
            Picture         =   "inicio.frx":C616
            Top             =   0
            Width           =   45
         End
      End
   End
   Begin VB.Frame frmMenu 
      Appearance      =   0  'Flat
      BackColor       =   &H0000008C&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4245
      Left            =   1245
      TabIndex        =   8
      Top             =   1995
      Width           =   4455
      Begin VB.Label lblSelecionado 
         Appearance      =   0  'Flat
         BackColor       =   &H00B9B9FB&
         ForeColor       =   &H80000008&
         Height          =   555
         Left            =   1485
         TabIndex        =   11
         Top             =   3195
         Visible         =   0   'False
         Width           =   2000
      End
      Begin VB.Label lblCodigoMenu 
         BackStyle       =   0  'Transparent
         Caption         =   "lblCodigoMenu"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   500
         Index           =   0
         Left            =   870
         TabIndex        =   9
         ToolTipText     =   "1020"
         Top             =   1725
         Width           =   3000
      End
      Begin VB.Image cmdMenu 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   1500
         Index           =   0
         Left            =   735
         Stretch         =   -1  'True
         ToolTipText     =   "Sorvete"
         Top             =   660
         Width           =   3255
      End
      Begin VB.Label lblCodigoMenusSombra 
         BackStyle       =   0  'Transparent
         Caption         =   "lblCodigoMenuSombra"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   500
         Index           =   0
         Left            =   915
         TabIndex        =   10
         Top             =   2475
         Width           =   3000
      End
   End
   Begin VB.Label lblMSGProduto 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Nenhum produto foi encontrado"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00EBEBF5&
      Height          =   420
      Left            =   6765
      TabIndex        =   23
      Top             =   2310
      Width           =   5025
   End
   Begin VB.Line linhaLimite 
      BorderColor     =   &H000000FF&
      Visible         =   0   'False
      X1              =   15390
      X2              =   15390
      Y1              =   -90
      Y2              =   30030
   End
End
Attribute VB_Name = "frmControle"
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

Private abilitaMovimento As Boolean
Private botaoSelecionado As Boolean

Private posicaoToque As Double
Private posicaoBotaoToque As Double
Private subFamiliaSelecionado As Boolean
Private ultimoGrupoClick As Byte

Private qtdeBotaoVisivel As Byte
Private MSGBotaoCPF As String

Private RSFamilia As New ADODB.Recordset
Private RSProdSabores As New ADODB.Recordset

Dim ADO_CN_local As New ADODB.Connection
Dim ADO_RS_local As New ADODB.Recordset

'Private Const nBUFFER As Long = 1024

Public Sub carregaUltimoGrupo()
    cmdGrupo_Click (ultimoGrupoClick)
End Sub

Private Sub cmdLogo_KeyPress(KeyAscii As Integer)
    If digitoNumerico(KeyAscii) Then
        abilitaControle False
        Call frmPesquisa.pesquisar(Chr(KeyAscii))
    End If
End Sub

Private Sub Form_Activate()
    Call selecionaMenu(cmdMenu(0), False)
    abilitaControle True
    timerFoca.Enabled = True
End Sub



Private Sub Form_Load()
    
    propriedadeBotaoInicial

End Sub

Private Sub propriedadeBotaoInicial()

    Dim i As Byte
    Dim margem As Double
    
    cmdMenu(0).Width = BOTAOLARGURA
    cmdMenu(0).Height = BOTAOALTURA
    
    lblSelecionado.Width = BOTAOLARGURASELEC - 10
    lblSelecionado.Height = BOTAOALTURASELEC - 10
    
    frmTopo.Visible = True
    frmTopo.left = 0
    frmTopo.Top = 0
    
    frmGrupo.Visible = True
    frmGrupo.left = 0
    frmGrupo.Top = (Screen.Height - frmGrupo.Height)
    
    frmPagamento.left = Screen.Width - frmPagamento.Width
    frmCPF.left = frmPagamento.left - frmCPF.Width
    
    MSGBotaoCPF = "Toque aqui para adicionar"

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''''''''''''''''''''''' RETIRAR  '''''''''''''''''''''''
    criaNovoBotao 5, cmdGrupo
    For i = 0 To cmdGrupo.UBound
        margem = (((Screen.Width) / (cmdGrupo.UBound + 1)) * i) + 900
        cmdGrupo(i).left = margem
        imgDivisaoGrupo(i).left = margem + 2200
        lblGrupo(i).left = cmdGrupo(i).left
    Next i
    imgDivisaoGrupo(cmdGrupo.UBound).Visible = False
    ''''''''''''''''''''''' RETIRAR  '''''''''''''''''''''''
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    centroFormulario frmCarregando
    centroFormularioHeight frmCarregando
    frmCarregando.Visible = False
    
    lblMSGProduto.Visible = False
    centroFormulario lblMSGProduto
    centroFormularioHeight lblMSGProduto
    
    'Call CupomVerificaStatus
    cmdGrupo_Click 0
    limpaVariavel True
    
End Sub

Private Sub abilitaControle(ativa As Boolean)
    frmMenu.Enabled = ativa
    frmTopo.Enabled = ativa
    frmGrupo.Enabled = ativa
End Sub

Public Sub limpaVariavel(limpezaCompleta As Boolean)
    
    abilitaControle True
    
    frmMenu.left = MENUPOSICAOESQUERDA
    frmMenu.Top = MENUPOSICAOTOPO

    'If CupomStatus = True Then
        If vendaAberta = False Then
            lblMSGCPF.Caption = "Nota Fiscal Paulista"
            lblCPF.Caption = MSGBotaoCPF
            lblMSGCPF.ForeColor = CORFONTE
            lblCPF.Caption = MSGBotaoCPF
            lblCPF.ForeColor = CORFONTE
        'End If
    Else
        lblMSGCPF.Caption = "Toque aqui para cancelar"
        'lblMSGCPF.ForeColor = vbRed
        'lblCPF.Caption = "Toque aqui para reconectar"
        'lblCPF.ForeColor = vbRed
    End If
    'lblCPF.Font.Underline = True
    lblMSGProduto.Caption = "Nenhum produto foi encontrado"
    
    If limpezaCompleta Then
    
        lblPrecoTotal.Caption = "0,00"
        lblQTDETotal.Caption = "0"
        
        ultimoGrupoClick = 1
        carregaUltimoGrupo
        vendaAberta = False
    
    End If
    
    statusVendaAberta
    
End Sub

Private Function validaMovimento()

    If frmMenu.Width > Screen.Width And abilitaMovimento = True Then
        validaMovimento = True
    End If

End Function

Private Sub chameleonButton4_Click(Index As Integer)
    End
End Sub

Private Sub cmdCarinho_Click()
    'frmFinalizaCompra.carregaTamanho wNumeroCupom
        frmPagamento_Click
    
End Sub

Private Sub cmdGrupo_Click(Index As Integer)

    Dim i As Byte
    
    If RSProdSabores.State <> 0 Then RSProdSabores.Close
    subFamiliaSelecionado = False
    carregaIMGGrupo Index

    Select Case Index
    
    Case 0 'CARDAPIO
        'carregaIMGGrupo Index
        If precoEspecial <> 1 Then
            precoEspecial = 1
        Else
            precoEspecial = 2
        End If
        
    Case 1, 2, 3
    
        frmMenu.Visible = False
        frmControle.Refresh
        frmCarregando.Visible = True
        frmControle.Refresh
    
        limpaVariavel False
        'carregaGrupo Format(Index, "00")
        carregaFamilia ""
        'carregaIMGGrupo Index
                
        frmControle.Refresh
        frmCarregando.Visible = False
        frmControle.Refresh
        
    Case 4
        carregaFamilia "999999999999"
        
    Case 5 'PESQUISA
        frmTopo.Enabled = False
        frmMenu.Enabled = False
        frmGrupo.Enabled = False
        
        frmPesquisa.pesquisar ("")
                
    End Select
    
    ultimoGrupoClick = Index
    
    lblSelecionado.Visible = False
    'frmMenu.Visible = True
    
End Sub

Private Sub carregaIMGGrupo(Index As Integer)
    Dim i As Byte
    
    If Index = 0 Then
        If precoEspecial <> 1 Then
            cmdGrupo(0).Picture = LoadPicture(pastaAtual & ENDPASTAIMG & "btGrupo00N")
        Else
            cmdGrupo(0).Picture = LoadPicture(pastaAtual & ENDPASTAIMG & "btGrupo00P")
        End If
    Else
    
        For i = 1 To cmdGrupo.UBound
            cmdGrupo(i).Picture = LoadPicture(pastaAtual & ENDPASTAIMG & "btGrupo0" & i & "N")
        Next i
        cmdGrupo(Index).Picture = LoadPicture(pastaAtual & ENDPASTAIMG & "btGrupo0" & Index & "P")
    
    End If
    
End Sub

Public Sub carregaFamilia(whereGrupo As String)

    Dim sql As scriptSQL
    Dim i As Byte
    
    sql.select = "select rtrim(FAP_CodigoFamilia) as codigoFamilia," & _
                 vbNewLine & "Upper(rtrim(FAP_Descricao)) as Descricao, " & _
                 vbNewLine & "fap_complementoAdicional as complementoAdicional, " & _
                 vbNewLine & "rtrim(FAP_SubFamilia) as SubFamilia"
    sql.from = "from FamiliaProduto"
    
    If whereGrupo <> Empty Then
        sql.where = "where FAP_CodigoFamilia in (" & whereGrupo & ")"
    End If
    
    If RSFamilia.State <> 0 Then RSFamilia.Close
    comandoSQL RSFamilia, sql
    
    qtdeBotaoVisivel = RSFamilia.RecordCount
    
    If qtdeBotaoVisivel > 0 Then
        lblMSGProduto.Visible = False
        frmMenu.Visible = True
        
        carregaPosicaoBotao qtdeBotaoVisivel - 1
        
        For i = 0 To qtdeBotaoVisivel - 1
            carregaInformacaoBotao RSFamilia("Descricao"), "", i
            carregaIMGBotao cmdMenu(i), RSFamilia("codigoFamilia")
            RSFamilia.MoveNext
        Next i
    Else
        frmMenu.Visible = False
        lblMSGProduto.Visible = True
    End If
    
    'RSFamilia.Close
    
End Sub

Private Sub carregaInformacaoBotao(descricao As String, codigo As String, Index As Byte)
       
    If Len(descricao) > 27 Then
        lblCodigoMenu(Index).Height = lblCodigoMenu(Index).Height + 250
        lblCodigoMenu(Index).Top = lblCodigoMenu(Index).Top - 250
        lblCodigoMenusSombra(Index).Top = lblCodigoMenusSombra(Index).Top - 250
    End If
            
    lblCodigoMenu(Index).Caption = codigo & vbNewLine & descricao
    lblCodigoMenu(Index).ToolTipText = descricao
    
    lblCodigoMenusSombra(Index).Caption = lblCodigoMenu(Index).Caption
    lblCodigoMenusSombra(Index).ToolTipText = lblCodigoMenu(Index).ToolTipText
    
    cmdMenu(Index).ToolTipText = lblCodigoMenu(Index).Caption
    
End Sub

Private Sub carregaGrupo(grupo As String)
    
    carregaFamilia grupo
    
    
    
End Sub

Private Sub cmdLogo_Click()
    frmInicio.Show
End Sub

Private Sub cmdMenu_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    'Toque Simples
    botaoSelecionado = True
    lblSelecionado.Visible = False

    'Movimentação
    posicaoToque = (Posicao.X * 15)
    posicaoBotaoToque = frmMenu.left
    abilitaMovimento = True
    
End Sub

Private Sub cmdMenu_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    'If Button = 1 Or Button Then
        GetCursorPos Posicao
        
        If validaMovimento Then
            rolagemBotao
        End If
        
        botaoSelecionado = False
    'End If
 
End Sub

Private Sub cmdMenu_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    abilitaMovimento = False
    
    If botaoSelecionado Then
        
        If subFamiliaSelecionado Then
        
            RSProdSabores.MoveFirst
            RSProdSabores.Move Index
        
        Else
        
            RSFamilia.MoveFirst
            RSFamilia.Move Index
        
        End If
        
                carregaMenuOpcao RSFamilia("Descricao"), _
                         RSFamilia("codigoFamilia"), _
                         RSFamilia("SubFamilia"), _
                         RSFamilia("ComplementoAdicional"), _
                         Index
        
    End If
    
    
End Sub

Private Sub carregaMenuOpcao(Nome As String, codigoFamilia As String, subFamilia As String, _
                            complementoAdicional As Integer, Index As Integer)
    
    If complementoAdicional = 0 Then
            Call selecionaMenu(cmdMenu(Index), True)
            frmMenuOpcao.defineDescricao Nome
            frmMenuOpcao.carregaTamanho codigoFamilia
            
            ativaFormulario frmMenuOpcao
    End If
    
    If subFamiliaSelecionado = False Then
    
        frmMenu.Visible = False
        frmCarregando.Visible = True
        frmControle.Refresh
    
        If complementoAdicional = 1 Then
            If Not carregaSabores(subFamilia, complementoAdicional) Then
                lblMSGProduto.Caption = "Erro de cadastro! Produto " & Chr(34) & codigoFamilia & Chr(34) & " não foi encontrado"
                lblMSGProduto.Visible = True
                frmMenu.Visible = False
            End If
        End If
        
        frmMenu.Visible = True
        frmCarregando.Visible = False
        
    Else
            Call selecionaMenu(cmdMenu(Index), True)
            frmMenuOpcao.defineDescricao Nome
            frmMenuOpcao.carregaTamanho codigoFamilia
            
            ativaFormulario frmMenuOpcao
    End If
    
End Sub

Private Function carregaSabores(codigoFamilia As String, codigoComplementoAdicional As Integer) As Boolean

    Dim sql As scriptSQL
    Dim i As Byte
    
    subFamiliaSelecionado = True
    
    sql.select = "select LOWER(rtrim(pro_descricao)) as Descricao," & vbNewLine & _
                 "PRO_PrecoVenda1 as PrecoVenda, " & vbNewLine & _
                 "PRO_codigo as codigoFamilia"
    sql.from = "from produtoLoja"
    sql.where = "where PRO_Familia = '" & codigoFamilia & "'"
    
    comandoSQL RSProdSabores, sql
    
    If RSProdSabores.RecordCount > 0 Then
        carregaPosicaoBotao RSProdSabores.RecordCount - 1
        
        For i = 0 To RSProdSabores.RecordCount - 1
            carregaInformacaoBotao RSProdSabores("Descricao"), RSProdSabores("CodigoFamilia"), i
            carregaIMGBotao cmdMenu(i), RSFamilia("codigoFamilia")
            RSProdSabores.MoveNext
        Next i
        
        carregaSabores = True
        
    Else
    
        carregaSabores = False
        
    End If
    
End Function

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

Public Sub statusVendaAberta()
    If vendaAberta Then
        cmdCarinho.Picture = LoadPicture(pastaAtual & ENDPASTAIMG & "btCarrinhoP")
        If lblCPF <> MSGBotaoCPF Then
            lblMSGCPF.Caption = "Toque aqui para cancelar"
        End If
    Else
        cmdCarinho.Picture = LoadPicture(pastaAtual & ENDPASTAIMG & "btCarrinhoN")
    End If
End Sub

Public Sub adicionaValoresVenda(preco As Double)

    Dim valorAtual As Double
    
    valorAtual = lblPrecoTotal.Caption
    
    lblPrecoTotal.Caption = formataPreco(valorAtual + preco)
    lblQTDETotal.Caption = lblQTDETotal + 1
    
End Sub

Public Sub adicionaCPF(CPF As String)
    lblCPF.ToolTipText = CPF
    lblCPF.Caption = Format(CPF, "###\.###\.###\-##")
    lblCPF.Font.Underline = False
End Sub

Public Sub ativaFormulario(formulario As Form)
    abilitaControle False
    formulario.Show
End Sub


Private Sub carregaIMGBotao(botao, codigoFamilia As String)
    
    Dim Arquivo As String
    Dim enderecoArquivo As String
    
    enderecoArquivo = pastaAtual & ENDPASTATEMP & codigoFamilia
    Arquivo = Dir(enderecoArquivo, vbDirectory)
    
    If subFamiliaSelecionado Then
        enderecoArquivo = pastaAtual & ENDPASTAIMG & nomeIMGPadrao & "2"
    ElseIf Arquivo = Empty Then
        enderecoArquivo = pastaAtual & ENDPASTAIMG & nomeIMGPadrao
    Else
        enderecoArquivo = enderecoArquivo
    End If
    
    botao.Picture = LoadPicture(enderecoArquivo)
    
End Sub

'Public Function LoadPictureFromDB(RS As ADODB.Recordset)
'
'    Dim i As Byte
'
'    For i = 0 To 18
'        Set Me.cmdMenu(i) = ExibeImagensGrandes(RS.Fields("imagem"))
'        RS.MoveNext
'    Next i
'
''    If RS Is Nothing Then
''        GoTo procNoPicture
''    End If
''
''    Set strStream = New ADODB.Stream
''    strStream.Type = adTypeBinary
''    strStream.Open
''
''    strStream.Write RS.Fields("imagem").Value
''
''
''    strStream.SaveToFile "C:\Temp.bmp", adSaveCreateOverWrite
''    Image1.Picture = LoadPicture("C:\Temp.bmp")
''    Kill ("C:\Temp.bmp")
''    LoadPictureFromDB = True
'
'End Function

'Public Function ExibeImagensGrandes(F As ADODB.Field) As StdPicture
'
'    Dim B()      As Byte
'    Dim ff       As Long
'    Dim File     As String
'    Dim i        As Long
'    Dim FileLen  As Long
'    Dim Blocks   As Long
'    Dim LeftOver As Long
'
'    On Error GoTo ErrHandler
'    File = "temp.jpg"
'    ff = FreeFile
'    Open File For Binary Access Write As ff
'    Blocks = Int(F.ActualSize / nBUFFER)
'    LeftOver = F.ActualSize Mod nBUFFER
'    B() = F.GetChunk(LeftOver)
'    Put ff, , B()
'    For i = 1 To Blocks
'        B() = F.GetChunk(nBUFFER)
'        Put ff, , B()
'    Next
'    Close ff
'    Erase B
'    Set ExibeImagensGrandes = LoadPicture(File)
'    Kill File
'    Exit Function
'
'ErrHandler:
'    MsgBox "ERROR: " & Err.Description
'End Function

Private Sub abrirConexao()
    ADO_CN_local.Provider = "SQLOLEDB"
    ADO_CN_local.Properties("Data Source").Value = "svdmac"
    ADO_CN_local.Properties("Initial Catalog").Value = "desenv_dmac_loja"
    ADO_CN_local.Properties("User ID").Value = "felipelima"
    ADO_CN_local.Properties("Password").Value = "fl"
    ADO_CN_local.Open
End Sub

Private Function RolagemLimitador(limitador As Integer) As Boolean
    If limitador >= 800 Then
        RolagemLimitador = True
    Else
        RolagemLimitador = False
    End If
End Function

Private Sub rolagemBotao()

    Dim i As Byte
    Dim movimenta As Double
    'Dim leftAnterior As Double
    
    movimenta = (((Posicao.X) * 15) - posicaoToque) - (frmMenu.left - posicaoBotaoToque)
    If (frmMenu.left + movimenta) > MENUPOSICAOESQUERDA Then
        movimenta = (MENUPOSICAOESQUERDA - frmMenu.left)
    ElseIf (frmMenu.left + frmMenu.Width) + movimenta < (Screen.Width) Then
        movimenta = (Screen.Width - (frmMenu.left + frmMenu.Width))
    End If
        
    frmMenu.left = frmMenu.left + movimenta
     
End Sub




Private Sub carregaPosicaoBotao(qtdeBotao As Byte)

    Dim colunaBotao As Byte
    Dim botao As Byte
    Dim left As Double
    
    Dim i As Byte
    Dim j As Byte
    
    criaNovoBotao qtdeBotao, lblCodigoMenu
    criaNovoBotao qtdeBotao, lblCodigoMenusSombra
    criaNovoBotao qtdeBotao, cmdMenu
    
    frmMenu.Width = (MARGEMLINHA * (((qtdeBotao) \ 4) + 1)) + 200
    frmMenu.Height = (MARGEMCOLUNA * (BOTAONUMLINHAS)) + 100
    
    For j = 0 To qtdeBotao Step BOTAONUMLINHAS
        
        For i = 0 To (BOTAONUMLINHAS - 1)
            botao = i + (colunaBotao * BOTAONUMLINHAS)
            If botao <= qtdeBotao Then
            
                left = MARGEMLINHA
                left = left * colunaBotao
                left = left + BOTAOPOSICAOESQUERDA
                
                cmdMenu(botao).left = left
                cmdMenu(botao).Top = BOTAOPOSICAOTOPO + (MARGEMCOLUNA * i)
                
                lblCodigoMenu(botao).left = cmdMenu(botao).left + BOTAOTEXTOPOSICAOESQUERDA
                lblCodigoMenu(botao).Top = cmdMenu(botao).Top + BOTAOTEXTOPOSICAOTOPO
                
                lblCodigoMenusSombra(botao).left = lblCodigoMenu(botao).left + 15
                lblCodigoMenusSombra(botao).Top = lblCodigoMenu(botao).Top + 15
                
            End If
        Next i
        
        colunaBotao = colunaBotao + 1
    Next j
    
End Sub


Private Sub abilitaFrameControle(ativa As Boolean)
    frmMenu.Enabled = ativa
    frmTopo.Enabled = ativa
    frmGrupo.Enabled = ativa
End Sub

Private Sub Form_LostFocus()
    timerFoca.Enabled = False
End Sub

Private Sub frmCPF_Click()

    'If CupomStatus = False Then
        'lblCPF.Caption = "TESTANDO CONEXÃO"
        'lblCPF.Font.Underline = False
        'lblCPF.Refresh
        'Call CupomVerificaStatus
        'limpaVariavel False
        
    If vendaAberta Then
        If frmLogin.validaLogin(glbUsuarioCodigo, True, True, _
        "Cancelar venda atual", True) Then
            'Form_Deactivate
            CupomCancela
            limpaVariavel True
        End If
    ElseIf lblCPF.Caption = MSGBotaoCPF Then
        ativaFormulario frmCPFNota
    End If
        
    'If
        
    'End If
    
    'End If
End Sub

Private Sub frmPagamento_Click()
    If vendaAberta Then
        If Val(lblQTDETotal.Caption) > 0 Then
            frmFinalizaCompra.carregaTamanho wNumeroCupom
            ativaFormulario frmFinalizaCompra
       End If
       'ativaFormulario frmFinalizaCompra
    Else
        'ativaFormulario frmFinalizaCompra
       ativaFormulario frmFuncoes
    End If
End Sub

Private Sub Label3_Click()
End Sub

Private Sub frmSubMenu_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub Label2_Click()

End Sub

Private Sub lblCodigoMenu_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call cmdMenu_MouseDown(Index, Button, Shift, X, Y)
End Sub

Private Sub lblCodigoMenu_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call cmdMenu_MouseMove(Index, Button, Shift, X, Y)
End Sub

Private Sub lblCodigoMenu_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call cmdMenu_MouseUp(Index, Button, Shift, X, Y)
End Sub

Private Sub lblCPF_Click()
    frmCPF_Click
End Sub

Private Sub lblMsgCPF_Click()
    frmCPF_Click
End Sub

Private Sub lblPrecoTotal_Click()
    frmPagamento_Click
End Sub

Private Sub lblQuantidadeTotal_Click()
    frmPagamento_Click
End Sub

Private Sub lblQTDETotal_Click()
    frmPagamento_Click
End Sub

'Private Sub timerMovimentoToque_Timer()
'    GetCursorPos posicao
'    CurrentX = 0
'    CurrentY = 0
'    logPonteiroTela.Caption = (posicao.x) * 15
'End Sub


Private Sub timerFoca_Timer()
    
End Sub
