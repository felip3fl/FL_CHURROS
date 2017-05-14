VERSION 5.00
Object = "{D76D7130-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7d.ocx"
Begin VB.Form frmMenuOpcao 
   Appearance      =   0  'Flat
   BackColor       =   &H0004047B&
   BorderStyle     =   0  'None
   Caption         =   "frmMenuOpcao"
   ClientHeight    =   11130
   ClientLeft      =   11805
   ClientTop       =   480
   ClientWidth     =   5550
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   ScaleHeight     =   11130
   ScaleWidth      =   5550
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmrAnimacao 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   3450
      Top             =   360
   End
   Begin VB.Frame frmQuantidade 
      Appearance      =   0  'Flat
      BackColor       =   &H0004047B&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1620
      Left            =   295
      TabIndex        =   4
      Top             =   2200
      Width           =   5000
      Begin VB.Label cmdQTDE 
         Alignment       =   2  'Center
         BackColor       =   &H0004047B&
         Caption         =   "+ 1"
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
         Height          =   735
         Index           =   5
         Left            =   3975
         TabIndex        =   15
         ToolTipText     =   "1"
         Top             =   705
         Width           =   600
      End
      Begin VB.Label cmdQTDE 
         Alignment       =   2  'Center
         BackColor       =   &H0004047B&
         Caption         =   "2"
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
         Height          =   735
         Index           =   1
         Left            =   870
         TabIndex        =   14
         ToolTipText     =   "2"
         Top             =   705
         Width           =   600
      End
      Begin VB.Label cmdQTDE 
         Alignment       =   2  'Center
         BackColor       =   &H0004047B&
         Caption         =   "4"
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
         Height          =   735
         Index           =   3
         Left            =   2430
         TabIndex        =   13
         ToolTipText     =   "4"
         Top             =   705
         Width           =   600
      End
      Begin VB.Label cmdQTDE 
         Alignment       =   2  'Center
         BackColor       =   &H0004047B&
         Caption         =   "3"
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
         Height          =   735
         Index           =   2
         Left            =   1650
         TabIndex        =   12
         ToolTipText     =   "3"
         Top             =   705
         Width           =   600
      End
      Begin VB.Label cmdQTDE 
         Alignment       =   2  'Center
         BackColor       =   &H0004047B&
         Caption         =   "5"
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
         Height          =   735
         Index           =   4
         Left            =   3195
         TabIndex        =   11
         ToolTipText     =   "5"
         Top             =   705
         Width           =   600
      End
      Begin VB.Label cmdQTDE 
         Alignment       =   2  'Center
         BackColor       =   &H0004047B&
         Caption         =   "1"
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
         Height          =   735
         Index           =   0
         Left            =   100
         TabIndex        =   10
         ToolTipText     =   "1"
         Top             =   700
         Width           =   600
      End
      Begin VB.Label lblExpandirQuantidade 
         BackStyle       =   0  'Transparent
         Caption         =   "Quantidade"
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
         Left            =   105
         TabIndex        =   5
         Top             =   195
         Width           =   4000
      End
      Begin VB.Line Line4 
         BorderColor     =   &H003333A4&
         X1              =   105
         X2              =   4800
         Y1              =   105
         Y2              =   105
      End
   End
   Begin VB.Frame frmTamanho 
      Appearance      =   0  'Flat
      BackColor       =   &H0004047B&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2865
      Left            =   295
      TabIndex        =   2
      Top             =   4200
      Width           =   5000
      Begin VB.Label cmdTamanhoPreco 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0004047B&
         Caption         =   "1,99"
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
         Height          =   735
         Index           =   0
         Left            =   4005
         TabIndex        =   16
         Top             =   705
         Width           =   795
      End
      Begin VB.Label cmdTamanho 
         BackColor       =   &H0004047B&
         Caption         =   "1,99"
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
         Height          =   735
         Index           =   0
         Left            =   105
         TabIndex        =   9
         Top             =   705
         Width           =   3915
      End
      Begin VB.Line Line3 
         BorderColor     =   &H003333A4&
         X1              =   100
         X2              =   4800
         Y1              =   100
         Y2              =   100
      End
      Begin VB.Label lblExpandirTamanho 
         BackStyle       =   0  'Transparent
         Caption         =   "Tamanho"
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
         Left            =   105
         TabIndex        =   3
         Top             =   195
         Width           =   4000
      End
   End
   Begin VB.Frame frmSabores 
      Appearance      =   0  'Flat
      BackColor       =   &H0004047B&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1995
      Left            =   295
      TabIndex        =   0
      Top             =   7000
      Visible         =   0   'False
      Width           =   5000
      Begin VB.Label cmdSaboresPreco 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0004047B&
         Caption         =   "1,99"
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
         Height          =   735
         Index           =   0
         Left            =   4005
         TabIndex        =   8
         Top             =   700
         Width           =   795
      End
      Begin VB.Label cmdSabores 
         BackColor       =   &H0004047B&
         Caption         =   "Label1"
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
         Height          =   735
         Index           =   0
         Left            =   105
         TabIndex        =   7
         Top             =   700
         Width           =   3915
      End
      Begin VB.Label lblExpandirSabores 
         BackStyle       =   0  'Transparent
         Caption         =   "Sabores (Qtde.: 1)"
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
         Left            =   105
         TabIndex        =   1
         Top             =   195
         Width           =   4000
      End
      Begin VB.Line Line5 
         BorderColor     =   &H003333A4&
         X1              =   100
         X2              =   4800
         Y1              =   100
         Y2              =   100
      End
   End
   Begin VSFlex7DAOCtl.VSFlexGrid grdPreco 
      Height          =   1140
      Left            =   390
      TabIndex        =   17
      Top             =   700
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
      FormatString    =   $"frmMenuOpcao.frx":0000
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
      Left            =   0
      TabIndex        =   18
      Top             =   9735
      Width           =   3435
   End
   Begin VB.Label lblMenuNomeItem 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MenuNomeItem"
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
      Left            =   395
      TabIndex        =   6
      Top             =   200
      Width           =   2685
   End
   Begin VB.Line borda1 
      BorderColor     =   &H00040460&
      X1              =   45
      X2              =   45
      Y1              =   90
      Y2              =   30210
   End
End
Attribute VB_Name = "frmMenuOpcao"
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

Dim RSProdTamanho As New ADODB.Recordset
Dim RSProdSabores As New ADODB.Recordset
Dim complementoAdicional As Byte

Public Sub carregaTamanho(subFamilia As String)

    Dim sql As scriptSQL
    
    If RSProdTamanho.State <> 0 Then RSProdTamanho.Close
    
    sql.select = "select LOWER(rtrim(pro_descricaoMenu)) as DescricaoMenu," & vbNewLine & _
                 "LOWER(rtrim(pro_descricao)) as Descricao, " & vbNewLine & _
                 "PRO_PrecoVenda" & precoEspecial & " as PrecoVenda, " & vbNewLine & _
                 "PRO_codigo as codigo"
    sql.from = "from produtoLoja"
    sql.where = "where PRO_Familia = '" & subFamilia & "'"
    
    comandoSQL RSProdTamanho, sql
    
    If RSProdTamanho.RecordCount > 0 Then
        Call carregaPosicaoBotaoMenu(cmdTamanho, RSProdTamanho.RecordCount - 1)
        Call carregaPosicaoBotaoMenu(cmdTamanhoPreco, RSProdTamanho.RecordCount - 1)
        carregaInfTamanho
    End If
    
End Sub

Private Sub defineQuantidade(QTDE As Integer)
    lblExpandirSabores.Caption = "Sabores (Qtde.: " & QTDE & ")"
End Sub

Private Sub carregaPosicaoBotaoMenu(botao, qtdeBotaoProduto As Byte)

    Dim i As Byte
    'Dim j As Byte

    criaNovoBotao qtdeBotaoProduto, botao
    
    For i = 1 To qtdeBotaoProduto
        botao(i).Top = botao(0).Top + ((botao(0).Height + 15) * i)
        botao(i).left = botao(0).left
    Next i
    
End Sub

Private Sub carregaInfTamanho()
    
    Dim i As Byte
    Dim descricao As String
    Dim precoVenda As String
    
    Do While Not RSProdTamanho.EOF
                
        descricao = RSProdTamanho("DescricaoMenu")
        precoVenda = RSProdTamanho("precoVenda")
                
        adicionaDescricaoMenu cmdTamanho(i), descricao, ""
        If precoVenda <> 0 Then _
            adicionaDescricaoMenu cmdTamanhoPreco(i), formataPreco(precoVenda), ""
        
        i = i + 1
        RSProdTamanho.MoveNext
    Loop
    
    RSProdTamanho.MoveFirst
    
End Sub

Private Sub carregaInfSabores()
    
    Dim i As Byte
    Dim descricao As String
    Dim precoVenda As String
    Dim codigo As String
    
    Do While Not RSProdSabores.EOF
                
        descricao = RSProdSabores("Descricao")
        precoVenda = RSProdSabores("PrecoVenda")
        codigo = RSProdSabores("codigo")

        adicionaDescricaoMenu cmdSabores(i), descricao, codigo
        'If precoVenda <> 0 Then _
            'adicionaDescricaoMenu cmdSaboresPreco(i), formataPreco(precoVenda), ""
        
        i = i + 1
        RSProdSabores.MoveNext
    Loop
    
End Sub


Private Sub adicionaDescricaoMenu(botao, ByRef descricao As String, codigo As String)
    
    Dim espacoEsq As Byte
    Dim espacoDir As Byte
    
    If botao.Alignment = 0 Then espacoEsq = 3
    If botao.Alignment = 1 Then espacoDir = 3
    
    If codigo <> Empty Then codigo = codigo & " "
    botao.Caption = vbNewLine & Space(espacoEsq) & codigo & descricao & Space(espacoDir)
                    
    botao.ToolTipText = codigo
    
End Sub

Private Sub cmdAdiciona_Click()
    
    frmMenuOpcao.Enabled = False
    
    Dim botaoSelecionado As Integer
    Dim vendaValida As Boolean
    
    If vendaAberta = False Then
        MSGBotaoCarregando Me, cmdAdiciona, "ABRINDO VENDA"
        If CupomAbertura("") Then
            vendaAberta = True
            frmControle.statusVendaAberta
        End If
    End If
    
    MSGBotaoCarregando Me, cmdAdiciona, "ENVIANDO CUPOM"
    
'    Select Case complementoAdicional
'    Case 0
        CupomCriaItem RSProdTamanho("Descricao"), RSProdTamanho("Codigo"), "FF", _
              QTDESelecionado, RSProdTamanho("PrecoVenda")
        
        gravaVendaBD wNumeroCupom, "CF", obterItem, RSProdTamanho("Codigo"), _
             CStr(precoEspecial), QTDESelecionado, RSProdTamanho("PrecoVenda"), ""
             
        frmControle.adicionaValoresVenda grdPreco.TextMatrix(2, 1)
        frmControle.carregaUltimoGrupo
        
'    Case 1
'        botaoSelecionado = menuSelecionado(0, cmdSabores)
'        If botaoSelecionado >= 0 Then
'
'
'            RSProdSabores.MoveFirst
'            RSProdSabores.Move botaoSelecionado
'
'            CupomCriaItem RSProdTamanho("Descricao") & " " & RSProdSabores("Descricao"), _
'                  RSProdTamanho("Codigo"), "FF", _
'                  QTDESelecionado, Replace(RSProdTamanho("PrecoVenda"), ",", ".")
'
'            gravaVendaBD wNumeroCupom, "CF", obterItem, RSProdTamanho("Codigo"), _
'                CStr(precoEspecial), QTDESelecionado, RSProdTamanho("PrecoVenda"), RSProdSabores("Codigo")
'
'            frmControle.adicionaValoresVenda RSProdTamanho("PrecoVenda")
'
'        End If
'    Case 9
'
'    End Select
    
    
    frmMenuOpcao.Enabled = True
    Form_Deactivate
    
End Sub

Private Function obterItem() As Integer
    obterItem = Val(frmControle.lblQTDETotal.Caption) + 1
End Function

Private Sub gravaVendaBD(notaFiscal As String, serie As String, _
                         Item As String, codigoProduto As String, _
                         codigoCardapio As String, QTDE As String, _
                         preco As String, subProduto As String)
                         
    Dim sql As scriptSQL
                         
    sql.insert = "insert ItensVenda ("
    sql.insert = sql.insert & vbNewLine & "itv_NotaFiscal,"
    sql.insert = sql.insert & vbNewLine & "itv_Serie,"
    sql.insert = sql.insert & vbNewLine & "itv_Item,"
    sql.insert = sql.insert & vbNewLine & "itv_Loja,"
    sql.insert = sql.insert & vbNewLine & "itv_CodigoProduto,"
    sql.insert = sql.insert & vbNewLine & "itv_CodigoCardapio,"
    sql.insert = sql.insert & vbNewLine & "itv_Quantidade,"
    sql.insert = sql.insert & vbNewLine & "itv_PrecoUnitario,"
    sql.insert = sql.insert & vbNewLine & "itv_SubProduto)"
    
    sql.insert = sql.insert & vbNewLine & "values ("
    sql.insert = sql.insert & vbNewLine & "" & notaFiscal & ","
    sql.insert = sql.insert & vbNewLine & "'" & serie & "',"
    sql.insert = sql.insert & vbNewLine & "" & Item & ","
    sql.insert = sql.insert & vbNewLine & "'" & glbCodigoLoja & "',"
    sql.insert = sql.insert & vbNewLine & "" & codigoProduto & ","
    sql.insert = sql.insert & vbNewLine & "" & codigoCardapio & ","
    sql.insert = sql.insert & vbNewLine & "" & QTDE & ","
    sql.insert = sql.insert & vbNewLine & "" & formataPrecoBD(preco) & ","
    sql.insert = sql.insert & vbNewLine & "'" & subProduto & "')"
    
    Call insercaoSQL(sql)
    
End Sub

Private Function QTDESelecionado() As Integer
    QTDESelecionado = grdPreco.TextMatrix(0, 1)
End Function

Private Sub cmdQTDE_Click(Index As Integer)
    selecionaBotao cmdQTDE, Index, False
    If Index = 5 Then
        grdPreco.TextMatrix(0, 1) = grdPreco.TextMatrix(0, 1) + 1
    Else
        grdPreco.TextMatrix(0, 1) = cmdQTDE(Index).ToolTipText
    End If
    calculaGrid
End Sub

Private Sub cmdSabores_Click(Index As Integer)
    selecionaBotao cmdSabores, Index, True
    selecionaBotao cmdSaboresPreco, Index, True
End Sub

Private Sub cmdTamanho_Click(Index As Integer)
    selecionaBotao cmdTamanho, Index, False
    selecionaBotao cmdTamanhoPreco, Index, False
    
    RSProdTamanho.MoveFirst
    RSProdTamanho.Move Index
    
    grdPreco.TextMatrix(1, 1) = formataPreco(RSProdTamanho("PrecoVenda"))
    
    calculaGrid
End Sub

Private Sub calculaGrid()
    Dim QTDE As Double
    Dim Valor As Double
    
    QTDE = grdPreco.TextMatrix(0, 1)
    Valor = grdPreco.TextMatrix(1, 1)
    grdPreco.TextMatrix(2, 1) = formataPreco(QTDE * Valor)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'RSProdTamanho.Close
    'RSProdSabores.Close
End Sub

Private Sub lblExpandirSabores_Click()
    expandirMenu frmSabores, Me
End Sub

Private Sub Form_Activate()
    ativaFormulario = True
    tmrAnimacao.Enabled = True
    limpaCampos
    grdPreco.SetFocus
End Sub

Private Function menuSelecionado(inicioPesquisa As Byte, ByRef botao) As Integer
    Dim i As Byte
    
    menuSelecionado = -1
    For i = 0 To botao.UBound
        If botao(i).BackColor = CORBOTAOPRESIONADO Then
            menuSelecionado = i
            Exit For
        End If
    Next i
End Function

Private Sub Form_Deactivate()

    'If RSProdTamanho.State <> 0 Then RSProdTamanho.Close
    'If RSProdSabores.State <> 0 Then RSProdSabores.Close

    ativaFormulario = False
    tmrAnimacao.Enabled = True
End Sub

Private Sub Form_Load()
    ajustaMenu Me
    ajustaMenuComponentes
End Sub

Public Sub defineDescricao(descricao As String)
    lblMenuNomeItem.Caption = descricao
End Sub

Private Sub ajustaMenuComponentes()

    Dim i As Byte
    
    For i = 0 To cmdQTDE.UBound - 1
        cmdQTDE(i).Caption = vbNewLine & (i + 1)
    Next i
    cmdQTDE(cmdQTDE.UBound).Caption = vbNewLine & "+ 1"

    cmdAdiciona.left = 0
    cmdAdiciona.Top = Me.Height - cmdAdiciona.Height
    cmdAdiciona.Width = Me.Width
    
    cmdTamanho(0).Caption = ""
    cmdTamanhoPreco(0).Caption = ""
    cmdSabores(0).Caption = ""
    cmdSaboresPreco(0).Caption = ""
End Sub

Private Sub tmrAnimacao_Timer()
    'If ativaFormulario = True Then
        'abilitaFormularioAnima Me
    'Else
        abilitaFormularioAnima ativaFormulario, Me, False
    'End If
End Sub

Private Sub limpaCampos()
    cmdQTDE_Click (0)
    cmdTamanho_Click (0)
    MSGBotaoNormal Me, cmdAdiciona, "Adiciona"
    'cmdAdiciona.Caption = vbNewLine & "Adiciona"
    'cmdAdiciona.ForeColor = CORFONTEBOTAO
End Sub
