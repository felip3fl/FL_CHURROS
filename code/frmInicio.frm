VERSION 5.00
Begin VB.Form frmInicio 
   BackColor       =   &H003333A4&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4380
   ClientLeft      =   7980
   ClientTop       =   3795
   ClientWidth     =   7515
   ControlBox      =   0   'False
   Icon            =   "frmInicio.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "frmInicio.frx":628A
   ScaleHeight     =   4380
   ScaleWidth      =   7515
   ShowInTaskbar   =   0   'False
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Projeto Churros v 0.27"
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
      Height          =   270
      Left            =   2370
      TabIndex        =   0
      Top             =   3150
      Width           =   2355
   End
End
Attribute VB_Name = "frmInicio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

Private Sub Form_DblClick()
    Unload Me
End Sub

Private Sub Form_Load()

    Me.Height = 4000
    Me.Width = 7000

    centroFormulario Me
    centroFormularioHeight Me
    
    Me.Caption = App.ProductName
    
    'imgLogo.left = (Me.Width / 2) - (imgLogo.Width / 2)
    'imgLogo.top = (Me.Height / 2) - (imgLogo.Height / 2)
    
    lblInfo.Top = Me.Height - 1000
    lblInfo.left = (Me.Width / 2) - (lblInfo.Width / 2)
    
End Sub



