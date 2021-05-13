VERSION 5.00
Object = "{6FBA474E-43AC-11CE-9A0E-00AA0062BB4C}#1.0#0"; "SYSINFO.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.MDIForm frmMain 
   AutoShowChildren=   0   'False
   BackColor       =   &H00E0E0E0&
   Caption         =   "Pedidos de Venda"
   ClientHeight    =   8295
   ClientLeft      =   75
   ClientTop       =   585
   ClientWidth     =   11880
   Icon            =   "frmMain.frx":0000
   LinkMode        =   1  'Source
   LinkTopic       =   "Main"
   Picture         =   "frmMain.frx":000C
   WhatsThisHelp   =   -1  'True
   WindowState     =   2  'Maximized
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   7680
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   1085
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Text            =   ""
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog dlgMain 
      Left            =   3000
      Top             =   7380
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer tmrMain 
      Interval        =   60000
      Left            =   2400
      Top             =   7440
   End
   Begin SysInfoLib.SysInfo sysScreen 
      Left            =   1680
      Top             =   7320
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin ComctlLib.ImageList ImageList2 
      Left            =   960
      Top             =   7260
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   12632256
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   15
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":5F090
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":5F1A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":5F2B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":5F3C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":5F4D8
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":5F5EA
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":5F6F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":5F806
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":5F918
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":5FA2A
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":5FB34
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":5FC7A
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":7AD30
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":7B04A
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":7B364
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&Cliente"
      WindowList      =   -1  'True
      Begin VB.Menu mnCadCliente 
         Caption         =   "&Cadastro"
      End
   End
   Begin VB.Menu mnuProduto 
      Caption         =   "&Produto"
      Begin VB.Menu mnProdCadastro 
         Caption         =   "&Cadastro"
      End
   End
   Begin VB.Menu mnuAtend 
      Caption         =   "P&edidos"
      Begin VB.Menu mnuPVenda 
         Caption         =   "&Pedidos de Venda"
      End
   End
   Begin VB.Menu mnuSair 
      Caption         =   "&Sair"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MDIForm_Load()
    Me.NewDimensions
    DoEvents
    Me.Show
    DoEvents
End Sub

Public Sub NewDimensions()
    If Me.WindowState = 0 Then
        Me.Top = sysScreen.WorkAreaTop
        Me.Left = sysScreen.WorkAreaLeft
        Me.Width = sysScreen.WorkAreaWidth
        Me.Height = sysScreen.WorkAreaHeight
    End If
End Sub

Private Sub mnCadCliente_Click()
    frmCliente.Show
End Sub

Private Sub mnProdCadastro_Click()
    frmProduto.Show
End Sub

Private Sub mnuPVenda_Click()
    frmPedidoVenda.Show
End Sub

Private Sub mnuSair_Click()
    dbPedidos.Close
    Unload Me
End Sub
