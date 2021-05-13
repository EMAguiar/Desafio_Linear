VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmPassword 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   8070
   ClientLeft      =   4920
   ClientTop       =   3930
   ClientWidth     =   11445
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8070
   ScaleWidth      =   11445
   ShowInTaskbar   =   0   'False
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.TextBox txtPassword 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      IMEMode         =   3  'DISABLE
      Left            =   5200
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   4200
      Width           =   2000
   End
   Begin VB.TextBox txtUser 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   5200
      TabIndex        =   0
      Top             =   3000
      Width           =   2000
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   450
      Left            =   0
      TabIndex        =   4
      Top             =   7620
      Width           =   11445
      _ExtentX        =   20188
      _ExtentY        =   794
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdEsc 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   6500
      TabIndex        =   3
      Top             =   5000
      WhatsThisHelpID =   17
      Width           =   1215
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Logar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   5000
      TabIndex        =   2
      Top             =   5000
      WhatsThisHelpID =   16
      Width           =   1170
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Senha:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   5200
      TabIndex        =   7
      Top             =   3800
      Width           =   870
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Usuário:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   5200
      TabIndex        =   6
      Top             =   2520
      Width           =   1005
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Login Sistema"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   5200
      TabIndex        =   5
      Top             =   1680
      Width           =   1725
   End
End
Attribute VB_Name = "frmPassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

On Error GoTo Err_Load

    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2
    
Ok_Load:
Exit Sub

Err_Load:
MsgBox Err.Description

End Sub

Private Sub cmdEsc_Click()
    Unload Me
    End
End Sub

Private Sub cmdOk_Click()
Dim dsLogin As New ADODB.Recordset

If Len(txtUser) > 0 And Len(txtPassword) > 0 Then
   dsLogin.Open "Select * from tbUsuarios where usuario = '" & UCase(txtUser) & "' and Senha='" & UCase(txtPassword) & "'", dbPedidos
    
   If Not dsLogin.RecordCount > 0 Then
      MsgBox "Usuário o senha inválido."
   
   Else
      Unload frmPassword
      frmMain.Show
   End If
End If
   
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       KeyAscii = 0
       cmdOk_Click
    End If
End Sub
