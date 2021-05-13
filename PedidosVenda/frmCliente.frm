VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmCliente 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Clientes"
   ClientHeight    =   6030
   ClientLeft      =   45
   ClientTop       =   2250
   ClientWidth     =   7515
   ControlBox      =   0   'False
   FillColor       =   &H00404040&
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6030
   ScaleWidth      =   7515
   Begin MSFlexGridLib.MSFlexGrid grdClientes 
      Height          =   2655
      Left            =   120
      TabIndex        =   11
      Top             =   2400
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   4683
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   120
      TabIndex        =   10
      Top             =   5160
      Width           =   7335
      Begin VB.CommandButton cmdGravar 
         Caption         =   "&Gravar"
         Height          =   420
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   1000
      End
      Begin VB.CommandButton cmdExclusao 
         Caption         =   "&Excluir"
         Height          =   420
         Left            =   1200
         TabIndex        =   5
         Top             =   240
         Width           =   1000
      End
      Begin VB.CommandButton cmdSaida 
         Caption         =   "&Sair"
         Height          =   420
         Left            =   6240
         TabIndex        =   6
         Top             =   240
         Width           =   1000
      End
   End
   Begin VB.TextBox txtNome 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   135
      TabIndex        =   0
      Top             =   280
      Width           =   7200
   End
   Begin MSMask.MaskEdBox txtTelefone 
      Height          =   330
      Left            =   135
      TabIndex        =   1
      Top             =   1020
      Width           =   7200
      _ExtentX        =   12700
      _ExtentY        =   582
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   16
      Format          =   "(##) # ####-####"
      Mask            =   "(##) # ####-####"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox txtLimiteCredito 
      Height          =   330
      Left            =   135
      TabIndex        =   2
      Top             =   1860
      Width           =   3600
      _ExtentX        =   6350
      _ExtentY        =   582
      _Version        =   393216
      PromptInclude   =   0   'False
      Format          =   "###,##0.00"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox txtCreditoDisponivel 
      Height          =   330
      Left            =   3840
      TabIndex        =   3
      Top             =   1845
      Width           =   3480
      _ExtentX        =   6138
      _ExtentY        =   582
      _Version        =   393216
      PromptInclude   =   0   'False
      Enabled         =   0   'False
      Format          =   "###,##0.00"
      PromptChar      =   " "
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Crédito Disponível"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3840
      TabIndex        =   12
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Nome:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   135
      TabIndex        =   9
      Top             =   75
      Width           =   555
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Limite Crédito"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   135
      TabIndex        =   8
      Top             =   1575
      Width           =   1155
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Telefone"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   135
      TabIndex        =   7
      Top             =   730
      Width           =   765
   End
End
Attribute VB_Name = "frmCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tblPedidos   As New ADODB.Recordset 'Tabela para pesquisa de pedidos do cliente
Dim tblClientes  As New ADODB.Recordset 'Tabela de Clientes
Dim CodCliente   As Integer 'Código do Cliente

Private Sub Form_Load()
    Call Montar_grdClientes
    Call Carregar_grdClientes
    CodCliente = 0
End Sub

Private Sub cmdNovo_Click()
    Call LimparCampos
    txtNome.SetFocus
End Sub

Private Sub cmdExclusao_Click()

On Error GoTo ErrExcluiCliente

   If CodCliente = 0 Then
      MsgBox "Não foi selecionado nenhum cliente para exclusão", vbInformation
      Exit Sub
   End If
   
   If MsgBox("Deseja realmente excluir este cliente ?", vbExclamation + vbYesNo, "Exclusão") = vbYes Then
      Set tblPedidos = New ADODB.Recordset
      tblPedidos.Open "Select * from tbPedidos where idCliente = " & CodCliente, dbPedidos, adOpenStatic
      If Not tblPedidos.EOF Then
         MsgBox "Não é possível excluir o cliente, pois ele possui pedidos no sistema.", vbInformation
         Exit Sub
      End If
      tblPedidos.Close
      
      Set tblClientes = New ADODB.Recordset
      tblClientes.Open "select * from tbclientes where idCliente = " & CodCliente, dbPedidos, adOpenKeyset, adLockPessimistic
      tblClientes.Delete
      Call Carregar_grdClientes
      CodCliene = 0
      MsgBox "Cliente excluído com sucesso.", vbInformation
   End If
   
   Exit Sub ' Evita que o bloco para tratamento de erros seja executado quando _
            ' a exclusão foi bem sucedida
      
ErrExcluiCliente:
   MsgBox MensErro, 16, "Falha na exclusão do Cliente", Titulo
   Exit Sub
End Sub

Private Sub cmdGravar_Click()
    GravarCliente
End Sub

Private Sub cmdSaida_Click()
   Unload Me
End Sub

Private Sub LimparCampos()
    txtNome = ""
    txtTelefone = ""
    txtLimiteCredito = ""
End Sub

Private Sub Montar_grdClientes()
    With grdClientes
         .Row = 0
         .Cols = 5
         .ColWidth(0) = 600:  .TextMatrix(0, 0) = "Cód" 'IdCliente
         .ColWidth(1) = 2500: .TextMatrix(0, 1) = "Nome"
         .ColWidth(2) = 1000: .TextMatrix(0, 2) = "Telefone"
         .ColWidth(3) = 1200: .TextMatrix(0, 3) = "Limite Crédito"
         .ColWidth(4) = 1500: .TextMatrix(0, 4) = "Crédito Disponivel"
         .Rows = 1
    End With
End Sub

Private Sub Carregar_grdClientes()
    Set tblClientes = New ADODB.Recordset
    tblClientes.Open "Select * from tbClientes order by nome", dbPedidos, adOpenStatic, adLockPessimistic
        
    grdClientes.Rows = 1
    
    If tblClientes.RecordCount > 0 Then
        grdClientes.Redraw = False
        Do While Not tblClientes.EOF
           grdClientes.AddItem tblClientes!idcliente & vbTab & tblClientes!nome & vbTab & Format(tblClientes!telefone, "(##) # ####-####") & vbTab & Format(tblClientes!LimiteCredito, "#,###,##0.00") & vbTab & Format(tblClientes!CreditoDisponivel, "#,###,##0.00")
           tblClientes.MoveNext
        Loop
        grdClientes.Redraw = True
    End If
    tblClientes.Close
End Sub

Private Sub grdClientes_DblClick()
    If grdClientes.Rows - 1 > 0 And CodCliente = 0 Then
       grdClientes.Row = grdClientes.RowSel
       grdClientes.Col = 1
       grdClientes.ColSel = grdClientes.Cols - 1
       grdClientes.CellBackColor = vbYellow
       CodCliente = grdClientes.TextMatrix(grdClientes.RowSel, 0)
    Else
       CodCliente = 0
       grdClientes.Row = grdClientes.RowSel
       grdClientes.Col = 1
       grdClientes.ColSel = grdClientes.Cols - 1
       grdClientes.CellBackColor = vbWhite
    End If
End Sub

Function GravarCliente()
   On Error GoTo ErrGravaCliente
   
    If Len(txtNome) <= 0 Then
       MsgBox "Não é possível gravar os dados sem o nome do cliente.", vbInformation
       txtNome.SetFocus
       Exit Function
    End If
    
    If Len(txtTelefone) <= 0 Then
       MsgBox "Não é possível gravar os dados sem o telefone do cliente.", vbInformation
       txtTelefone.SetFocus
       Exit Function
    End If
    
    If CDbl(txtLimiteCredito) = 0 Then
       MsgBox "Informe um limite de crédito para o cliente.", vbInformation
       txtLimiteCredito.SetFocus
       Exit Function
    End If
    
    Set tblClientes = New ADODB.Recordset
    tblClientes.Open "tbClientes", dbPedidos, adOpenKeyset, adLockPessimistic

    tblClientes.AddNew
    tblClientes("Nome") = txtNome
    tblClientes("Telefone") = txtTelefone
    tblClientes("LimiteCredito") = CDbl(txtLimiteCredito)
    tblClientes("CreditoDisponivel") = CDbl(txtLimiteCredito)
    tblClientes.Update
    
    LimparCampos
    Carregar_grdClientes
    MsgBox "Dados gravados com sucesso.", vbInformation
    txtNome.SetFocus
    Exit Function
    
ErrGravaCliente:
    MsgBox "Não foi possível gravar os dados. " & Err.Description, vbCritical
    
End Function
