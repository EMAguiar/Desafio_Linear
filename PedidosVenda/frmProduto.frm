VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmProduto 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Produtos"
   ClientHeight    =   6030
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7515
   ControlBox      =   0   'False
   FillColor       =   &H00404040&
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6030
   ScaleWidth      =   7515
   Begin MSFlexGridLib.MSFlexGrid grdProdutos 
      Height          =   2655
      Left            =   120
      TabIndex        =   10
      Top             =   2400
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   4683
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   120
      TabIndex        =   9
      Top             =   5160
      Width           =   7335
      Begin VB.CommandButton cmdGravar 
         Caption         =   "&Gravar"
         Height          =   420
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   1000
      End
      Begin VB.CommandButton cmdExclusao 
         Caption         =   "&Excluir"
         Height          =   420
         Left            =   1200
         TabIndex        =   4
         Top             =   240
         Width           =   1000
      End
      Begin VB.CommandButton cmdSaida 
         Caption         =   "&Sair"
         Height          =   420
         Left            =   6240
         TabIndex        =   5
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
   Begin MSMask.MaskEdBox txtPreco 
      Height          =   330
      Left            =   135
      TabIndex        =   1
      Top             =   1020
      Width           =   7200
      _ExtentX        =   12700
      _ExtentY        =   582
      _Version        =   393216
      PromptInclude   =   0   'False
      Format          =   "#,##0.00"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox txtCodProduto 
      Height          =   330
      Left            =   135
      TabIndex        =   2
      Top             =   1860
      Width           =   7200
      _ExtentX        =   12700
      _ExtentY        =   582
      _Version        =   393216
      PromptInclude   =   0   'False
      PromptChar      =   " "
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Nome"
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
      Top             =   75
      Width           =   495
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Código"
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
      Top             =   1575
      Width           =   585
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Preço"
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
      TabIndex        =   6
      Top             =   735
      Width           =   495
   End
End
Attribute VB_Name = "frmProduto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tblPedidos   As New ADODB.Recordset 'Tabela para pesquisa de pedidos do cliente
Dim tblProdutos  As New ADODB.Recordset 'Tabela de Produtos
Dim CodProduto   As String  'Código do Produto

Private Sub Form_Load()
    Call Montar_grdProdutos
    Call Carregar_grdProdutos
    CodProduto = ""
End Sub

Private Sub cmdExclusao_Click()

On Error GoTo ErrExcluirProduto

   If Len(CodProduto) = 0 Then
      MsgBox "Não foi selecionado nenhum produto para exclusão", vbInformation
      Exit Sub
   End If
   
   If MsgBox("Deseja realmente excluir este produto ?", vbExclamation + vbYesNo, "Exclusão") = vbYes Then
      Set tblPedidos = New ADODB.Recordset
      tblPedidos.Open "Select * from tbpedidos where idProduto = " & CodProduto, dbPedidos, adOpenStatic
      If Not tblPedidos.EOF Then
         MsgBox "Não é possível excluir o produto, pois ele já foi utilizado em algum pedido no sistema.", vbInformation
         Exit Sub
      End If
      tblPedidos.Close
      
      Set tblProdutos = New ADODB.Recordset
      tblProdutos.Open "select * from tbprodutos where idproduto = " & CodProduto, dbPedidos, adOpenKeyset, adLockPessimistic
      tblProdutos.Delete
      Call Carregar_grdProdutos
      CodProduto = 0
      MsgBox "Produto excluído com sucesso.", vbInformation
   End If
   
   Exit Sub ' Evita que o bloco para tratamento de erros seja executado quando _
            ' a exclusão foi bem sucedida
      
ErrExcluirProduto:
   MsgBox MensErro, 16, "Falha na exclusão do Produto.", Titulo
   Exit Sub
End Sub

Private Sub cmdGravar_Click()
    GravarProduto
End Sub

Private Sub cmdSaida_Click()
   Unload Me
End Sub

Private Sub LimparCampos()
    txtNome = ""
    txtPreco = ""
    txtCodProduto = ""
End Sub

Private Sub Montar_grdProdutos()
    With grdProdutos
         .Row = 0
         .Cols = 4
         .ColWidth(0) = 1 'IdProduto
         .ColWidth(1) = 3500: .TextMatrix(0, 1) = "Nome"
         .ColWidth(2) = 1700: .TextMatrix(0, 2) = "Preço"
         .ColWidth(3) = 1500: .TextMatrix(0, 3) = "Código"
         .Rows = 1
    End With
End Sub

Private Sub Carregar_grdProdutos()
    Set tblProdutos = New ADODB.Recordset
    tblProdutos.Open "Select * from tbprodutos order by nome", dbPedidos, adOpenStatic, adLockPessimistic
        
    grdProdutos.Rows = 1
    
    If tblProdutos.RecordCount > 0 Then
        grdProdutos.Redraw = False
        Do While Not tblProdutos.EOF
           grdProdutos.AddItem tblProdutos!idProduto & vbTab & tblProdutos!nome & vbTab & Format(tblProdutos!Preco, "#,###,##0.00") & vbTab & tblProdutos!idProduto
           tblProdutos.MoveNext
        Loop
        grdProdutos.Redraw = True
    End If
    tblProdutos.Close
End Sub

Private Sub grdProdutos_DblClick()
    If grdProdutos.Rows - 1 > 0 And Len(CodProduto) = 0 Then
       grdProdutos.Row = grdProdutos.RowSel
       grdProdutos.Col = 1
       grdProdutos.ColSel = grdProdutos.Cols - 1
       grdProdutos.CellBackColor = vbYellow
       CodProduto = grdProdutos.TextMatrix(grdProdutos.RowSel, 0)
    Else
       CodProduto = ""
       grdProdutos.Row = grdProdutos.RowSel
       grdProdutos.Col = 1
       grdProdutos.ColSel = grdProdutos.Cols - 1
       grdProdutos.CellBackColor = vbWhite
    End If
End Sub

Function GravarProduto()
   On Error GoTo ErrGravaProduto
   
    If Len(txtNome) <= 0 Then
       MsgBox "Não é possível gravar os dados sem o nome do produto.", vbInformation
       txtNome.SetFocus
       Exit Function
    End If
    
    If Len(txtPreco) = 0 Then
       MsgBox "Não é possível gravar os dados sem o preço do produto.", vbInformation
       txtPreco.SetFocus
       Exit Function
    
    Else
       If CDbl(txtPreco) = 0 Then
          MsgBox "Não é possível gravar os dados com o preço do produto zerado.", vbInformation
          txtPreco.SetFocus
          Exit Function
       End If
    End If
    
    If Len(txtCodProduto) = 0 Then
       MsgBox "Não é possível grava os dados sem o código do produto.", vbInformation
       txtCodProduto.SetFocus
       Exit Function
       
    Else
        If CDbl(txtCodProduto) = 0 Then
           MsgBox "Não é possível gravar os dados com o código do produto zerado.", vbInformation
           txtCodProduto.SetFocus
           Exit Function
        End If
    End If
    
    Set tblProdutos = New ADODB.Recordset
    tblProdutos.Open "tbprodutos", dbPedidos, adOpenKeyset, adLockPessimistic

    tblProdutos.AddNew
    tblProdutos("Nome") = txtNome
    tblProdutos("Preco") = CDec(txtPreco)
    tblProdutos("idProduto") = txtCodProduto
    tblProdutos.Update
    
    LimparCampos
    Carregar_grdProdutos
    MsgBox "Dados gravados com sucesso.", vbInformation
    txtNome.SetFocus
    Exit Function
    
ErrGravaProduto:
    MsgBox "Não foi possível gravar os dados. " & Err.Description, vbCritical
    
End Function
