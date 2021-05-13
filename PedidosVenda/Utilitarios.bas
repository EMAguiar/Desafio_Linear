Attribute VB_Name = "Utilitarios"
Public dbPedidos     As New ADODB.Connection    'Banco de dados de Pedidos de Clientes

Sub Main()
On Error GoTo Err_Main

If App.PrevInstance = 0 Then

   dbPedidos.ConnectionString = "Driver={MySQL ODBC 8.0 ANSI Driver};Server=localhost;Database=dbPedidos;uid=root;pwd=123456;"
   dbPedidos.CursorLocation = adUseClient
   dbPedidos.Open
   
   frmPassword.Show
   Exit Sub
   
End If

Ok_Main:
    Exit Sub

Err_Main:

Select Case Err.Number
       Case 3024
       MsgBox "Não foi possível iniciar o sistema devido a falha de conexão"
       Case Else
       MsgBox Err.Description

End Select

End Sub
