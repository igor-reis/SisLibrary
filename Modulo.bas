Attribute VB_Name = "Modulo"
Global cn As ADODB.Connection
Global OP As ADODB.Command
'Para saber qual formulario esta aberto
Global NomeFormulario As String
Public Rs As New ADODB.Recordset
Public SQL As String
Global confirmarcmd As Integer
Global usuario, senha, calendario As String


Public Function conexao(cn As ADODB.Connection)
'Função que conecta no banco de dados
On Error GoTo Error

Set cn = New ADODB.Connection
    With cn
        .ConnectionString = "Provider=ZStyle IBOLE Provider;Password=masterkey;User ID=SYSDBA;SQL Dialect=3; " & _
        "Logging Level=0;Silent mode=True;CharacterSet=WIN1252;Data Source = C:\Users\Igor\Desktop\SisLibrary\bancodedados\sislibrary.GDB"
        .Open
    End With

Error:
If Err Then
    MsgBox "Banco de dados não econtrado!", vbCritical, "Mensagem do Sistema"
    Err.Clear
End If
    
End Function



