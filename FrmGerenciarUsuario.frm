VERSION 5.00
Object = "{CFAB6834-3B57-49FC-8770-CBA3667FE193}#1.0#0"; "ctextbox.ocx"
Object = "{61159A24-3E03-4E76-9CA9-2396C6822B8F}#1.0#0"; "chamaleonbtn.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form FrmGerenciarUsuario 
   BackColor       =   &H000080FF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Gerenciar Usuários"
   ClientHeight    =   4875
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7200
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4875
   ScaleWidth      =   7200
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin cTextBox.cText txtIdUsuario 
      Height          =   255
      Left            =   3600
      TabIndex        =   12
      Top             =   2760
      Visible         =   0   'False
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   450
      BackColor_MouseMove=   16709609
      Enabled         =   0   'False
      Appearance      =   0
      Alignment       =   0
      FontBold        =   0   'False
      FontSize        =   8,25
      FontName        =   "MS Sans Serif"
      BackColor       =   16777215
      AutoSelect      =   -1  'True
      DateFormat      =   "dd/mm/yyyy"
      FormatoExibData =   "__/__/____"
      Calendar_FormBackcolor=   16777215
      Calendar_BackColor=   14671839
      Calendar_ColorWeekDay=   8421376
      Calendar_Selected=   12640511
   End
   Begin ChamaleonBtn.chameleonButton cmdVoltar 
      Height          =   495
      Left            =   5880
      TabIndex        =   10
      Top             =   4200
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "Voltar"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Hobo Std"
         Size            =   8.25
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmGerenciarUsuario.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonBtn.chameleonButton cmdEditar 
      Height          =   495
      Left            =   4800
      TabIndex        =   9
      Top             =   4200
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "Editar"
      ENAB            =   0   'False
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Hobo Std"
         Size            =   8.25
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmGerenciarUsuario.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonBtn.chameleonButton cmdExcluir 
      Height          =   495
      Left            =   3720
      TabIndex        =   8
      Top             =   4200
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "Excluir"
      ENAB            =   0   'False
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Hobo Std"
         Size            =   8.25
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmGerenciarUsuario.frx":0038
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonBtn.chameleonButton cmdCancelar 
      Height          =   495
      Left            =   2640
      TabIndex        =   7
      Top             =   4200
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "Cancelar"
      ENAB            =   0   'False
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Hobo Std"
         Size            =   8.25
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmGerenciarUsuario.frx":0054
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonBtn.chameleonButton cmdGravar 
      Height          =   495
      Left            =   1560
      TabIndex        =   6
      Top             =   4200
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "Gravar"
      ENAB            =   0   'False
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Hobo Std"
         Size            =   8.25
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmGerenciarUsuario.frx":0070
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonBtn.chameleonButton cmdNovo 
      Height          =   495
      Left            =   480
      TabIndex        =   5
      Top             =   4200
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "Novo"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Hobo Std"
         Size            =   8.25
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmGerenciarUsuario.frx":008C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin cTextBox.cText txtSenha 
      Height          =   375
      Left            =   3120
      TabIndex        =   1
      Top             =   3600
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      BackColor_MouseMove=   16709609
      PasswordChar    =   "*"
      Appearance      =   0
      Alignment       =   0
      FontBold        =   0   'False
      FontSize        =   8,25
      FontName        =   "MS Sans Serif"
      BackColor       =   16777215
      AutoSelect      =   -1  'True
      DateFormat      =   "dd/MM/yy"
      FormatoExibData =   "__/__/____"
      tipoLetra       =   2
      Calendar_FormBackcolor=   16777215
      Calendar_BackColor=   14671839
      Calendar_ColorWeekDay=   8421376
      Calendar_Selected=   12640511
   End
   Begin cTextBox.cText txtUsuario 
      Height          =   375
      Left            =   3120
      TabIndex        =   0
      Top             =   3120
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      BackColor_MouseMove=   16709609
      Appearance      =   0
      Alignment       =   0
      FontBold        =   0   'False
      FontSize        =   8,25
      FontName        =   "MS Sans Serif"
      BackColor       =   16777215
      AutoSelect      =   -1  'True
      DateFormat      =   "dd/MM/yy"
      FormatoExibData =   "__/__/____"
      tipoLetra       =   2
      Calendar_FormBackcolor=   16777215
      Calendar_BackColor=   14671839
      Calendar_ColorWeekDay=   8421376
      Calendar_Selected=   12640511
   End
   Begin MSFlexGridLib.MSFlexGrid grid 
      Height          =   1515
      Left            =   2400
      TabIndex        =   13
      Top             =   480
      Width           =   2715
      _ExtentX        =   4789
      _ExtentY        =   2672
      _Version        =   393216
      Rows            =   6
      Cols            =   3
      FixedCols       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Hobo Std"
         Size            =   8.25
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H000080FF&
      BeginProperty Font 
         Name            =   "Hobo Std"
         Size            =   9.75
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   270
      Left            =   120
      TabIndex        =   4
      Top             =   2520
      Width           =   60
   End
   Begin VB.Label lblIdUsuario 
      AutoSize        =   -1  'True
      BackColor       =   &H000080FF&
      Caption         =   "ID Usuário:"
      BeginProperty Font 
         Name            =   "Hobo Std"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   2400
      TabIndex        =   11
      Top             =   2760
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.Label lblUsuario 
      AutoSize        =   -1  'True
      BackColor       =   &H000080FF&
      Caption         =   "Usuário:"
      BeginProperty Font 
         Name            =   "Hobo Std"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2040
      TabIndex        =   3
      Top             =   3120
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Label lblSenha 
      AutoSize        =   -1  'True
      BackColor       =   &H000080FF&
      Caption         =   "Senha:"
      BeginProperty Font 
         Name            =   "Hobo Std"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2160
      TabIndex        =   2
      Top             =   3600
      Visible         =   0   'False
      Width           =   780
   End
End
Attribute VB_Name = "FrmGerenciarUsuario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim limpaF As Control
Dim i As Integer

Private Sub cmdVoltar_Click()
Unload Me
FrmPrincipal.Show 1
End Sub

Private Sub cmdCancelar_Click()

Call LimpaControles

 Unload Me

 FrmGerenciarUsuario.Show
    
End Sub


Private Sub cmdEditar_Click()
FrmConfirmarSenhaAtual.Show 1

If confirmarcmd = 1 Then

    Rs.Close

    Set OP = New ADODB.Command
    With OP
      .ActiveConnection = cn
      .CommandText = "UPDATE AUTENTICACAO SET USUARIO = '" & txtUsuario.Text & "',SENHA = '" & senha & "' WHERE ID_USUARIO = " & CInt(txtIdUsuario.Text)
      .Execute
    End With
     MsgBox ("Registro alterado com sucesso!"), vbInformation, "Mensagem do Sistema"
     
     
     Call LimpaControles
    
     Unload Me
    
     FrmGerenciarUsuario.Show
     
 End If
 
End Sub

Private Sub cmdExcluir_Click()
FrmConfirmarSenha.Show 1

If confirmarcmd = 1 Then
    
    Rs.Close
 
    Rs.Open "DELETE FROM AUTENTICACAO WHERE ID_USUARIO = " & CInt(txtIdUsuario.Text)
    MsgBox "Registro excluido com sucesso!", vbExclamation, "Mensagem do Sistema"
    
    Call LimpaControles
     
    Unload Me
    
    FrmGerenciarUsuario.Show
     
End If
 
End Sub

Private Sub cmdGravar_Click()
    Set OP = New ADODB.Command
    With OP
            .ActiveConnection = cn
            .CommandText = "insert into AUTENTICACAO(ID_USUARIO,USUARIO,SENHA) values (" & CInt(txtIdUsuario.Text) & ",'" & txtUsuario.Text & "','" & txtSenha.Text & "');"
            .Execute
    End With
    MsgBox "Registro salvo com sucesso!", vbInformation, "Mensagem do Sistema"
    
    Call LimpaControles
       
    Unload Me
    
    FrmGerenciarUsuario.Show
End Sub

Public Function GeraID(ByVal GEN_ID_USUARIO As String) As Long
Call conexao(cn)
    Set Rs = New ADODB.Recordset
        'Use a tabela RDB$DATABASE, pois ela sempre retorna um único registro
        Rs.Open "Select GEN_ID(" & GEN_ID_USUARIO & ", 1) From RDB$DATABASE", cn
        GeraID = Rs(0)
        Rs.Close
    Set Rs = Nothing
End Function

Private Sub cmdNovo_Click()
    lblIdUsuario.Visible = True
    txtIdUsuario.Visible = True
    lblUsuario.Visible = True
    txtUsuario.Visible = True
    lblSenha.Visible = True
    txtSenha.Visible = True
    txtUsuario.SetFocus
    
    cmdGravar.Enabled = True
    cmdCancelar.Enabled = True
    
    txtIdUsuario.Text = GeraID("GEN_ID_USUARIO")
        
End Sub


 Public Function LimpaControles()
 'Limpa todos campos dos Ctext
 For Each limpaF In Controls
    If TypeOf limpaF Is cText Then
    limpaF.Text = ""
    End If
Next limpaF

End Function



Function CarregaGrid()

Call conexao(cn)

Rs.Open SQL, cn

If Rs.RecordCount > "0" Then
    
    grid.TextMatrix(0, 0) = "ID Usuário"
    grid.TextMatrix(0, 1) = "Usuário"
    grid.TextMatrix(0, 2) = "Senha"

    grid.Rows = Rs.RecordCount + 1
    Rs.MoveFirst

    For i = 1 To Rs.RecordCount
        grid.TextMatrix(i, 0) = Rs!ID_USUARIO
        grid.TextMatrix(i, 1) = Rs!usuario
        grid.TextMatrix(i, 2) = "*******"
        Rs.MoveNext
    Next i
    
    Call AjustaGrid
    
End If
    
Rs.Close
    
'Ordenar em ordem crescente
grid.Col = 0
grid.Sort = flexSortGenericAscending
End Function

Private Sub Form_Load()
    Label1.Caption = "Conectado como: " & usuario & ""
    
    
    SQL = "SELECT * FROM AUTENTICACAO"
    
    Call CarregaGrid
    Call AjustaGrid
End Sub

Public Function AjustaGrid()
    'Função Ajusta as colunas do grid para o tamanho do texto contido nas células
    Dim Max_Wid As Single
    Dim Wid As Single
    Dim Max_Row As Integer
    Dim R As Integer
    Dim c As Integer
    Screen.MousePointer = vbHourglass
    'Ajusta as colunas do grid para o tamanho do texto contido nas células
    Max_Row = grid.Rows - 1
    For c = 0 To grid.Cols - 1
      Max_Wid = 0
      For R = 0 To Max_Row
        Wid = TextWidth(grid.TextMatrix(R, c))
        If Max_Wid < Wid Then Max_Wid = Wid
      Next R
       grid.ColWidth(c) = Max_Wid + 240
    Next c
    Screen.MousePointer = vbDefault
End Function

Private Sub grid_DblClick()
    lblIdUsuario.Visible = True
    txtIdUsuario.Visible = True
    lblUsuario.Visible = True
    txtUsuario.Visible = True
    lblSenha.Visible = True
    txtSenha.Visible = True

Call conexao(cn)
     
     grid.Col = 0
     Rs.Open "SELECT * FROM AUTENTICACAO WHERE ID_USUARIO = " & CInt(grid.Text), cn
     
     If IsNull(Rs!ID_USUARIO) = False Then txtIdUsuario.Text = Rs!ID_USUARIO
     If IsNull(Rs!usuario) = False Then txtUsuario.Text = Rs!usuario
     If IsNull(Rs!senha) = False Then txtSenha.Text = Rs!senha
          
     Rs.Close
        
     cmdCancelar.Enabled = True
     cmdExcluir.Enabled = True
     cmdEditar.Enabled = True
     cmdNovo.Enabled = False
     txtSenha.Enabled = False
     
     End Sub


