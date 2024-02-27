VERSION 5.00
Object = "{CFAB6834-3B57-49FC-8770-CBA3667FE193}#1.0#0"; "ctextbox.ocx"
Object = "{61159A24-3E03-4E76-9CA9-2396C6822B8F}#1.0#0"; "chamaleonbtn.ocx"
Begin VB.Form FrmLogin 
   BackColor       =   &H000080FF&
   Caption         =   "Autenticação - SisLibrary"
   ClientHeight    =   5625
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7770
   Icon            =   "FrmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5625
   ScaleWidth      =   7770
   StartUpPosition =   2  'CenterScreen
   Begin cTextBox.cText txtSenha 
      Height          =   375
      Left            =   2880
      TabIndex        =   1
      Top             =   3360
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      BackColorGotFocus=   8454016
      PasswordChar    =   "*"
      Appearance      =   0
      Alignment       =   0
      FontBold        =   0   'False
      FontSize        =   9,75
      FontName        =   "Verdana"
      BackColor       =   8454016
      FontItalic      =   -1  'True
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
      Left            =   2880
      TabIndex        =   0
      Top             =   2640
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      BackColorGotFocus=   8454016
      Appearance      =   0
      Alignment       =   0
      FontBold        =   0   'False
      FontSize        =   9,75
      FontName        =   "Verdana"
      BackColor       =   8454016
      FontItalic      =   -1  'True
      DateFormat      =   "dd/MM/yy"
      FormatoExibData =   "__/__/____"
      tipoLetra       =   2
      Calendar_FormBackcolor=   16777215
      Calendar_BackColor=   14671839
      Calendar_ColorWeekDay=   8421376
      Calendar_Selected=   12640511
   End
   Begin VB.PictureBox logo 
      BorderStyle     =   0  'None
      Height          =   2295
      Left            =   2760
      Picture         =   "FrmLogin.frx":324A
      ScaleHeight     =   2295
      ScaleWidth      =   2295
      TabIndex        =   6
      Top             =   120
      Width           =   2295
   End
   Begin ChamaleonBtn.chameleonButton cmdLogar 
      DragIcon        =   "FrmLogin.frx":5749
      Height          =   1095
      Left            =   2160
      TabIndex        =   2
      Top             =   4080
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   1931
      BTYPE           =   3
      TX              =   "&Acessar"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Hobo Std"
         Size            =   9.75
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmLogin.frx":7413
      PICN            =   "FrmLogin.frx":742F
      UMCOL           =   -1  'True
      SOFT            =   -1  'True
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   2
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonBtn.chameleonButton cmdSair 
      DragIcon        =   "FrmLogin.frx":9109
      Height          =   1095
      Left            =   4080
      TabIndex        =   3
      Top             =   4080
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   1931
      BTYPE           =   3
      TX              =   "&Sair"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Hobo Std"
         Size            =   9.75
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmLogin.frx":9DD3
      PICN            =   "FrmLogin.frx":9DEF
      UMCOL           =   -1  'True
      SOFT            =   -1  'True
      PICPOS          =   1
      NGREY           =   0   'False
      FX              =   2
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label lblRestantes 
      AutoSize        =   -1  'True
      BackColor       =   &H000080FF&
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Hobo Std"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   2640
      TabIndex        =   8
      Top             =   5280
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Label lblTentativas 
      AutoSize        =   -1  'True
      BackColor       =   &H000080FF&
      Caption         =   "Tentativas restantes:"
      BeginProperty Font 
         Name            =   "Hobo Std"
         Size            =   12
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   0
      TabIndex        =   7
      Top             =   5280
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Label lblSenha 
      AutoSize        =   -1  'True
      BackColor       =   &H000080FF&
      Caption         =   " Senha:"
      BeginProperty Font 
         Name            =   "Hobo Std"
         Size            =   12
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1875
      TabIndex        =   5
      Top             =   3360
      Width           =   810
   End
   Begin VB.Label lblUsuario 
      AutoSize        =   -1  'True
      BackColor       =   &H000080FF&
      Caption         =   "Usuário:"
      BeginProperty Font 
         Name            =   "Hobo Std"
         Size            =   12
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1800
      TabIndex        =   4
      Top             =   2640
      Width           =   915
   End
End
Attribute VB_Name = "FrmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tentativas As Integer

Private Sub cmdLogar_Click()
'verificar campos vazios

If txtUsuario.Text = "" Or txtSenha.Text = "" Then
    MsgBox "Por favor, complete todos os campos e tente novamente!", vbInformation, "Mensagem do Sistema"
    txtSenha.SetFocus
    
Else
    'verificar o usuário e senha
    
    Call conexao(cn)

    Rs.Open "SELECT * FROM AUTENTICACAO WHERE USUARIO= '" & txtUsuario.Text & "' AND SENHA='" & txtSenha.Text & "'", cn
    
        
    If Rs.RecordCount > 0 Then
                
        Rs.Close
        Set Rs = Nothing
        cn.Close
        Set cn = Nothing
        
        usuario = txtUsuario.Text
        
        Unload Me
        
        FrmCarregamento.Show 1
        
        
    Else
        
        MsgBox "Usuário ou senha inválidos!", vbInformation, "Mensagem do Sistema"
        txtUsuario.Text = ""
        txtSenha.Text = ""
        txtUsuario.SetFocus
        
        Rs.Close
        Set Rs = Nothing
        cn.Close
        Set cn = Nothing
        
        tentativas = tentativas + 1
        
        lblRestantes.Caption = lblRestantes.Caption - 1
        
        If tentativas >= 1 Then
        
            lblTentativas.Visible = True
            lblRestantes.Visible = True
            
        End If
        
          If tentativas = 4 Then
        
              'Exibe uma mensagem informando que o numero de tentativas foi ultrapassado
              MsgBox "Você ultrapassou o número de tentativas de acesso, o sistema será fechado!", vbCritical, "Mensagem do Sistema"
          End
          End If
    End If
End If
End Sub

Private Sub cmdSair_Click()
If MsgBox("Você tem certeza que deseja sair do sistema?", vbYesNo, "Mensagem do Sistema") = vbYes Then
     'exclui
     End
Else
     'cancela
End If
End Sub


Private Sub txtSenha_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdLogar_Click
End Sub

