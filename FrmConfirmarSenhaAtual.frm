VERSION 5.00
Object = "{CFAB6834-3B57-49FC-8770-CBA3667FE193}#1.0#0"; "ctextbox.ocx"
Object = "{61159A24-3E03-4E76-9CA9-2396C6822B8F}#1.0#0"; "chamaleonbtn.ocx"
Begin VB.Form FrmConfirmarSenhaAtual 
   BorderStyle     =   0  'None
   ClientHeight    =   1725
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3015
   LinkTopic       =   "Form1"
   ScaleHeight     =   1725
   ScaleWidth      =   3015
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin cTextBox.cText txtSenhaAtual 
      Height          =   255
      Left            =   1200
      TabIndex        =   0
      Top             =   360
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   450
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
   Begin cTextBox.cText txtNovaSenha 
      Height          =   255
      Left            =   1200
      TabIndex        =   1
      Top             =   720
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   450
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
   Begin ChamaleonBtn.chameleonButton cmdConfirmar 
      Height          =   375
      Left            =   600
      TabIndex        =   2
      Top             =   1200
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Confirmar"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
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
      MICON           =   "FrmConfirmarSenhaAtual.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonBtn.chameleonButton cmdVoltar 
      Height          =   375
      Left            =   1560
      TabIndex        =   3
      Top             =   1200
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Voltar"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
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
      MICON           =   "FrmConfirmarSenhaAtual.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Nova senha:"
      Height          =   195
      Index           =   1
      Left            =   240
      TabIndex        =   5
      Top             =   720
      Width           =   915
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Senha atual:"
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   4
      Top             =   360
      Width           =   900
   End
End
Attribute VB_Name = "FrmConfirmarSenhaAtual"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdConfirmar_Click()
If txtNovaSenha.Text = "" Or txtSenhaAtual.Text = "" Then
    MsgBox "Por favor, complete todos os campos e tente novamente!", vbInformation, "Mensagem do Sistema"
        
Else

    Call conexao(cn)
    
        Rs.Open "SELECT * FROM AUTENTICACAO WHERE SENHA = '" & txtSenhaAtual.Text & "' and ID_USUARIO = '" & FrmGerenciarUsuario.txtIdUsuario.Text & "'", cn
        
            
        If Rs.RecordCount > 0 Then
        
            If txtNovaSenha = "" Then
            
                senha = txtSenhaAtual.Text
                
            Else
             
                senha = txtNovaSenha.Text
                
            End If
                                  
        Unload Me
        confirmarcmd = 1
        
    Else
    
        MsgBox "Dados não conferem, tente novamente!", vbExclamation, "Mensagem do Sistema"
        txtSenhaAtual = ""
        txtNovaSenha = ""
        txtSenhaAtual.SetFocus
        
        Rs.Close
        
    End If
End If
End Sub

Private Sub cmdVoltar_Click()
Unload Me

confirmarcmd = 0
End Sub

Private Sub txtNovaSenha_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdConfirmar_Click
End Sub
