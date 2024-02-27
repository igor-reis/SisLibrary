VERSION 5.00
Object = "{61159A24-3E03-4E76-9CA9-2396C6822B8F}#1.0#0"; "chamaleonbtn.ocx"
Object = "{71D1ACD0-5AE6-4E74-A1A3-219049211E99}#1.0#0"; "alphaimagecontrol.ocx"
Begin VB.Form FrmPrincipal 
   Caption         =   "Bem vindo ao SisLibrary"
   ClientHeight    =   5040
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6405
   Icon            =   "FrmPrincipal.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5040
   ScaleWidth      =   6405
   StartUpPosition =   2  'CenterScreen
   Begin ChamaleonBtn.chameleonButton cmdGerenciar 
      Height          =   1215
      Left            =   120
      TabIndex        =   4
      Top             =   3600
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   2143
      BTYPE           =   3
      TX              =   "&Gerenciar Usuários"
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
      MICON           =   "FrmPrincipal.frx":324A
      PICN            =   "FrmPrincipal.frx":3266
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Timer TimerImagem 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   3480
      Top             =   3480
   End
   Begin VB.Timer TimerLabel 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   3960
      Top             =   3480
   End
   Begin ChamaleonBtn.chameleonButton cmdEmprestimos 
      Height          =   1335
      Left            =   2280
      TabIndex        =   0
      Top             =   240
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   2355
      BTYPE           =   3
      TX              =   "&Empréstimos"
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
      MICON           =   "FrmPrincipal.frx":603B
      PICN            =   "FrmPrincipal.frx":6057
      UMCOL           =   -1  'True
      SOFT            =   -1  'True
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   2
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonBtn.chameleonButton cmdAlunos 
      Height          =   1335
      Left            =   360
      TabIndex        =   1
      Top             =   240
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   2355
      BTYPE           =   3
      TX              =   "&Alunos"
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
      MICON           =   "FrmPrincipal.frx":680D
      PICN            =   "FrmPrincipal.frx":6829
      UMCOL           =   -1  'True
      SOFT            =   -1  'True
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   2
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonBtn.chameleonButton cmdLivros 
      Height          =   1335
      Left            =   4200
      TabIndex        =   2
      Top             =   240
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   2355
      BTYPE           =   3
      TX              =   "&Livros"
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
      MICON           =   "FrmPrincipal.frx":9C03
      PICN            =   "FrmPrincipal.frx":9C1F
      UMCOL           =   -1  'True
      SOFT            =   -1  'True
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   2
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonBtn.chameleonButton cmdSair 
      Height          =   855
      Left            =   1680
      TabIndex        =   3
      Top             =   3960
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   1508
      BTYPE           =   3
      TX              =   "&Sair"
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
      MICON           =   "FrmPrincipal.frx":102A7
      PICN            =   "FrmPrincipal.frx":102C3
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label lblAtencao 
      Caption         =   "       ATENÇÃO       Existem       empréstimos       atrasados!"
      BeginProperty Font 
         Name            =   "Hobo Std"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   4440
      TabIndex        =   6
      Top             =   4560
      Visible         =   0   'False
      Width           =   1680
   End
   Begin VB.Label Label1 
      Caption         =   "Olá, clique na opção desejada e aguarde!"
      BeginProperty Font 
         Name            =   "Hobo Std"
         Size            =   14.25
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   615
      Left            =   720
      TabIndex        =   5
      Top             =   2400
      Width           =   5295
   End
   Begin AlphaImageControl.aicAlphaImage imgAtencao 
      Height          =   2355
      Left            =   4440
      Top             =   2760
      Visible         =   0   'False
      Width           =   1755
      _ExtentX        =   3096
      _ExtentY        =   3096
      Image           =   "FrmPrincipal.frx":13A0A
      Props           =   5
   End
End
Attribute VB_Name = "FrmPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Atencao As Integer
Private Sub cmdAlunos_Click()
Unload Me
FrmAlunos.Show 1
End Sub

Private Sub cmdEmprestimos_Click()
Unload Me
FrmEmprestimos.Show 1
End Sub

Private Sub cmdGerenciar_Click()
Unload Me
FrmGerenciarUsuario.Show
tentativas = 0
End Sub

Private Sub cmdLivros_Click()
Unload Me
FrmLivros.Show 1
End Sub

Private Sub cmdSair_Click()
If MsgBox("Você tem certeza que deseja sair do sistema?", vbYesNo, "Mensagem do Sistema") = vbYes Then
     'exclui
     End
Else
     'cancela
End If
End Sub


Private Sub Form_Load()
Call conexao(cn)

    Rs.Open "SELECT * FROM EMPRESTIMOS WHERE STATUS = 'ATRASADO'", cn
    
        
    If Rs.RecordCount > 0 Then
                
        Rs.Close
        Set Rs = Nothing
        cn.Close
        Set cn = Nothing
        
        imgAtencao.Visible = True
        lblAtencao.Visible = True
        TimerImagem.Enabled = True
        TimerLabel.Enabled = True
        
    Else
             
        Rs.Close
        Set Rs = Nothing
        cn.Close
        Set cn = Nothing
        
        imgAtencao.Visible = False
        lblAtencao.Visible = False
        TimerImagem.Enabled = False
        TimerLabel.Enabled = False
    End If
    FrmPrincipal.Caption = "Bem vindo ao SisLibrary, " & usuario & ""
End Sub

Private Sub TimerImagem_Timer()
 Atencao = Atencao + 1
        
        If Atencao = 1 Then
            imgAtencao.Visible = False
        End If
    
        If Atencao = 2 Then
            imgAtencao.Visible = True
            Atencao = 0
        End If
End Sub

Private Sub TimerLabel_Timer()
lblAtencao.Caption = Right(lblAtencao.Caption, Len(lblAtencao.Caption) - 1) & Left(lblAtencao.Caption, 1)
End Sub
