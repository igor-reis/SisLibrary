VERSION 5.00
Object = "{CFAB6834-3B57-49FC-8770-CBA3667FE193}#1.0#0"; "ctextbox.ocx"
Object = "{61159A24-3E03-4E76-9CA9-2396C6822B8F}#1.0#0"; "chamaleonbtn.ocx"
Begin VB.Form FrmDescricaoRelAlunos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Descrição Relatório"
   ClientHeight    =   9360
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9870
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9360
   ScaleWidth      =   9870
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrmConsStatus 
      Caption         =   "Consulta por Status:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   240
      TabIndex        =   15
      Top             =   6120
      Width           =   9255
      Begin VB.CheckBox chkInativo 
         Caption         =   "Inativo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4680
         TabIndex        =   6
         Top             =   600
         Width           =   975
      End
      Begin VB.CheckBox chkAtivo 
         Caption         =   "Ativo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3600
         TabIndex        =   5
         Top             =   600
         Width           =   855
      End
   End
   Begin VB.Frame FrmConsLivro 
      Caption         =   "Consulta por Endereço:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   240
      TabIndex        =   13
      Top             =   3960
      Width           =   9255
      Begin VB.CheckBox chkEndereco 
         Caption         =   "Check1"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   255
      End
      Begin cTextBox.cText txtEndereco 
         Height          =   255
         Left            =   360
         TabIndex        =   4
         Top             =   1080
         Width           =   8655
         _ExtentX        =   15266
         _ExtentY        =   450
         BackColorGotFocus=   8454016
         BackColor_MouseMove=   16709609
         Appearance      =   0
         Alignment       =   0
         FontBold        =   0   'False
         FontSize        =   8,25
         FontName        =   "MS Sans Serif"
         BackColor       =   8454016
         AutoSelect      =   -1  'True
         DateFormat      =   "dd/MM/yy"
         FormatoExibData =   "__/__/____"
         tipoLetra       =   1
         Calendar_FormBackcolor=   16777215
         Calendar_BackColor=   14671839
         Calendar_ColorWeekDay=   8421376
         Calendar_Selected=   12640511
      End
      Begin VB.Label lblEndereco 
         AutoSize        =   -1  'True
         Caption         =   "Endereço:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   360
         TabIndex        =   14
         Top             =   600
         Width           =   930
      End
   End
   Begin VB.Frame FrmConsAluno 
      Caption         =   "Consulta por Aluno:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   240
      TabIndex        =   11
      Top             =   1920
      Width           =   9255
      Begin VB.CheckBox chkNomeAluno 
         Caption         =   "Check1"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   255
      End
      Begin cTextBox.cText txtNomeAluno 
         Height          =   255
         Left            =   360
         TabIndex        =   2
         Top             =   840
         Width           =   8655
         _ExtentX        =   15266
         _ExtentY        =   450
         BackColorGotFocus=   8454016
         BackColor_MouseMove=   16709609
         Appearance      =   0
         Alignment       =   0
         FontBold        =   0   'False
         FontSize        =   8,25
         FontName        =   "MS Sans Serif"
         BackColor       =   8454016
         AutoSelect      =   -1  'True
         DateFormat      =   "dd/MM/yy"
         FormatoExibData =   "__/__/____"
         tipoLetra       =   1
         Calendar_FormBackcolor=   16777215
         Calendar_BackColor=   14671839
         Calendar_ColorWeekDay=   8421376
         Calendar_Selected=   12640511
      End
      Begin VB.Label lblNomeAluno 
         AutoSize        =   -1  'True
         Caption         =   "Nome:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   360
         TabIndex        =   12
         Top             =   480
         Width           =   600
      End
   End
   Begin VB.Frame FrmConsData 
      Caption         =   "Consulta por data:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   240
      TabIndex        =   9
      Top             =   840
      Width           =   9255
      Begin cTextBox.cText txtDataCadastro 
         Height          =   345
         Left            =   4920
         TabIndex        =   0
         Top             =   360
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   609
         Text            =   "__/__/____"
         BackColorGotFocus=   8454016
         BackColor_MouseMove=   16709609
         MaxLength       =   10
         Appearance      =   0
         Alignment       =   2
         FontBold        =   0   'False
         FontSize        =   8,25
         FontName        =   "MS Sans Serif"
         BackColor       =   8454016
         TextType        =   2
         AutoSelect      =   -1  'True
         Mask            =   "__/__/____"
         DateFormat      =   "dd/MM/yyyy"
         FormatoExibData =   "__/__/____"
         Calendar_FormBackcolor=   16777215
         Calendar_BackColor=   14671839
         Calendar_ColorWeekDay=   8421376
         Calendar_Selected=   12640511
      End
      Begin ChamaleonBtn.chameleonButton cmdCalendarioDataCA 
         Height          =   375
         Left            =   6000
         TabIndex        =   17
         Top             =   360
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "?"
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
         MICON           =   "FrmDescricaoRelAlunos.frx":0000
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label lblDataCadastro 
         Caption         =   "Data de Cadastro:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   2880
         TabIndex        =   10
         Top             =   360
         Width           =   1995
      End
   End
   Begin ChamaleonBtn.chameleonButton cmdImprimir 
      DragIcon        =   "FrmDescricaoRelAlunos.frx":001C
      Height          =   1215
      Left            =   3600
      TabIndex        =   7
      Top             =   7560
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   2143
      BTYPE           =   3
      TX              =   "&Imprimir"
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
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmDescricaoRelAlunos.frx":0CE6
      PICN            =   "FrmDescricaoRelAlunos.frx":0D02
      UMCOL           =   -1  'True
      SOFT            =   -1  'True
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   2
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonBtn.chameleonButton cmdVoltar 
      Height          =   1215
      Left            =   5040
      TabIndex        =   8
      Top             =   7560
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   2143
      BTYPE           =   3
      TX              =   "&Voltar"
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
      MICON           =   "FrmDescricaoRelAlunos.frx":14DC
      PICN            =   "FrmDescricaoRelAlunos.frx":14F8
      UMCOL           =   -1  'True
      SOFT            =   -1  'True
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   2
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label lblMensagem 
      AutoSize        =   -1  'True
      Caption         =   "Marque as opções que deseja incluir no relatório"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   345
      Left            =   2160
      TabIndex        =   16
      Top             =   240
      Width           =   5535
   End
End
Attribute VB_Name = "FrmDescricaoRelAlunos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCalendarioDataCA_Click()
calendario = "dataca"
FrmCalendario.Show 1
End Sub

Private Sub cmdImprimir_Click()
SQL = "SELECT * FROM CAD_ALUNO"

SQL = SQL + " WHERE "



If txtDataCadastro.Text = "__/__/____" Then

    MsgBox "Por favor, digite a data corretamente!", vbInformation, "Mensagem do Sistema"
    
    Unload Me
    FrmDescricaoRelAlunos.Show 1
    
Else

    SQL = SQL + " DATA_CADASTRO =' " & Format(txtDataCadastro.Text, "yyyy-mm-dd") & " ' "
                      
End If
    

If chkNomeAluno.Value = Checked Then
  
    If txtNomeAluno.Text = "" Then
    
        MsgBox "Por favor, digite um nome no campo para continuar!", vbInformation, "Mensagem do Sistema"
        
        Unload Me
        FrmDescricaoRelAlunos.Show 1
        
    Else

     SQL = SQL + " AND NOME_ALUNO like UPPER('" & UCase(txtNomeAluno.Text) & "%')"
    
    End If
    
End If

If chkEndereco.Value = Checked Then
    
    If txtEndereco.Text = "" Then
    
        MsgBox "Por favor, digite um endereço no campo para continuar!", vbInformation, "Mensagem do Sistema"
        
        Unload Me
        FrmDescricaoRelAlunos.Show 1
        
    Else

     SQL = SQL + " AND ENDERECO like UPPER('" & UCase(txtEndereco.Text) & "%')"
    
    End If
    
End If

If chkAtivo.Value = Checked Then

    SQL = SQL + " AND STATUS_ALUNO like 'ATIVO'"
    
End If

If chkInativo.Value = Checked Then

    SQL = SQL + " AND STATUS_ALUNO like 'INATIVO'"
    
End If

Unload Me
RelAlunos.Config
Me.Show 1
End Sub

Private Sub cmdVoltar_Click()
Unload Me
FrmAlunos.Show 1
End Sub
