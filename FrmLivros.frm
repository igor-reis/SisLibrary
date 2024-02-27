VERSION 5.00
Object = "{CFAB6834-3B57-49FC-8770-CBA3667FE193}#1.0#0"; "ctextbox.ocx"
Object = "{61159A24-3E03-4E76-9CA9-2396C6822B8F}#1.0#0"; "chamaleonbtn.ocx"
Begin VB.Form FrmLivros 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Livros"
   ClientHeight    =   8970
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11685
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8970
   ScaleWidth      =   11685
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer timerPrenObrigatorio 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   1440
      Top             =   3000
   End
   Begin VB.ComboBox cmbStatus 
      BackColor       =   &H80000002&
      Height          =   315
      ItemData        =   "FrmLivros.frx":0000
      Left            =   4080
      List            =   "FrmLivros.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   6480
      Width           =   1815
   End
   Begin ChamaleonBtn.chameleonButton cmdSair 
      Height          =   855
      Left            =   4440
      TabIndex        =   7
      Top             =   7800
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   1508
      BTYPE           =   3
      TX              =   "&Sair"
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
      MICON           =   "FrmLivros.frx":0028
      PICN            =   "FrmLivros.frx":0044
      UMCOL           =   -1  'True
      SOFT            =   -1  'True
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   2
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin cTextBox.cText txtIdLivro 
      Height          =   375
      Left            =   1200
      TabIndex        =   9
      Top             =   1920
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BackColorGotFocus=   -2147483633
      Enabled         =   0   'False
      Appearance      =   0
      Alignment       =   0
      FontBold        =   0   'False
      FontSize        =   8,25
      FontName        =   "MS Sans Serif"
      BackColor       =   -2147483633
      DateFormat      =   "dd/mm/yyyy"
      FormatoExibData =   "__/__/____"
      Calendar_FormBackcolor=   16777215
      Calendar_BackColor=   14671839
      Calendar_ColorWeekDay=   8421376
      Calendar_Selected=   12640511
   End
   Begin cTextBox.cText txtEditora 
      Height          =   375
      Left            =   4080
      TabIndex        =   5
      Top             =   5880
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   661
      BackColorGotFocus=   8454016
      Appearance      =   0
      Alignment       =   0
      FontBold        =   0   'False
      FontSize        =   8,25
      FontName        =   "MS Sans Serif"
      BackColor       =   8454016
      DateFormat      =   "dd/MM/yy"
      FormatoExibData =   "__/__/____"
      tipoLetra       =   1
      Calendar_FormBackcolor=   16777215
      Calendar_BackColor=   14671839
      Calendar_ColorWeekDay=   8421376
      Calendar_Selected=   12640511
   End
   Begin cTextBox.cText txtGenero 
      Height          =   375
      Left            =   4080
      TabIndex        =   4
      Top             =   5280
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   661
      BackColorGotFocus=   8454016
      Appearance      =   0
      Alignment       =   0
      FontBold        =   0   'False
      FontSize        =   8,25
      FontName        =   "MS Sans Serif"
      BackColor       =   8454016
      DateFormat      =   "dd/MM/yy"
      FormatoExibData =   "__/__/____"
      tipoLetra       =   1
      Calendar_FormBackcolor=   16777215
      Calendar_BackColor=   14671839
      Calendar_ColorWeekDay=   8421376
      Calendar_Selected=   12640511
   End
   Begin cTextBox.cText txtNPaginas 
      Height          =   375
      Left            =   8400
      TabIndex        =   3
      Top             =   4680
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      BackColorGotFocus=   8454016
      MaxLength       =   11
      Appearance      =   0
      Alignment       =   2
      FontBold        =   0   'False
      FontSize        =   8,25
      FontName        =   "MS Sans Serif"
      BackColor       =   8454016
      TextType        =   1
      InserirZeros    =   0   'False
      DateFormat      =   "dd/MM/yy"
      FormatoExibData =   "__/__/____"
      Calendar_FormBackcolor=   16777215
      Calendar_BackColor=   14671839
      Calendar_ColorWeekDay=   8421376
      Calendar_Selected=   12640511
   End
   Begin cTextBox.cText txtAno 
      Height          =   375
      Left            =   4080
      TabIndex        =   2
      Top             =   4680
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      BackColorGotFocus=   8454016
      MaxLength       =   4
      Appearance      =   0
      Alignment       =   2
      FontBold        =   0   'False
      FontSize        =   8,25
      FontName        =   "MS Sans Serif"
      BackColor       =   8454016
      TextType        =   1
      InserirZeros    =   0   'False
      DateFormat      =   "dd/MM/yy"
      FormatoExibData =   "__/__/____"
      Calendar_FormBackcolor=   16777215
      Calendar_BackColor=   14671839
      Calendar_ColorWeekDay=   8421376
      Calendar_Selected=   12640511
   End
   Begin cTextBox.cText txtAutor 
      Height          =   375
      Left            =   4080
      TabIndex        =   1
      Top             =   4080
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   661
      BackColorGotFocus=   8454016
      Appearance      =   0
      Alignment       =   0
      FontBold        =   0   'False
      FontSize        =   8,25
      FontName        =   "MS Sans Serif"
      BackColor       =   8454016
      DateFormat      =   "dd/MM/yy"
      FormatoExibData =   "__/__/____"
      tipoLetra       =   1
      Calendar_FormBackcolor=   16777215
      Calendar_BackColor=   14671839
      Calendar_ColorWeekDay=   8421376
      Calendar_Selected=   12640511
   End
   Begin cTextBox.cText txtNome 
      Height          =   375
      Left            =   4080
      TabIndex        =   0
      Top             =   3480
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   661
      BackColorGotFocus=   8454016
      Appearance      =   0
      Alignment       =   0
      FontBold        =   0   'False
      FontSize        =   8,25
      FontName        =   "MS Sans Serif"
      BackColor       =   8454016
      DateFormat      =   "dd/MM/yy"
      FormatoExibData =   "__/__/____"
      tipoLetra       =   1
      Calendar_FormBackcolor=   16777215
      Calendar_BackColor=   14671839
      Calendar_ColorWeekDay=   8421376
      Calendar_Selected=   12640511
   End
   Begin ChamaleonBtn.chameleonButton cmdVoltar 
      Height          =   1215
      Left            =   10200
      TabIndex        =   28
      Top             =   240
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
      MICON           =   "FrmLivros.frx":378B
      PICN            =   "FrmLivros.frx":37A7
      UMCOL           =   -1  'True
      SOFT            =   -1  'True
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   2
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonBtn.chameleonButton cmdConsultar 
      Height          =   1215
      Left            =   7320
      TabIndex        =   29
      Top             =   240
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   2143
      BTYPE           =   3
      TX              =   "C&onsultar"
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
      MICON           =   "FrmLivros.frx":6DD4
      PICN            =   "FrmLivros.frx":6DF0
      UMCOL           =   -1  'True
      SOFT            =   -1  'True
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   2
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonBtn.chameleonButton cmdAlterar 
      Height          =   1215
      Left            =   5880
      TabIndex        =   30
      Top             =   240
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   2143
      BTYPE           =   3
      TX              =   "&Alterar"
      ENAB            =   0   'False
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
      MICON           =   "FrmLivros.frx":9F38
      PICN            =   "FrmLivros.frx":9F54
      UMCOL           =   -1  'True
      SOFT            =   -1  'True
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   2
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonBtn.chameleonButton cmdExcluir 
      Height          =   1215
      Left            =   4440
      TabIndex        =   31
      Top             =   240
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   2143
      BTYPE           =   3
      TX              =   "&Excluir"
      ENAB            =   0   'False
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
      MICON           =   "FrmLivros.frx":A628
      PICN            =   "FrmLivros.frx":A644
      UMCOL           =   -1  'True
      SOFT            =   -1  'True
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   2
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonBtn.chameleonButton cmdCancelar 
      Height          =   1215
      Left            =   3000
      TabIndex        =   32
      Top             =   240
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   2143
      BTYPE           =   3
      TX              =   "&Cancelar"
      ENAB            =   0   'False
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
      MICON           =   "FrmLivros.frx":ACF0
      PICN            =   "FrmLivros.frx":AD0C
      UMCOL           =   -1  'True
      SOFT            =   -1  'True
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   2
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonBtn.chameleonButton cmdNovo 
      Height          =   1215
      Left            =   120
      TabIndex        =   33
      Top             =   240
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   2143
      BTYPE           =   3
      TX              =   "&Novo"
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
      MICON           =   "FrmLivros.frx":E375
      PICN            =   "FrmLivros.frx":E391
      UMCOL           =   -1  'True
      SOFT            =   -1  'True
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   2
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonBtn.chameleonButton cmdGravar 
      Height          =   1215
      Left            =   1560
      TabIndex        =   34
      Top             =   240
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   2143
      BTYPE           =   3
      TX              =   "&Gravar Dados"
      ENAB            =   0   'False
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
      MICON           =   "FrmLivros.frx":11D74
      PICN            =   "FrmLivros.frx":11D90
      UMCOL           =   -1  'True
      SOFT            =   -1  'True
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   2
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonBtn.chameleonButton cmdRelatorio 
      DragIcon        =   "FrmLivros.frx":170F8
      Height          =   1215
      Left            =   8760
      TabIndex        =   35
      Top             =   240
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   2143
      BTYPE           =   3
      TX              =   "&Relatório"
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
      MICON           =   "FrmLivros.frx":17DC2
      PICN            =   "FrmLivros.frx":17DDE
      UMCOL           =   -1  'True
      SOFT            =   -1  'True
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   2
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin cTextBox.cText txtDataCadastro 
      Height          =   345
      Left            =   8520
      TabIndex        =   37
      Top             =   6480
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   609
      Text            =   "__/__/____"
      BackColorGotFocus=   8454016
      BackColor_MouseMove=   16709609
      MaxLength       =   10
      Enabled         =   0   'False
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
      Height          =   300
      Left            =   6480
      TabIndex        =   38
      Top             =   6480
      Width           =   2055
   End
   Begin VB.Label Label1 
      Height          =   1455
      Index           =   0
      Left            =   0
      TabIndex        =   36
      Top             =   240
      Width           =   12015
   End
   Begin VB.Label lblObrigatorio 
      AutoSize        =   -1  'True
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   360
      Index           =   7
      Left            =   8040
      TabIndex        =   27
      Top             =   4560
      Width           =   105
   End
   Begin VB.Label lblObrigatorio 
      AutoSize        =   -1  'True
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   360
      Index           =   6
      Left            =   2760
      TabIndex        =   26
      Top             =   6360
      Width           =   105
   End
   Begin VB.Label lblObrigatorio 
      AutoSize        =   -1  'True
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   360
      Index           =   4
      Left            =   2760
      TabIndex        =   25
      Top             =   5760
      Width           =   105
   End
   Begin VB.Label lblObrigatorio 
      AutoSize        =   -1  'True
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   360
      Index           =   3
      Left            =   3480
      TabIndex        =   24
      Top             =   3360
      Width           =   105
   End
   Begin VB.Label lblObrigatorio 
      AutoSize        =   -1  'True
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   360
      Index           =   2
      Left            =   2640
      TabIndex        =   23
      Top             =   3960
      Width           =   105
   End
   Begin VB.Label lblObrigatorio 
      AutoSize        =   -1  'True
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Index           =   1
      Left            =   2520
      TabIndex        =   22
      Top             =   4560
      Width           =   105
   End
   Begin VB.Label lblObrigatorio 
      AutoSize        =   -1  'True
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   360
      Index           =   0
      Left            =   2880
      TabIndex        =   21
      Top             =   5160
      Width           =   105
   End
   Begin VB.Label lblObrigatorio 
      AutoSize        =   -1  'True
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   360
      Index           =   5
      Left            =   1920
      TabIndex        =   20
      Top             =   3000
      Width           =   105
   End
   Begin VB.Label lblPrenObrigatório 
      AutoSize        =   -1  'True
      Caption         =   "Campo de preenchimento obrigatório!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   2040
      TabIndex        =   19
      Top             =   3120
      Width           =   3165
   End
   Begin VB.Label lblMensagem 
      AutoSize        =   -1  'True
      Caption         =   "Peencha os dados corretamente e clique em Gravar Dados"
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
      Left            =   2280
      TabIndex        =   18
      Top             =   2640
      Visible         =   0   'False
      Width           =   6660
   End
   Begin VB.Label lblNomeLivro 
      Caption         =   "Nome do livro:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   1920
      TabIndex        =   17
      Top             =   3480
      Width           =   2055
   End
   Begin VB.Label lblAno 
      Caption         =   "Ano:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   1920
      TabIndex        =   16
      Top             =   4680
      Width           =   2055
   End
   Begin VB.Label lblAutor 
      Caption         =   "Autor:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   1920
      TabIndex        =   15
      Top             =   4080
      Width           =   2055
   End
   Begin VB.Label lblGenero 
      Caption         =   "Genêro:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   1920
      TabIndex        =   14
      Top             =   5280
      Width           =   2055
   End
   Begin VB.Label lblEditora 
      Caption         =   "Editora:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   1920
      TabIndex        =   13
      Top             =   5880
      Width           =   2055
   End
   Begin VB.Label lblStatus 
      Caption         =   "Status:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   1920
      TabIndex        =   12
      Top             =   6480
      Width           =   2055
   End
   Begin VB.Label lblPaginas 
      Caption         =   "Número de Páginas:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   5880
      TabIndex        =   11
      Top             =   4680
      Width           =   2295
   End
   Begin VB.Label lbIdLivro 
      Caption         =   "Código:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   360
      TabIndex        =   10
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label Label1 
      Height          =   1215
      Index           =   1
      Left            =   0
      TabIndex        =   8
      Top             =   7680
      Width           =   11535
   End
End
Attribute VB_Name = "FrmLivros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim limpaF As Control
Dim PrenObrigatorio As Integer

Private Sub cmdAlterar_Click()
If txtNome.Text = "" Or txtAutor.Text = "" Or txtAno.Text = "" Or txtNPaginas.Text = "" Or txtGenero.Text = "" Or txtEditora.Text = "" Or cmbStatus.Text = "" Then
   MsgBox "Preencha todos os campos necessários!", vbInformation, "Mensagem do Sistema"
   timerPrenObrigatorio.Enabled = True
Else

Set OP = New ADODB.Command
With OP
  .ActiveConnection = cn
  .CommandText = "UPDATE CAD_LIVRO SET NOME_LIVRO = '" & txtNome.Text & "',AUTOR = '" & txtAutor.Text & "',ANO = '" & txtAno.Text & "',NUMERO_PAGINAS = '" & txtNPaginas.Text & "',GENERO = '" & txtGenero.Text & "',EDITORA = '" & txtEditora.Text & "',STATUS_LIVRO = '" & cmbStatus.Text & "' WHERE ID_LIVRO = " & CInt(txtIdLivro.Text)
  .Execute
End With
 MsgBox ("Registro alterado com sucesso!"), vbInformation, "Mensagem do Sistema"
 
 cmdAlterar.Enabled = False
 cmdExcluir.Enabled = False
 cmdCancelar.Enabled = False
 cmdConsultar.Enabled = True
 cmdNovo.Enabled = True
 
 Unload Me

 FrmLivros.Show 1
    
End If
End Sub

Private Sub cmdCancelar_Click()
cmdGravar.Enabled = False
cmdCancelar.Enabled = False
cmdNovo.Enabled = True
cmdConsultar.Enabled = True
cmdVoltar.Enabled = True
cmdAlterar.Enabled = False
cmdExcluir.Enabled = False

Unload Me

FrmLivros.Show 1
End Sub

Private Sub cmdConsultar_Click()
NomeFormulario = "Livros"
FrmConsultarLivros.Show 1
End Sub


Private Sub cmdExcluir_Click()
If MsgBox("Você tem certeza que deseja excluir esse registro?", vbYesNo, "Mensagem do Sistema") = vbNo Then
Else
    Call conexao(cn)
    Rs.Open "DELETE FROM CAD_LIVRO WHERE ID_LIVRO= " & CInt(txtIdLivro.Text)
    MsgBox "Registro excluido com sucesso!", vbExclamation, "Mensagem do Sistema"
    
    Unload Me
    
    FrmLivros.Show 1
    
    cmdExcluir.Enabled = False
    cmdAlterar.Enabled = False
    cmdCancelar.Enabled = False
    cmdNovo.Enabled = True
    cmdConsultar.Enabled = True
End If
End Sub

Private Sub cmdGravar_Click()
If txtNome.Text = "" Or txtAutor.Text = "" Or txtAno.Text = "" Or txtNPaginas.Text = "" Or txtGenero.Text = "" Or txtEditora.Text = "" Or cmbStatus.Text = "" Then
   MsgBox "Preencha todos os campos necessários!", vbInformation, "Mensagem do Sistema"
   timerPrenObrigatorio.Enabled = True
Else

Call conexao(cn)

Set OP = New ADODB.Command
    With OP
            .ActiveConnection = cn
            .CommandText = "insert into CAD_LIVRO(ID_LIVRO,NOME_LIVRO,AUTOR,ANO,NUMERO_PAGINAS,GENERO,EDITORA,STATUS_LIVRO,DATA_CADASTRO) values ('" & txtIdLivro.Text & "','" & txtNome.Text & "','" & txtAutor.Text & "','" & txtAno.Text & "','" & txtNPaginas.Text & "','" & txtGenero.Text & "','" & txtEditora & "','" & cmbStatus.Text & "','" & Format(txtDataCadastro.Text, "yyyy-mm-dd") & "');"
            .Execute
    End With
    MsgBox "Registro salvo com sucesso!", vbInformation, "Mensagem do Sistema"
    
    cmdGravar.Enabled = False
    cmdCancelar.Enabled = False
    cmdConsultar.Enabled = True
    cmdNovo.Enabled = True
    
    Unload Me

    FrmLivros.Show 1
    
End If
End Sub


Private Sub cmdNovo_Click()
txtIdLivro.Text = GeraID("GEN_CAD_LIVRO")
txtDataCadastro.Text = Date
txtNome.SetFocus

lblMensagem.Visible = True
cmdGravar.Enabled = True
cmdCancelar.Enabled = True
cmdConsultar.Enabled = False
cmdNovo.Enabled = False

End Sub


Private Sub cmdRelatorio_Click()
Unload Me
FrmDescricaoRelLivros.Show 1
End Sub

Private Sub cmdSair_Click()
If MsgBox("Você tem certeza que deseja sair do sistema?", vbYesNo, "Mensagem do Sistema") = vbYes Then
     'exclui
     End
Else
     'cancela
End If
End Sub


Public Function GeraID(ByVal GEN_CAD_LIVRO As String) As Long
Call conexao(cn)
    Set Rs = New ADODB.Recordset
        'Use a tabela RDB$DATABASE, pois ela sempre retorna um único registro
        Rs.Open "Select GEN_ID (" & GEN_CAD_LIVRO & ", 1) From RDB$DATABASE", cn
        GeraID = Rs(0)
        Rs.Close
    Set Rs = Nothing
End Function

 Public Function LimpaControles()
 'Limpa todos campos dos Ctext
 For Each limpaF In Controls
    If TypeOf limpaF Is cText Then
    limpaF.Text = ""
    End If
Next limpaF

End Function

Private Sub cmdVoltar_Click()
Unload Me
FrmPrincipal.Show 1
End Sub

Private Sub timerPrenObrigatorio_Timer()
 PrenObrigatorio = PrenObrigatorio + 1
        
        If PrenObrigatorio = 1 Then
            lblPrenObrigatório.Visible = False
        End If
    
        If PrenObrigatorio = 2 Then
            lblPrenObrigatório.Visible = True
            PrenObrigatorio = 0
        End If
        
End Sub
