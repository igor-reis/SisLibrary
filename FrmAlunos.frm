VERSION 5.00
Object = "{CFAB6834-3B57-49FC-8770-CBA3667FE193}#1.0#0"; "ctextbox.ocx"
Object = "{61159A24-3E03-4E76-9CA9-2396C6822B8F}#1.0#0"; "chamaleonbtn.ocx"
Begin VB.Form FrmAlunos 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Alunos"
   ClientHeight    =   8850
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11655
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8850
   ScaleWidth      =   11655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer timerPrenObrigatorio 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   1440
      Top             =   2880
   End
   Begin cTextBox.cText txtNumero 
      Height          =   375
      Left            =   4320
      TabIndex        =   2
      Top             =   4560
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      BackColorGotFocus=   8454016
      BackColor_MouseMove=   16709609
      Appearance      =   0
      Alignment       =   0
      CaractLiberados =   "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ"
      FontBold        =   0   'False
      FontSize        =   8,25
      FontName        =   "MS Sans Serif"
      BackColor       =   8454016
      AutoSelect      =   -1  'True
      InserirZeros    =   0   'False
      DateFormat      =   "dd/MM/yy"
      FormatoExibData =   "__/__/____"
      tipoLetra       =   1
      TextoLivre      =   0   'False
      Calendar_FormBackcolor=   16777215
      Calendar_BackColor=   14671839
      Calendar_ColorWeekDay=   8421376
      Calendar_Selected=   12640511
   End
   Begin VB.ComboBox cmbStatus 
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      Height          =   315
      ItemData        =   "FrmAlunos.frx":0000
      Left            =   4320
      List            =   "FrmAlunos.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   6960
      Width           =   1815
   End
   Begin cTextBox.cText txtTelefone 
      Height          =   375
      Left            =   4320
      TabIndex        =   6
      Top             =   6360
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      Text            =   "(__) ____-_____"
      BackColorGotFocus=   8454016
      BackColor_MouseMove=   16709609
      MaxLength       =   15
      Appearance      =   0
      Alignment       =   2
      FontBold        =   0   'False
      FontSize        =   8,25
      FontName        =   "MS Sans Serif"
      BackColor       =   8454016
      TextType        =   3
      AutoSelect      =   -1  'True
      Mask            =   "(__) ____-_____"
      MaskRules       =   "(##) ####-#####"
      MaskSomenteNumeros=   -1  'True
      DateFormat      =   "dd/MM/yy"
      FormatoExibData =   "__/__/____"
      Calendar_FormBackcolor=   16777215
      Calendar_BackColor=   14671839
      Calendar_ColorWeekDay=   8421376
      Calendar_Selected=   12640511
   End
   Begin cTextBox.cText txtCidade 
      Height          =   375
      Left            =   4320
      TabIndex        =   4
      Top             =   5160
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   661
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
   Begin cTextBox.cText txtEndereco 
      Height          =   375
      Left            =   4320
      TabIndex        =   1
      Top             =   3960
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   661
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
   Begin cTextBox.cText txtNome 
      Height          =   375
      Left            =   4320
      TabIndex        =   0
      Top             =   3360
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   661
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
   Begin VB.ComboBox cmbEstado 
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      Height          =   315
      ItemData        =   "FrmAlunos.frx":001E
      Left            =   4320
      List            =   "FrmAlunos.frx":0073
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   5760
      Width           =   1815
   End
   Begin cTextBox.cText txtIdAluno 
      Height          =   375
      Left            =   1080
      TabIndex        =   8
      Top             =   1800
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BackColorGotFocus=   -2147483633
      BackColor_MouseMove=   16709609
      Enabled         =   0   'False
      Appearance      =   0
      Alignment       =   0
      FontBold        =   0   'False
      FontSize        =   8,25
      FontName        =   "MS Sans Serif"
      BackColor       =   -2147483633
      AutoSelect      =   -1  'True
      DateFormat      =   "dd/mm/yyyy"
      FormatoExibData =   "__/__/____"
      Calendar_FormBackcolor=   16777215
      Calendar_BackColor=   14671839
      Calendar_ColorWeekDay=   8421376
      Calendar_Selected=   12640511
   End
   Begin ChamaleonBtn.chameleonButton cmdSair 
      Height          =   855
      Left            =   4440
      TabIndex        =   9
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
      MICON           =   "FrmAlunos.frx":00E3
      PICN            =   "FrmAlunos.frx":00FF
      UMCOL           =   -1  'True
      SOFT            =   -1  'True
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   2
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
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
      MICON           =   "FrmAlunos.frx":3846
      PICN            =   "FrmAlunos.frx":3862
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
      MICON           =   "FrmAlunos.frx":6E8F
      PICN            =   "FrmAlunos.frx":6EAB
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
      MICON           =   "FrmAlunos.frx":9FF3
      PICN            =   "FrmAlunos.frx":A00F
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
      MICON           =   "FrmAlunos.frx":A6E3
      PICN            =   "FrmAlunos.frx":A6FF
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
      MICON           =   "FrmAlunos.frx":ADAB
      PICN            =   "FrmAlunos.frx":ADC7
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
      MICON           =   "FrmAlunos.frx":E430
      PICN            =   "FrmAlunos.frx":E44C
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
      MICON           =   "FrmAlunos.frx":11E2F
      PICN            =   "FrmAlunos.frx":11E4B
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
      DragIcon        =   "FrmAlunos.frx":171B3
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
      MICON           =   "FrmAlunos.frx":17E7D
      PICN            =   "FrmAlunos.frx":17E99
      UMCOL           =   -1  'True
      SOFT            =   -1  'True
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   2
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin cTextBox.cText txtBairro 
      Height          =   375
      Left            =   6840
      TabIndex        =   3
      Top             =   4560
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   661
      BackColorGotFocus=   8454016
      Appearance      =   0
      Alignment       =   0
      FontBold        =   0   'False
      FontSize        =   8,25
      FontName        =   "MS Sans Serif"
      BackColor       =   8454016
      InserirZeros    =   0   'False
      DateFormat      =   "dd/MM/yy"
      FormatoExibData =   "__/__/____"
      tipoLetra       =   1
      Calendar_FormBackcolor=   16777215
      Calendar_BackColor=   14671839
      Calendar_ColorWeekDay=   8421376
      Calendar_Selected=   12640511
   End
   Begin cTextBox.cText txtDataCadastro 
      Height          =   345
      Left            =   8400
      TabIndex        =   39
      Top             =   6960
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
      Left            =   6360
      TabIndex        =   40
      Top             =   6960
      Width           =   2055
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
      Left            =   6600
      TabIndex        =   38
      Top             =   4560
      Width           =   105
   End
   Begin VB.Label lblBairro 
      Caption         =   "Bairro:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5880
      TabIndex        =   37
      Top             =   4560
      Width           =   855
   End
   Begin VB.Label Label1 
      Height          =   1455
      Index           =   0
      Left            =   0
      TabIndex        =   36
      Top             =   240
      Width           =   12015
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
      TabIndex        =   27
      Top             =   3000
      Width           =   3165
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
      TabIndex        =   26
      Top             =   2880
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
      Index           =   8
      Left            =   2760
      TabIndex        =   25
      Top             =   5640
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
      Left            =   2880
      TabIndex        =   24
      Top             =   4440
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
      TabIndex        =   23
      Top             =   6840
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
      Left            =   2760
      TabIndex        =   22
      Top             =   5040
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
      Index           =   1
      Left            =   3000
      TabIndex        =   21
      Top             =   3840
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
      Left            =   3720
      TabIndex        =   20
      Top             =   3240
      Width           =   105
   End
   Begin VB.Label lblNumero 
      Caption         =   "Número:"
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
      TabIndex        =   19
      Top             =   4560
      Width           =   2055
   End
   Begin VB.Label lbIdAluno 
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
      Left            =   240
      TabIndex        =   18
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label Label1 
      Height          =   1215
      Index           =   1
      Left            =   0
      TabIndex        =   17
      Top             =   7680
      Width           =   11535
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
      TabIndex        =   16
      Top             =   2520
      Visible         =   0   'False
      Width           =   6660
   End
   Begin VB.Label lblNome 
      Caption         =   "Nome Completo:"
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
      TabIndex        =   15
      Top             =   3360
      Width           =   2055
   End
   Begin VB.Label lblCidade 
      Caption         =   "Cidade:"
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
      TabIndex        =   14
      Top             =   5160
      Width           =   2055
   End
   Begin VB.Label lblEndereco 
      Caption         =   "Endereço:"
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
      TabIndex        =   13
      Top             =   3960
      Width           =   2055
   End
   Begin VB.Label lblEstado 
      Caption         =   "Estado:"
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
      TabIndex        =   12
      Top             =   5760
      Width           =   2055
   End
   Begin VB.Label lblTelefone 
      Caption         =   "Telefone:"
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
      TabIndex        =   11
      Top             =   6360
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
      TabIndex        =   10
      Top             =   6960
      Width           =   2055
   End
End
Attribute VB_Name = "FrmAlunos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim limpaF As Control
Dim PrenObrigatorio As Integer

Private Sub cmdAlterar_Click()
If txtNome.Text = "" Or txtEndereco.Text = "" Or txtNumero.Text = "" Or txtBairro.Text = "" Or txtCidade.Text = "" Or cmbEstado.Text = "" Or cmbStatus.Text = "" Then
   MsgBox "Preencha todos os campos necessários!", vbInformation, "Mensagem do Sistema"
   timerPrenObrigatorio.Enabled = True
Else

Set OP = New ADODB.Command
With OP
  .ActiveConnection = cn
  .CommandText = "UPDATE CAD_ALUNO SET NOME_ALUNO = '" & txtNome.Text & "',ENDERECO = '" & txtEndereco.Text & "',NUMERO = '" & txtNumero.Text & "',BAIRRO = '" & txtBairro.Text & "',CIDADE = '" & txtCidade.Text & "',ESTADO = '" & cmbEstado.Text & "',TELEFONE = '" & txtTelefone.Value & "',STATUS_ALUNO = '" & cmbStatus.Text & "' WHERE ID_ALUNO = " & CInt(txtIdAluno.Text)
  .Execute
End With
 MsgBox ("Registro alterado com sucesso!"), vbInformation, "Mensagem do Sistema"
 
 cmdAlterar.Enabled = False
 cmdExcluir.Enabled = False
 cmdCancelar.Enabled = False
 cmdConsultar.Enabled = True
 cmdNovo.Enabled = True
 
 Unload Me

 FrmAlunos.Show 1
 
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

FrmAlunos.Show 1

End Sub

Private Sub cmdConsultar_Click()
NomeFormulario = "Alunos"
FrmConsultarAlunos.Show 1
End Sub


Private Sub cmdExcluir_Click()
If MsgBox("Você tem certeza que deseja excluir esse registro?", vbYesNo, "Mensagem do Sistema") = vbNo Then
Else
    Call conexao(cn)
    Rs.Open "DELETE FROM CAD_ALUNO WHERE ID_ALUNO= " & CInt(txtIdAluno.Text)
    MsgBox "Registro excluido com sucesso!", vbExclamation, "Mensagem do Sistema"
    
    Unload Me
    
    FrmAlunos.Show 1
    
    cmdExcluir.Enabled = False
    cmdAlterar.Enabled = False
    cmdCancelar.Enabled = False
    cmdNovo.Enabled = True
    cmdConsultar.Enabled = True
End If

End Sub

Private Sub cmdGravar_Click()
If txtNome.Text = "" Or txtEndereco.Text = "" Or txtNumero.Text = "" Or txtCidade.Text = "" Or cmbEstado.Text = "" Or cmbStatus.Text = "" Then
   MsgBox "Preencha todos os campos necessários!", vbInformation, "Mensagem do Sistema"
   timerPrenObrigatorio.Enabled = True
Else

Call conexao(cn)

Set OP = New ADODB.Command
    With OP
            .ActiveConnection = cn
            .CommandText = "insert into CAD_ALUNO(ID_ALUNO,NOME_ALUNO,ENDERECO,BAIRRO,NUMERO,CIDADE,ESTADO,TELEFONE,STATUS_ALUNO,DATA_CADASTRO) values ('" & txtIdAluno.Text & "','" & txtNome.Text & "','" & txtEndereco.Text & "','" & txtNumero.Text & "','" & txtBairro.Text & "','" & txtCidade.Text & "','" & cmbEstado.Text & "','" & txtTelefone.Value & "','" & cmbStatus.Text & "','" & Format(txtDataCadastro.Text, "yyyy-mm-dd") & "');"
            .Execute
    End With
    MsgBox "Registro salvo com sucesso!", vbInformation, "Mensagem do Sistema"
    
    cmdGravar.Enabled = False
    cmdCancelar.Enabled = False
    cmdConsultar.Enabled = True
    cmdNovo.Enabled = True
    
    Unload Me

    FrmAlunos.Show 1
    
End If
End Sub


Private Sub cmdNovo_Click()
txtIdAluno.Text = GeraID("GEN_CAD_ALUNO")
txtDataCadastro = Date
txtNome.SetFocus

lblMensagem.Visible = True
cmdGravar.Enabled = True
cmdCancelar.Enabled = True
cmdConsultar.Enabled = False
cmdNovo.Enabled = False

End Sub


Private Sub cmdRelatorio_Click()
Unload Me
FrmDescricaoRelAlunos.Show 1
End Sub

Private Sub cmdSair_Click()
If MsgBox("Você tem certeza que deseja sair do sistema?", vbYesNo, "Mensagem do Sistema") = vbYes Then
     'exclui
     End
Else
     'cancela
End If
End Sub


Public Function GeraID(ByVal GEN_CAD_ALUNO As String) As Long
Call conexao(cn)
    Set Rs = New ADODB.Recordset
        'Use a tabela RDB$DATABASE, pois ela sempre retorna um único registro
        Rs.Open "Select GEN_ID (" & GEN_CAD_ALUNO & ", 1) From RDB$DATABASE", cn
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

