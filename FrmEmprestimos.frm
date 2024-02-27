VERSION 5.00
Object = "{CFAB6834-3B57-49FC-8770-CBA3667FE193}#1.0#0"; "ctextbox.ocx"
Object = "{61159A24-3E03-4E76-9CA9-2396C6822B8F}#1.0#0"; "chamaleonbtn.ocx"
Begin VB.Form FrmEmprestimos 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Empréstimos"
   ClientHeight    =   7275
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   13020
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7275
   ScaleWidth      =   13020
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin ChamaleonBtn.chameleonButton cmdCalendarioDataE 
      Height          =   375
      Left            =   3720
      TabIndex        =   47
      Top             =   3120
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
      MICON           =   "FrmEmprestimos.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonBtn.chameleonButton cmdCalendarioDataM 
      Height          =   375
      Left            =   8040
      TabIndex        =   46
      Top             =   3120
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
      MICON           =   "FrmEmprestimos.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Timer TimerChk 
      Interval        =   1000
      Left            =   10920
      Top             =   2880
   End
   Begin cTextBox.cText txtDataDevolucao 
      Height          =   345
      Left            =   10320
      TabIndex        =   45
      Top             =   3600
      Visible         =   0   'False
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
   Begin VB.CheckBox chkDevolvido 
      Caption         =   "Devolvido"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   9360
      TabIndex        =   42
      Top             =   3120
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Timer timerPrenObrigatorio 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   2280
      Top             =   2640
   End
   Begin cTextBox.cText txtIdLivro 
      Height          =   255
      Left            =   480
      TabIndex        =   28
      Top             =   5640
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   450
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
   Begin cTextBox.cText txtDataMDevolucao 
      Height          =   345
      Left            =   6960
      TabIndex        =   2
      Top             =   3120
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
   Begin cTextBox.cText txtDataEmprestimo 
      Height          =   345
      Left            =   2640
      TabIndex        =   1
      Top             =   3120
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
   Begin cTextBox.cText txtIdEmprestimo 
      Height          =   375
      Left            =   1080
      TabIndex        =   12
      Top             =   1920
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
   Begin cTextBox.cText txtNomeLivro 
      Height          =   255
      Left            =   1920
      TabIndex        =   13
      Top             =   5640
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   450
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
   Begin cTextBox.cText txtNomeAluno 
      Height          =   255
      Left            =   1920
      TabIndex        =   14
      Top             =   4320
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   450
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
   Begin ChamaleonBtn.chameleonButton cmdProcurarLivro 
      Height          =   255
      Left            =   1440
      TabIndex        =   4
      Top             =   5640
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   450
      BTYPE           =   3
      TX              =   "?"
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
      MICON           =   "FrmEmprestimos.frx":0038
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonBtn.chameleonButton cmdProcurarAluno 
      Height          =   255
      Left            =   1440
      TabIndex        =   3
      Top             =   4320
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   450
      BTYPE           =   3
      TX              =   "?"
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
      MICON           =   "FrmEmprestimos.frx":0054
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin cTextBox.cText txtIdAluno 
      Height          =   255
      Left            =   480
      TabIndex        =   15
      Top             =   4320
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   450
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
   Begin ChamaleonBtn.chameleonButton cmdVoltar 
      Height          =   1215
      Left            =   10680
      TabIndex        =   10
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
      MICON           =   "FrmEmprestimos.frx":0070
      PICN            =   "FrmEmprestimos.frx":008C
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
      Left            =   7800
      TabIndex        =   7
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
      MICON           =   "FrmEmprestimos.frx":36B9
      PICN            =   "FrmEmprestimos.frx":36D5
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
      Left            =   6360
      TabIndex        =   8
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
      MICON           =   "FrmEmprestimos.frx":681D
      PICN            =   "FrmEmprestimos.frx":6839
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
      Left            =   4920
      TabIndex        =   9
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
      MICON           =   "FrmEmprestimos.frx":6F0D
      PICN            =   "FrmEmprestimos.frx":6F29
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
      Left            =   3480
      TabIndex        =   6
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
      MICON           =   "FrmEmprestimos.frx":75D5
      PICN            =   "FrmEmprestimos.frx":75F1
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
      Left            =   600
      TabIndex        =   0
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
      MICON           =   "FrmEmprestimos.frx":AC5A
      PICN            =   "FrmEmprestimos.frx":AC76
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
      Left            =   2040
      TabIndex        =   5
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
      MICON           =   "FrmEmprestimos.frx":E659
      PICN            =   "FrmEmprestimos.frx":E675
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
      Left            =   5280
      TabIndex        =   11
      Top             =   6120
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
      MICON           =   "FrmEmprestimos.frx":139DD
      PICN            =   "FrmEmprestimos.frx":139F9
      UMCOL           =   -1  'True
      SOFT            =   -1  'True
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   2
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonBtn.chameleonButton cmdRelatorio 
      DragIcon        =   "FrmEmprestimos.frx":17140
      Height          =   1215
      Left            =   9240
      TabIndex        =   43
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
      MICON           =   "FrmEmprestimos.frx":17E0A
      PICN            =   "FrmEmprestimos.frx":17E26
      UMCOL           =   -1  'True
      SOFT            =   -1  'True
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   2
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label lblDataDevolucao 
      Caption         =   "Data Devolução:"
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
      Left            =   8400
      TabIndex        =   44
      Top             =   3600
      Visible         =   0   'False
      Width           =   1830
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
      Height          =   255
      Left            =   8400
      TabIndex        =   41
      Top             =   3120
      Visible         =   0   'False
      Width           =   855
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
      Left            =   6840
      TabIndex        =   40
      Top             =   3000
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
      Left            =   2520
      TabIndex        =   39
      Top             =   3000
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
      Left            =   2520
      TabIndex        =   38
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
      Index           =   2
      Left            =   1200
      TabIndex        =   37
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
      Index           =   1
      Left            =   2520
      TabIndex        =   36
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
      Left            =   1200
      TabIndex        =   35
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
      Index           =   5
      Left            =   2760
      TabIndex        =   34
      Top             =   2640
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
      Left            =   2880
      TabIndex        =   33
      Top             =   2760
      Width           =   3165
   End
   Begin VB.Label lblIdLivro 
      AutoSize        =   -1  'True
      Caption         =   "Código:"
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
      Index           =   3
      Left            =   480
      TabIndex        =   32
      Top             =   5280
      Width           =   705
   End
   Begin VB.Label lblIdAluno 
      AutoSize        =   -1  'True
      Caption         =   "Código:"
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
      Index           =   2
      Left            =   480
      TabIndex        =   31
      Top             =   3960
      Width           =   705
   End
   Begin VB.Label lblNomeLivro 
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
      Index           =   1
      Left            =   1920
      TabIndex        =   30
      Top             =   5280
      Width           =   600
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
      Left            =   1920
      TabIndex        =   29
      Top             =   3960
      Width           =   600
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
      Left            =   2880
      TabIndex        =   27
      Top             =   2280
      Visible         =   0   'False
      Width           =   6660
   End
   Begin VB.Label Label1 
      Height          =   1215
      Index           =   2
      Left            =   840
      TabIndex        =   26
      Top             =   6000
      Width           =   11535
   End
   Begin VB.Label Label1 
      Height          =   1455
      Index           =   0
      Left            =   480
      TabIndex        =   25
      Top             =   240
      Width           =   12015
   End
   Begin VB.Label lblIdEmprestimo 
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
      Height          =   315
      Index           =   1
      Left            =   240
      TabIndex        =   24
      Top             =   1920
      Width           =   795
   End
   Begin VB.Label Label2 
      Caption         =   "Aluno"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   0
      Left            =   480
      TabIndex        =   23
      Top             =   3600
      Width           =   12495
   End
   Begin VB.Label Label3 
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
      Height          =   375
      Index           =   0
      Left            =   480
      TabIndex        =   22
      Top             =   3960
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Livro"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   1
      Left            =   480
      TabIndex        =   21
      Top             =   4920
      Width           =   12615
   End
   Begin VB.Label Label3 
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
      Height          =   375
      Index           =   1
      Left            =   480
      TabIndex        =   20
      Top             =   5280
      Width           =   975
   End
   Begin VB.Label lblDataEmprestimo 
      Caption         =   "Data Empréstimo:"
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
      Left            =   600
      TabIndex        =   19
      Top             =   3120
      Width           =   2115
   End
   Begin VB.Label lblDataMDevolucao 
      Caption         =   "Data Marcada Devolução:"
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
      Index           =   1
      Left            =   4080
      TabIndex        =   18
      Top             =   3120
      Width           =   2955
   End
   Begin VB.Label Label3 
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
      Height          =   375
      Index           =   2
      Left            =   1920
      TabIndex        =   17
      Top             =   3960
      Width           =   975
   End
   Begin VB.Label Label3 
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
      Height          =   375
      Index           =   3
      Left            =   1920
      TabIndex        =   16
      Top             =   5280
      Width           =   975
   End
End
Attribute VB_Name = "FrmEmprestimos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim limpaF As Control
Dim PrenObrigatorio As Integer
Dim status As String



Private Sub cmdAlterar_Click()
If txtDataEmprestimo.Text = "" Or txtDataMDevolucao.Text = "" Or txtIdAluno.Text = "" Or txtNomeAluno.Text = "" Or txtIdLivro.Text = "" Or txtNomeLivro.Text = "" Then
   MsgBox "Preencha todos os campos necessários!", vbInformation, "Mensagem do Sistema"
   timerPrenObrigatorio.Enabled = True
Else

   If chkDevolvido.Value = Checked Then
   
      status = "DEVOLVIDO"
   
   Else
      
      Rs.Open
      
      If ((Date) >= (Rs!DATA_M_DEVOLUCAO)) Then status = "ATRASADO" Else status = "PENDENTE"
      
      Rs.Close
      
   End If
   
   If txtDataDevolucao.Text <> "__/__/____" Then
     Set OP = New ADODB.Command
        With OP
          .ActiveConnection = cn
          .CommandText = "UPDATE EMPRESTIMOS SET ID_ALUNO = '" & txtIdAluno.Text & "',ID_LIVRO = '" & txtIdLivro.Text & "',DATA_EMPRESTIMO = '" & Format(txtDataEmprestimo.Text, "yyyy-mm-dd") & "',DATA_M_DEVOLUCAO = '" & Format(txtDataMDevolucao.Text, "yyyy-mm-dd") & "',DATA_DEVOLUCAO = '" & Format(txtDataDevolucao.Text, "yyyy-mm-dd") & "', STATUS = '" & status & "' WHERE ID_EMPRESTIMO = " & CInt(txtIdEmprestimo.Text)
          .Execute
        End With
        
    Else

     Set OP = New ADODB.Command
        With OP
          .ActiveConnection = cn
          .CommandText = "UPDATE EMPRESTIMOS SET ID_ALUNO = '" & txtIdAluno.Text & "',ID_LIVRO = '" & txtIdLivro.Text & "',DATA_EMPRESTIMO = '" & Format(txtDataEmprestimo.Text, "yyyy-mm-dd") & "',DATA_M_DEVOLUCAO = '" & Format(txtDataMDevolucao.Text, "yyyy-mm-dd") & "',STATUS = '" & status & "' WHERE ID_EMPRESTIMO = " & CInt(txtIdEmprestimo.Text)
          .Execute
        End With
        
    End If
 MsgBox ("Registro alterado com sucesso!"), vbInformation, "Mensagem do Sistema"
 
 cmdAlterar.Enabled = False
 cmdExcluir.Enabled = False
 cmdCancelar.Enabled = False
 cmdConsultar.Enabled = True
 cmdNovo.Enabled = True
 chkDevolvido.Visible = False
 lblStatus.Visible = False
 
 Call LimpaControles
 
 chkDevolvido.Value = Unchecked
 
End If
End Sub

Private Sub cmdCalendarioDataE_Click()
calendario = "datae"
FrmCalendario.Show 1
End Sub

Private Sub cmdCalendarioDataM_Click()
calendario = "datam"
FrmCalendario.Show 1
End Sub

Private Sub cmdCancelar_Click()
cmdGravar.Enabled = False
cmdCancelar.Enabled = False
cmdNovo.Enabled = True
cmdConsultar.Enabled = True
cmdVoltar.Enabled = True
cmdAlterar.Enabled = False
cmdExcluir.Enabled = False
cmdProcurarAluno.Enabled = False
cmdProcurarLivro.Enabled = False
chkDevolvido.Visible = False
lblStatus.Visible = False
timerPrenObrigatorio.Enabled = False

Unload Me

FrmEmprestimos.Show 1

End Sub

Private Sub cmdConsultar_Click()
FrmConsultarEmprestimos.Show 1
End Sub


Private Sub cmdExcluir_Click()
If MsgBox("Você tem certeza que deseja excluir esse registro?", vbYesNo, "Mensagem do Sistema") = vbNo Then
Else
    Call conexao(cn)
    Rs.Open "DELETE FROM EMPRESTIMOS WHERE ID_EMPRESTIMO= " & CInt(txtIdEmprestimo.Text)
    MsgBox "Registro excluido com sucesso!", vbExclamation, "Mensagem do Sistema"
    
    Unload Me
    
    FrmEmprestimos.Show 1
 
End If
End Sub

Private Sub cmdGravar_Click()
If txtDataEmprestimo.Text = "" Or txtDataMDevolucao.Text = "" Or txtIdAluno.Text = "" Or txtNomeAluno.Text = "" Or txtIdLivro.Text = "" Or txtNomeLivro.Text = "" Then
   MsgBox "Preencha todos os campos necessários!", vbInformation, "Mensagem do Sistema"
   timerPrenObrigatorio.Enabled = True
Else

Call conexao(cn)

Set OP = New ADODB.Command
    With OP
            .ActiveConnection = cn
            .CommandText = "insert into EMPRESTIMOS(ID_EMPRESTIMO,ID_ALUNO,ID_LIVRO,DATA_EMPRESTIMO,DATA_M_DEVOLUCAO) values ('" & txtIdEmprestimo.Text & "','" & txtIdAluno.Text & "','" & txtIdLivro.Text & "','" & Format(txtDataEmprestimo.Text, "yyyy-mm-dd") & "','" & Format(txtDataMDevolucao.Text, "yyyy-mm-dd") & "');"
            .Execute
    End With
    MsgBox "Registro salvo com sucesso!", vbInformation, "Mensagem do Sistema"
    
    cmdGravar.Enabled = False
    cmdCancelar.Enabled = False
    cmdConsultar.Enabled = True
    cmdNovo.Enabled = True
    
    Unload Me
    
    FrmEmprestimos.Show 1
    
End If
End Sub


Private Sub cmdRelatorio_Click()
Unload Me
FrmDescricaoRelEmprestimos.Show 1
End Sub

Private Sub cmdNovo_Click()
cmdProcurarAluno.Enabled = True
cmdProcurarLivro.Enabled = True

txtDataEmprestimo.Text = Date
'txtDataDevolucao.Text = Date + 15
txtDataMDevolucao.SetFocus
txtIdEmprestimo.Text = GeraID("GEN_EMPRESTIMO")

lblMensagem.Visible = True
cmdGravar.Enabled = True
cmdCancelar.Enabled = True
cmdConsultar.Enabled = False
cmdNovo.Enabled = False
End Sub


Private Sub cmdProcurarAluno_Click()
NomeFormulario = "Empréstimos"

FrmConsultarAlunos.Show 1
End Sub

Private Sub cmdProcurarLivro_Click()
NomeFormulario = "Empréstimos"

FrmConsultarLivros.Show 1
End Sub

Private Sub cmdSair_Click()
If MsgBox("Você tem certeza que deseja sair do sistema?", vbYesNo, "Mensagem do Sistema") = vbYes Then
     'exclui
     End
Else
     'cancela
End If
End Sub


Public Function GeraID(ByVal GEN_EMPRESTIMO As String) As Long
Call conexao(cn)
    Set Rs = New ADODB.Recordset
        'Use a tabela RDB$DATABASE, pois ela sempre retorna um único registro
        Rs.Open "Select GEN_ID (" & GEN_EMPRESTIMO & ", 1) From RDB$DATABASE", cn
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
Call conexao(cn)

    Rs.Open "SELECT * FROM EMPRESTIMOS WHERE STATUS = 'ATRASADO'", cn
    
        
    If Rs.RecordCount > 0 Then
                
        Rs.Close
        Set Rs = Nothing
        cn.Close
        Set cn = Nothing
        
        FrmPrincipal.imgAtencao.Visible = True
        FrmPrincipal.lblAtencao.Visible = True
        FrmPrincipal.TimerImagem.Enabled = True
        FrmPrincipal.TimerLabel.Enabled = True
        
    Else
             
        Rs.Close
        Set Rs = Nothing
        cn.Close
        Set cn = Nothing
        
        FrmPrincipal.imgAtencao.Visible = False
        FrmPrincipal.lblAtencao.Visible = False
        FrmPrincipal.TimerImagem.Enabled = False
        FrmPrincipal.TimerLabel.Enabled = False
    End If
Unload Me
FrmPrincipal.Show 1
End Sub


Private Sub TimerChk_Timer()
If chkDevolvido.Value = Checked Then

lblDataDevolucao.Visible = True
txtDataDevolucao.Visible = True
txtDataDevolucao.Text = Date

Else

lblDataDevolucao.Visible = False
txtDataDevolucao.Visible = False
txtDataDevolucao.Text = ""

End If
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
