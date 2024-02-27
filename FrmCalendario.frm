VERSION 5.00
Object = "{61159A24-3E03-4E76-9CA9-2396C6822B8F}#1.0#0"; "chamaleonbtn.ocx"
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Begin VB.Form FrmCalendario 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3585
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5400
   LinkTopic       =   "Form1"
   ScaleHeight     =   3585
   ScaleWidth      =   5400
   ShowInTaskbar   =   0   'False
   Begin ChamaleonBtn.chameleonButton cmdFechar 
      Height          =   375
      Left            =   4680
      TabIndex        =   1
      Top             =   120
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "X"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   255
      FCOLO           =   255
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmCalendario.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSACAL.Calendar Calendar1 
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   5175
      _Version        =   524288
      _ExtentX        =   9128
      _ExtentY        =   4895
      _StockProps     =   1
      BackColor       =   -2147483633
      Year            =   2013
      Month           =   9
      Day             =   1
      DayLength       =   1
      MonthLength     =   1
      DayFontColor    =   0
      FirstDay        =   7
      GridCellEffect  =   1
      GridFontColor   =   10485760
      GridLinesColor  =   -2147483632
      ShowDateSelectors=   -1  'True
      ShowDays        =   -1  'True
      ShowHorizontalGrid=   -1  'True
      ShowTitle       =   -1  'True
      ShowVerticalGrid=   -1  'True
      TitleFontColor  =   10485760
      ValueIsNull     =   0   'False
      BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "FrmCalendario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Calendar1_Click()
If calendario = "datae" Then

FrmEmprestimos.txtDataEmprestimo.Text = Calendar1.Value

End If

If calendario = "datam" Then

FrmEmprestimos.txtDataMDevolucao.Text = Calendar1.Value

End If

If calendario = "dataca" Then

FrmDescricaoRelAlunos.txtDataCadastro.Text = Calendar1.Value

End If

If calendario = "datai" Then

FrmDescricaoRelEmprestimos.txtDataEmprestimoInicial.Text = Calendar1.Value

End If

If calendario = "dataf" Then

FrmDescricaoRelEmprestimos.txtDataEmprestimoFinal.Text = Calendar1.Value

End If

If calendario = "datacl" Then

FrmDescricaoRelLivros.txtDataCadastro.Text = Calendar1.Value

End If


End Sub

Private Sub cmdFechar_Click()
Unload Me
End Sub
