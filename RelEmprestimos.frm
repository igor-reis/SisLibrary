VERSION 5.00
Object = "{D2618305-B2BB-11D2-925E-444553540000}#1.4#0"; "ReportX.ocx"
Begin VB.Form RelEmprestimos 
   Caption         =   "Relatório Empréstimos"
   ClientHeight    =   5370
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11535
   LinkTopic       =   "Form1"
   ScaleHeight     =   94.721
   ScaleMode       =   6  'Millimeter
   ScaleWidth      =   203.465
   StartUpPosition =   3  'Windows Default
   Begin ReportX.ReportMain Relatorio 
      Height          =   480
      Left            =   120
      TabIndex        =   0
      Top             =   4800
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   847
      Pagina          =   9
      Titulo          =   ""
      Registrado      =   0   'False
   End
   Begin ReportX.ReportSection ReportSection3 
      Align           =   1  'Align Top
      Height          =   1095
      Left            =   0
      Top             =   3840
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   1931
      Tipo            =   7
      Ordem           =   1
      Begin ReportX.ReportField ReportField2 
         Height          =   390
         Left            =   10320
         TabIndex        =   2
         Top             =   600
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   688
         Linhas          =   2
         Campo           =   "=Página [Pagina]"
         Caption         =   ""
         Formula         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin ReportX.ReportSection ReportSection1 
      Align           =   1  'Align Top
      Height          =   1215
      Left            =   0
      Top             =   0
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   2143
      Tipo            =   2
      Begin ReportX.ReportField ReportField1 
         Height          =   435
         Left            =   4080
         TabIndex        =   1
         Top             =   360
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   767
         Caption         =   "Relatório Empréstimo"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin ReportX.ReportSection ReportSection2 
      Align           =   1  'Align Top
      Height          =   2625
      Left            =   0
      Top             =   1215
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   4630
      Begin ReportX.ReportField ReportField16 
         Height          =   300
         Left            =   7320
         TabIndex        =   16
         Top             =   120
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   529
         Caption         =   "Status"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   16711680
      End
      Begin ReportX.ReportField ReportField15 
         Height          =   240
         Left            =   3960
         TabIndex        =   15
         Top             =   480
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   423
         Campo           =   "DATA_M_DEVOLUCAO"
         Caption         =   "Data Marcada Devolução"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField ReportField11 
         Height          =   240
         Left            =   1560
         TabIndex        =   14
         Top             =   480
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   423
         Campo           =   "DATA_EMPRESTIMO"
         Caption         =   "Data Empréstimo"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField ReportField5 
         CausesValidation=   0   'False
         Height          =   300
         Left            =   3960
         TabIndex        =   13
         Top             =   120
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   529
         Caption         =   "Data Marcada Devolução"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   16711680
      End
      Begin ReportX.ReportField ReportField4 
         Height          =   300
         Left            =   1560
         TabIndex        =   12
         Top             =   120
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   529
         Caption         =   "Data Empréstimo"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   16711680
      End
      Begin ReportX.ReportField ReportField10 
         Height          =   240
         Left            =   7320
         TabIndex        =   3
         Top             =   480
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   423
         Campo           =   "STATUS"
         Caption         =   "Status"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField ReportField3 
         Height          =   300
         Left            =   240
         TabIndex        =   4
         Top             =   120
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   529
         Caption         =   "Código"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   16711680
      End
      Begin ReportX.ReportField ReportField6 
         Height          =   300
         Left            =   9000
         TabIndex        =   5
         Top             =   120
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   529
         Caption         =   "Data de Devolução"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   16711680
      End
      Begin ReportX.ReportField ReportField7 
         Height          =   300
         Left            =   240
         TabIndex        =   6
         Top             =   840
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   529
         Caption         =   "Aluno"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   16711680
      End
      Begin ReportX.ReportField ReportField8 
         Height          =   300
         Left            =   240
         TabIndex        =   7
         Top             =   1560
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   529
         Caption         =   "Livro"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   16711680
      End
      Begin ReportX.ReportField ReportField9 
         Height          =   240
         Left            =   240
         TabIndex        =   8
         Top             =   480
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   423
         Campo           =   "ID_EMPRESTIMO"
         Caption         =   "Código"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField ReportField12 
         Height          =   240
         Left            =   9000
         TabIndex        =   9
         Top             =   480
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   423
         Campo           =   "DATA_DEVOLUCAO"
         Caption         =   "Data de Devolução"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField ReportField13 
         Height          =   240
         Left            =   240
         TabIndex        =   10
         Top             =   1200
         Width           =   5565
         _ExtentX        =   9816
         _ExtentY        =   423
         Campo           =   "NOME_ALUNO"
         Caption         =   "Aluno"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField ReportField14 
         Height          =   240
         Left            =   240
         TabIndex        =   11
         Top             =   1920
         Width           =   5565
         _ExtentX        =   9816
         _ExtentY        =   423
         Campo           =   "NOME_LIVRO"
         Caption         =   "Livro"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ReportX.ReportField ReportField17 
         Height          =   300
         Left            =   7320
         TabIndex        =   17
         Top             =   1560
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   529
         Caption         =   "Gênero"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   16711680
      End
      Begin ReportX.ReportField ReportField18 
         Height          =   240
         Left            =   7320
         TabIndex        =   18
         Top             =   1920
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   423
         Campo           =   "GENERO"
         Caption         =   "Gênero"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Image Image1 
         Height          =   1650
         Left            =   9720
         Picture         =   "RelEmprestimos.frx":0000
         Stretch         =   -1  'True
         Top             =   720
         Width           =   1650
      End
      Begin VB.Line Line1 
         BorderStyle     =   3  'Dot
         X1              =   120
         X2              =   11520
         Y1              =   2400
         Y2              =   2400
      End
   End
End
Attribute VB_Name = "RelEmprestimos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Rs As ADODB.Recordset

' Método para chamar o relatorio.
' Dessa forma todo o codigo para o funcionamento
' do relatorio pode ficar encapsulado no proprio formulario
Public Sub Config()
    Call conexao(cn)
        
    ' Uso do ADO nesse exemplo
    Set Rs = New ADODB.Recordset
    
    ' Abre o recordset com os dados
    'MsgBox SQL
     Rs.Open SQL, cn
    
    '  Associa o recordset ao relatorio
    Set Relatorio.Recordset = Rs
    ' Inicia a geração do relatório.
    Relatorio.Ativar
    
    ' Fecha o recordset
    Rs.Close
    Set Rs = Nothing
    
    ' Retira o formulário de relatorio da memória
    Unload Me
    
End Sub

Private Sub Relatorio_Erro(ByVal Numero As Long)
    
    ' O componente trabalha no modo silencioso para
    ' erros. Ele dispara esse evento Erro e sai. Caso
    ' o seu relatório esteja iniciando e saindo sem
    ' apresentar erro, verifique se você colocou algum
    ' código nesse evento.

    Rpx_MsgErro Numero
    
End Sub


' Sub para apresentar mensagens de erro para o Visual ReportX
' Utilize sempre uma rotina no evento Erro do componente.
Private Sub Rpx_MsgErro(Numero As Long)

    Dim Msg$
    
    If Numero < 0 Then
    
        ' Mensagens de erro previstas
        Select Case Numero - vbObjectError
            Case 1001: Msg = "É necessário existir uma impressora instalada no Windows"
            Case 1002: Msg = "Não há registros a imprimir"
            Case 1003: Msg = "Não foi definida a seção de detalhe do relatório"
            Case 1004: Msg = "A configuração das seções de grupos está incorreta"
            Case 1005: Msg = "Foi definido um cursor do tipo Forward-Only para o recordset do relatório."
            Case 1006: Msg = "A página configurada para o relatório não possuí espaço suficiente para a impressão"
            Case 1007: Msg = "Já existe um relatório em andamento"
        End Select
        
        MsgBox Msg, vbInformation, "Impressão"
        
    Else
        
        ' Mensagens não previstas. Isso pode significar um erro
        ' interno no ReportX. Se isso acontecer, por favor reporte isso
        ' através de e-mail para ser corrigido.
        MsgBox "Erro não previsto:" & Numero & vbCrLf & Error(Numero) & _
            IIf(Err.Number <> 0, vbCrLf + Err.Description, ""), vbCritical, "Impressão"
        
    End If
    
End Sub










