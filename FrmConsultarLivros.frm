VERSION 5.00
Object = "{CFAB6834-3B57-49FC-8770-CBA3667FE193}#1.0#0"; "ctextbox.ocx"
Object = "{61159A24-3E03-4E76-9CA9-2396C6822B8F}#1.0#0"; "chamaleonbtn.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmConsultarLivros 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Consulta de Livros"
   ClientHeight    =   5610
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10470
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4818.158
   ScaleMode       =   0  'User
   ScaleWidth      =   10559.62
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin ChamaleonBtn.chameleonButton cmdConsultarS 
      Height          =   840
      Left            =   2400
      TabIndex        =   3
      Top             =   4560
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   1482
      BTYPE           =   3
      TX              =   "&Consultar Status"
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
      MICON           =   "FrmConsultarLivros.frx":0000
      PICN            =   "FrmConsultarLivros.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame Frame1 
      Caption         =   "Ações"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   120
      TabIndex        =   8
      Top             =   3480
      Width           =   10215
      Begin VB.OptionButton optIndisponivel 
         Caption         =   "Indisponível"
         Height          =   375
         Left            =   960
         TabIndex        =   2
         Top             =   1560
         Width           =   1215
      End
      Begin VB.OptionButton optDisponivel 
         Caption         =   "Disponível"
         Height          =   375
         Left            =   960
         TabIndex        =   1
         Top             =   1080
         Width           =   1335
      End
      Begin cTextBox.cText txtConsulta 
         Height          =   375
         Left            =   240
         TabIndex        =   0
         Top             =   360
         Width           =   9855
         _ExtentX        =   17383
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
         DragIcon        =   "FrmConsultarLivros.frx":3164
         Height          =   825
         Left            =   7320
         TabIndex        =   6
         Top             =   1080
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   1455
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
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   16777215
         FCOLO           =   16777215
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmConsultarLivros.frx":3E2E
         PICN            =   "FrmConsultarLivros.frx":3E4A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdConsultarN 
         DragIcon        =   "FrmConsultarLivros.frx":7477
         Height          =   825
         Left            =   3960
         TabIndex        =   4
         Top             =   1080
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   1455
         BTYPE           =   3
         TX              =   "C&onsultar Nome"
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
         FCOL            =   16777215
         FCOLO           =   16777215
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmConsultarLivros.frx":8141
         PICN            =   "FrmConsultarLivros.frx":815D
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonBtn.chameleonButton cmdLimpar 
         DragIcon        =   "FrmConsultarLivros.frx":B2A5
         Height          =   825
         Left            =   5640
         TabIndex        =   5
         Top             =   1080
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   1455
         BTYPE           =   3
         TX              =   "&Limpar"
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
         FCOL            =   16777215
         FCOLO           =   16777215
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmConsultarLivros.frx":BF6F
         PICN            =   "FrmConsultarLivros.frx":BF8B
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grid 
      Height          =   3450
      Left            =   105
      TabIndex        =   7
      Top             =   0
      Width           =   10260
      _ExtentX        =   18098
      _ExtentY        =   6085
      _Version        =   393216
      Cols            =   9
      FixedCols       =   0
   End
End
Attribute VB_Name = "FrmConsultarLivros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer
Dim opt As String

Function CarregaGrid()

Call conexao(cn)

Rs.Open SQL, cn

If Rs.RecordCount > "0" Then
    
    grid.TextMatrix(0, 0) = "ID Livro"
    grid.TextMatrix(0, 1) = "Nome"
    grid.TextMatrix(0, 2) = "Autor"
    grid.TextMatrix(0, 3) = "Ano"
    grid.TextMatrix(0, 4) = "Número de Páginas"
    grid.TextMatrix(0, 5) = "Genêro"
    grid.TextMatrix(0, 6) = "Editora"
    grid.TextMatrix(0, 7) = "Status"
    grid.TextMatrix(0, 8) = "Data de Cadastro"
    
    grid.Rows = Rs.RecordCount + 1
    Rs.MoveFirst

    For i = 1 To Rs.RecordCount
        grid.TextMatrix(i, 0) = Rs!ID_LIVRO
        grid.TextMatrix(i, 1) = Rs!NOME_LIVRO
        grid.TextMatrix(i, 2) = Rs!AUTOR
        grid.TextMatrix(i, 3) = Rs!ANO
        grid.TextMatrix(i, 4) = Rs!NUMERO_PAGINAS
        grid.TextMatrix(i, 5) = Rs!GENERO
        grid.TextMatrix(i, 6) = Rs!EDITORA
        grid.TextMatrix(i, 7) = Rs!STATUS_LIVRO
        grid.TextMatrix(i, 8) = Rs!DATA_CADASTRO
        Rs.MoveNext
    Next i
    
    Call AjustaGrid
    Rs.Close
    
Else

    MsgBox "Nenhum registro foi encontrado!", vbExclamation, "Mensagem do Sistema"
    
    optDisponivel.Value = False And optIndisponivel.Value = False
    txtConsulta.Text = ""
    
    Rs.Close
    
End If

'Ordenar em ordem crescente
grid.Col = 0
grid.Sort = flexSortGenericAscending
End Function

Private Sub cmdConsultarN_Click()
If txtConsulta.Text = "" Then
   MsgBox "Por favor, digite um nome no campo para continuar!", vbInformation, "Mensagem do Sistema"
   txtConsulta.SetFocus
Else
    SQL = "SELECT * FROM CAD_LIVRO WHERE NOME_LIVRO like UPPER('" & UCase(txtConsulta.Text) & "%')"
    
    Call CarregaGrid
End If
End Sub

Private Sub cmdConsultarS_Click()
If optDisponivel.Value = False And optIndisponivel.Value = False Then
    
    MsgBox "Por favor, selecione uma opção ao lado para continuar!", vbInformation, "Mensagem do Sistema"

Else
    
    If optDisponivel.Value = True Then
    
        opt = "DISPONIVEL"
    
    Else
    
        opt = "INDISPONIVEL"
    
    End If
    
    SQL = "SELECT * FROM CAD_LIVRO WHERE STATUS_LIVRO like UPPER('" & opt & "%')"
    
    Call CarregaGrid
    
End If

End Sub

Private Sub cmdLimpar_Click()
Unload Me
FrmConsultarLivros.Show 1
End Sub


Private Sub cmdVoltar_Click()
Unload Me
End Sub

Private Sub Form_Load()
    SQL = "SELECT * FROM CAD_LIVRO"
    Call CarregaGrid
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
Call conexao(cn)
     grid.Col = 0
     Rs.Open "SELECT * FROM CAD_LIVRO WHERE ID_LIVRO = " & CInt(grid.Text), cn
     
If NomeFormulario = "Empréstimos" Then
    
    If IsNull(Rs!ID_LIVRO) = False Then FrmEmprestimos.txtIdLivro.Text = Rs!ID_LIVRO
    If IsNull(Rs!NOME_LIVRO) = False Then FrmEmprestimos.txtNomeLivro.Text = Rs!NOME_LIVRO

End If

If NomeFormulario = "Livros" Then
     
     If IsNull(Rs!ID_LIVRO) = False Then FrmLivros.txtIdLivro.Text = Rs!ID_LIVRO
     If IsNull(Rs!NOME_LIVRO) = False Then FrmLivros.txtNome.Text = Rs!NOME_LIVRO
     If IsNull(Rs!AUTOR) = False Then FrmLivros.txtAutor.Text = Rs!AUTOR
     If IsNull(Rs!ANO) = False Then FrmLivros.txtAno.Text = Rs!ANO
     If IsNull(Rs!NUMERO_PAGINAS) = False Then FrmLivros.txtNPaginas.Text = Rs!NUMERO_PAGINAS
     If IsNull(Rs!GENERO) = False Then FrmLivros.txtGenero.Text = Rs!GENERO
     If IsNull(Rs!EDITORA) = False Then FrmLivros.txtEditora.Text = Rs!EDITORA
     If IsNull(Rs!STATUS_LIVRO) = False Then FrmLivros.cmbStatus.Text = Rs!STATUS_LIVRO
     If IsNull(Rs!DATA_CADASTRO) = False Then FrmLivros.txtDataCadastro.Text = Rs!DATA_CADASTRO
     
End If
     
     Rs.Close
     
     Unload Me
     
     FrmLivros.cmdCancelar.Enabled = True
     FrmLivros.cmdAlterar.Enabled = True
     FrmLivros.cmdExcluir.Enabled = True
     FrmLivros.cmdNovo.Enabled = False
     
End Sub


