VERSION 5.00
Object = "{CFAB6834-3B57-49FC-8770-CBA3667FE193}#1.0#0"; "ctextbox.ocx"
Object = "{61159A24-3E03-4E76-9CA9-2396C6822B8F}#1.0#0"; "chamaleonbtn.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmConsultarEmprestimos 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Consulta de Empréstimos"
   ClientHeight    =   5685
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10485
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5685
   ScaleWidth      =   10485
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Height          =   2175
      Left            =   120
      TabIndex        =   5
      Top             =   3480
      Width           =   10215
      Begin VB.OptionButton optEmprestimoA 
         Caption         =   "Empréstimo atrasado"
         Height          =   375
         Left            =   1320
         TabIndex        =   8
         Top             =   1680
         Width           =   1815
      End
      Begin VB.OptionButton optNomeL 
         Caption         =   "Nome Livro"
         Height          =   495
         Left            =   1320
         TabIndex        =   7
         Top             =   1200
         Width           =   1815
      End
      Begin VB.OptionButton optNomeA 
         Caption         =   "Nome Aluno"
         Height          =   375
         Left            =   1320
         TabIndex        =   6
         Top             =   840
         Width           =   1815
      End
      Begin ChamaleonBtn.chameleonButton cmdConsultar 
         Height          =   825
         Left            =   3360
         TabIndex        =   1
         Top             =   1080
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   1455
         BTYPE           =   3
         TX              =   "&Consultar"
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
         MICON           =   "FrmConsultarEmprestimos.frx":0000
         PICN            =   "FrmConsultarEmprestimos.frx":001C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
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
      Begin ChamaleonBtn.chameleonButton cmdVoltar 
         DragIcon        =   "FrmConsultarEmprestimos.frx":3164
         Height          =   825
         Left            =   6720
         TabIndex        =   3
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
         MICON           =   "FrmConsultarEmprestimos.frx":3E2E
         PICN            =   "FrmConsultarEmprestimos.frx":3E4A
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
         DragIcon        =   "FrmConsultarEmprestimos.frx":7477
         Height          =   825
         Left            =   5040
         TabIndex        =   2
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
         MICON           =   "FrmConsultarEmprestimos.frx":8141
         PICN            =   "FrmConsultarEmprestimos.frx":815D
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
      Height          =   3495
      Left            =   120
      TabIndex        =   4
      Top             =   0
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   6165
      _Version        =   393216
      Cols            =   9
      FixedCols       =   0
   End
End
Attribute VB_Name = "FrmConsultarEmprestimos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer

Function CarregaGrid()

Call conexao(cn)

Rs.Open SQL, cn

If Rs.RecordCount > "0" Then
    
    grid.TextMatrix(0, 0) = "ID Empréstimo"
    grid.TextMatrix(0, 1) = "ID Aluno"
    grid.TextMatrix(0, 2) = "Nome Aluno"
    grid.TextMatrix(0, 3) = "ID Livro"
    grid.TextMatrix(0, 4) = "Nome Livro"
    grid.TextMatrix(0, 5) = "Data Empréstimo"
    grid.TextMatrix(0, 6) = "Data Marcada Devolução"
    grid.TextMatrix(0, 7) = "Status"
    grid.TextMatrix(0, 8) = "Data Devolução"
        
    grid.Rows = Rs.RecordCount + 1
    Rs.MoveFirst

    For i = 1 To Rs.RecordCount
        grid.TextMatrix(i, 0) = Rs!ID_EMPRESTIMO
        grid.TextMatrix(i, 1) = Rs!ID_ALUNO
        grid.TextMatrix(i, 2) = Rs!NOME_ALUNO
        grid.TextMatrix(i, 3) = Rs!ID_LIVRO
        grid.TextMatrix(i, 4) = Rs!NOME_LIVRO
        grid.TextMatrix(i, 5) = Rs!DATA_EMPRESTIMO
        grid.TextMatrix(i, 6) = Rs!DATA_M_DEVOLUCAO
        
    
        If Rs!status = "DEVOLVIDO" Then
           
            grid.TextMatrix(i, 7) = Rs!status
           
        Else
        
            If ((Date) > (Rs!DATA_M_DEVOLUCAO)) Then
            
                grid.TextMatrix(i, 7) = "ATRASADO"
                
                Set OP = New ADODB.Command
                With OP
                  .ActiveConnection = cn
                  .CommandText = "UPDATE EMPRESTIMOS SET STATUS = 'ATRASADO' WHERE ID_EMPRESTIMO = " & CInt(grid.TextMatrix(i, 0))
                  .Execute
                End With
            
            Else
            
                grid.TextMatrix(i, 7) = "PENDENTE"
                
                Set OP = New ADODB.Command
                With OP
                  .ActiveConnection = cn
                  .CommandText = "UPDATE EMPRESTIMOS SET STATUS = 'PENDENTE' WHERE ID_EMPRESTIMO = " & CInt(grid.TextMatrix(i, 0))
                  .Execute
                End With
                
            End If
        
        End If
        
        If grid.TextMatrix(i, 7) = "DEVOLVIDO" = True Then grid.TextMatrix(i, 8) = Rs!DATA_DEVOLUCAO
        Rs.MoveNext
        
    Next i
    Call AjustaGrid
    Rs.Close
    
Else

    MsgBox "Nenhum registro foi encontrado!", vbExclamation, "Mensagem do Sistema"
    
    optNomeA.Value = False And optNomeL.Value = False And optEmprestimoA.Value = False
    txtConsulta.Text = ""
    
    Rs.Close
    
End If

'Ordenar em ordem crescente
grid.Col = 0
grid.Sort = flexSortGenericAscending
End Function


Private Sub cmdConsultarNL_Click()
If txtConsulta.Text = "" Then
   MsgBox "Por favor, digite um nome no campo para continuar!", vbInformation, "Mensagem do Sistema"
   txtConsulta.SetFocus
Else
    SQL = "SELECT * FROM EMPRESTIMOS WHERE NOME_LIVRO like ('" & UCase(txtConsulta.Text) & "%')"
    
    Call CarregaGrid
End If
End Sub

Private Sub cmdConsultar_Click()
If optNomeA.Value = False And optNomeL.Value = False And optEmprestimoA.Value = False Then
   
   MsgBox "Por favor, selecione uma opção ao lado para continuar!", vbInformation, "Mensagem do Sistema"
   
Else

    SQL = "SELECT A.* ,B.NOME_ALUNO AS NOME_ALUNO, C.NOME_LIVRO AS NOME_LIVRO FROM ((EMPRESTIMOS A INNER JOIN CAD_ALUNO B ON A.ID_ALUNO=B.ID_ALUNO) INNER JOIN CAD_LIVRO C ON A.ID_LIVRO=C.ID_LIVRO)"
    
    If optNomeA.Value = True Then
       
        If txtConsulta.Text = "" Then
       
            MsgBox "Por favor, digite um nome no campo para continuar!", vbInformation, "Mensagem do Sistema"
            txtConsulta.SetFocus
            
        Else
       
            SQL = SQL + "WHERE B.NOME_ALUNO like UPPER ('" & (txtConsulta.Text) & "%')"
            
        End If
       
    End If
    
    If optNomeL.Value = True Then
    
        If txtConsulta.Text = "" Then
       
            MsgBox "Por favor, digite algo para continuar!", vbInformation, "Mensagem do Sistema"
            txtConsulta.SetFocus
            
        Else
        
            SQL = SQL + "WHERE C.NOME_LIVRO like UPPER ('" & (txtConsulta.Text) & "%')"
            
        End If
            
     End If
    
    If optEmprestimoA.Value = True Then
              SQL = SQL + "WHERE STATUS like 'ATRASADO'"
    End If

Call CarregaGrid

End If
End Sub

Private Sub cmdLimpar_Click()
Unload Me
FrmConsultarEmprestimos.Show 1
End Sub

Private Sub cmdVoltar_Click()
Unload Me
End Sub

Private Sub Form_Load()
    SQL = "SELECT A.* ,B.NOME_ALUNO AS NOME_ALUNO, C.NOME_LIVRO AS NOME_LIVRO FROM ((EMPRESTIMOS A INNER JOIN CAD_ALUNO B ON A.ID_ALUNO=B.ID_ALUNO) INNER JOIN CAD_LIVRO C ON A.ID_LIVRO=C.ID_LIVRO)"
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
     Rs.Open "SELECT A.* ,B.NOME_ALUNO AS NOME_ALUNO, C.NOME_LIVRO AS NOME_LIVRO FROM ((EMPRESTIMOS A INNER JOIN CAD_ALUNO B ON A.ID_ALUNO=B.ID_ALUNO) INNER JOIN CAD_LIVRO C ON A.ID_LIVRO=C.ID_LIVRO) WHERE ID_EMPRESTIMO = " & CInt(grid.Text), cn
     
     If IsNull(Rs!ID_EMPRESTIMO) = False Then FrmEmprestimos.txtIdEmprestimo.Text = Rs!ID_EMPRESTIMO
     If IsNull(Rs!ID_ALUNO) = False Then FrmEmprestimos.txtIdAluno.Text = Rs!ID_ALUNO
     If IsNull(Rs!NOME_ALUNO) = False Then FrmEmprestimos.txtNomeAluno.Text = Rs!NOME_ALUNO
     If IsNull(Rs!ID_LIVRO) = False Then FrmEmprestimos.txtIdLivro.Text = Rs!ID_LIVRO
     If IsNull(Rs!NOME_LIVRO) = False Then FrmEmprestimos.txtNomeLivro.Text = Rs!NOME_LIVRO
     If IsNull(Rs!DATA_EMPRESTIMO) = False Then FrmEmprestimos.txtDataEmprestimo.Text = Rs!DATA_EMPRESTIMO
     If IsNull(Rs!DATA_M_DEVOLUCAO) = False Then FrmEmprestimos.txtDataMDevolucao.Text = Rs!DATA_M_DEVOLUCAO
     If Rs!status = "DEVOLVIDO" Then FrmEmprestimos.chkDevolvido.Value = Checked
                  
     Rs.Close
     
     Unload Me
     
     FrmEmprestimos.lblStatus.Visible = True
     FrmEmprestimos.chkDevolvido.Visible = True
     FrmEmprestimos.cmdCancelar.Enabled = True
     FrmEmprestimos.cmdProcurarAluno.Enabled = True
     FrmEmprestimos.cmdProcurarLivro.Enabled = True
     FrmEmprestimos.cmdAlterar.Enabled = True
     FrmEmprestimos.cmdExcluir.Enabled = True
     FrmEmprestimos.cmdNovo.Enabled = False
          
End Sub




