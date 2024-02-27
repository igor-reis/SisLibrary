VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MsComCtl.ocx"
Begin VB.Form FrmCarregamento 
   BackColor       =   &H000080FF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4695
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9735
   LinkTopic       =   "Form1"
   ScaleHeight     =   4695
   ScaleWidth      =   9735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4455
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9495
      Begin VB.Timer TimerCarregamento 
         Enabled         =   0   'False
         Interval        =   60
         Left            =   240
         Top             =   3840
      End
      Begin MSComctlLib.ProgressBar pbCarregamento 
         Height          =   375
         Left            =   840
         TabIndex        =   7
         Top             =   3840
         Width           =   7695
         _ExtentX        =   13573
         _ExtentY        =   661
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
      End
      Begin VB.Label lblCopyright 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Desenvolvido por Igor Reis"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FF80&
         Height          =   210
         Left            =   5400
         TabIndex        =   6
         Top             =   3000
         Width           =   2220
      End
      Begin VB.Label lblAlerta 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Aviso: Este software foi adaptado para fins educacionais e não pode ser vendido!"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   210
         Left            =   1200
         TabIndex        =   5
         Top             =   3360
         Width           =   6690
      End
      Begin VB.Label lblVersao 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Versão Ultimate"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C000&
         Height          =   285
         Left            =   6555
         TabIndex        =   4
         Top             =   2520
         Width           =   1800
      End
      Begin VB.Label lblNomeSoftaware 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SisLibrary"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   32.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   765
         Left            =   3240
         TabIndex        =   3
         Top             =   1560
         Width           =   3165
      End
      Begin VB.Label lblEmpresa 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Escola Estadual Dona Augusta"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   435
         Left            =   2040
         TabIndex        =   2
         Top             =   840
         Width           =   5310
      End
      Begin VB.Label lblLicenca 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Licenciado para"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   7560
         TabIndex        =   1
         Top             =   240
         Width           =   1155
      End
   End
End
Attribute VB_Name = "FrmCarregamento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim X As Integer

Private Sub Form_Load()
TimerCarregamento.Enabled = True
End Sub

Private Sub TimerCarregamento_Timer()
If pbCarregamento.Value <> 100 Then
    pbCarregamento.Value = pbCarregamento + 1
Else
   TimerCarregamento.Enabled = False
   Unload Me
   FrmPrincipal.Show 1
End If
End Sub


Private Sub TimerTempo_Timer()

End Sub
