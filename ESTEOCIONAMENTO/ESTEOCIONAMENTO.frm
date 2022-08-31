VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   5985
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5760
   LinkTopic       =   "Form1"
   ScaleHeight     =   5985
   ScaleMode       =   0  'User
   ScaleWidth      =   2938.775
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Alterador 
      Caption         =   "Alterador"
      Height          =   2415
      Left            =   6360
      TabIndex        =   24
      Top             =   3240
      Visible         =   0   'False
      Width           =   4695
      Begin VB.CommandButton ConfirmaPreco 
         Caption         =   "Confirma"
         Height          =   615
         Left            =   2160
         TabIndex        =   32
         Top             =   1680
         Width           =   2295
      End
      Begin VB.CheckBox CheckCarro 
         Caption         =   "Check2"
         Height          =   255
         Left            =   960
         TabIndex        =   31
         Top             =   2040
         Width           =   255
      End
      Begin VB.CheckBox CheckMoto 
         Caption         =   "Check1"
         Height          =   195
         Left            =   960
         TabIndex        =   30
         Top             =   1680
         Width           =   255
      End
      Begin VB.TextBox NOVOVALOR 
         Height          =   405
         Left            =   840
         TabIndex        =   26
         Top             =   960
         Width           =   3495
      End
      Begin VB.Label Carro 
         Caption         =   "Carro"
         Height          =   255
         Left            =   1320
         TabIndex        =   29
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label Moto 
         Caption         =   "Moto"
         Height          =   255
         Left            =   1320
         TabIndex        =   28
         Top             =   1680
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "R$"
         Height          =   255
         Left            =   360
         TabIndex        =   27
         Top             =   1080
         Width           =   375
      End
      Begin VB.Label NValor 
         Caption         =   "Novo Valor"
         Height          =   375
         Left            =   360
         TabIndex        =   25
         Top             =   600
         Width           =   975
      End
   End
   Begin VB.TextBox Senha 
      BeginProperty Font 
         Name            =   "MS Outlook"
         Size            =   8.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6600
      TabIndex        =   20
      Top             =   1680
      Visible         =   0   'False
      Width           =   4215
   End
   Begin VB.Frame Login 
      Caption         =   "Login"
      Height          =   2055
      Left            =   6360
      TabIndex        =   18
      Top             =   480
      Visible         =   0   'False
      Width           =   4695
      Begin VB.CommandButton Confirma 
         Caption         =   "Confirma"
         Height          =   255
         Left            =   3240
         TabIndex        =   21
         Top             =   1680
         Width           =   1215
      End
      Begin VB.TextBox User 
         Height          =   285
         Left            =   240
         TabIndex        =   19
         Top             =   600
         Width           =   4215
      End
      Begin VB.Label Label4 
         Caption         =   "Senha"
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Usuário"
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.CommandButton Menu 
      Caption         =   "Menu Administrativo"
      Height          =   495
      Left            =   360
      TabIndex        =   17
      Top             =   5160
      Width           =   5000
   End
   Begin VB.TextBox CARRROSDIA 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000001&
      Height          =   285
      Left            =   3120
      TabIndex        =   12
      Text            =   "Carros   - >"
      Top             =   3840
      Width           =   1095
   End
   Begin VB.TextBox MotoMomento 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000001&
      Height          =   285
      Left            =   480
      TabIndex        =   8
      Text            =   "Motos   - >"
      Top             =   4560
      Width           =   1095
   End
   Begin VB.TextBox CarrosMomento 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000001&
      Height          =   285
      Left            =   480
      TabIndex        =   7
      Text            =   "Carros   - >"
      Top             =   3840
      Width           =   855
   End
   Begin VB.TextBox Lucro 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000001&
      Height          =   285
      Left            =   3120
      TabIndex        =   6
      Text            =   "LUCRO   - >"
      Top             =   4680
      Width           =   1095
   End
   Begin VB.TextBox MotoDia 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000001&
      Height          =   285
      Left            =   3120
      TabIndex        =   5
      Text            =   "Motos   - >"
      Top             =   4200
      Width           =   1095
   End
   Begin VB.CommandButton SMoto 
      Caption         =   "Saída de Moto"
      Height          =   615
      Left            =   3000
      TabIndex        =   4
      Top             =   2520
      Width           =   2175
   End
   Begin VB.CommandButton SCarro 
      Caption         =   "Saída de Carro"
      Height          =   615
      Left            =   360
      TabIndex        =   3
      Top             =   2520
      Width           =   2175
   End
   Begin VB.CommandButton EMoto 
      Caption         =   "Entrada de Moto"
      Height          =   615
      Left            =   3000
      TabIndex        =   2
      Top             =   1680
      Width           =   2175
   End
   Begin VB.CommandButton ECarro 
      Caption         =   "Entrada de Carro"
      Height          =   615
      Left            =   360
      TabIndex        =   1
      Top             =   1680
      Width           =   2175
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      Caption         =   "Label9"
      Height          =   255
      Left            =   4680
      TabIndex        =   36
      Top             =   1320
      Width           =   495
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      Caption         =   "Label8"
      Height          =   255
      Left            =   2040
      TabIndex        =   35
      Top             =   1320
      Width           =   495
   End
   Begin VB.Label Label7 
      Caption         =   "PRECO MOTO R$"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3000
      TabIndex        =   34
      Top             =   1320
      Width           =   1815
   End
   Begin VB.Label Label6 
      Caption         =   "PRECO CARRO R$"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   33
      Top             =   1320
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "DIA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3120
      TabIndex        =   16
      Top             =   3480
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "ATUAL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   15
      Top             =   3480
      Width           =   615
   End
   Begin VB.Label CTotaisN 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      Height          =   255
      Left            =   4440
      TabIndex        =   14
      Top             =   3840
      Width           =   855
   End
   Begin VB.Label MTotaisN 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      Height          =   255
      Left            =   4440
      TabIndex        =   13
      Top             =   4200
      Width           =   855
   End
   Begin VB.Label LucroN 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      Height          =   255
      Left            =   4440
      TabIndex        =   11
      Top             =   4680
      Width           =   855
   End
   Begin VB.Label CAtuaisN 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      Height          =   255
      Left            =   1560
      TabIndex        =   10
      Top             =   3840
      Width           =   855
   End
   Begin VB.Label MAtuaisN 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      Height          =   255
      Left            =   1560
      TabIndex        =   9
      Top             =   4560
      Width           =   855
   End
   Begin VB.Label Titulo 
      Alignment       =   2  'Center
      Caption         =   "ESTEOCINAMENTO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   22.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   4920
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public CarrosA, CarrosT, CarrosM, MotosA, MotosT, MotosM As Integer
Public CarrosP, MotosP, Lucros As Double

Private Sub Confirma_Click()

    If User.Text = "Teo" And Senha.Text = "Teo" Then
        Alterador.Visible = True
    Else
        Login.Visible = False
        Senha.Visible = False
        User.Text = ""
        Senha.Text = ""
        Form1.Width = 6000
    End If
            
End Sub

Private Sub ConfirmaPreco_Click()

    Dim Novo As Double
    
   Novo = CDbl(NOVOVALOR.Text)
    
        If Novo > 0 Then
            If CheckCarro = True And CheckMoto = True Then
                Login.Visible = False
                Senha.Visible = False
                User.Text = ""
                Senha.Text = ""
                Alterador.Visible = False
                Form1.Width = 6000
                ERRO.Visible = True
            ElseIf CheckCarro = True Then
                CarrosP = Novo
                Login.Visible = False
                Senha.Visible = False
                User.Text = ""
                Senha.Text = ""
                Alterador.Visible = False
                Form1.Width = 6000
                SUCESSO.Visible = True
                Label8.Caption = CarrosP
            Else
                MotosP = Novo
                Login.Visible = False
                Senha.Visible = False
                User.Text = ""
                Senha.Text = ""
                Alterador.Visible = False
                Form1.Width = 6000
                SUCESSO.Visible = True
                Label9.Caption = MotosP
            End If
        Else
            ERRO.Visible = True
            Login.Visible = False
            Senha.Visible = False
            User.Text = ""
            Senha.Text = ""
            Alterador.Visible = False
            Form1.Width = 6000
        End If


End Sub

Public Sub ECarro_Click()
    
    If CarrosA < CarrosM Then
        CarrosA = CarrosA + 1
        CarrosT = CarrosT + 1
        Lucros = Lucros + CarrosP
        CTotaisN.Caption = CarrosT
        CAtuaisN.Caption = CarrosA
        LucroN.Caption = Lucros
    Else
        Carro.Visible = True
    End If
    
        
End Sub
Public Sub EMoto_Click()
    
    If MotosA < MotosM Then
        MotosA = MotosA + 1
        MotosT = MotosT + 1
        Lucros = Lucros + MotosP
        MTotaisN.Caption = MotosT
        MAtuaisN.Caption = MotosA
        LucroN.Caption = Lucros
    Else
        Moto.Visible = True
    End If
        
End Sub
Public Sub Form_Load()

    CarrosA = 0
    CarrosT = 0
    MotosA = 0
    CarrosM = 5
    MotosM = 10
    CarrosP = 15.5
    MotosP = 5.5
    Lucros = 0
    Label8.Caption = CarrosP
    Label9.Caption = MotosP
    
End Sub

Private Sub Menu_Click()

    Form1.Width = 12000
    Login.Visible = True
    Senha.Visible = True
    
End Sub

Public Sub SCarro_Click()
    
    If CarrosA > 0 Then
    
        CarrosA = CarrosA - 1
        CAtuaisN.Caption = CarrosA
    
    End If
    
End Sub
Public Sub SMoto_Click()

    If MotosA > 0 Then
        MotosA = MotosA - 1
        MAtuaisN.Caption = MotosA
    
    End If
   
End Sub

