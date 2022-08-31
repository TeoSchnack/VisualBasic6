VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2700
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   2700
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame4 
      Height          =   2055
      Left            =   240
      TabIndex        =   10
      Top             =   240
      Visible         =   0   'False
      Width           =   3975
      Begin VB.CommandButton Command4 
         Caption         =   "OK"
         Height          =   615
         Left            =   1200
         TabIndex        =   11
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
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
         Left            =   240
         TabIndex        =   12
         Top             =   480
         Width           =   3375
      End
   End
   Begin VB.Frame Frame3 
      Height          =   2055
      Left            =   240
      TabIndex        =   7
      Top             =   240
      Width           =   3975
      Begin VB.CommandButton Command5 
         Caption         =   "CONFIRMA"
         Height          =   615
         Left            =   1200
         TabIndex        =   8
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Esse programa desetina-se para o calculo de áre de terrenos triangulares"
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
         Left            =   240
         TabIndex        =   9
         Top             =   480
         Width           =   3375
      End
   End
   Begin VB.Frame Frame2 
      Height          =   2055
      Left            =   240
      TabIndex        =   4
      Top             =   240
      Visible         =   0   'False
      Width           =   3975
      Begin VB.CommandButton Command2 
         Caption         =   "CONFIRMA"
         Height          =   615
         Left            =   1200
         TabIndex        =   5
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Deseja fazer um novo cálculo?"
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
         Left            =   840
         TabIndex        =   6
         Top             =   480
         Width           =   2175
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2055
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Visible         =   0   'False
      Width           =   3975
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   240
         TabIndex        =   14
         Top             =   1080
         Width           =   3495
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   240
         TabIndex        =   13
         Top             =   720
         Width           =   3495
      End
      Begin VB.CommandButton Command1 
         Caption         =   "CONFIRMA"
         Height          =   375
         Left            =   2280
         TabIndex        =   2
         Top             =   1440
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   3495
      End
      Begin VB.Label Label1 
         Caption         =   "Digite o Lado do Terreno"
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   120
         Width           =   2175
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

    Dim Area As Double
    Dim L(2) As Double
    
    L(0) = Text1.Text
    L(1) = Text2.Text
    L(2) = Text3.Text
    
    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
    
    
    
        
    Area = (L(0) * L(1) * L(2)) / 3
    Label4.Caption = CStr(L(0)) + " * " + CStr(L(1)) + " * " + CStr(L(2)) + " /3 = " + CStr(Area)
    Frame4.Visible = True
    Frame1.Visible = False

    
End Sub
Private Sub Command2_Click()

    Frame2.Visible = False
    Frame1.Visible = True

End Sub

Private Sub Command4_Click()
    
    Frame4.Visible = False
    Frame2.Visible = True
    
End Sub

Private Sub Command5_Click()

    Frame3.Visible = False
    Frame1.Visible = True
    
End Sub

