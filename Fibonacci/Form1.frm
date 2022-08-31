VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Fibonacci"
   ClientHeight    =   1170
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4785
   LinkTopic       =   "Form1"
   ScaleHeight     =   1170
   ScaleWidth      =   4785
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Executar"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   4215
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   1  'Horizontal
      TabIndex        =   0
      Text            =   "Form1.frx":0000
      Top             =   1200
      Visible         =   0   'False
      Width           =   4215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Gerador da Sequencia de Fibonacci"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   4215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

    Dim Na, Nb, Nc As Integer
    Na , Nb, Nc = 1
    
    Text1.Text = Nb
    
    For I = 0 To 10
        If I Mod 2 = 0 Then
            Nc = Na
            Na = Nb + Nc
            Text1.Text = Text1.Text + " --- " + CStr(Na)
        Else
            Nc = Nb
            Nb = Na + Nc
            Text1.Text = Text1.Text + " --- " + CStr(Nb)
        End If

    Next
    
    Form1.Height = 2730
    Text1.Visible = True

End Sub
