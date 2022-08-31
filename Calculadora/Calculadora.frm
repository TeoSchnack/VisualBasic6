VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2580
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8610
   LinkTopic       =   "Form1"
   ScaleHeight     =   2580
   ScaleWidth      =   8610
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox C1 
      Height          =   615
      Left            =   360
      TabIndex        =   10
      Top             =   600
      Width           =   1935
   End
   Begin VB.TextBox C2 
      Height          =   615
      Left            =   3240
      TabIndex        =   9
      Top             =   600
      Width           =   1935
   End
   Begin VB.CommandButton LIMP 
      Caption         =   "LIMPAR"
      Height          =   735
      Index           =   6
      Left            =   3600
      TabIndex        =   6
      Top             =   1440
      Width           =   1695
   End
   Begin VB.CommandButton DIV 
      Caption         =   "/"
      Height          =   735
      Index           =   5
      Left            =   2760
      TabIndex        =   5
      Top             =   1440
      Width           =   615
   End
   Begin VB.CommandButton MULTI 
      Caption         =   "X"
      Height          =   735
      Index           =   4
      Left            =   1920
      TabIndex        =   4
      Top             =   1440
      Width           =   615
   End
   Begin VB.CommandButton SUB 
      Caption         =   "-"
      Height          =   735
      Index           =   3
      Left            =   1080
      TabIndex        =   3
      Top             =   1440
      Width           =   615
   End
   Begin VB.CommandButton SOMA 
      Caption         =   "+"
      Height          =   735
      Index           =   2
      Left            =   240
      TabIndex        =   2
      Top             =   1440
      Width           =   615
   End
   Begin VB.Frame Resultado 
      Caption         =   "Resultado"
      Height          =   975
      Left            =   6360
      TabIndex        =   1
      Top             =   360
      Width           =   1935
      Begin VB.Label RES 
         Height          =   495
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.CommandButton IGUAL 
      Caption         =   "="
      Height          =   735
      Index           =   1
      Left            =   5520
      TabIndex        =   0
      Top             =   480
      Width           =   615
   End
   Begin VB.Label Ope 
      Height          =   255
      Left            =   2640
      TabIndex        =   8
      Top             =   720
      Width           =   255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub DIV_Click(Index As Integer)

    Ope.Caption = "/"

End Sub

Private Sub IGUAL_Click(Index As Integer)

    Dim V1 As Single
    Dim V2 As Single
    
    V1 = C1.Text
    V2 = C2.Text
    
    If Ope.Caption = "+" Then
        V3 = V1 + V2
    ElseIf Ope.Caption = "-" Then
        V3 = V1 - V2
    ElseIf Ope.Caption = "X" Then
        V3 = V1 * V2
    ElseIf Ope.Caption = "/" Then
        V3 = V1 / V2
    End If
    
    RES.Caption = V3
    
End Sub

Private Sub LIMP_Click(Index As Integer)

    Ope.Caption = ""
    C1.Text = ""
    C2.Text = ""
    RES.Caption = ""

End Sub

Private Sub MULTI_Click(Index As Integer)

    Ope.Caption = "X"
    
End Sub

Private Sub SOMA_Click(Index As Integer)

    Ope.Caption = "+"

End Sub

Private Sub SUB_Click(Index As Integer)

    Ope.Caption = "-"

End Sub

