VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3210
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6825
   LinkTopic       =   "Form1"
   ScaleHeight     =   3210
   ScaleWidth      =   6825
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Height          =   2655
      Left            =   3600
      TabIndex        =   7
      Top             =   240
      Width           =   2895
      Begin VB.CommandButton REPETE 
         Caption         =   "REPETE"
         Height          =   375
         Left            =   1440
         TabIndex        =   8
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Label N1 
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   600
         Width           =   495
      End
      Begin VB.Label N4 
         Height          =   255
         Left            =   2040
         TabIndex        =   11
         Top             =   1680
         Width           =   495
      End
      Begin VB.Label N3 
         Height          =   255
         Left            =   1440
         TabIndex        =   10
         Top             =   1320
         Width           =   375
      End
      Begin VB.Label N2 
         Height          =   255
         Left            =   840
         TabIndex        =   9
         Top             =   960
         Width           =   375
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2655
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2895
      Begin VB.CommandButton Ordenar 
         Caption         =   "Ordenar"
         Height          =   375
         Left            =   1440
         TabIndex        =   5
         Top             =   2160
         Width           =   1215
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   120
         TabIndex        =   4
         Text            =   "4"
         Top             =   1680
         Width           =   2535
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   120
         TabIndex        =   3
         Text            =   "5"
         Top             =   1320
         Width           =   2535
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Text            =   "2"
         Top             =   960
         Width           =   2535
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Text            =   "6"
         Top             =   600
         Width           =   2535
      End
      Begin VB.Label Label1 
         Caption         =   "Digite os numeros a serem ordenados"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   2655
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Ordenar_Click()
    
    Dim N(4) As Integer
    N(0) = CInt(Text1.Text)
    N(1) = CInt(Text2.Text)
    N(2) = CInt(Text3.Text)
    N(3) = CInt(Text4.Text)
    
    Dim Aux As Integer
    Aux = N(1)
    Cont = 1
    
    Do While Cont = 1
        
        Cont = 2
        
        If N(0) > N(1) Then
            Aux = N(1)
            N(1) = N(0)
            N(0) = Aux
            Cont = 1
        End If
        If N(1) > N(2) Then
            Aux = N(2)
            N(2) = N(1)
            N(1) = Aux
            Cont = 1
        End If
        If N(2) > N(3) Then
            Aux = N(3)
            N(3) = N(2)
            N(2) = Aux
            Cont = 1
        End If
        
    Loop
    

    
    N1.Caption = N(0)
    N2.Caption = N(1)
    N3.Caption = N(2)
    N4.Caption = N(3)
    
    Frame2.Visible = True
    Frame1.Visible = False
    
End Sub
Private Sub REPETE_Click()

    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
    Text4.Text = ""
    
    Frame1.Visible = True
    Frame2.Visible = False
End Sub

