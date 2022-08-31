VERSION 5.00
Begin VB.Form Multiplos 
   Caption         =   "Multiplos"
   ClientHeight    =   1710
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3405
   LinkTopic       =   "Form1"
   ScaleHeight     =   1710
   ScaleWidth      =   3405
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox CC 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1320
      TabIndex        =   6
      Text            =   "100000"
      Top             =   1200
      Width           =   735
   End
   Begin VB.TextBox CB 
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   5
      Text            =   "5"
      Top             =   720
      Width           =   735
   End
   Begin VB.TextBox CA 
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   4
      Text            =   "3"
      Top             =   240
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   615
      HideSelection   =   0   'False
      Left            =   360
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   1  'Horizontal
      TabIndex        =   1
      Text            =   "Form1.frx":0000
      Top             =   1680
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Gerar"
      Height          =   1215
      Left            =   2280
      TabIndex        =   0
      Top             =   240
      Width           =   855
   End
   Begin VB.Label Label5 
      Caption         =   "Até"
      Height          =   255
      Left            =   360
      TabIndex        =   9
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "Multiplo de"
      Height          =   255
      Left            =   360
      TabIndex        =   8
      Top             =   720
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "Multiplo de"
      Height          =   255
      Left            =   360
      TabIndex        =   7
      Top             =   240
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Soma dos Multiplos ->"
      Height          =   495
      Left            =   360
      TabIndex        =   3
      Top             =   2400
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   495
      Left            =   2040
      TabIndex        =   2
      Top             =   2400
      Visible         =   0   'False
      Width           =   1095
   End
End
Attribute VB_Name = "Multiplos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public A, B As Integer
Public Soma, C As Double
Public Multiplo, Test As String
Sub Calcula()

    Multiplo = "0"

    For I = 1 To C
        
        If StrComp("0", Mid(CStr(I), Len(CStr(I))), 0) = 0 Then
            
            If I Mod A = 0 Then
                
                Multiplo = Multiplo + "----" + CStr(I)
                Soma = Soma + I
                
            End If
            
        ElseIf StrComp("5", Mid(CStr(I), Len(CStr(I))), 0) = 0 Then
            
            If I Mod A = 0 Then

                Multiplo = Multiplo + "----" + CStr(I)
                Soma = Soma + I
                
            End If
            
        End If
        
    Next
    
End Sub
Private Sub Command1_Click()

    A = CA.Text
    B = CB.Text
    C = CC.Text

    Call Calcula
    Text1.Text = Multiplo
    Label1.Caption = Soma
    
    Text1.Visible = True
    Label1.Visible = True
    Label2.Visible = True
    Multiplos.Height = 3855
    Multiplos.Width = 3645
    

End Sub

