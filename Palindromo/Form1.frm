VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1755
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5760
   LinkTopic       =   "Form1"
   ScaleHeight     =   1755
   ScaleWidth      =   5760
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Descobrir numero palindromo da multiplicação dos dois maiores numeros de 3 casas"
      Height          =   975
      Left            =   240
      TabIndex        =   3
      Top             =   360
      Width           =   5175
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   480
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   600
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   600
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   3840
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "="
      Height          =   255
      Left            =   3600
      TabIndex        =   5
      Top             =   720
      Width           =   135
   End
   Begin VB.Label Label1 
      Caption         =   "X"
      Height          =   255
      Left            =   1920
      TabIndex        =   4
      Top             =   720
      Width           =   135
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command2_Click()

    Dim Res, Inv, Aux, Aux1, Aux2 As String
    Dim N1, N2, T, A, B, Cont As Integer
    
    Command2.Visible = False
    A = 1
    
    For N1 = 999 To 1 Step -1
        For N2 = 999 To 1 Step -1
            Res = CStr(N1 * N2)
            Aux = Res
            For Cont = 0 To Len(Aux)
                
                Aux1 = Mid(Aux, A, A)
                Aux2 = Mid(Aux, Len(Aux), Len(Aux))
                
                If StrComp(Aux1, Aux2, 1) = 0 Then
                    Aux = Mid(Aux, 2, (Len(Aux) - 2))
                    
                    If Len(Aux) = 1 Then
                        Aux = ""
                        T = 1
                        Exit For
                    ElseIf StrComp(Aux, "", 1) = 0 Then
                        T = 1
                        Exit For
                    End If
                Else
                    Exit For
                End If
                
            Next
            
            If T = 1 Then
                Exit For
            End If
            
        Next
       
        If T = 1 Then
            Exit For
        End If
    Next

    Text1.Text = Res
    Text2.Text = CStr(N2)
    Text3.Text = CStr(N1)

End Sub
