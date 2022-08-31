VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Fatoracao"
   ClientHeight    =   7305
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10500
   LinkTopic       =   "Form1"
   ScaleHeight     =   7305
   ScaleWidth      =   10500
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command11 
      Caption         =   "Geral1"
      Height          =   615
      Left            =   480
      TabIndex        =   2
      Top             =   840
      Width           =   9495
   End
   Begin VB.TextBox Text1 
      Height          =   4095
      Left            =   600
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   2040
      Width           =   9375
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Descubra os 10.001 primeiros numeros primos"
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
      Left            =   480
      TabIndex        =   1
      Top             =   240
      Width           =   9375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public I As Double
Private Sub Parte1()
    
    Dim Res As String
    Dim Cont As Double
    
    Cont = 0
    
    Do While Cont <= 1000
        
        Nt = 0
        For E = 2 To (I - 1)
            
            If InStr(1, CStr(I / E), ",", 0) <> 0 Then
                Nt = 1
            Else
                Nt = 0
                Exit For
            End If

         Next
        
        If (Nt > 0) Then
            Res = Res + CStr(I) + " --- "
            Cont = Cont + 1
        End If
            
            
        I = I + 1
    
    Loop
    
    Text1.Text = Text1.Text + Res
    
End Sub
Private Sub Command11_Click()

    Text1.Text = "2 --- "
    I = 3
    
        
    For P = 1 To 10
        Call Parte1
    Next
        
    
    Text1.Text = RTrim(Mid(Text1.Text, 1, (Len(Text1.Text) - 6)))

End Sub

