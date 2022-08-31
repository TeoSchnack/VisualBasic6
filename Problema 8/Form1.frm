VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5610
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   ScaleHeight     =   5610
   ScaleWidth      =   6030
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   4920
      Width           =   5535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Executar"
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   3960
      Width           =   5535
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   3495
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "Form1.frx":0000
      Top             =   120
      Width           =   5415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim No, Nb, SNm, SNt As String
    Dim N(15), Nr, Nf As Double
    Dim T As Boolean
    
    No = Text1.Text
    T = True
    
    Do While T
        Nb = Mid(No, 1, 15)
        If StrComp(No, "", 1) Then
            Nr = 1
            Do While StrComp(Nb, "", 1)
            
                    N(E) = CDbl(Mid(Nb, 1, 1))
                    SNt = SNt + CStr(N(E)) + " X "
                    Nr = Nr * N(E)
                    Nb = Mid(Nb, 2, Len(Nb))
                    Nb = Trim(Nb)

            Loop
            
            SNt = Mid(SNt, 1, (Len(SNt) - 3)) + " = " + CStr(Nr)
            
             If (Nr >= Nf) Then
                Nf = Nr
                SNm = SNt
            End If
            
            SNt = ""
            No = Mid(No, 2, Len(No))
            
            If (Len(No) = 16) Then
                Exit Do
            End If
            
        Else
            Exit Do
        End If
    Loop
    
    Text2.Text = SNm
    
    
End Sub

