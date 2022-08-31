VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()

    Dim SW As Integer, SH As Integer
    SW = ScaleWidth
    SH = ScaleHeight
    Line (SW / 2, 50)-(50, 2 * SH / 3)
    Line Step(0, 0)-(SW / 2, SH / 2), RGB(255, 0, 0)
    Line Step(0, 0)-Step(SW / 2 - 50, SH / 6), RGB(0, 0, 255)
    Line (SW - 50, 2 * SH / 3)-(SW / 2, 50), RGB(226, 0, 127)

End Sub

