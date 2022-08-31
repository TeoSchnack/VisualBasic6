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
      Caption         =   "Iniciar"
      Height          =   735
      Left            =   600
      TabIndex        =   0
      Top             =   2040
      Width           =   2055
   End
   Begin VB.Label lblmensagem 
      Height          =   735
      Left            =   600
      TabIndex        =   1
      Top             =   960
      Width           =   3135
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

    lblmensagem.Caption = "Voce clicou em"

End Sub

Private Sub Label1_Click()

End Sub
