VERSION 5.00
Begin VB.Form NO 
   Caption         =   "Form1"
   ClientHeight    =   6030
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7515
   LinkTopic       =   "Form1"
   ScaleHeight     =   6030
   ScaleWidth      =   7515
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton NO 
      Caption         =   "NO"
      Height          =   735
      Left            =   3960
      TabIndex        =   7
      Top             =   3600
      Width           =   2055
   End
   Begin VB.CommandButton NO3 
      Caption         =   "NO"
      Height          =   735
      Left            =   720
      TabIndex        =   6
      Top             =   480
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton NO4 
      Caption         =   "NO"
      Height          =   735
      Left            =   5040
      TabIndex        =   5
      Top             =   5040
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton NO2 
      Caption         =   "NO"
      Height          =   735
      Left            =   5160
      TabIndex        =   4
      Top             =   360
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton NO1 
      Caption         =   "NO"
      Height          =   735
      Left            =   1920
      TabIndex        =   3
      Top             =   4920
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton YES 
      Caption         =   "YES"
      Height          =   735
      Left            =   1080
      TabIndex        =   0
      Top             =   3600
      Width           =   2055
   End
   Begin VB.Label Title2 
      Alignment       =   2  'Center
      Caption         =   "I KNEW IT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   840
      TabIndex        =   2
      Top             =   2640
      Visible         =   0   'False
      Width           =   5775
   End
   Begin VB.Label Title 
      Alignment       =   2  'Center
      Caption         =   "ARE YOU A FOOL?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   840
      TabIndex        =   1
      Top             =   1560
      Width           =   5775
   End
End
Attribute VB_Name = "NO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label1_Click()

End Sub

Private Sub NO_Click()

    NO1.Visible = True
    NO.Visible = False
    
End Sub

Private Sub NO1_Click()
    
    NO2.Visible = True
    NO1.Visible = False

End Sub

Private Sub NO2_Click()
    
    NO3.Visible = True
    NO2.Visible = False
    
End Sub

Private Sub NO3_Click()
    
    NO4.Visible = True
    NO3.Visible = False

End Sub

Private Sub NO4_Click()
    
    NO1.Visible = True
    NO4.Visible = False

End Sub

Private Sub YES_Click()

    NO.Visible = False
    NO1.Visible = False
    NO2.Visible = False
    NO3.Visible = False
    NO4.Visible = False
    YES.Visible = False
    Title.Visible = False
    Title2.Visible = YES
    
End Sub
