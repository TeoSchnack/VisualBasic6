VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2100
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5070
   LinkTopic       =   "Form1"
   ScaleHeight     =   2100
   ScaleWidth      =   5070
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Height          =   1575
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Visible         =   0   'False
      Width           =   4575
      Begin VB.CommandButton FIM 
         Caption         =   "FIM"
         Height          =   255
         Left            =   2160
         TabIndex        =   7
         Top             =   1200
         Width           =   975
      End
      Begin VB.CommandButton CONFIRMA 
         Caption         =   "CONFIRMA"
         Height          =   255
         Left            =   3360
         TabIndex        =   6
         Top             =   1200
         Width           =   975
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   4215
      End
      Begin VB.Label Label2 
         Caption         =   "Label2"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   2295
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1215
      Left            =   240
      TabIndex        =   1
      Top             =   360
      Width           =   4575
      Begin VB.CommandButton OK 
         Caption         =   "OK"
         Height          =   255
         Left            =   1560
         TabIndex        =   4
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Nesse programa o novo salário dos coleaboradores da empresa é caclculado"
         Height          =   495
         Left            =   360
         TabIndex        =   3
         Top             =   240
         Width           =   3735
      End
   End
   Begin VB.Frame Frame1 
      Height          =   4695
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Visible         =   0   'False
      Width           =   5175
      Begin VB.Label Label3 
         Height          =   4095
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   4815
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public T As Boolean
Private Sub CONFIRMA_Click()
    
    If T Then
    
        Dim Sal As Double
        Dim Ab As String
        
        Sal = (CDbl(Text2.Text))
        If Sal <= 1200 Then
            Sal = Sal * 1.1
        ElseIf Sal <= 2500 Then
            Sal = Sal * 1.08
        Else
            Sal = Sal * 1.05
        End If
        
        Sal = Sal / 1.1
        
        If Sal >= 2500 Then
            Sal = Sal / 1.18
        End If
        
        Ab = CStr(Sal)
        
        Label3.Caption = Label3.Caption + Text2.Text + "      ->      " + Ab + vbCrLf + vbCrLf
        
        
        Label2.Caption = "Qual o nome do colaborador?"
        Text2.Text = ""
        FIM.Visible = True
        T = False
    
    Else
        
        Label3.Caption = Label3.Caption + Text2.Text + vbCrLf
        
        Label2.Caption = "Qual o salario do colaborador?"
        Text2.Text = ""
        FIM.Visible = False
        T = True

    
    End If


End Sub
Private Sub FIM_Click()
    
    Frame1.Visible = True
    Frame3.Visible = False
    Form1.Height = 6150
    Form1.Width = 5970
    
End Sub

Private Sub OK_Click()

    Frame2.Visible = False
    Frame3.Visible = True
    Label2.Caption = "Qual o nome do colaborador?"
    T = False

End Sub

Private Sub Text2_Change()

End Sub
