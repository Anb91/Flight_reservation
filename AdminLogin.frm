VERSION 5.00
Begin VB.Form AdminLogin 
   Caption         =   "Form2"
   ClientHeight    =   5475
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5910
   LinkTopic       =   "Form2"
   ScaleHeight     =   5475
   ScaleWidth      =   5910
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Log In"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1920
      TabIndex        =   5
      Top             =   4560
      Width           =   3135
   End
   Begin VB.TextBox Text2 
      Height          =   615
      Left            =   2160
      TabIndex        =   3
      Top             =   2760
      Width           =   2895
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   2160
      TabIndex        =   2
      Top             =   1920
      Width           =   2895
   End
   Begin VB.Label Label3 
      Caption         =   "Admin "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   615
      Left            =   1680
      TabIndex        =   4
      Top             =   840
      Width           =   2535
   End
   Begin VB.Label Label2 
      Caption         =   "Password:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   480
      TabIndex        =   1
      Top             =   2880
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Username:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      TabIndex        =   0
      Top             =   1920
      Width           =   1575
   End
End
Attribute VB_Name = "AdminLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Text1.Text = "admin" And Text2.Text = "admin" Then
Me.Hide
Else
MsgBox "Invalid Username and Password"
End If

End Sub

Private Sub Label2_Click()

End Sub

Private Sub Text1_Change()

End Sub
