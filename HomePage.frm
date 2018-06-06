VERSION 5.00
Begin VB.Form HomePage 
   Caption         =   "Form3"
   ClientHeight    =   4695
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7200
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   30
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form3"
   ScaleHeight     =   4695
   ScaleWidth      =   7200
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "User"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3720
      TabIndex        =   2
      Top             =   2160
      Width           =   3135
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Admin"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   720
      TabIndex        =   1
      Top             =   2160
      Width           =   2775
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "AIR INDIA"
      ForeColor       =   &H8000000D&
      Height          =   735
      Left            =   600
      TabIndex        =   0
      Top             =   600
      Width           =   6015
   End
End
Attribute VB_Name = "HomePage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
AdminLogin.Show
Me.Hide
End Sub

Private Sub Command2_Click()
Userpage.Show
Me.Hide

End Sub

