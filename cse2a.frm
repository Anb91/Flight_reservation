VERSION 5.00
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Begin VB.Form Userpage 
   Caption         =   "Form1"
   ClientHeight    =   8700
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10905
   LinkTopic       =   "Form1"
   ScaleHeight     =   8700
   ScaleWidth      =   10905
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1680
      TabIndex        =   21
      Top             =   6960
      Width           =   4575
   End
   Begin MSACAL.Calendar Calendar2 
      Height          =   2175
      Left            =   5400
      TabIndex        =   18
      Top             =   2760
      Width           =   3735
      _Version        =   524288
      _ExtentX        =   6588
      _ExtentY        =   3836
      _StockProps     =   1
      BackColor       =   -2147483633
      Year            =   2018
      Month           =   5
      Day             =   11
      DayLength       =   1
      MonthLength     =   1
      DayFontColor    =   0
      FirstDay        =   7
      GridCellEffect  =   1
      GridFontColor   =   10485760
      GridLinesColor  =   -2147483632
      ShowDateSelectors=   -1  'True
      ShowDays        =   -1  'True
      ShowHorizontalGrid=   -1  'True
      ShowTitle       =   -1  'True
      ShowVerticalGrid=   -1  'True
      TitleFontColor  =   10485760
      ValueIsNull     =   0   'False
      BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSACAL.Calendar Calendar1 
      Height          =   2175
      Left            =   480
      TabIndex        =   17
      Top             =   2760
      Width           =   3735
      _Version        =   524288
      _ExtentX        =   6588
      _ExtentY        =   3836
      _StockProps     =   1
      BackColor       =   -2147483633
      Year            =   2018
      Month           =   5
      Day             =   11
      DayLength       =   1
      MonthLength     =   1
      DayFontColor    =   0
      FirstDay        =   7
      GridCellEffect  =   1
      GridFontColor   =   10485760
      GridLinesColor  =   -2147483632
      ShowDateSelectors=   -1  'True
      ShowDays        =   -1  'True
      ShowHorizontalGrid=   -1  'True
      ShowTitle       =   -1  'True
      ShowVerticalGrid=   -1  'True
      TitleFontColor  =   10485760
      ValueIsNull     =   0   'False
      BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.ComboBox Combo2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      ItemData        =   "cse2a.frx":0000
      Left            =   6000
      List            =   "cse2a.frx":0013
      TabIndex        =   16
      Text            =   "Destination"
      Top             =   1560
      Width           =   4215
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      ItemData        =   "cse2a.frx":0043
      Left            =   6000
      List            =   "cse2a.frx":0053
      TabIndex        =   15
      Text            =   "From Where"
      Top             =   1080
      Width           =   4215
   End
   Begin VB.TextBox Text5 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5280
      TabIndex        =   14
      Top             =   6120
      Width           =   3495
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3240
      TabIndex        =   13
      Top             =   6120
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1320
      TabIndex        =   12
      Top             =   6120
      Width           =   615
   End
   Begin VB.CommandButton Command7 
      Caption         =   "CONFIRM"
      BeginProperty Font 
         Name            =   "AR BLANCA"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3960
      MaskColor       =   &H00C00000&
      TabIndex        =   7
      Top             =   7680
      Width           =   3375
   End
   Begin VB.Frame Frame1 
      Caption         =   "   Type Of Journey"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   360
      TabIndex        =   0
      Top             =   960
      Width           =   3375
      Begin VB.OptionButton Option2 
         Caption         =   "Round Trip"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1560
         TabIndex        =   9
         Top             =   600
         Width           =   1695
      End
      Begin VB.OptionButton Option1 
         Caption         =   "One Way"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   600
         Width           =   1455
      End
   End
   Begin VB.Label Label5 
      Caption         =   "Mail Id"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   20
      Top             =   6960
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Welcome To IIR INDIA Flight Reservation "
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   615
      Left            =   2040
      TabIndex        =   19
      Top             =   240
      Width           =   6855
   End
   Begin VB.Label Label10 
      Caption         =   "Infant"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      TabIndex        =   11
      Top             =   6120
      Width           =   1215
   End
   Begin VB.Label Label7 
      Caption         =   "Adult"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   10
      Top             =   6120
      Width           =   855
   End
   Begin VB.Label Label15 
      Caption         =   "Class:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6600
      TabIndex        =   6
      Top             =   5280
      Width           =   975
   End
   Begin VB.Label Label11 
      Caption         =   "No. of Passengers:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      TabIndex        =   5
      Top             =   5280
      Width           =   3375
   End
   Begin VB.Label Label8 
      Caption         =   "Return On:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6720
      TabIndex        =   4
      Top             =   2160
      Width           =   1695
   End
   Begin VB.Label Label4 
      Caption         =   "To:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   375
      Left            =   4800
      TabIndex        =   3
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   "From:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   375
      Left            =   4560
      TabIndex        =   2
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Date Of Journey"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      TabIndex        =   1
      Top             =   2160
      Width           =   2535
   End
End
Attribute VB_Name = "Userpage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim adult As String
Dim infant As String
Dim class As String
Dim c1 As String
Dim c2 As String
Dim op1 As String
Dim mail As String
Dim x As String
Dim y As String
Dim i As Integer
Dim j As Integer

Private Sub Command1_Click()
Text1.Text = Text1.Text + 1
End Sub

Private Sub Command2_Click()
If Text1.Text = 1 Then
   Text1.Text = 1
 Else
   Text1.Text = Text1.Text - 1
End If
End Sub

Private Sub Command3_Click()
If Text2.Text = 0 Then
   Text2.Text = 0
 Else
   Text2.Text = Text2.Text - 1
End If
End Sub

Private Sub Command4_Click()
  Text2.Text = Text2.Text + 1
End Sub

Private Sub Command5_Click()
If Text3.Text = 0 Then
   Text3.Text = 0
 Else
   Text3.Text = Text3.Text - 1
End If
End Sub

Private Sub Command6_Click()
Text3.Text = Text3.Text + 1
End Sub

Private Sub Calendar1_Click()
c1 = Calendar1.Value
End Sub

Private Sub Calendar2_Click()
c2 = Calendar2.Value
End Sub

Private Sub Command7_Click()
x = Combo1.ListIndex
y = Combo2.ListIndex
MsgBox ("Welcome to AIR INDIA..." & _
vbCrLf + "Type of journey : " & op1 & _
vbCrLf + "From - " & Combo1.List(x) & _
vbCrLf + "To - " & Combo2.List(y) & _
vbCrLf + "Date of journey : " & c1 & _
vbCrLf + "Date of return : " & c2 & _
vbCrLf + "Adult : " & adult & _
vbCrLf + "Infant : " & infant & _
vbCrLf + "Class : " & class & _
vbCrLf + "Time of flight and other detals will be mailed to your mail id : " & mail & _
vbCrLf + "Happy Journey...")

End Sub

Private Sub Option1_Click()
op1 = Option1.Caption()
End Sub

Private Sub Option2_Click()
op1 = Option2.Caption()
End Sub

Private Sub Text1_Change()
adult = Text1.Text
End Sub

Private Sub Text2_Change()
infant = Text2.Text
End Sub

Private Sub Text3_Change()
mail = Text3.Text
End Sub

Private Sub Text5_Change()
class = Text5.Text
End Sub
