VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "thewardog's easy but cool thingys"
   ClientHeight    =   5010
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   7155
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5010
   ScaleWidth      =   7155
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1053
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Left            =   240
      TabIndex        =   51
      Text            =   "5"
      Top             =   3720
      Width           =   1335
   End
   Begin VB.CommandButton Command28 
      Caption         =   "push here"
      Height          =   495
      Left            =   120
      TabIndex        =   50
      Top             =   4200
      Width           =   1695
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   240
      Top             =   3120
   End
   Begin VB.CommandButton Command27 
      Caption         =   "4"
      Height          =   255
      Left            =   1440
      TabIndex        =   40
      Top             =   2760
      Width           =   255
   End
   Begin VB.CommandButton Command26 
      Caption         =   "3"
      Height          =   255
      Left            =   1080
      TabIndex        =   39
      Top             =   2760
      Width           =   255
   End
   Begin VB.CommandButton Command25 
      Caption         =   "2"
      Height          =   255
      Left            =   720
      TabIndex        =   38
      Top             =   2760
      Width           =   255
   End
   Begin VB.CommandButton Command24 
      Caption         =   "1"
      Height          =   255
      Left            =   360
      TabIndex        =   37
      Top             =   2760
      Width           =   255
   End
   Begin VB.CommandButton Command23 
      Caption         =   "4"
      Height          =   255
      Left            =   1440
      TabIndex        =   36
      Top             =   2400
      Width           =   255
   End
   Begin VB.CommandButton Command4 
      Caption         =   "3"
      Height          =   255
      Left            =   1080
      TabIndex        =   35
      Top             =   2400
      Width           =   255
   End
   Begin VB.CommandButton Command3 
      Caption         =   "2"
      Height          =   255
      Left            =   720
      TabIndex        =   34
      Top             =   2400
      Width           =   255
   End
   Begin VB.CommandButton Command2 
      Caption         =   "1"
      Height          =   255
      Left            =   360
      TabIndex        =   33
      Top             =   2400
      Width           =   255
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      ItemData        =   "Form1.frx":0ECA
      Left            =   4200
      List            =   "Form1.frx":0EDA
      TabIndex        =   32
      Text            =   "and this"
      Top             =   1680
      Width           =   1335
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Form1.frx":0F17
      Left            =   4200
      List            =   "Form1.frx":0F27
      TabIndex        =   31
      Text            =   "this "
      Top             =   1320
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "push here"
      Height          =   495
      Left            =   3960
      TabIndex        =   30
      Top             =   2160
      Width           =   2655
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   3600
      TabIndex        =   25
      Text            =   "label here"
      Top             =   3360
      Width           =   1935
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   3600
      TabIndex        =   24
      Text            =   "text here"
      Top             =   3840
      Width           =   1935
   End
   Begin VB.CommandButton Command22 
      Caption         =   "push here"
      Height          =   375
      Left            =   3840
      TabIndex        =   23
      Top             =   4320
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   3960
      TabIndex        =   21
      Text            =   "text here"
      Top             =   480
      Width           =   1335
   End
   Begin VB.CommandButton Command21 
      Caption         =   "push here"
      Height          =   255
      Left            =   3960
      TabIndex        =   20
      Top             =   840
      Width           =   1575
   End
   Begin VB.CommandButton Command20 
      Caption         =   "4"
      Height          =   255
      Left            =   1440
      TabIndex        =   19
      Top             =   2040
      Width           =   255
   End
   Begin VB.CommandButton Command19 
      Caption         =   "3"
      Height          =   255
      Left            =   1080
      TabIndex        =   18
      Top             =   2040
      Width           =   255
   End
   Begin VB.CommandButton Command18 
      Caption         =   "2"
      Height          =   255
      Left            =   720
      TabIndex        =   17
      Top             =   2040
      Width           =   255
   End
   Begin VB.CommandButton Command17 
      Caption         =   "1"
      Height          =   255
      Left            =   360
      TabIndex        =   16
      Top             =   2040
      Width           =   255
   End
   Begin VB.CommandButton Command16 
      Caption         =   "4"
      Height          =   255
      Left            =   1440
      TabIndex        =   15
      Top             =   1680
      Width           =   255
   End
   Begin VB.CommandButton Command15 
      Caption         =   "3"
      Height          =   255
      Left            =   1080
      TabIndex        =   14
      Top             =   1680
      Width           =   255
   End
   Begin VB.CommandButton Command14 
      Caption         =   "2"
      Height          =   255
      Left            =   720
      TabIndex        =   13
      Top             =   1680
      Width           =   255
   End
   Begin VB.CommandButton Command13 
      Caption         =   "1"
      Height          =   255
      Left            =   360
      TabIndex        =   12
      Top             =   1680
      Width           =   255
   End
   Begin VB.CommandButton Command12 
      Caption         =   "4"
      Height          =   255
      Left            =   1440
      TabIndex        =   11
      Top             =   1320
      Width           =   255
   End
   Begin VB.CommandButton Command11 
      Caption         =   "3"
      Height          =   255
      Left            =   1080
      TabIndex        =   10
      Top             =   1320
      Width           =   255
   End
   Begin VB.CommandButton Command10 
      Caption         =   "2"
      Height          =   255
      Left            =   720
      TabIndex        =   9
      Top             =   1320
      Width           =   255
   End
   Begin VB.CommandButton Command9 
      Caption         =   "1"
      Height          =   255
      Left            =   360
      TabIndex        =   8
      Top             =   1320
      Width           =   255
   End
   Begin VB.CommandButton Command8 
      Caption         =   "4"
      Height          =   255
      Left            =   1440
      TabIndex        =   7
      Top             =   960
      Width           =   255
   End
   Begin VB.CommandButton Command7 
      Caption         =   "3"
      Height          =   255
      Left            =   1080
      TabIndex        =   6
      Top             =   960
      Width           =   255
   End
   Begin VB.CommandButton Command6 
      Caption         =   "2"
      Height          =   255
      Left            =   720
      TabIndex        =   5
      Top             =   960
      Width           =   255
   End
   Begin VB.CommandButton Command5 
      Caption         =   "1"
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   960
      Width           =   255
   End
   Begin VB.OptionButton Option4 
      Caption         =   "Option4"
      Height          =   255
      Left            =   1440
      TabIndex        =   3
      Top             =   600
      Width           =   255
   End
   Begin VB.OptionButton Option3 
      Caption         =   "Option3"
      Height          =   255
      Left            =   1080
      TabIndex        =   2
      Top             =   600
      Width           =   255
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Option2"
      Height          =   255
      Left            =   720
      TabIndex        =   1
      Top             =   600
      Width           =   255
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Option1"
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   600
      Width           =   255
   End
   Begin VB.Label Label16 
      Caption         =   "interval (seconds)"
      Height          =   375
      Left            =   1920
      TabIndex        =   53
      Top             =   3840
      Width           =   1335
   End
   Begin VB.Label Label15 
      Caption         =   "timed message"
      Height          =   375
      Left            =   240
      TabIndex        =   52
      Top             =   3480
      Width           =   1575
   End
   Begin VB.Label Label14 
      Caption         =   "visible"
      Height          =   255
      Left            =   2280
      TabIndex        =   49
      Top             =   2640
      Width           =   975
   End
   Begin VB.Label Label13 
      Caption         =   "False"
      Height          =   255
      Left            =   1800
      TabIndex        =   48
      Top             =   2760
      Width           =   615
   End
   Begin VB.Label Label12 
      Caption         =   "True"
      Height          =   255
      Left            =   1800
      TabIndex        =   47
      Top             =   2400
      Width           =   615
   End
   Begin VB.Label Label11 
      Caption         =   "Checked"
      Height          =   375
      Left            =   2280
      TabIndex        =   46
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label Label7 
      Caption         =   "False"
      Height          =   255
      Left            =   1800
      TabIndex        =   45
      Top             =   2040
      Width           =   615
   End
   Begin VB.Label Label5 
      Caption         =   "True"
      Height          =   255
      Left            =   1800
      TabIndex        =   44
      Top             =   1680
      Width           =   615
   End
   Begin VB.Label Label4 
      Caption         =   "Enabled"
      Height          =   375
      Left            =   2280
      TabIndex        =   43
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "True"
      Height          =   255
      Left            =   1800
      TabIndex        =   42
      Top             =   960
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "False"
      Height          =   255
      Left            =   1800
      TabIndex        =   41
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "change options"
      Height          =   495
      Left            =   360
      TabIndex        =   29
      Top             =   360
      Width           =   1935
   End
   Begin VB.Label Label10 
      Caption         =   "make a popup"
      Height          =   255
      Left            =   3840
      TabIndex        =   28
      Top             =   3000
      Width           =   2055
   End
   Begin VB.Label Label9 
      Caption         =   "message text"
      Height          =   255
      Left            =   5640
      TabIndex        =   27
      Top             =   3960
      Width           =   1095
   End
   Begin VB.Label Label8 
      Caption         =   "message Label"
      Height          =   375
      Left            =   5640
      TabIndex        =   26
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "change caption on button"
      Height          =   255
      Left            =   3720
      TabIndex        =   22
      Top             =   120
      Width           =   2295
   End
   Begin VB.Menu about 
      Caption         =   "&About"
   End
   Begin VB.Menu other 
      Caption         =   "&Other thingys"
      Begin VB.Menu story 
         Caption         =   "&Story"
      End
      Begin VB.Menu navigate 
         Caption         =   "&Navigate window"
      End
   End
   Begin VB.Menu exit 
      Caption         =   "exit"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'this will open the about screen when pressing about in the menu
Private Sub about_Click()
frmAbout.Visible = True
End Sub

'this will change the button caption to "This and this" or "do you yahoo" etc..
Private Sub Command1_Click()
Command1.Caption = Combo1.Text & Combo2.Text
End Sub

'this will make option 1 visible
Private Sub Command2_Click()
Option1.Visible = True
End Sub

'this will make option 2 visible
Private Sub Command23_Click()
Option4.Visible = True
End Sub

'this will make option 1 invisible
Private Sub Command24_Click()
Option1.Visible = False
End Sub

'this will make option 2 invisible
Private Sub Command25_Click()
Option2.Visible = False
End Sub

'this will make option 3 invisible
Private Sub Command26_Click()
Option3.Visible = False
End Sub

'this will make option 4 invisible
Private Sub Command27_Click()
Option4.Visible = False
End Sub

'this will start the timer and make the timer interval to eg 5 seconds
'but becouse of one value in the interval is 1 millisecond i hade to multiply it with 1000
Private Sub Command28_Click()
Timer1.Enabled = True
Timer1.Interval = Text4.Text * 1000 'as you can see here.
Text4.Text = "please wait..."
End Sub

'this will make option 2 visible
Private Sub Command3_Click()
Option2.Visible = True
End Sub

'this will make option 3 visible
Private Sub Command4_Click()
Option3.Visible = True
End Sub

'this will change the button caption to the text above
Private Sub Command21_Click()
Command21.Caption = Text1.Text
End Sub

'this will make the popup popup
Private Sub Command22_Click()
Form2.Visible = True
End Sub

'this will make option 1 enabled
Private Sub Command5_Click()
Option1.Enabled = True
End Sub

'this will make option 2 enabled
Private Sub Command6_Click()
Option2.Enabled = True
End Sub

'this will make option 3 enabled
Private Sub Command7_Click()
Option3.Enabled = True
End Sub

'this will make option 4 enabled
Private Sub Command8_Click()
Option4.Enabled = True
End Sub

'this will make option 1 unabled
Private Sub Command9_Click()
Option1.Enabled = False
End Sub

'this will make option 2 unabled
Private Sub Command10_Click()
Option2.Enabled = False
End Sub

'this will make option 3 unabled
Private Sub Command11_Click()
Option3.Enabled = False
End Sub

'this will make option 4 unabled
Private Sub Command12_Click()
Option4.Enabled = False
End Sub

'this will make option 1 checked
Private Sub Command13_Click()
Option1.Value = True
End Sub

'this will make option 2 checked
Private Sub Command14_Click()
Option2.Value = True
End Sub

'this will make option 3 checked
Private Sub Command15_Click()
Option3.Value = True
End Sub

'this will make option 4 checked
Private Sub Command16_Click()
Option4.Value = True
End Sub

'this will make option 1 unchecked
Private Sub Command17_Click()
Option1.Value = False
End Sub

'this will make option 2 unchecked
Private Sub Command18_Click()
Option2.Value = False
End Sub

'this will make option 3 unchecked
Private Sub Command19_Click()
Option3.Value = False
End Sub

'this will make option 4 unchecked
Private Sub Command20_Click()
Option4.Value = False
End Sub

'exit the programm trough the menu
Private Sub exit_Click()
 Unload Me
End Sub

'this will open the navigatewindow script trough the menu
Private Sub navigate_Click()
Form5.Visible = True
End Sub

'this will open the story script trough the menu
Private Sub story_Click()
Form3.Visible = True
End Sub

'this will open the timed popup after the interval (above the button)
Private Sub Timer1_Timer()
Text4.Text = "5"
Form7.Visible = True
Timer1.Enabled = False
End Sub

'this script makes it unable to write something else then numeric in the interval
Private Sub Text4_KeyPress(KeyAscii As Integer)
    If Not IsNumeric(Chr(KeyAscii)) Then KeyAscii = 0
End Sub

