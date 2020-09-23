VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "make a story"
   ClientHeight    =   2550
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6900
   Icon            =   "Form3.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2550
   ScaleWidth      =   6900
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "exit"
      Height          =   375
      Left            =   1200
      TabIndex        =   9
      Top             =   2040
      Width           =   1335
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   480
      TabIndex        =   4
      Text            =   "a thing"
      Top             =   1560
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "submit"
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   2040
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Form3.frx":0ECA
      Left            =   480
      List            =   "Form3.frx":0ED4
      TabIndex        =   2
      Text            =   "likes"
      Top             =   1200
      Width           =   1815
   End
   Begin VB.TextBox Text2 
      Height          =   315
      Left            =   480
      TabIndex        =   1
      Text            =   "he,she or it"
      Top             =   840
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   480
      TabIndex        =   0
      Text            =   "a thing"
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label Label4 
      Caption         =   "a thing eg: a banana, bananas, the banana or the bananas"
      Height          =   255
      Left            =   2280
      TabIndex        =   8
      Top             =   1560
      Width           =   4575
   End
   Begin VB.Label Label3 
      Caption         =   "likes or dosnt likes"
      Height          =   255
      Left            =   2400
      TabIndex        =   7
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "he, she or it"
      Height          =   255
      Left            =   2400
      TabIndex        =   6
      Top             =   840
      Width           =   3015
   End
   Begin VB.Label Label1 
      Caption         =   "a thing eg: a dog"
      Height          =   375
      Left            =   2400
      TabIndex        =   5
      Top             =   480
      Width           =   2655
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'this will open the window with the finished story by pressing the "submit" button
Private Sub Command1_Click()
Form4.Visible = True
End Sub
'this will close the window
Private Sub Command2_Click()
Form3.Visible = False
End Sub

