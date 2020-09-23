VERSION 5.00
Begin VB.Form Form5 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1080
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   1095
   Icon            =   "Form5.frx":0000
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1080
   ScaleWidth      =   1095
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Up"
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   0
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Lt"
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   360
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Rt"
      Height          =   375
      Left            =   720
      TabIndex        =   1
      Top             =   360
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Dn"
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   720
      Width           =   375
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'this will move the window up 40 pixels
Private Sub Command1_Click()
Form5.Top = Form5.Top - 40
End Sub

'this will move the window left 40 pixels
Private Sub Command2_Click()
Form5.Left = Form5.Left - 40
End Sub

'this will move the window right 40 pixels
Private Sub Command3_Click()
Form5.Left = Form5.Left + 40
End Sub

'this will move the window down 40 pixels
Private Sub Command4_Click()
Form5.Top = Form5.Top + 40
End Sub

