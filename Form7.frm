VERSION 5.00
Begin VB.Form Form7 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "timed message"
   ClientHeight    =   270
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3150
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   270
   ScaleWidth      =   3150
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   1455
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3015
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'this will show how long time it took to wait for the window to open
Private Sub Form_Load()
Label1.Caption = "it took " & Form1.Timer1.Interval / 1000 & " seconds to open this window"
End Sub
