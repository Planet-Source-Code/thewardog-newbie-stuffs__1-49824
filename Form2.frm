VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form2"
   ClientHeight    =   2175
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3450
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2175
   ScaleWidth      =   3450
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   1320
      TabIndex        =   1
      Top             =   1680
      Width           =   735
   End
   Begin VB.Label Labe89 
      Height          =   2175
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3615
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'this will close the window when pressing the "ok" button
Private Sub Command1_Click()
Unload Me
End Sub

'this will change the windows caption and the text from the
'"label here" and "text" here
Private Sub Form_Load()
Labe89.Caption = Form1.Text2
Form2.Caption = Form1.Text3
End Sub

