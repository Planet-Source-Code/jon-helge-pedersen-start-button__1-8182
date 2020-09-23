VERSION 5.00
Begin VB.Form fStartButton 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Start Button"
   ClientHeight    =   540
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5160
   Icon            =   "fStartButton.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   540
   ScaleWidth      =   5160
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command5 
      Caption         =   ">"
      Height          =   330
      Left            =   4725
      TabIndex        =   4
      Top             =   105
      Width           =   330
   End
   Begin VB.CommandButton Command4 
      Caption         =   "<"
      Height          =   330
      Left            =   4410
      TabIndex        =   3
      Top             =   105
      Width           =   330
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Restore"
      Height          =   330
      Left            =   2835
      TabIndex        =   2
      Top             =   105
      Width           =   1380
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Clear"
      Height          =   330
      Left            =   1470
      TabIndex        =   1
      Top             =   105
      Width           =   1380
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Change bitmap"
      Height          =   330
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   1380
   End
End
Attribute VB_Name = "fStartButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Start As New cStartButton

Private Sub Command1_Click()
  Start.Bitmap = "Start.bmp"
End Sub

Private Sub Command2_Click()
  Start.Bitmap = ""
End Sub

Private Sub Command3_Click()
  Start.RestoreAll
End Sub

Private Sub Command4_Click()
  Start.Left = Start.Left - 100
End Sub

Private Sub Command5_Click()
  Start.Left = Start.Left + 100
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set Start = Nothing
End Sub
