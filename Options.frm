VERSION 5.00
Begin VB.Form Options 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Options"
   ClientHeight    =   6045
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3705
   Icon            =   "Options.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6045
   ScaleWidth      =   3705
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   1200
      TabIndex        =   12
      Text            =   "Mortal Kombat"
      Top             =   4800
      Width           =   1335
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1200
      TabIndex        =   10
      Text            =   "Lighthouse"
      Top             =   3960
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1200
      TabIndex        =   8
      Text            =   "Yellow"
      Top             =   3240
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1200
      TabIndex        =   5
      Text            =   "Blue"
      Top             =   2640
      Width           =   1335
   End
   Begin VB.CommandButton Apply 
      Caption         =   "Apply"
      Height          =   375
      Left            =   840
      TabIndex        =   3
      Top             =   5400
      Width           =   1935
   End
   Begin VB.CheckBox playfx 
      Caption         =   "Play Sound Effects"
      Height          =   315
      Left            =   1200
      TabIndex        =   2
      Top             =   1680
      Width           =   2295
   End
   Begin VB.CheckBox showframe 
      Caption         =   "Display Frame Rate"
      Height          =   315
      Left            =   1200
      TabIndex        =   1
      Top             =   2040
      Width           =   2295
   End
   Begin VB.CheckBox playmusic 
      Caption         =   "Play Music"
      Height          =   315
      Left            =   1200
      TabIndex        =   0
      Top             =   1320
      Width           =   2295
   End
   Begin VB.Label Label5 
      Caption         =   "Sound Track:"
      Height          =   255
      Left            =   1200
      TabIndex        =   11
      Top             =   4440
      Width           =   1575
   End
   Begin VB.Label Label4 
      Caption         =   "Fighting Ground:"
      Height          =   255
      Left            =   1200
      TabIndex        =   9
      Top             =   3600
      Width           =   1575
   End
   Begin VB.Label Label3 
      Caption         =   "Player2 Name:"
      Height          =   255
      Left            =   1200
      TabIndex        =   7
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Player 1 Name:"
      Height          =   255
      Left            =   1200
      TabIndex        =   6
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Blood Match Options"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1215
      Left            =   240
      TabIndex        =   4
      Top             =   120
      Width           =   3255
   End
End
Attribute VB_Name = "Options"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Private Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long 'This API call will be used to stop the normal windows cursor from appearing.

Private Sub Apply_Click()
Open "options.ini" For Output As #1
If playmusic.Value = 1 Then
  Print #1, "Playmusic = T"
ElseIf playmusic.Value = 0 Then
  Print #1, "Playmusic = F"
End If
If playfx.Value = 0 Then
  Print #1, "PlayFX = F"
ElseIf playfx.Value = 1 Then
  Print #1, "PlayFX = T"
End If
If showframe.Value = 1 Then
  Print #1, "ShowFrame = T"
ElseIf showframe.Value = 0 Then
  Print #1, "ShowFrame = F"
End If

Print #1, Text1.Text
Print #1, Text2.Text


If Combo1.Text = "Sunset" Then
  Print #1, "Battleground = Sunset    "
ElseIf Combo1.Text = "Lighthouse" Then
  Print #1, "Battleground = Lighthouse"
End If

If Combo2.Text = "Mortal Kombat" Then
  Print #1, "Music = Mortal Kombat"
ElseIf Combo2.Text = "Terminator 2" Then
  Print #1, "Music = Terminator 2 "
End If

Close #1

Unload game
game.Hide
Load game
game.Show

End Sub



Private Sub Form_Load()
Combo1.AddItem "Sunset"
Combo1.AddItem "Lighthouse"

Combo2.AddItem "Mortal Kombat"
Combo2.AddItem "Terminator 2"

Open "options.ini" For Input As #1
Dim tempvar As String
Line Input #1, tempvar
If Right(tempvar, 1) = "T" Then playmusic.Value = 1
Line Input #1, tempvar
If Right(tempvar, 1) = "T" Then playfx.Value = 1
Line Input #1, tempvar
If Right(tempvar, 1) = "T" Then showframe.Value = 1

Line Input #1, tempvar
Text1.Text = tempvar
Line Input #1, tempvar
Text2.Text = tempvar
Line Input #1, tempvar
If Right(tempvar, 10) = "Sunset    " Then
  Combo1.Text = "Sunset    "
Else
  Combo1.Text = "Lighthouse"
End If
Line Input #1, tempvar
If Right(tempvar, 13) = "Mortal Kombat" Then
  Combo2.Text = "Mortal Kombat"
Else
  Combo2.Text = "Terminator 2 "
End If


Close #1






End Sub
