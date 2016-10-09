VERSION 5.00
Begin VB.Form Main 
   Caption         =   "My MEMO"
   ClientHeight    =   5910
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7740
   LinkTopic       =   "Form1"
   Picture         =   "Main.frx":0000
   ScaleHeight     =   5910
   ScaleWidth      =   7740
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picChoices 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   855
      Index           =   3
      Left            =   4560
      MouseIcon       =   "Main.frx":C006
      MousePointer    =   99  'Custom
      Picture         =   "Main.frx":C158
      ScaleHeight     =   855
      ScaleWidth      =   705
      TabIndex        =   5
      Top             =   3360
      Width           =   705
   End
   Begin VB.PictureBox picChoices 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   855
      Index           =   2
      Left            =   4560
      MouseIcon       =   "Main.frx":CBC3
      MousePointer    =   99  'Custom
      Picture         =   "Main.frx":CD15
      ScaleHeight     =   855
      ScaleWidth      =   705
      TabIndex        =   4
      Top             =   2160
      Width           =   705
   End
   Begin VB.PictureBox picChoices 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   855
      Index           =   1
      Left            =   1800
      MouseIcon       =   "Main.frx":DAD1
      MousePointer    =   99  'Custom
      Picture         =   "Main.frx":DC23
      ScaleHeight     =   855
      ScaleWidth      =   705
      TabIndex        =   3
      Top             =   3360
      Width           =   705
   End
   Begin VB.PictureBox picChoices 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   855
      Index           =   0
      Left            =   1800
      MouseIcon       =   "Main.frx":E942
      MousePointer    =   99  'Custom
      Picture         =   "Main.frx":EA94
      ScaleHeight     =   855
      ScaleWidth      =   705
      TabIndex        =   2
      Top             =   2160
      Width           =   705
   End
   Begin VB.PictureBox picQuit 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   6240
      MouseIcon       =   "Main.frx":F893
      MousePointer    =   99  'Custom
      Picture         =   "Main.frx":F9E5
      ScaleHeight     =   435
      ScaleWidth      =   1230
      TabIndex        =   0
      Top             =   5280
      Width           =   1230
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Quit"
         Height          =   210
         Index           =   4
         Left            =   480
         MouseIcon       =   "Main.frx":1163F
         MousePointer    =   99  'Custom
         TabIndex        =   1
         Top             =   120
         Width           =   405
      End
   End
   Begin VB.Label lblUser 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   5400
      Width           =   3975
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Profile"
      Height          =   210
      Index           =   3
      Left            =   5520
      MouseIcon       =   "Main.frx":11791
      MousePointer    =   99  'Custom
      TabIndex        =   11
      Top             =   3675
      Width           =   450
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Reminders"
      Height          =   210
      Index           =   2
      Left            =   5520
      MouseIcon       =   "Main.frx":118E3
      MousePointer    =   99  'Custom
      TabIndex        =   10
      Top             =   2475
      Width           =   765
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Notes"
      Height          =   210
      Index           =   1
      Left            =   2760
      MouseIcon       =   "Main.frx":11A35
      MousePointer    =   99  'Custom
      TabIndex        =   9
      Top             =   3675
      Width           =   420
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Diary"
      Height          =   210
      Index           =   0
      Left            =   2760
      MouseIcon       =   "Main.frx":11B87
      MousePointer    =   99  'Custom
      TabIndex        =   8
      Top             =   2475
      Width           =   375
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "by Aishi Rana Putiara"
      Height          =   255
      Left            =   1560
      TabIndex        =   7
      Top             =   960
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "My MEMO"
      BeginProperty Font 
         Name            =   "Jokerman"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      TabIndex        =   6
      Top             =   240
      Width           =   2895
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 
If (X >= 1155 And Y >= 0) And (X <= 6720 And Y <= 300) Then
    If Button = vbLeftButton Then Call DragIt(Me.hwnd)
End If

End Sub

Sub Reset_Captions()

For i = 0 To 4
    lblCaption(i).FontUnderline = False
Next

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Reset_Captions

End Sub

Private Sub lblCaption_Click(Index As Integer)
Dim ans

Select Case Index

    Case 0
        Load Diary
        Diary.Show , Me
    Case 1
        Load Notes
        Notes.Show , Me
    Case 2
        Load Reminders
        Reminders.Show , Me
    Case 3
        Load Profile
        Profile.Show , Me
        
    Case 4
    
        ans = MsgBox("Are you sure?", vbQuestion + vbYesNo, "Quit My MEMO")
        
        If ans = vbNo Then Exit Sub
    
        End

End Select

End Sub

Private Sub lblCaption_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

If lblCaption(Index).FontUnderline = True Then Exit Sub

Reset_Captions
lblCaption(Index).FontUnderline = True

End Sub

Private Sub picChoices_Click(Index As Integer)

lblCaption_Click Index

End Sub

Private Sub picChoices_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

If lblCaption(Index).FontUnderline = True Then Exit Sub

Reset_Captions
lblCaption(Index).FontUnderline = True

End Sub

Private Sub picPick_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Reset_Captions

End Sub

