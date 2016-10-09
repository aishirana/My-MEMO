VERSION 5.00
Begin VB.Form WriteDiary 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6345
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   7905
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "WriteDiary.frx":0000
   MousePointer    =   99  'Custom
   Picture         =   "WriteDiary.frx":030A
   ScaleHeight     =   6345
   ScaleWidth      =   7905
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3600
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3120
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.TextBox txtTemp 
      Appearance      =   0  'Flat
      Height          =   1815
      Left            =   1320
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   2160
      Visible         =   0   'False
      Width           =   5295
   End
   Begin VB.PictureBox picCancel 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   5400
      MouseIcon       =   "WriteDiary.frx":C310
      MousePointer    =   99  'Custom
      Picture         =   "WriteDiary.frx":C462
      ScaleHeight     =   435
      ScaleWidth      =   1725
      TabIndex        =   4
      Top             =   5280
      Width           =   1725
      Begin VB.Label lblCancel 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Close / Cancel"
         Height          =   210
         Left            =   330
         MouseIcon       =   "WriteDiary.frx":EC10
         MousePointer    =   99  'Custom
         TabIndex        =   5
         Top             =   120
         Width           =   1065
      End
   End
   Begin VB.PictureBox picSave 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   3600
      MouseIcon       =   "WriteDiary.frx":ED62
      MousePointer    =   99  'Custom
      Picture         =   "WriteDiary.frx":EEB4
      ScaleHeight     =   435
      ScaleWidth      =   1725
      TabIndex        =   2
      Top             =   5280
      Width           =   1725
      Begin VB.Label lblSave 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Save Diary"
         Height          =   210
         Left            =   435
         MouseIcon       =   "WriteDiary.frx":11662
         MousePointer    =   99  'Custom
         TabIndex        =   3
         Top             =   120
         Width           =   825
      End
   End
   Begin VB.TextBox txtDiary 
      Appearance      =   0  'Flat
      Height          =   3975
      Left            =   840
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   1080
      Width           =   6255
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Write to Diary"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "WriteDiary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

With Data1
    
    .DatabaseName = App.Path & "\db.mdb"
    .RecordSource = _
        "SELECT * FROM Diary WHERE Username ='" & Main.Tag & "'"
    .Refresh

End With

Icons.Show , Me

End Sub

Sub Convert_HTML()

If txtDiary.Text = "" Then
    
    MsgBox "Diary details is empty.", vbCritical, "Required"
    txtDiary.SetFocus
    Exit Sub
    
End If

txtTemp.Text = Replace(txtDiary.Text, "<", "&lt")
txtTemp.Text = Replace(txtTemp.Text, ">", "&gt")
txtTemp.Text = Replace(txtTemp.Text, """", "&quot")
txtTemp.Text = Replace(txtTemp.Text, vbCrLf, "<br>")

For i = 0 To 25
    txtTemp = Replace(txtTemp, "[" & CStr(Chr(i + 97)) & "]", _
        "<img src=""" & "GIF\" & CStr(Chr(i + 97)) & ".gif"">")
Next

With Data1.Recordset

    .AddNew
    .Fields("Username") = Main.Tag
    .Fields("Date") = Me.Tag
    .Fields("Diary") = txtDiary
    .Fields("HTML") = txtTemp
    .Update

    MsgBox "Diary has been saved.", vbInformation, "Message"

End With

lblCancel_Click

End Sub

Private Sub lblCancel_Click()

Diary.Enabled = True
Diary.SetFocus
Unload Me

End Sub

Private Sub lblSave_Click()

Convert_HTML

End Sub

Private Sub picCancel_Click()

Diary.Enabled = True
Diary.SetFocus
Unload Me

End Sub

Private Sub picSave_Click()

Convert_HTML

End Sub

Private Sub picSpell_Click()

End Sub
