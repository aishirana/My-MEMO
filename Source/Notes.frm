VERSION 5.00
Begin VB.Form Notes 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6315
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   7920
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
   MouseIcon       =   "Notes.frx":0000
   MousePointer    =   99  'Custom
   Picture         =   "Notes.frx":030A
   ScaleHeight     =   6315
   ScaleWidth      =   7920
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
      Left            =   2640
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2640
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.PictureBox picClose 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   435
      Left            =   5760
      MouseIcon       =   "Notes.frx":C310
      MousePointer    =   99  'Custom
      Picture         =   "Notes.frx":C462
      ScaleHeight     =   435
      ScaleWidth      =   1515
      TabIndex        =   9
      Top             =   5280
      Width           =   1515
      Begin VB.Label lblClose 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Close"
         Height          =   210
         Left            =   495
         MouseIcon       =   "Notes.frx":E714
         MousePointer    =   99  'Custom
         TabIndex        =   19
         Top             =   120
         Width           =   435
      End
   End
   Begin VB.PictureBox picNext 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   435
      Left            =   2160
      MouseIcon       =   "Notes.frx":E866
      MousePointer    =   99  'Custom
      Picture         =   "Notes.frx":E9B8
      ScaleHeight     =   435
      ScaleWidth      =   1515
      TabIndex        =   8
      Top             =   5280
      Width           =   1515
      Begin VB.Label lblNext 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ne&xt"
         Height          =   210
         Left            =   540
         MouseIcon       =   "Notes.frx":10C6A
         MousePointer    =   99  'Custom
         TabIndex        =   18
         Top             =   120
         Width           =   345
      End
   End
   Begin VB.PictureBox picPrev 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   435
      Left            =   600
      MouseIcon       =   "Notes.frx":10DBC
      MousePointer    =   99  'Custom
      Picture         =   "Notes.frx":10F0E
      ScaleHeight     =   435
      ScaleWidth      =   1515
      TabIndex        =   7
      Top             =   5280
      Width           =   1515
      Begin VB.Label lblPrev 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Previous"
         Height          =   210
         Left            =   390
         MouseIcon       =   "Notes.frx":131C0
         MousePointer    =   99  'Custom
         TabIndex        =   17
         Top             =   120
         Width           =   645
      End
   End
   Begin VB.PictureBox picDelete 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   435
      Left            =   5760
      MouseIcon       =   "Notes.frx":13312
      MousePointer    =   99  'Custom
      Picture         =   "Notes.frx":13464
      ScaleHeight     =   435
      ScaleWidth      =   1515
      TabIndex        =   6
      Top             =   2760
      Width           =   1515
      Begin VB.Label lblDelete 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Delete"
         Height          =   210
         Left            =   540
         MouseIcon       =   "Notes.frx":15716
         MousePointer    =   99  'Custom
         TabIndex        =   16
         Top             =   120
         Width           =   465
      End
   End
   Begin VB.PictureBox picEdit 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   435
      Left            =   5760
      MouseIcon       =   "Notes.frx":15868
      MousePointer    =   99  'Custom
      Picture         =   "Notes.frx":159BA
      ScaleHeight     =   435
      ScaleWidth      =   1515
      TabIndex        =   5
      Top             =   2160
      Width           =   1515
      Begin VB.Label lblEdit 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Edit Notes"
         Height          =   210
         Left            =   405
         MouseIcon       =   "Notes.frx":17C6C
         MousePointer    =   99  'Custom
         TabIndex        =   14
         Top             =   120
         Width           =   735
      End
   End
   Begin VB.PictureBox picNew 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   435
      Left            =   5760
      MouseIcon       =   "Notes.frx":17DBE
      MousePointer    =   99  'Custom
      Picture         =   "Notes.frx":17F10
      ScaleHeight     =   435
      ScaleWidth      =   1515
      TabIndex        =   4
      Top             =   1560
      Width           =   1515
      Begin VB.Label lblNew 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&New Notes"
         Height          =   210
         Left            =   360
         MouseIcon       =   "Notes.frx":1A1C2
         MousePointer    =   99  'Custom
         TabIndex        =   10
         Top             =   120
         Width           =   810
      End
   End
   Begin VB.TextBox txtTitle 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   600
      Locked          =   -1  'True
      MaxLength       =   100
      MousePointer    =   3  'I-Beam
      TabIndex        =   1
      Tag             =   "Title"
      Top             =   1200
      Width           =   4935
   End
   Begin VB.TextBox txtNotes 
      Appearance      =   0  'Flat
      Height          =   3495
      Left            =   600
      Locked          =   -1  'True
      MousePointer    =   3  'I-Beam
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Tag             =   "Notes"
      Top             =   1605
      Width           =   4935
   End
   Begin VB.PictureBox picCancel 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   435
      Left            =   5760
      MouseIcon       =   "Notes.frx":1A314
      MousePointer    =   99  'Custom
      Picture         =   "Notes.frx":1A466
      ScaleHeight     =   435
      ScaleWidth      =   1515
      TabIndex        =   11
      Top             =   2400
      Visible         =   0   'False
      Width           =   1515
      Begin VB.Label lblCancel 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Cancel"
         Height          =   210
         Left            =   510
         MouseIcon       =   "Notes.frx":1C718
         MousePointer    =   99  'Custom
         TabIndex        =   15
         Top             =   120
         Width           =   525
      End
   End
   Begin VB.PictureBox picSave 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   435
      Left            =   5760
      MouseIcon       =   "Notes.frx":1C86A
      MousePointer    =   99  'Custom
      Picture         =   "Notes.frx":1C9BC
      ScaleHeight     =   435
      ScaleWidth      =   1515
      TabIndex        =   12
      Top             =   1800
      Visible         =   0   'False
      Width           =   1515
      Begin VB.Label lblSave 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Save"
         Height          =   210
         Left            =   570
         MouseIcon       =   "Notes.frx":1EC6E
         MousePointer    =   99  'Custom
         TabIndex        =   13
         Top             =   120
         Width           =   405
      End
   End
   Begin VB.Label lblNotes 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Notes"
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
      TabIndex        =   3
      Top             =   120
      Width           =   480
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Title:"
      Height          =   210
      Left            =   600
      TabIndex        =   2
      Top             =   960
      Width           =   330
   End
End
Attribute VB_Name = "Notes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

With Data1
    
    .DatabaseName = App.Path & "\db.mdb"
    .RecordSource = _
        "SELECT * FROM Notes WHERE UserName ='" & _
        Main.Tag & "' ORDER BY Date,Title"
    .Refresh
        
    Call Check_Record
    
    With .Recordset
    
        If .RecordCount > 0 Then Call Display_Record
    
    End With
    
End With

End Sub

Sub Display_Error(ctrl As Control)

msg = "Please fill-up the following:" & vbCrLf & ctrl.Tag
buttons = vbInformation
titles = "Message"

MsgBox msg, buttons, titles
ctrl.SetFocus

End Sub

Sub Prev_Click()

With Data1.Recordset
    
    .MovePrevious
    If .BOF = True Then .MoveFirst
    
    Call Display_Record
    
End With

End Sub

Sub Next_Click()

With Data1.Recordset
    
    .MoveNext
    If .EOF = True Then .MoveLast
    
    Call Display_Record
    
End With

End Sub

Sub New_Click()

Call Hide_Buttons
picPrev.Enabled = False
picNext.Enabled = False

lblNotes.Caption = "Notes - Date: " & Date
Me.Tag = "New"

txtTitle.Text = Empty
txtNotes.Text = Empty

txtTitle.SetFocus

End Sub

Sub Edit_Click()

With Data1.Recordset

    If .RecordCount = 0 Then
        MsgBox "Notes is empty.", vbExclamation, "Message"
        Exit Sub
    End If

End With

Call Hide_Buttons
picPrev.Enabled = False
picNext.Enabled = False

Me.Tag = "Edit"
Call Display_Record

txtTitle.SetFocus

End Sub

Sub Delete_Click()

With Data1.Recordset

    If .RecordCount = 0 Then
        MsgBox "Notes is empty.", vbExclamation, "Message"
        Exit Sub
    End If

    NoteTitle = .Fields("Title")
    ans = MsgBox("Are you sure?", vbQuestion + vbYesNo, _
            "Delete " & NoteTitle & " ?")
    
    If ans = vbNo Then Exit Sub
    
    .Delete
    MsgBox NoteTitle & " has been deleted.", , "Message"
    
    If .RecordCount <> 0 Then .MoveFirst
    
End With

Data1.Refresh
Call Check_Record
Call Display_Record

End Sub

Sub Save_Click()

If txtTitle.Text = Empty Then
    Call Display_Error(txtTitle)
    Exit Sub
ElseIf txtNotes.Text = Empty Then
    Call Display_Error(txtNotes)
    Exit Sub
End If

With Data1.Recordset

    If Me.Tag = "New" Then
        .AddNew
        .Fields("UserName") = Main.Tag
        lblNotes.Tag = Date
    ElseIf Me.Tag = "Edit" Then
        .Edit
    End If
    
    .Fields("Title") = txtTitle.Text
    .Fields("Date") = lblNotes.Tag
    .Fields("Notes") = txtNotes.Text
    .Update
    
    MsgBox "Notes has been saved.", vbInformation, "Message"

End With

Data1.Refresh

Call Hide_Buttons
Call Check_Record
Call Display_Record

Me.Tag = Empty

End Sub

Sub Cancel_Click()

Call Hide_Buttons
Call Check_Record
Call Display_Record

Me.Tag = Empty

End Sub

Sub Hide_Buttons()

picNew.Visible = Not picNew.Visible
picEdit.Visible = Not picEdit.Visible
picDelete.Visible = Not picDelete.Visible
picSave.Visible = Not picSave.Visible
picCancel.Visible = Not picCancel.Visible
picClose.Visible = Not picClose.Visible

txtTitle.Locked = Not txtTitle.Locked
txtNotes.Locked = Not txtNotes.Locked

picNext.Visible = Not picNext.Visible
picPrev.Visible = Not picPrev.Visible

End Sub

Sub Display_Record()

With Data1.Recordset
    
    If .RecordCount = 0 Then
        lblNotes.Caption = "Notes"
        lblNotes.Tag = Empty
        txtTitle.Text = Empty
        txtNotes.Text = Empty
        Exit Sub
    End If

    lblNotes.Caption = "Notes - Date: " & .Fields("Date")
    lblNotes.Tag = .Fields("Date")
    txtTitle.Text = .Fields("Title")
    txtNotes.Text = .Fields("Notes")

End With

End Sub

Sub Check_Record()
    
With Data1.Recordset
    
    If .RecordCount = 0 Then
                        
        picEdit.Enabled = False
        picDelete.Enabled = False
        picPrev.Enabled = False
        picNext.Enabled = False
    
    Else
                        
        picEdit.Enabled = True
        picDelete.Enabled = True
        picPrev.Enabled = True
        picNext.Enabled = True
    
    End If
    
End With


End Sub

Private Sub lblCancel_Click()

Cancel_Click

End Sub

Private Sub lblClose_Click()

Unload Me

End Sub

Private Sub lblDelete_Click()

Delete_Click

End Sub

Private Sub lblEdit_Click()

Edit_Click

End Sub

Private Sub lblNew_Click()

New_Click

End Sub

Private Sub lblNext_Click()

Next_Click

End Sub

Private Sub lblPrev_Click()

Prev_Click

End Sub

Private Sub lblSave_Click()

Save_Click

End Sub

Private Sub picCancel_Click()

Cancel_Click

End Sub

Private Sub picClose_Click()

Unload Me

End Sub

Private Sub picDelete_Click()

Delete_Click

End Sub

Private Sub picEdit_Click()

Edit_Click

End Sub

Private Sub picNew_Click()

New_Click

End Sub

Private Sub picNext_Click()

Next_Click

End Sub

Private Sub picPrev_Click()

Prev_Click

End Sub

Private Sub picSave_Click()

Save_Click

End Sub


