VERSION 5.00
Begin VB.Form Profile 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5625
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   7245
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
   MouseIcon       =   "Profile.frx":0000
   MousePointer    =   99  'Custom
   Picture         =   "Profile.frx":030A
   ScaleHeight     =   5625
   ScaleWidth      =   7245
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picChange 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000004&
      FillColor       =   &H80000016&
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   3705
      Left            =   960
      ScaleHeight     =   3675
      ScaleWidth      =   4860
      TabIndex        =   21
      Top             =   960
      Visible         =   0   'False
      Width           =   4890
      Begin VB.PictureBox picSavePass 
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
         Left            =   600
         MouseIcon       =   "Profile.frx":C310
         MousePointer    =   99  'Custom
         Picture         =   "Profile.frx":C462
         ScaleHeight     =   435
         ScaleWidth      =   1725
         TabIndex        =   27
         Top             =   2760
         Width           =   1725
         Begin VB.Label lblSavePass 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Save Password"
            Height          =   210
            Left            =   255
            MouseIcon       =   "Profile.frx":EC10
            MousePointer    =   99  'Custom
            TabIndex        =   28
            Top             =   120
            Width           =   1185
         End
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
         Left            =   2400
         MouseIcon       =   "Profile.frx":ED62
         MousePointer    =   99  'Custom
         Picture         =   "Profile.frx":EEB4
         ScaleHeight     =   435
         ScaleWidth      =   1725
         TabIndex        =   25
         Top             =   2760
         Width           =   1725
         Begin VB.Label lblCancel 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cancel"
            Height          =   210
            Left            =   585
            MouseIcon       =   "Profile.frx":11662
            MousePointer    =   99  'Custom
            TabIndex        =   26
            Top             =   120
            Width           =   525
         End
      End
      Begin VB.TextBox txtOld 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   2160
         MaxLength       =   15
         PasswordChar    =   "*"
         TabIndex        =   24
         Tag             =   "Old Password"
         Top             =   1200
         Width           =   1815
      End
      Begin VB.TextBox txtNew 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   2160
         MaxLength       =   15
         PasswordChar    =   "*"
         TabIndex        =   23
         Tag             =   "New Password"
         Top             =   1680
         Width           =   1815
      End
      Begin VB.TextBox txtConfirm 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   2160
         MaxLength       =   15
         PasswordChar    =   "*"
         TabIndex        =   22
         Tag             =   "Confirm Password"
         Top             =   2160
         Width           =   1815
      End
      Begin VB.Label lblPass 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Old Password:"
         Height          =   210
         Left            =   870
         TabIndex        =   32
         Top             =   1275
         Width           =   1080
      End
      Begin VB.Label lblConfirm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "New Password:"
         Height          =   210
         Left            =   780
         TabIndex        =   31
         Top             =   1755
         Width           =   1185
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Confirm Password:"
         Height          =   210
         Left            =   480
         TabIndex        =   30
         Top             =   2235
         Width           =   1395
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Change Password"
         Height          =   210
         Left            =   120
         TabIndex        =   29
         Top             =   120
         Width           =   1350
      End
   End
   Begin VB.PictureBox picProfile 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5535
      Left            =   0
      Picture         =   "Profile.frx":117B4
      ScaleHeight     =   5535
      ScaleWidth      =   7455
      TabIndex        =   5
      Top             =   0
      Width           =   7455
      Begin VB.PictureBox picSaveProfile 
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
         Left            =   1680
         MouseIcon       =   "Profile.frx":1D7BA
         MousePointer    =   99  'Custom
         Picture         =   "Profile.frx":1D90C
         ScaleHeight     =   435
         ScaleWidth      =   1725
         TabIndex        =   19
         Top             =   4080
         Visible         =   0   'False
         Width           =   1725
         Begin VB.Label lblSaveProfile 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Save Profile"
            Height          =   210
            Left            =   405
            MouseIcon       =   "Profile.frx":200BA
            MousePointer    =   99  'Custom
            TabIndex        =   20
            Top             =   120
            Width           =   885
         End
      End
      Begin VB.PictureBox picCancelChange 
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
         Left            =   3480
         MouseIcon       =   "Profile.frx":2020C
         MousePointer    =   99  'Custom
         Picture         =   "Profile.frx":2035E
         ScaleHeight     =   435
         ScaleWidth      =   1725
         TabIndex        =   17
         Top             =   4080
         Visible         =   0   'False
         Width           =   1725
         Begin VB.Label lblCancelChange 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cancel  Changes"
            Height          =   210
            Left            =   225
            MouseIcon       =   "Profile.frx":22B0C
            MousePointer    =   99  'Custom
            TabIndex        =   18
            Top             =   120
            Width           =   1245
         End
      End
      Begin VB.Data dbUsers 
         Caption         =   "Users"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2280
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   4800
         Visible         =   0   'False
         Width           =   2340
      End
      Begin VB.PictureBox picClose 
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
         Left            =   4320
         MouseIcon       =   "Profile.frx":22C5E
         MousePointer    =   99  'Custom
         Picture         =   "Profile.frx":22DB0
         ScaleHeight     =   435
         ScaleWidth      =   1725
         TabIndex        =   15
         Top             =   4080
         Width           =   1725
         Begin VB.Label lblClose 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Close"
            Height          =   210
            Left            =   630
            MouseIcon       =   "Profile.frx":2555E
            MousePointer    =   99  'Custom
            TabIndex        =   16
            Top             =   120
            Width           =   435
         End
      End
      Begin VB.PictureBox picEditProfile 
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
         Left            =   720
         MouseIcon       =   "Profile.frx":256B0
         MousePointer    =   99  'Custom
         Picture         =   "Profile.frx":25802
         ScaleHeight     =   435
         ScaleWidth      =   1725
         TabIndex        =   13
         Top             =   4080
         Width           =   1725
         Begin VB.Label lblEditProfile 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Edit Profile"
            Height          =   210
            Left            =   465
            MouseIcon       =   "Profile.frx":27FB0
            MousePointer    =   99  'Custom
            TabIndex        =   14
            Top             =   120
            Width           =   765
         End
      End
      Begin VB.PictureBox picChangePass 
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
         Left            =   2520
         MouseIcon       =   "Profile.frx":28102
         MousePointer    =   99  'Custom
         Picture         =   "Profile.frx":28254
         ScaleHeight     =   435
         ScaleWidth      =   1725
         TabIndex        =   11
         Top             =   4080
         Width           =   1725
         Begin VB.Label lblChangePass 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Change &Password"
            Height          =   210
            Left            =   165
            MouseIcon       =   "Profile.frx":2AA02
            MousePointer    =   99  'Custom
            TabIndex        =   12
            Top             =   120
            Width           =   1365
         End
      End
      Begin VB.TextBox txtMobile 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   2400
         Locked          =   -1  'True
         MaxLength       =   20
         MousePointer    =   3  'I-Beam
         TabIndex        =   4
         Tag             =   "Mobile"
         Top             =   3120
         Width           =   2295
      End
      Begin VB.TextBox txtPhone 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   2415
         Locked          =   -1  'True
         MaxLength       =   20
         MousePointer    =   3  'I-Beam
         TabIndex        =   3
         Tag             =   "Home Phone"
         Top             =   2640
         Width           =   2295
      End
      Begin VB.TextBox txtAddress 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   2400
         Locked          =   -1  'True
         MousePointer    =   3  'I-Beam
         TabIndex        =   2
         Tag             =   "Home Address"
         Top             =   2160
         Width           =   3855
      End
      Begin VB.TextBox txtName 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   2400
         Locked          =   -1  'True
         MaxLength       =   50
         MousePointer    =   3  'I-Beam
         TabIndex        =   1
         Tag             =   "Name"
         Top             =   1680
         Width           =   3015
      End
      Begin VB.TextBox txtUser 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   2400
         Locked          =   -1  'True
         MaxLength       =   20
         MousePointer    =   3  'I-Beam
         TabIndex        =   0
         Tag             =   "User Name"
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label lblProfile 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Profile"
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
         TabIndex        =   33
         Top             =   120
         Width           =   540
      End
      Begin VB.Label lblMobile 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mobile:"
         Height          =   210
         Left            =   1710
         TabIndex        =   10
         Top             =   3195
         Width           =   495
      End
      Begin VB.Label lblPhone 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Home Phone:"
         Height          =   210
         Left            =   1260
         TabIndex        =   9
         Top             =   2715
         Width           =   945
      End
      Begin VB.Label lblAdd 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Home Address:"
         Height          =   210
         Left            =   1080
         TabIndex        =   8
         Top             =   2235
         Width           =   1125
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name:"
         Height          =   210
         Left            =   1755
         TabIndex        =   7
         Top             =   1755
         Width           =   450
      End
      Begin VB.Label lblUser 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "User Name:"
         Height          =   210
         Left            =   1365
         TabIndex        =   6
         Top             =   1275
         Width           =   840
      End
   End
End
Attribute VB_Name = "Profile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()

With dbUsers

    .DatabaseName = App.Path & "\db.mdb"
    .RecordSource = _
        "SELECT * FROM Users WHERE User='" & Main.Tag & "'"
    .Refresh
    
    Call Display_Record

End With

End Sub

Private Sub lblCancel_Click()

picProfile.Enabled = True
picChange.Visible = False
txtUser.SetFocus

End Sub

Private Sub lblCancelChange_Click()

Form_Controls
txtUser.SetFocus

End Sub

Private Sub lblChangePass_Click()

picProfile.Enabled = False
picChange.Visible = True

Load ChangePass
ChangePass.Show , Me

End Sub

Private Sub lblClose_Click()

Unload Me

End Sub

Private Sub lblEditProfile_Click()

Form_Controls
txtAddress.SetFocus

End Sub

Private Sub lblSavePass_Click()

Call Save_Pass

End Sub

Private Sub lblSaveProfile_Click()

Call SavePro

End Sub

Private Sub picCancel_Click()

picProfile.Enabled = True
picChange.Visible = False
txtUser.SetFocus

End Sub

Private Sub picCancelChange_Click()

Form_Controls
txtUser.SetFocus

End Sub

Private Sub picChangePass_Click()

picProfile.Enabled = False
picChange.Visible = True

txtOld.Text = ""
txtNew.Text = ""
txtConfirm.Text = ""

txtOld.SetFocus

End Sub

Private Sub picClose_Click()

Unload Me

End Sub

Private Sub picEditProfile_Click()

Form_Controls
txtAddress.SetFocus

End Sub

Private Sub picSavePass_Click()

Call SavePro

End Sub

Private Sub picSaveProfile_Click()

Call SavePro
picCancel_Click

End Sub



Private Sub txtAddress_KeyPress(KeyAscii As Integer)

Select Case KeyAscii

    Case 97 To 122                      'a-z
        KeyAscii = KeyAscii - 32        'Convert to capital letters
    Case vbKeyReturn
        txtPhone.SetFocus

End Select

End Sub

Private Sub txtAddress_LostFocus()

If txtAddress.Text = Empty Then
    txtAddress.Text = "N/A"
End If

End Sub



Private Sub txtConfirm_KeyPress(KeyAscii As Integer)

If KeyAscii = vbKeyReturn Then Call Save_Pass

End Sub



Private Sub txtMobile_KeyPress(KeyAscii As Integer)

Select Case KeyAscii

    Case 45                             'Hyphen (-)
    Case 48 To 57                       '0-9
    Case vbKeyBack                      'Backspace key
    Case vbKeySpace                     'Spacebar key
    Case vbKeyReturn                    'Enter key
        Call SavePro
    Case Else
        KeyAscii = 0                    'Display nothing
    
End Select

End Sub

Private Sub txtMobile_LostFocus()

If txtMobile.Text = Empty Then
    txtMobile.Text = "N/A"
End If

End Sub



Private Sub txtNew_KeyPress(KeyAscii As Integer)

If KeyAscii = vbKeyReturn Then
    txtConfirm.SetFocus
End If

End Sub

Private Sub txtOld_KeyPress(KeyAscii As Integer)

If KeyAscii = vbKeyReturn Then
    txtNew.SetFocus
End If

End Sub


Private Sub txtPhone_KeyPress(KeyAscii As Integer)

Select Case KeyAscii

    Case 45                             'Hyphen (-)
    Case 48 To 57                       '0-9
    Case vbKeyBack                      'Backspace key
    Case vbKeySpace                     'Spacebar key
    Case vbKeyReturn                    'Enter key
        txtMobile.SetFocus             'Set focus to txtPager
    Case Else
        KeyAscii = 0                    'Display nothing
    
End Select

End Sub

Private Sub txtPhone_LostFocus()

If txtPhone.Text = Empty Then
    txtPhone.Text = "N/A"
End If

End Sub

Sub Form_Controls()

txtAddress.Locked = Not txtAddress.Locked
txtPhone.Locked = Not txtPhone.Locked
txtMobile.Locked = Not txtMobile.Locked

picEditProfile.Visible = Not picEditProfile.Visible
picChangePass.Visible = Not picChangePass.Visible
picClose.Visible = Not picClose.Visible

picSaveProfile.Visible = Not picSaveProfile.Visible
picCancelChange.Visible = Not picCancelChange.Visible

End Sub

Sub SavePro()

If txtAddress.Text = Empty Then txtAddress.Text = "N/A"
If txtPhone.Text = Empty Then txtPhone.Text = "N/A"
If txtMobile.Text = Empty Then txtMobile.Text = "N/A"

With dbUsers.Recordset

    .Edit

    .Fields("Address") = txtAddress.Text
    .Fields("Home Phone") = txtPhone.Text
    .Fields("Mobile") = txtMobile.Text
    
    .Update
    
    MsgBox "Profile has been changed.", vbInformation, "Message"

End With

Call Form_Controls
txtUser.SetFocus

End Sub

Sub Display_Record()

With dbUsers.Recordset

    txtUser.Text = .Fields("User")
    txtName.Text = .Fields("Name")
    txtAddress.Text = .Fields("Address")
    txtPhone.Text = .Fields("Home Phone")
    txtMobile.Text = .Fields("Mobile")

    picChange.Tag = .Fields("Password")

End With

End Sub

Sub Save_Pass()

If txtOld.Text = Empty Then
    MsgBox "Please enter your old password.", vbExclamation, "Message"
    txtOld.SetFocus
    Exit Sub
ElseIf txtOld.Text <> picChange.Tag Then
    MsgBox "Old password is incorrect.", vbExclamation, "Message"
    txtOld.SetFocus
    Exit Sub
ElseIf txtNew.Text = Empty Then
    MsgBox "Please enter your new password.", vbExclamation, "Message"
    txtOld.SetFocus
    Exit Sub
ElseIf txtConfirm.Text = Empty Then
    MsgBox "Please enter your confirm password.", vbExclamation, "Message"
    txtConfirm.SetFocus
    Exit Sub
ElseIf txtNew.Text <> txtConfirm Then
    MsgBox "New and confirm password must be the same.", _
        vbExclamation, "Message"
    txtNew.SetFocus
    Exit Sub
End If

With dbUsers.Recordset

    .Edit
    
    .Fields("Password") = txtNew.Text
    .Update
    
    MsgBox "Password has been changed."

End With

picCancel_Click

End Sub






