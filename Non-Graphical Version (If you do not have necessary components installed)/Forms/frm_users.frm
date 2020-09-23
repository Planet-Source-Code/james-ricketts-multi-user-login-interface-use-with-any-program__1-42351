VERSION 5.00
Begin VB.Form frm_users 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Users"
   ClientHeight    =   4500
   ClientLeft      =   150
   ClientTop       =   840
   ClientWidth     =   1455
   Icon            =   "frm_users.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4500
   ScaleWidth      =   1455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4305
      ItemData        =   "frm_users.frx":0442
      Left            =   0
      List            =   "frm_users.frx":0444
      TabIndex        =   0
      Top             =   0
      Width           =   1455
   End
   Begin VB.Menu mnu_users 
      Caption         =   "&Options"
      NegotiatePosition=   1  'Left
      Begin VB.Menu mnu_user_new 
         Caption         =   "New User"
      End
      Begin VB.Menu mnu_bar 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_user_reset 
         Caption         =   "Reset All Users"
      End
   End
   Begin VB.Menu mnu_popup 
      Caption         =   "Popup"
      Visible         =   0   'False
      Begin VB.Menu mnu_popup_edit 
         Caption         =   "Edit User"
      End
      Begin VB.Menu mnu_popup_delete 
         Caption         =   "Delete User"
      End
   End
End
Attribute VB_Name = "frm_users"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim holduser As String

Private Sub List1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 And List1.ListCount <> 0 Then
    Me.PopupMenu mnu_popup
End If
End Sub

Private Sub ListView1_BeforeLabelEdit(Cancel As Integer)

End Sub

Private Sub mnu_popup_delete_Click()
Dim UserSelected As String

If List1.ListIndex = "-1" Then
    result = MsgBox("You must select a user first", vbExclamation, "Select user to delete")
    Exit Sub
End If

'Check that the person using the program is the one who logged into it
PasswordCheckCompleted = False
        
frm_passcheck.Show 1
If PasswordCheckCompleted = False Then
    Exit Sub
Else:
    PasswordCheckCompleted = False
End If

UserSelected = List1.List(List1.ListIndex)

If Trim(LCase(UserSelected)) = Trim(LCase(LoggedIn)) Then
    MsgBox "You cannot delete this user since this user is currently logged into the program", vbCritical, "Error..."
    Exit Sub
End If

If MsgBox("Are you sure you wish to delete the user " & holduser & "?", vbYesNo, "Are you sure?") = vbNo Then
    Exit Sub
End If
'''Checks completed

ReturnedValue = DeleteRecord(UserSelected)

If ReturnedValue = 0 Then   'Username does not exist (ERROR)
    MsgBox "Error, the username does not exist, the data file has been tampered with.", vbCritical, "Error..."
ElseIf ReturnedValue = 1 Then   'Completed
    List1.RemoveItem (List1.ListIndex)
    Call Refreshlist
End If

End Sub

Private Sub mnu_popup_edit_Click()
If List1.ListIndex = "-1" Then
    MsgBox "You must select a user first", vbExclamation, "Select user to edit"
    Exit Sub
End If

UserNameBuffer = List1.List(List1.ListIndex)
frm_useredit.Caption = "Edit User - " & UserNameBuffer
frm_useredit.txt_username.Text = UserNameBuffer
frm_useredit.Show 1
End Sub

Private Sub mnu_user_new_Click()

'Check that the person using the program is the one who logged into it
PasswordCheckCompleted = False
        
frm_passcheck.Show 1
If PasswordCheckCompleted = False Then
    Exit Sub
Else:
    PasswordCheckCompleted = False
    frm_newpassword.Show 1
End If
End Sub

Private Sub mnu_user_reset_Click()

'Check that the person using the program is the one who logged into it
PasswordCheckCompleted = False
        
frm_passcheck.Show 1
If PasswordCheckCompleted = False Then
    Exit Sub
Else:
    PasswordCheckCompleted = False
End If
 

Beep
If MsgBox(("WARNING..." & vbCrLf & vbCrLf & "This will delete all usernames stored by this program." & vbCrLf & _
vbCrLf & "To be able to create a new user you will need yo know the Authorisation Password supplied with this software" & _
vbCrLf & "Are you sure you want to continue?"), vbYesNo, "Warning...") = vbYes Then
    
    ResetUserList
    Unload Me
    Unload frm_useredit
    ResetAllUsers = True
    Unload frm_main
    frm_passwordscreen.txt_user.Text = ""
    frm_passwordscreen.txt_password.Text = ""
    frm_passwordscreen.Show 1

Else:
    Exit Sub
End If

End Sub

Private Sub Command5_Click()
Unload Me
End Sub

Private Sub Form_Load()
List1.Height = Me.Height
List1.Top = 0

Call Refreshlist

'position the users form to the left of the Main form
Me.Left = frm_main.Left - Me.Width
Me.Top = frm_main.Top
End Sub


Public Sub Refreshlist()
Dim HoldUserName As String

List1.Clear

HoldRecordCount = RecordCount

Open Filename For Random As #1 Len = UserDataLength

For i = 1 To HoldRecordCount
    Get #1, i, HOLDUSERDATA
    Decrypt HOLDUSERDATA.UserName
    HoldUserName = Trim(HOLDUSERDATA.UserName)
    List1.AddItem StrConv(HoldUserName, vbProperCase)
Next

Close #1

End Sub

Private Sub Form_Terminate()
Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Me
End Sub

