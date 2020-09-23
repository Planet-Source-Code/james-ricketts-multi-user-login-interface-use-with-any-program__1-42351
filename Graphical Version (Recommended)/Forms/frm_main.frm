VERSION 5.00
Begin VB.Form frm_main 
   Caption         =   "Main Form"
   ClientHeight    =   3150
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   8670
   Icon            =   "frm_main.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3150
   ScaleWidth      =   8670
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   2535
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Text            =   "frm_main.frx":0442
      Top             =   480
      Width           =   8415
   End
   Begin VB.Label Label1 
      Caption         =   "Developed and Programmed by James Ricketts"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2400
      TabIndex        =   0
      Top             =   120
      Width           =   4695
   End
   Begin VB.Menu mnu_file 
      Caption         =   "File"
      Begin VB.Menu mnu_logoff 
         Caption         =   "Log Off"
      End
   End
   Begin VB.Menu mnu_settings 
      Caption         =   "Settings"
      Begin VB.Menu mnu_users 
         Caption         =   "Users"
      End
   End
End
Attribute VB_Name = "frm_main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
UpdateUser
End Sub

Private Sub Form_Terminate()
End
End Sub

Private Sub Form_Unload(Cancel As Integer)
If ResetAllUsers = False Then
    End
Else:
    ResetAllUsers = False
End If
End Sub

Private Sub mnu_logoff_Click()
Unload frm_useredit         '
Unload frm_users

'unloading main form = END (so dont!)
frm_main.Hide
frm_passwordscreen.txt_user.Text = ""
frm_passwordscreen.txt_password.Text = ""
frm_passwordscreen.Show 1
End Sub

Private Sub mnu_users_Click()
frm_users.Show
End Sub
Public Sub UpdateUser()

frm_main.Caption = App.CompanyName & " " & App.ProductName & " v" & App.Major & "." & App.Minor & "." & App.Revision & _
" -   Logged in: " & StrConv(LoggedIn, vbProperCase)

'''Add code for the change to the program that you wish to occur when different users login.
'''The data for the changes can be stored inside the login data file.

End Sub
