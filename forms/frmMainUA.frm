VERSION 5.00
Begin VB.Form frmMainUA 
   Caption         =   "User Access Module"
   ClientHeight    =   7020
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   7980
   BeginProperty Font 
      Name            =   "Myriad Web"
      Size            =   7.5
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMainUA.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "frmMainUA.frx":F172
   ScaleHeight     =   7020
   ScaleWidth      =   7980
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuLogOut 
         Caption         =   "Log Out"
      End
      Begin VB.Menu mnuClose 
         Caption         =   "Close"
      End
   End
   Begin VB.Menu mnuUserAccess 
      Caption         =   "Set Up Users and Roles"
      Begin VB.Menu mnuRoles 
         Caption         =   "Roles"
      End
      Begin VB.Menu mnuUsers 
         Caption         =   "Users"
      End
      Begin VB.Menu mnuChangePass 
         Caption         =   "Change Password"
      End
   End
End
Attribute VB_Name = "frmMainUA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
'''''''Modified by Mahboob 06/06/2023 Change ID 16 work item 1 Stop visible change password mnuChangePass
mnuChangePass.Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
''''''''''''''''Modified by Mahboob 03/04/2023 Change ID 1 work item 13 unload other form
If IsLoadedAndVisible("frmChangePasswordNew") Then Unload frmChangePasswordNew
          If IsLoadedAndVisible("frmRoles") Then Unload frmRoles
          If IsLoadedAndVisible("frmUsersNew") Then Unload frmUsersNew
     'Code added by anol 2023-08-26
    Set Form = Nothing
End Sub

Private Sub mnuChangePass_Click()
''''''''''''''''Modified by Mahboob 03/04/2023 Change ID 4 work item 5 load change pass form
frmChangePasswordNew.Show vbModeless, frmMainUA
End Sub

Private Sub mnuClose_Click()
''''''''''''''''Modified by Mahboob 03/04/2023 Change ID 1 work item 11 unload the user moudle
 If MsgBox("Are you sure you wish to close?", vbYesNo + vbQuestion, "Close") = vbNo Then
       Exit Sub
   Else
      If IsLoadedAndVisible("frmChangePasswordNew") Then Unload frmChangePasswordNew
          If IsLoadedAndVisible("frmRoles") Then Unload frmRoles
          If IsLoadedAndVisible("frmUsersNew") Then Unload frmUsersNew
'      Unload frmMainUA
End
   End If

End Sub

Private Sub mnuLogOut_Click()
''''''''''''''''Modified by Mahboob 03/04/2023 Change ID 9 work item 13 logout
  If MsgBox("Are you sure you wish to log out?", vbYesNo + vbQuestion, "Log Out") = vbNo Then
       Exit Sub
   Else
      If IsLoadedAndVisible("frmChangePasswordNew") Then Unload frmChangePasswordNew
          If IsLoadedAndVisible("frmRoles") Then Unload frmRoles
          If IsLoadedAndVisible("frmUsersNew") Then Unload frmUsersNew
      Load frmLogin2
'      LastTenBackup
      Unload frmMainUA
      frmLogin2.Show
      frmLogin2.txtUserName.text = ""
      frmLogin2.txtPassword.text = ""
   End If
End Sub

Private Sub mnuRoles_Click()
''''''''''''''''Modified by Mahboob 03/04/2023 Change ID 2 work item 9 load role form
frmRoles.Show vbModeless, frmMainUA
End Sub

Private Sub mnuUsers_Click()
''''''''''''''''Modified by Mahboob 03/04/2023 Change ID 3 work item 10 User form load
frmUsersNew.Show vbModeless, frmMainUA
End Sub
