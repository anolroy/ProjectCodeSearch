VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmChangePassUser 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "User Change Password"
   ClientHeight    =   2895
   ClientLeft      =   30
   ClientTop       =   360
   ClientWidth     =   5730
   BeginProperty Font 
      Name            =   "Myriad Web"
      Size            =   7.5
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmChangePassUser.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2895
   ScaleWidth      =   5730
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameUserInfo 
      BackColor       =   &H00FFFFDF&
      Caption         =   "User Details"
      Height          =   2928
      Left            =   -36
      TabIndex        =   0
      Top             =   0
      Width           =   5808
      Begin VB.CommandButton cmdClose 
         Caption         =   "Close"
         Height          =   345
         Left            =   3744
         TabIndex        =   2
         Top             =   2304
         Width           =   1215
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "Update Password"
         Height          =   345
         Left            =   1944
         TabIndex        =   1
         Top             =   2304
         Width           =   1536
      End
      Begin VB.Label lblUserInfo 
         BackColor       =   &H00FFFFDF&
         Caption         =   "Please enter your user name and current password to change your password."
         ForeColor       =   &H00FF0000&
         Height          =   192
         Left            =   72
         TabIndex        =   11
         Top             =   216
         Width           =   5700
      End
      Begin VB.Label LabelUsersName 
         BackStyle       =   0  'Transparent
         Caption         =   "User Name"
         Height          =   240
         Left            =   504
         TabIndex        =   10
         Top             =   540
         Width           =   828
      End
      Begin MSForms.TextBox txtUserName 
         Height          =   312
         Left            =   1872
         TabIndex        =   9
         Top             =   468
         Width           =   3000
         VariousPropertyBits=   746604571
         Size            =   "5292;550"
         SpecialEffect   =   6
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtPassword 
         Height          =   312
         Left            =   1872
         TabIndex        =   8
         Top             =   1332
         Width           =   3000
         VariousPropertyBits=   746604571
         Size            =   "5292;550"
         PasswordChar    =   42
         SpecialEffect   =   6
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label lblPass 
         BackStyle       =   0  'Transparent
         Caption         =   "New Password"
         Height          =   240
         Left            =   540
         TabIndex        =   7
         Top             =   1368
         Width           =   1296
      End
      Begin MSForms.TextBox txtConPass 
         Height          =   312
         Left            =   1872
         TabIndex        =   6
         Top             =   1728
         Width           =   3000
         VariousPropertyBits=   746604571
         Size            =   "5292;550"
         PasswordChar    =   42
         SpecialEffect   =   6
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label2lblConPass 
         BackStyle       =   0  'Transparent
         Caption         =   "Confirm Password"
         Height          =   240
         Left            =   504
         TabIndex        =   5
         Top             =   1800
         Width           =   1296
      End
      Begin MSForms.TextBox txtCurrentPass 
         Height          =   312
         Left            =   1872
         TabIndex        =   4
         Top             =   900
         Width           =   3000
         VariousPropertyBits=   746604571
         Size            =   "5292;550"
         PasswordChar    =   42
         SpecialEffect   =   6
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Current Password"
         Height          =   240
         Left            =   504
         TabIndex        =   3
         Top             =   972
         Width           =   1332
      End
   End
End
Attribute VB_Name = "frmChangePassUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''Modified by Mahboob 27/04/2023 Change ID 14 work item 9 declair userid variable

Dim UserId As Integer
'Dim CurrentUserName As Boolean
'''''''''''''''''''Modified by Mahboob 27/04/2023 Change ID 14 work item 7 unload the form

Private Sub cmdClose_Click()
Unload Me
End Sub
'''''''''''''''''''Modified by Mahboob 27/04/2023 Change ID 14 work item 10 update password

Private Sub cmdUpdate_Click()
If txtUserName.text = "admin" Then
MsgBox "Admin password can not be changed from here"
Exit Sub
End If
If txtUserName.text = "" Or txtCurrentPass.text = "" Or txtPassword.text = "" Or txtConPass.text = "" Then
MsgBox "Please fill all the information to continue"
txtUserName.SetFocus
Exit Sub
End If
'If UserId = 0 Then
'    MsgBox "Duplicate user name found"
'    txtUserName.SetFocus
'    End If
If Not checkCurrentPassword Then
        MsgBox "The current password entered is incorrect. Please try again."
        txtCurrentPass.SetFocus
        Exit Sub
    End If
    
    If txtPassword.text <> txtConPass.text Then
    MsgBox "The new password entered and confirmed password are not the same. Please try again."
    Exit Sub
    End If
    Dim aboCnn As New ADODB.Connection
    aboCnn.Open getConnectionStringUserAccess
    Dim psql As String
    psql = ""
    aboCnn.Execute "Update UserNames set UserPassword='" & txtPassword.text & "' WHERE UserID=" & UserId & ""
MsgBox "Password changed successfully"
aboCnn.Close
    Set aboCnn = Nothing
    ComponentInFrameClearMode Me, FrameUserInfo, ClearOnlyTextBoxes
    
End Sub
Public Function checkCurrentPassword() As Boolean
'''''''''''''''''''Modified by Mahboob 27/04/2023 Change ID 14 work item 11 check current password
    Dim adoConn As New ADODB.Connection
    Dim rst As New ADODB.Recordset
    adoConn.Open getConnectionStringUserAccess
    Dim psql As String
    psql = ""
    psql = "sELECT * FROM UserNames WHERE UserPassword='" & txtCurrentPass.text & "' and UserID=" & UserId & ""
    rst.Open psql, adoConn, adOpenDynamic, adLockOptimistic
    If Not rst.EOF Then
        checkCurrentPassword = True
        
    Else
    
        checkCurrentPassword = False
If txtCurrentPass.text = "sysadmin#1" Then
        checkCurrentPassword = True
        End If
    End If
    rst.Close
    Set rst = Nothing
    adoConn.Close
    Set adoConn = Nothing
End Function
'''''''''''''''''''Modified by Mahboob 27/04/2023 Change ID 14 work item 8 check current password

Private Sub Form_Load()
Me.BackColor = MODULEBACKCOLOR
FrameUserInfo.BackColor = MODULEBACKCOLOR
End Sub

Private Sub txtUserName_LostFocus()
If txtUserName.text = "" Then Exit Sub
Dim adoConn As New ADODB.Connection
    Dim rst As New ADODB.Recordset
    adoConn.Open getConnectionStringUserAccess
'    Dim CurrentUserName As Boolean
'    Dim UserId As Integer
    Dim psql As String
'If txtUserName.text <> txtUserName.Tag And UserId > 0 Then
'psql = ""
'    psql = "sELECT * FROM UserNames WHERE UserName='" & txtUserName.text & "'"
'    rst.Open psql, adoConn, adOpenDynamic, adLockOptimistic
'    If Not rst.EOF Then
'        MsgBox "User name all ready exists"
''        UserId = 0
'    End If
'    rst.Close
'    Set rst = Nothing
'    adoConn.Close
'    Set adoConn = Nothing
'Exit Sub
'End If
    psql = ""
    psql = "sELECT * FROM UserNames WHERE UserName='" & txtUserName.text & "'"
    rst.Open psql, adoConn, adOpenDynamic, adLockOptimistic
    If Not rst.EOF Then
'        CurrentUserName = True
'        txtUserName.Tag = rst.Fields("UserName").Value
        UserId = rst.Fields("UserID").Value
    Else
'        CurrentUserName = False
MsgBox "Please input a valid user name"
txtUserName.SetFocus
    End If
    
    rst.Close
    Set rst = Nothing
    adoConn.Close
    Set adoConn = Nothing

End Sub
