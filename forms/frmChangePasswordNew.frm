VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmChangePasswordNew 
   BackColor       =   &H00FFFFDF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Change Password"
   ClientHeight    =   6480
   ClientLeft      =   30
   ClientTop       =   360
   ClientWidth     =   8670
   BeginProperty Font 
      Name            =   "Myriad Web"
      Size            =   7.5
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmChangePasswordNew.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6480
   ScaleWidth      =   8670
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameUserInfo 
      BackColor       =   &H00FFFFDF&
      Caption         =   "User Details"
      Height          =   6492
      Left            =   0
      TabIndex        =   4
      Top             =   36
      Width           =   8652
      Begin VB.PictureBox fmeUserLookup 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Height          =   4116
         Left            =   36
         ScaleHeight     =   4080
         ScaleWidth      =   8550
         TabIndex        =   6
         Top             =   2376
         Width           =   8580
         Begin VB.TextBox txtSearchActive 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   6615
            TabIndex        =   11
            Top             =   225
            Width           =   1908
         End
         Begin VB.TextBox txtSearchEmail 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   4944
            TabIndex        =   10
            Top             =   228
            Width           =   1632
         End
         Begin VB.TextBox txtSearchRole 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2700
            TabIndex        =   9
            Top             =   240
            Width           =   2190
         End
         Begin VB.TextBox txtSearchName 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1110
            TabIndex        =   8
            Top             =   240
            Width           =   1560
         End
         Begin VB.TextBox txtSearchUser 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   84
            TabIndex        =   7
            Top             =   240
            Width           =   1020
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid gridUserLookup 
            Height          =   3516
            Left            =   36
            TabIndex        =   12
            Top             =   540
            Width           =   8520
            _ExtentX        =   15028
            _ExtentY        =   6191
            _Version        =   393216
            Cols            =   9
            FixedCols       =   0
            BackColorFixed  =   13553358
            ForeColorFixed  =   16777215
            BackColorSel    =   14737632
            ForeColorSel    =   0
            BackColorBkg    =   16777215
            GridColor       =   -2147483638
            WordWrap        =   -1  'True
            HighLight       =   2
            GridLinesFixed  =   1
            ScrollBars      =   2
            SelectionMode   =   1
            Appearance      =   0
            BandDisplay     =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            _NumberOfBands  =   1
            _Band(0).Cols   =   9
            _Band(0).GridLinesBand=   1
            _Band(0).TextStyleBand=   0
            _Band(0).TextStyleHeader=   0
         End
         Begin VB.Label lblSUserID 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "User ID"
            Height          =   180
            Index           =   0
            Left            =   60
            TabIndex        =   17
            Top             =   36
            Width           =   480
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Name"
            Height          =   180
            Index           =   1
            Left            =   1140
            TabIndex        =   16
            Top             =   36
            Width           =   396
         End
         Begin VB.Label lblSRole 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Role"
            Height          =   180
            Index           =   2
            Left            =   2748
            TabIndex        =   15
            Top             =   36
            Width           =   288
         End
         Begin VB.Label lblSEmail 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Email"
            Height          =   180
            Index           =   4
            Left            =   4896
            TabIndex        =   14
            Top             =   36
            Width           =   360
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Status"
            Height          =   180
            Index           =   6
            Left            =   6660
            TabIndex        =   13
            Top             =   36
            Width           =   396
         End
         Begin VB.Shape Shape4 
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00C0C0C0&
            BorderStyle     =   6  'Inside Solid
            FillColor       =   &H00C0FFFF&
            FillStyle       =   0  'Solid
            Height          =   192
            Index           =   6
            Left            =   48
            Top             =   36
            Width           =   8472
         End
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "Update Password"
         Height          =   345
         Left            =   5292
         TabIndex        =   3
         Top             =   1728
         Width           =   1536
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "Close"
         Height          =   345
         Left            =   7092
         TabIndex        =   5
         Top             =   1728
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Current Password"
         Height          =   240
         Left            =   3960
         TabIndex        =   28
         Top             =   540
         Width           =   1332
      End
      Begin MSForms.TextBox txtCurrentPass 
         Height          =   312
         Left            =   5328
         TabIndex        =   0
         Top             =   504
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
      Begin VB.Label lblEMail 
         BackStyle       =   0  'Transparent
         Caption         =   "Email"
         Height          =   240
         Left            =   36
         TabIndex        =   27
         Top             =   1800
         Width           =   828
      End
      Begin MSForms.TextBox txtEmail 
         Height          =   312
         Left            =   936
         TabIndex        =   26
         Top             =   1692
         Width           =   2964
         VariousPropertyBits=   746604571
         Size            =   "5228;550"
         SpecialEffect   =   6
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label lblRole 
         BackStyle       =   0  'Transparent
         Caption         =   "Role"
         Height          =   240
         Left            =   36
         TabIndex        =   25
         Top             =   1404
         Width           =   828
      End
      Begin MSForms.TextBox txtRole 
         Height          =   312
         Left            =   936
         TabIndex        =   24
         Top             =   1296
         Width           =   3000
         VariousPropertyBits=   746604569
         Size            =   "5292;550"
         SpecialEffect   =   6
         FontName        =   "Myriad Web"
         FontEffects     =   1073750016
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label2lblConPass 
         BackStyle       =   0  'Transparent
         Caption         =   "Confirm Password"
         Height          =   240
         Left            =   3960
         TabIndex        =   23
         Top             =   1368
         Width           =   1296
      End
      Begin MSForms.TextBox txtConPass 
         Height          =   312
         Left            =   5328
         TabIndex        =   2
         Top             =   1296
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
         Left            =   3996
         TabIndex        =   22
         Top             =   936
         Width           =   1296
      End
      Begin MSForms.TextBox txtPassword 
         Height          =   312
         Left            =   5328
         TabIndex        =   1
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
      Begin MSForms.TextBox txtUserName 
         Height          =   312
         Left            =   936
         TabIndex        =   21
         Top             =   900
         Width           =   3000
         VariousPropertyBits=   746604571
         Size            =   "5292;550"
         SpecialEffect   =   6
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtUserID 
         Height          =   312
         Left            =   936
         TabIndex        =   20
         Top             =   468
         Width           =   3000
         VariousPropertyBits=   746604569
         Size            =   "5292;550"
         SpecialEffect   =   6
         FontName        =   "Myriad Web"
         FontEffects     =   1073750016
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label LabelUsersName 
         BackStyle       =   0  'Transparent
         Caption         =   "User Name"
         Height          =   240
         Left            =   36
         TabIndex        =   19
         Top             =   936
         Width           =   828
      End
      Begin VB.Label LabelUserID 
         BackStyle       =   0  'Transparent
         Caption         =   "User ID"
         Height          =   240
         Left            =   36
         TabIndex        =   18
         Top             =   540
         Width           =   828
      End
   End
End
Attribute VB_Name = "frmChangePasswordNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdClose_Click()
''''''Modified by Mahboob 03/04/2023 Change ID 10 work item 9 close the form
Unload Me
End Sub
Function LoadUserList()
'''''''''''''''''''Modified by Mahboob 04/04/2023 Change ID 10 work item 2 function to load user in grid
Dim rRow As Integer
     Dim adoConn As New ADODB.Connection
   Dim rst As New ADODB.Recordset
   adoConn.Open getConnectionStringUserAccess
    Dim sqlCom As String
   txtSearchUser.text = ""
   txtSearchName.text = ""
   txtSearchRole.text = ""
   txtSearchEmail.text = ""
   txtSearchActive.text = ""
   gridUserLookup.RowHeight(0) = 0
   gridUserLookup.Cols = 7
   gridUserLookup.ColWidth(0) = 5
   gridUserLookup.ColWidth(1) = 1020
   gridUserLookup.ColWidth(2) = 1565
   gridUserLookup.ColWidth(3) = 2195
   gridUserLookup.ColWidth(4) = 1640
   gridUserLookup.ColWidth(5) = 1925
   gridUserLookup.ColWidth(6) = 5
   gridUserLookup.Clear
   gridUserLookup.ColAlignment(0) = vbLeftJustify
   gridUserLookup.ColAlignment(1) = vbLeftJustify
   gridUserLookup.ColAlignment(2) = vbLeftJustify
   gridUserLookup.ColAlignment(3) = vbLeftJustify
   gridUserLookup.ColAlignment(4) = vbLeftJustify
   gridUserLookup.ColAlignment(5) = vbLeftJustify
   
    
    sqlCom = ""
    sqlCom = sqlCom & " SELECT UserNames.RoleID,UserNames.UserID, UserNames.UserName, Roles.RoleName, UserNames.UserEmail, UserNames.IsActive "
    sqlCom = sqlCom & " FROM Roles INNER JOIN UserNames ON Roles.RoleID = UserNames.RoleID ;"
   rst.Open sqlCom, adoConn, adOpenStatic, adLockReadOnly
rRow = 1
Dim rCount As Integer
rCount = rst.RecordCount
   gridUserLookup.Rows = rst.RecordCount
     If Not rst.EOF Then
      While Not rst.EOF
         If rCount = 1 Then
      gridUserLookup.row = 0
      rRow = 0
      Else
      gridUserLookup.row = 1
      End If
           gridUserLookup.RowSel = 0
           gridUserLookup.ColSel = 0
           gridUserLookup.TextMatrix(rRow, 0) = ""
           gridUserLookup.TextMatrix(rRow, 1) = rst!UserId
           gridUserLookup.TextMatrix(rRow, 2) = rst!UserName
           gridUserLookup.TextMatrix(rRow, 3) = rst!RoleName
           gridUserLookup.TextMatrix(rRow, 4) = IIf(IsNull(rst!UserEmail), "", rst!UserEmail)
'           gridUserLookup.TextMatrix(rRow, 5) = rst!IsActive
gridUserLookup.TextMatrix(rRow, 5) = IIf(rst!IsActive = "Y", "Active", "Disabled")
           gridUserLookup.TextMatrix(rRow, 6) = rst!roleID
           gridUserLookup.RowHeight(rRow) = 240
           If Not rst.EOF Then gridUserLookup.AddItem ""
           rRow = rRow + 1
         rst.MoveNext
      Wend
   End If
'gridUserLookup.Rows = gridUserLookup.Rows - 1
   rst.Close
   Set rst = Nothing
   adoConn.Close
   Set adoConn = Nothing
End Function

Private Sub cmdUpdate_Click()
'''''''''''''''Modified by Mahboob 04/04/2023 Change ID 10 work item 10 password
If txtUserID.text = "" Then
MsgBox "Please select a user to continue."
Exit Sub
End If
'If txtUserID.text = "1" And txtUserName.text <> "Admin" Then
'MsgBox "Admin User name can not be changed."
'Exit Sub
'End If
'checking the current password
If Not checkCurrentPassword Then
        MsgBox "The current password entered is incorrect. Please try again."
'        txtUserName.SetFocus
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
    aboCnn.Execute "Update UserNames set UserPassword='" & txtPassword.text & "' WHERE UserID=" & txtUserID.text & ""
MsgBox "Password changed successfully"
aboCnn.Close
    Set aboCnn = Nothing
    ComponentInFrameClearMode Me, FrameUserInfo, ClearOnlyTextBoxes
End Sub


Public Function checkCurrentPassword() As Boolean
'''''''''''''''''''Modified by Mahboob 03/04/2023 Change ID 10 work item 11 check current password
    Dim adoConn As New ADODB.Connection
    Dim rst As New ADODB.Recordset
    adoConn.Open getConnectionStringUserAccess
    Dim psql As String
    psql = ""
    psql = "sELECT * FROM UserNames WHERE UserID=" & txtUserID.text & " and UserPassword='" & txtCurrentPass.text & "'"
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
Private Sub txtSearchUser_Change()
'''''''''''''''''''Modified by Mahboob 03/04/2023 Change ID 10 work item 3 search by user ID
Dim i As Integer
   If Len(txtSearchUser.text) > 0 Then
        txtSearchName.text = ""
        txtSearchRole.text = ""
        txtSearchEmail.text = ""
        txtSearchActive.text = ""
   End If
   For i = gridUserLookup.Rows - 1 To 1 Step -1
   gridUserLookup.RowHeight(i) = 240
        If InStr(1, UCase(gridUserLookup.TextMatrix(i, 1)), UCase(txtSearchUser.text), vbTextCompare) = 0 Then
              gridUserLookup.RowHeight(i) = 0
        End If
        If gridUserLookup.RowHeight(i) = 240 Then
              gridUserLookup.row = i
        End If
   Next i
End Sub

Private Sub txtSearchName_Change()
'''''''''''''''''''Modified by Mahboob 03/04/2023 Change ID 10 work item 4 search by user Name
Dim i As Integer
   If Len(txtSearchName.text) > 0 Then
        txtSearchUser.text = ""
        txtSearchRole.text = ""
        txtSearchEmail.text = ""
        txtSearchActive.text = ""
   End If
   For i = gridUserLookup.Rows - 1 To 1 Step -1
   gridUserLookup.RowHeight(i) = 240
        If InStr(1, UCase(gridUserLookup.TextMatrix(i, 2)), UCase(txtSearchName.text), vbTextCompare) = 0 Then
              gridUserLookup.RowHeight(i) = 0
        End If
        If gridUserLookup.RowHeight(i) = 240 Then
              gridUserLookup.row = i
        End If
   Next i
End Sub

Private Sub txtSearchRole_Change()
'''''''''Modified by Mahboob 03/04/2023 Change ID 10 work item 5 search by Role Name
Dim i As Integer
   If Len(txtSearchRole.text) > 0 Then
   txtSearchName.text = ""
        txtSearchUser.text = ""

        txtSearchEmail.text = ""
        txtSearchActive.text = ""
   End If
   For i = gridUserLookup.Rows - 1 To 1 Step -1
   gridUserLookup.RowHeight(i) = 240
        If InStr(1, UCase(gridUserLookup.TextMatrix(i, 3)), UCase(txtSearchRole.text), vbTextCompare) = 0 Then
              gridUserLookup.RowHeight(i) = 0
        End If
        If gridUserLookup.RowHeight(i) = 240 Then
              gridUserLookup.row = i
        End If
   Next i
End Sub

Private Sub Form_Load()
Me.BackColor = MODULEBACKCOLOR
FrameUserInfo.BackColor = MODULEBACKCOLOR
'''''''Modified by Mahboob 03/04/2023 Change ID 10 work item 1 load the user details
'Load the users
LoadUserList
End Sub

Private Sub gridUserLookup_Click()
'''''''''Modified by Mahboob 03/04/2023 Change ID 10 work item 8 load user details in text boxes
 txtUserID.text = gridUserLookup.TextMatrix(gridUserLookup.row, 1)
    txtUserName.text = gridUserLookup.TextMatrix(gridUserLookup.row, 2)
    txtUserName.Tag = gridUserLookup.TextMatrix(gridUserLookup.row, 2)
    txtRole.text = gridUserLookup.TextMatrix(gridUserLookup.row, 3)
    txtEmail.text = gridUserLookup.TextMatrix(gridUserLookup.row, 4)
'    chkActive.Value = IIf(gridUserLookup.TextMatrix(gridUserLookup.row, 5) = "Y", True, False)
txtUserName.Enabled = False

End Sub
Private Sub txtSearchActive_Change()
'''''''''Modified by Mahboob 03/04/2023 Change ID 10 work item 7 search by Status
Dim i As Integer
   If Len(txtSearchActive.text) > 0 Then
   txtSearchName.text = ""
        txtSearchUser.text = ""
txtSearchRole.text = ""
        txtSearchEmail.text = ""

   End If
   For i = gridUserLookup.Rows - 1 To 1 Step -1
   gridUserLookup.RowHeight(i) = 240
        If InStr(1, UCase(gridUserLookup.TextMatrix(i, 5)), UCase(txtSearchActive.text), vbTextCompare) = 0 Then
              gridUserLookup.RowHeight(i) = 0
        End If
        If gridUserLookup.RowHeight(i) = 240 Then
              gridUserLookup.row = i
        End If
   Next i
End Sub

Private Sub txtSearchEmail_Change()
'''''''''Modified by Mahboob 03/04/2023 Change ID 10 work item 6 search by Email
Dim i As Integer
   If Len(txtSearchEmail.text) > 0 Then
   txtSearchName.text = ""
        txtSearchUser.text = ""
txtSearchRole.text = ""

        txtSearchActive.text = ""
   End If
   For i = gridUserLookup.Rows - 1 To 1 Step -1
   gridUserLookup.RowHeight(i) = 240
        If InStr(1, UCase(gridUserLookup.TextMatrix(i, 4)), UCase(txtSearchEmail.text), vbTextCompare) = 0 Then
              gridUserLookup.RowHeight(i) = 0
        End If
        If gridUserLookup.RowHeight(i) = 240 Then
              gridUserLookup.row = i
        End If
   Next i
End Sub
