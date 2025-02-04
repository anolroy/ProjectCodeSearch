VERSION 5.00
Begin VB.Form frmUsers 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Change User Name"
   ClientHeight    =   3120
   ClientLeft      =   150
   ClientTop       =   240
   ClientWidth     =   4860
   Icon            =   "frmUsers.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3120
   ScaleWidth      =   4860
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00800000&
      Height          =   100
      Left            =   -120
      ScaleHeight     =   45
      ScaleWidth      =   5235
      TabIndex        =   16
      Top             =   540
      Width           =   5295
   End
   Begin VB.CommandButton cmdCancelNew 
      Caption         =   "&Cancel New User"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3240
      TabIndex        =   7
      Top             =   4080
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton cmdSaveNew 
      Caption         =   "&Save New User"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   8
      Top             =   4080
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancelEdit 
      Caption         =   "&Cancel Changes"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3240
      TabIndex        =   6
      Top             =   3600
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton cmdSaveEdit 
      Caption         =   "&Save Changes"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   5
      Top             =   3600
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2400
      MaxLength       =   15
      PasswordChar    =   "#"
      TabIndex        =   3
      Top             =   1200
      Width           =   2295
   End
   Begin VB.CommandButton cmdDeleteUser 
      Caption         =   "&Delete User"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3240
      TabIndex        =   15
      Top             =   2160
      Width           =   1455
   End
   Begin VB.CommandButton cmdCloseScreen 
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   13
      Top             =   2640
      Width           =   1455
   End
   Begin VB.TextBox Text3 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2400
      MaxLength       =   15
      PasswordChar    =   "#"
      TabIndex        =   4
      Top             =   1680
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2400
      MaxLength       =   15
      TabIndex        =   2
      Top             =   720
      Width           =   2295
   End
   Begin VB.ComboBox cboUserName 
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2400
      TabIndex        =   1
      Top             =   120
      Width           =   2295
   End
   Begin VB.CommandButton cmdAddNewUser 
      Caption         =   "&Add New User"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   2160
      Width           =   1455
   End
   Begin VB.CommandButton cmdEditUser 
      Caption         =   "&Edit User"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   0
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Select Existing User:"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   14
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Re-enter Password:"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   12
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   11
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "User Name:"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   10
      Top             =   720
      Width           =   1455
   End
End
Attribute VB_Name = "frmUsers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Conn As New ADODB.Connection
Dim Rst As New ADODB.Recordset
Dim SQLStr As String
Dim CurrentUser As String

Private Enum OperationMode
   EditMode = 0
   AddNewMode = 1
   DefaultMode = 3
End Enum

Private Sub CommandButtonHandling(ByVal mode As OperationMode)
   Select Case mode
      Case OperationMode.AddNewMode
         cmdAddNewUser.Enabled = False
         cmdSaveNew.Visible = True
         cmdSaveNew.Top = cmdEditUser.Top
         cmdEditUser.Visible = False
         cmdCancelNew.Visible = True
         cmdCancelNew.Top = cmdDeleteUser.Top
         cmdDeleteUser.Visible = False

      Case OperationMode.EditMode
         cmdEditUser.Left = cmdAddNewUser.Left
         cmdEditUser.Enabled = False
         cmdAddNewUser.Visible = False
         cmdSaveEdit.Top = cmdEditUser.Top
         cmdSaveEdit.Visible = True
         cmdCancelEdit.Top = cmdDeleteUser.Top
         cmdDeleteUser.Visible = False
         cmdCancelEdit.Visible = True

      Case OperationMode.DefaultMode
         cmdAddNewUser.Enabled = True
         cmdAddNewUser.Visible = True
         cmdEditUser.Visible = True
         cmdEditUser.Left = cmdCloseScreen.Left
         cmdEditUser.Enabled = True
         cmdDeleteUser.Visible = True
         cmdSaveEdit.Visible = False
         cmdCancelEdit.Visible = False
         cmdSaveNew.Visible = False
         cmdCancelNew.Visible = False
   End Select
End Sub

Private Sub cboUserName_Click()

Dim i, j, match As Integer
match = 0

If cboUserName.text = "" Then
    ShowMsgInTaskBar "You must select a User to view!", , "N"
    Exit Sub
End If
j = cboUserName.ListCount - 1
For i = 0 To j
    If cboUserName.List(i) = cboUserName.text Then
        match = 1
        Exit For
    End If
Next i
If match = 0 Then
    ShowMsgInTaskBar "User selected is invalid.", , "N"
    cboUserName.text = ""
    Exit Sub
End If
Call GetRecord

End Sub

Private Sub cmdAddNewUser_Click()
   Call AddNewUser
   CommandButtonHandling AddNewMode
End Sub

Private Sub cmdCancelEdit_Click()
   CommandButtonHandling DefaultMode
   Call EnableMenu
   Call EmptyBoxes
   Call DisableBoxes
   Call GetRecord
   Call FillcboUserName
   cboUserName.text = Text1.text
End Sub

Private Sub cmdCancelNew_Click()
   CommandButtonHandling DefaultMode

'   cmdCancelNew.Visible = False
'   cmdSaveNew.Visible = False
'   cmdEditUser.Visible = True
'   cmdAddNewUser.Visible = True
'   cmdDeleteUser.Visible = True
'   Label3.Visible = False
'   Text3.Visible = False
   Call EnableMenu
   Call EmptyBoxes
   Call DisableBoxes
   Call FillcboUserName
   cboUserName.text = Text1.text
End Sub

Private Sub cmdCloseScreen_Click()

Call CloseScreen

End Sub

Private Sub cmdDeleteUser_Click()

Call DeleteUser

End Sub

Private Sub cmdEditUser_Click()

Call EditUser

End Sub

Private Sub cmdSaveEdit_Click()
   
Dim match As Integer
match = 0

'check password entered the same in both text2 and text3
If Text2.text <> Text3.text Then ' not same
    ShowMsgInTaskBar "Password re-entered incorrectly. Please re-enter.", , "N"
    Text2.SetFocus
    'select the text
    SelTxtInCtrl Text2
    Exit Sub
Else
    If Text2.text = "" Then
        ShowMsgInTaskBar "You must enter a Password.", , "N"
        Text2.SetFocus
        Exit Sub
    End If
End If

'check username different
If Text1.text <> cboUserName.text Then 'new user name entered and need to check if it is unique.
    Conn.Open getConnectionString
    SQLStr = "SELECT UserName FROM UserNames"
    Rst.Open SQLStr, Conn, adOpenStatic, adLockReadOnly
    
    While Rst.EOF = False
        If Rst!UserName = Text1.text Then match = 1
        Rst.MoveNext
    Wend
    Rst.Close
    If match = 1 Then 'user with this username already exists
        If MsgBox("User with User Name: " & Text1.text & " already exitst. Do you want to save changes to this existing User?", vbYesNo + vbQuestion, "User already exists") = vbNo Then
            Conn.Close
            Exit Sub
        End If
    Else
        CurrentUser = Text1.text
        Call SaveChanges
    End If
End If

Call DisableBoxes
CommandButtonHandling DefaultMode
Call EnableMenu
Call FillcboUserName
cboUserName.text = Text1.text

End Sub

Private Sub cmdSaveNew_Click()
Dim match As Integer
match = 0

'check password entered the same in both text2 and text3
If Text2.text <> Text3.text Then ' not same
    ShowMsgInTaskBar "Password re-entered incorrectly. Please re-enter.", , "N"
    Text2.SetFocus
    'select the text
    SelTxtInCtrl Text2
    Exit Sub
Else
    If Text2.text = "" Then
        ShowMsgInTaskBar "You must enter a Password.", , "N"
        Text2.SetFocus
        Exit Sub
    End If
End If

'check username entered and does not already exist in table
If Text1.text = "" Then
    ShowMsgInTaskBar "You must enter a User Name.", , "N"
    Exit Sub
Else
    Conn.Open getConnectionString

    SQLStr = "SELECT UserName FROM UserNames"
    Rst.Open SQLStr, Conn, adOpenStatic, adLockReadOnly
    
    While Rst.EOF = False
        If Rst!UserName = Text1.text Then match = 1
        Rst.MoveNext
    Wend
    Rst.Close
    If match = 1 Then 'user with this username already exists
        If MsgBox("User with User Name: " & Text1.text & " already exitst. Do you want to save changes to this existing User?", vbYesNo + vbQuestion, "User already exists") = vbNo Then
            Conn.Close
            Exit Sub
        Else
            CurrentUser = Text1.text
            Call SaveChanges
        End If
    End If
End If

'save new user
SQLStr = "SELECT * FROM UserNames WHERE UserName"
Rst.Open SQLStr, Conn, adOpenDynamic, adLockOptimistic

Rst.AddNew
Rst!UserName = Text1.text
Rst!Password = Text2.text
Rst.Update
Rst.Close
Conn.Close

Call DisableBoxes
CommandButtonHandling DefaultMode
Call EnableMenu
Call FillcboUserName
cboUserName.text = Text1.text

End Sub

Private Sub Form_Load()
   frmMMain.Arrange vbCascade
   Me.ZOrder 0
   Me.BackColor = MODULEBACKCOLOR

   Call EmptyBoxes
   Call DisableBoxes
   Call FillcboUserName

   CommandButtonHandling DefaultMode
End Sub

Public Sub AddNewUser()
'   cmdAddNewUser.Visible = False
'   cmdEditUser.Visible = False
'   cmdDeleteUser.Visible = False
'   cmdSaveNew.Visible = True
'   cmdCancelNew.Visible = True
'   Label3.Visible = True
'   Text3.Visible = True
   Call DisableMenu
   Call EmptyBoxes
   cboUserName.text = ""
   Call EnableBoxes
End Sub

Public Sub DisableMenu()

'mnuAddNewUser.Enabled = False
'mnuEditUser.Enabled = False
'mnuDeleteUser.Enabled = False

End Sub

Public Sub EnableMenu()

'mnuAddNewUser.Enabled = True
'mnuEditUser.Enabled = True
'mnuDeleteUser.Enabled = True

End Sub

Private Sub mnuAddNewUser_Click()

Call AddNewUser

End Sub

Public Sub FillcboUserName()

    cboUserName.Clear
    
    Conn.Open getConnectionString
    
    SQLStr = "SELECT UserName FROM UserNames"
    Rst.Open SQLStr, Conn, adOpenStatic, adLockReadOnly
    
    If Rst.EOF = False Then
        While Rst.EOF = False
            cboUserName.AddItem Rst!UserName
            Rst.MoveNext
        Wend
    End If
    
    Rst.Close
    Conn.Close

End Sub

Public Sub EditUser()
   If cboUserName.text = "" Then
       ShowMsgInTaskBar "You must select a User!", , "N"
       Exit Sub
   End If

   CommandButtonHandling EditMode
   Call DisableMenu
   Call EnableBoxes
   Text2.SetFocus
   'select the text
   SelTxtInCtrl Text2
End Sub

Private Sub mnuCloseScreen_Click()
   Call CloseScreen
End Sub

Private Sub mnuDeleteUser_Click()
   Call DeleteUser
End Sub

Private Sub mnuEditUser_Click()
   Call EditUser
End Sub

Public Sub EmptyBoxes()
   Text1.text = ""
   Text2.text = ""
   Text3.text = ""
End Sub

Public Sub EnableBoxes()
   cboUserName.Enabled = False
   Text1.Enabled = True
   Text2.Enabled = True
   Text3.Enabled = True
End Sub

Public Sub DisableBoxes()
   cboUserName.Enabled = True
   Text1.Enabled = False
   Text2.Enabled = False
   Text3.Enabled = False
End Sub

Public Sub GetRecord()
   CurrentUser = cboUserName.text

   Conn.Open getConnectionString

   SQLStr = "SELECT * FROM UserNames WHERE UserName = '" & CurrentUser & "'"
   Rst.Open SQLStr, Conn, adOpenStatic, adLockReadOnly

   Text1.text = Rst!UserName
   Text2.text = Rst!Password

   Rst.Close
   Conn.Close
End Sub

Public Sub SaveChanges()
   Conn.Open getConnectionString

   SQLStr = "SELECT * FROM UserNames WHERE UserName = '" & CurrentUser & "'"
   Rst.Open SQLStr, Conn, adOpenDynamic, adLockOptimistic

   Rst!UserName = Text1.text
   Rst!Password = Text2.text
   Rst.Update
   Rst.Close
   Conn.Close
End Sub

Public Sub DeleteUser()
   If cboUserName.text = "" Then
      ShowMsgInTaskBar "You must select a user to delete!", , "N"
      Exit Sub
   End If

   Conn.Open getConnectionString

   SQLStr = "SELECT * FROM UserNames"
   Rst.Open SQLStr, Conn, adOpenStatic, adLockReadOnly

   If Rst.EOF = False Then
      Rst.MoveLast
      Rst.MoveFirst
      If Rst.RecordCount = 1 Then ' cannot delete user because only one user left.
         ShowMsgInTaskBar "This is the only user.  You cannot delete this user.", , "N"
         Rst.Close
         Conn.Close
         Exit Sub
      End If
   Rst.Close
   End If

   If MsgBox("Are you sure you want to delete user: " & Text1.text & "?", vbOKCancel + vbQuestion, "Delete User") = vbCancel Then
      Conn.Close
      Exit Sub
   End If

   SQLStr = "SELECT * FROM UserNames WHERE UserName = '" & Text1.text & "'"
   Rst.Open SQLStr, Conn, adOpenDynamic, adLockOptimistic

   Rst.Delete
   Rst.Close
   Conn.Close

   Call EmptyBoxes
   Call FillcboUserName
End Sub

Public Sub CloseScreen()
   Unload Me
End Sub
