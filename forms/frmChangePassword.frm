VERSION 5.00
Begin VB.Form frmChangPws 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Change User Name"
   ClientHeight    =   6885
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   8760
   Icon            =   "frmChangePassword.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6885
   ScaleWidth      =   8760
   Begin VB.CommandButton cmdCancelNew 
      Caption         =   "&Cancel New User"
      Height          =   375
      Left            =   960
      TabIndex        =   7
      Top             =   3480
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton cmdSaveNew 
      Caption         =   "&Save New User"
      Height          =   375
      Left            =   2760
      TabIndex        =   8
      Top             =   3480
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancelEdit 
      Caption         =   "&Cancel Changes"
      Height          =   375
      Left            =   2880
      TabIndex        =   6
      Top             =   2880
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton cmdSaveEdit 
      Caption         =   "&Save Changes"
      Height          =   375
      Left            =   960
      TabIndex        =   5
      Top             =   2880
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2400
      MaxLength       =   15
      TabIndex        =   3
      Top             =   1200
      Width           =   1575
   End
   Begin VB.CommandButton cmdDeleteUser 
      Caption         =   "&Delete User"
      Height          =   375
      Left            =   1800
      TabIndex        =   15
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton cmdCloseScreen 
      Caption         =   "&Close Screen"
      Height          =   375
      Left            =   2040
      TabIndex        =   13
      Top             =   3960
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   2400
      MaxLength       =   15
      TabIndex        =   4
      Top             =   1680
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2400
      MaxLength       =   15
      TabIndex        =   2
      Top             =   720
      Width           =   1575
   End
   Begin VB.ComboBox cboUserName 
      Height          =   315
      Left            =   2400
      TabIndex        =   1
      Top             =   240
      Width           =   1575
   End
   Begin VB.CommandButton cmdAddNewUser 
      Caption         =   "&Add New User"
      Height          =   375
      Left            =   3240
      TabIndex        =   9
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton cmdEditUser 
      Caption         =   "&Edit User"
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Select Existing User:"
      Height          =   255
      Left            =   360
      TabIndex        =   14
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Re-enter Password:"
      Height          =   255
      Left            =   840
      TabIndex        =   12
      Top             =   1680
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Password:"
      Height          =   255
      Left            =   1080
      TabIndex        =   11
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "User Name:"
      Height          =   255
      Left            =   840
      TabIndex        =   10
      Top             =   720
      Width           =   1455
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuCloseScreen 
         Caption         =   "&Close Screen"
      End
   End
   Begin VB.Menu mnuChange 
      Caption         =   "Change User Details"
      Begin VB.Menu mnuEditUser 
         Caption         =   "Edit User"
      End
      Begin VB.Menu mnuAddNewUser 
         Caption         =   "Add New User"
      End
      Begin VB.Menu mnuDeleteUser 
         Caption         =   "Delete User"
      End
   End
End
Attribute VB_Name = "frmChangPws"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*************************************
'       This Form needed to modify
'*************************************
' This is the copy of the copy of the frmUsers form. This form
' will be used for only changing password purpose. frmUsers form
' is only for change/edit/add user.
'*************************************


Option Explicit

Dim Conn As New RDO.rdoConnection
Dim Rst As rdoResultset
Dim SQLStr As String
Dim CurrentUser As String

Private Sub cboUserName_Click()

Dim i, j, match As Integer
match = 0

If cboUserName.Text = "" Then
    MsgBox "You must select a User to view!", vbOKOnly + vbCritical, "No user selected"
    Exit Sub
End If
j = cboUserName.ListCount - 1
For i = 0 To j
    If cboUserName.List(i) = cboUserName.Text Then
        match = 1
        Exit For
    End If
Next i
If match = 0 Then
    MsgBox "User selected is invalid.", vbOKOnly + vbCritical, "Invalid User"
    cboUserName.Text = ""
    Exit Sub
End If
Call GetRecord

End Sub

Private Sub cmdAddNewUser_Click()

Call AddNewUser

End Sub

Private Sub cmdCancelEdit_Click()

cmdCancelEdit.Visible = False
cmdSaveEdit.Visible = False
cmdEditUser.Visible = True
cmdDeleteUser.Visible = True
cmdAddNewUser.Visible = True
Label3.Visible = False
Text3.Visible = False
Call EnableMenu
Call EmptyBoxes
Call DisableBoxes
Call GetRecord
Call FillcboUserName
cboUserName.Text = Text1.Text

End Sub

Private Sub cmdCancelNew_Click()

cmdCancelNew.Visible = False
cmdSaveNew.Visible = False
cmdEditUser.Visible = True
cmdAddNewUser.Visible = True
cmdDeleteUser.Visible = True
Label3.Visible = False
Text3.Visible = False
Call EnableMenu
Call EmptyBoxes
Call DisableBoxes
Call FillcboUserName
cboUserName.Text = Text1.Text

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
If Text2.Text <> Text3.Text Then ' not same
    MsgBox "Password re-entered incorrectly. Please re-enter.", vbOKOnly + vbCritical, "Incorrect Password"
    Text2.SetFocus
    'select the text
    Text2.SelStart = 0
    Text2.SelLength = Len(Text2.Text)
    Exit Sub
Else
    If Text2.Text = "" Then
        MsgBox "You must enter a Password.", vbOKOnly + vbCritical, "No Password entered"
        Text2.SetFocus
        Exit Sub
    End If
End If

'check username different
If Text1.Text <> cboUserName.Text Then 'new user name entered and need to check if it is unique.
    Conn.Connect = "DSN=" & Adsn & ";UID=;PWD="
    Conn.CursorDriver = rdUseIfNeeded
    Conn.EstablishConnection rdDriverNoPrompt
    SQLStr = "SELECT UserName FROM UserNames"
    Set Rst = Conn.OpenResultset(SQLStr, rdOpenStatic, rdConcurReadOnly)
    
    While Rst.EOF = False
        If Rst!UserName = Text1.Text Then match = 1
        Rst.MoveNext
    Wend
    Rst.Close
    If match = 1 Then 'user with this username already exists
        If MsgBox("User with User Name: " & Text1.Text & " already exitst. Do you want to save changes to this existing User?", vbYesNo + vbQuestion, "User already exists") = vbNo Then
            Conn.Close
            Exit Sub
        End If
    Else
        CurrentUser = Text1.Text
        Call SaveChanges
    End If
End If

Call DisableBoxes
cmdSaveEdit.Visible = False
cmdCancelEdit.Visible = False
cmdAddNewUser.Visible = True
cmdEditUser.Visible = True
cmdDeleteUser.Visible = True
Label3.Visible = False
Text3.Visible = False
Call EnableMenu
Call FillcboUserName
cboUserName.Text = Text1.Text

End Sub

Private Sub cmdSaveNew_Click()
Dim match As Integer
match = 0

'check password entered the same in both text2 and text3
If Text2.Text <> Text3.Text Then ' not same
    MsgBox "Password re-entered incorrectly. Please re-enter.", vbOKOnly + vbCritical, "Incorrect Password"
    Text2.SetFocus
    'select the text
    Text2.SelStart = 0
    Text2.SelLength = Len(Text2.Text)
    Exit Sub
Else
    If Text2.Text = "" Then
        MsgBox "You must enter a Password.", vbOKOnly + vbCritical, "No Password entered"
        Text2.SetFocus
        Exit Sub
    End If
End If

'check username entered and does not already exist in table
If Text1.Text = "" Then
    MsgBox "You must enter a User Name.", vbOKOnly + vbCritical, "No User Name Entered"
    Exit Sub
Else
    Conn.Connect = "DSN=" & Adsn & ";UID=;PWD="
    Conn.CursorDriver = rdUseIfNeeded
    Conn.EstablishConnection rdDriverNoPrompt

    SQLStr = "SELECT UserName FROM UserNames"
    Set Rst = Conn.OpenResultset(SQLStr, rdOpenStatic, rdConcurReadOnly)
    
    While Rst.EOF = False
        If Rst!UserName = Text1.Text Then match = 1
        Rst.MoveNext
    Wend
    Rst.Close
    If match = 1 Then 'user with this username already exists
        If MsgBox("User with User Name: " & Text1.Text & " already exitst. Do you want to save changes to this existing User?", vbYesNo + vbQuestion, "User already exists") = vbNo Then
            Conn.Close
            Exit Sub
        Else
            CurrentUser = Text1.Text
            Call SaveChanges
        End If
    End If
End If

'save new user
SQLStr = "SELECT * FROM UserNames WHERE UserName"
Set Rst = Conn.OpenResultset(SQLStr, rdOpenDynamic, rdConcurRowVer)

Rst.AddNew
Rst!UserName = Text1.Text
Rst!Password = Text2.Text
Rst.Update
Rst.Close
Conn.Close

Call DisableBoxes
cmdSaveNew.Visible = False
cmdCancelNew.Visible = False
cmdAddNewUser.Visible = True
cmdEditUser.Visible = True
cmdDeleteUser.Visible = True
Label3.Visible = False
Text3.Visible = False
Call EnableMenu
Call FillcboUserName
cboUserName.Text = Text1.Text

End Sub

Private Sub Form_Load()
    Me.Top = 50
    Me.Left = 50

Call EmptyBoxes
Call DisableBoxes
Call FillcboUserName

End Sub

Public Sub AddNewUser()

cmdAddNewUser.Visible = False
cmdEditUser.Visible = False
cmdDeleteUser.Visible = False
cmdSaveNew.Visible = True
cmdCancelNew.Visible = True
Label3.Visible = True
Text3.Visible = True
Call DisableMenu
Call EmptyBoxes
cboUserName.Text = ""
Call EnableBoxes

End Sub

Public Sub DisableMenu()

mnuAddNewUser.Enabled = False
mnuEditUser.Enabled = False
mnuDeleteUser.Enabled = False

End Sub

Public Sub EnableMenu()

mnuAddNewUser.Enabled = True
mnuEditUser.Enabled = True
mnuDeleteUser.Enabled = True

End Sub

Private Sub mnuAddNewUser_Click()

Call AddNewUser

End Sub

Public Sub FillcboUserName()

    cboUserName.Clear
    
    Conn.Connect = "DSN=" & Adsn & ";UID=;PWD="
    Conn.CursorDriver = rdUseIfNeeded
    Conn.EstablishConnection rdDriverNoPrompt
    
    SQLStr = "SELECT UserName FROM UserNames"
    Set Rst = Conn.OpenResultset(SQLStr, rdOpenStatic, rdConcurReadOnly)
    
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

    If cboUserName.Text = "" Then
        MsgBox "You must select a User!", vbOKOnly + vbExclamation, "No User Selected"
        Exit Sub
    End If
    
    cmdAddNewUser.Visible = False
    cmdEditUser.Visible = False
    cmdDeleteUser.Visible = False
    cmdSaveEdit.Visible = True
    cmdCancelEdit.Visible = True
    Label3.Visible = True
    Text3.Visible = True
    Call DisableMenu
    Call EnableBoxes
    Text2.SetFocus
    'select the text
    Text2.SelStart = 0
    Text2.SelLength = Len(Text2.Text)
    
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

    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""

End Sub

Public Sub EnableBoxes()

    cboUserName.Enabled = False
    Text1.Enabled = True
    Text2.Enabled = True

End Sub

Public Sub DisableBoxes()
    
    cboUserName.Enabled = True
    Text1.Enabled = False
    Text2.Enabled = False

End Sub

Public Sub GetRecord()
    
    CurrentUser = cboUserName.Text
    
    Conn.Connect = "DSN=" & Adsn & ";UID=;PWD="
    Conn.CursorDriver = rdUseIfNeeded
    Conn.EstablishConnection rdDriverNoPrompt
    
    SQLStr = "SELECT * FROM UserNames WHERE UserName = '" & CurrentUser & "'"
    Set Rst = Conn.OpenResultset(SQLStr, rdOpenStatic, rdConcurReadOnly)
    
    Text1.Text = Rst!UserName
    Text2.Text = Rst!Password
    
    Rst.Close
    Conn.Close

End Sub

Public Sub SaveChanges()
    Conn.Connect = "DSN=" & Adsn & ";UID=;PWD="
    Conn.CursorDriver = rdUseOdbc
    Conn.EstablishConnection rdDriverNoPrompt
    
    SQLStr = "SELECT * FROM UserNames WHERE UserName = '" & CurrentUser & "'"
    Set Rst = Conn.OpenResultset(SQLStr, rdOpenDynamic, rdConcurRowVer)
        
    Rst.Edit
    Rst!UserName = Text1.Text
    Rst!Password = Text2.Text
    Rst.Update
    Rst.Close
    Conn.Close

End Sub

Public Sub DeleteUser()

If cboUserName.Text = "" Then
    MsgBox "You must select a user to delete!", vbOKOnly + vbExclamation, "No User Selected"
    Exit Sub
End If

Conn.Connect = "DSN=" & Adsn & ";UID=;PWD="
Conn.CursorDriver = rdUseOdbc
Conn.EstablishConnection rdDriverNoPrompt

SQLStr = "SELECT * FROM UserNames"
Set Rst = Conn.OpenResultset(SQLStr, rdOpenStatic, rdConcurReadOnly)

If Rst.EOF = False Then
    Rst.MoveLast
    Rst.MoveFirst
    If Rst.RowCount = 1 Then ' cannot delete user because only one user left.
        MsgBox "This is the only user.  You cannot delete this user.", vbOKOnly + vbCritical, "Cannot Delete User"
        Rst.Close
        Conn.Close
        Exit Sub
    End If
Rst.Close
End If

If MsgBox("Are you sure you want to delete user: " & Text1.Text & "?", vbOKCancel + vbQuestion, "Delete User") = vbCancel Then
    Conn.Close
    Exit Sub
End If

SQLStr = "SELECT * FROM UserNames WHERE UserName = '" & Text1.Text & "'"
Set Rst = Conn.OpenResultset(SQLStr, rdOpenDynamic, rdConcurRowVer)

Rst.Delete
Rst.Close
Conn.Close

Call EmptyBoxes
Call FillcboUserName

End Sub

Public Sub CloseScreen()

Unload Me
'Load frmMain
'frmMain.Show

End Sub
