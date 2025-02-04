VERSION 5.00
Begin VB.Form frmChangePassword1 
   BackColor       =   &H00FFDFC0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Change Password"
   ClientHeight    =   3135
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   4320
   Icon            =   "frmChangePassword1.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   4320
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFDFC0&
      Height          =   1215
      Left            =   120
      TabIndex        =   8
      Top             =   1200
      Width           =   3975
      Begin VB.TextBox Text3 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   2610
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox Text4 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   2610
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFCFD0&
         BackStyle       =   0  'Transparent
         Caption         =   "New Password:"
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   165
         TabIndex        =   10
         Top             =   240
         Width           =   1380
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFCFD0&
         BackStyle       =   0  'Transparent
         Caption         =   "Re-enter New Password:"
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   165
         TabIndex        =   9
         Top             =   720
         Width           =   2190
      End
   End
   Begin VB.TextBox Text2 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2760
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   600
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2760
      TabIndex        =   5
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save "
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   3
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      TabIndex        =   4
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFDFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Current Password:"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   315
      TabIndex        =   7
      Top             =   600
      Width           =   1665
   End
   Begin VB.Label kabel1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFDFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "User Name:"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   315
      TabIndex        =   6
      Top             =   120
      Width           =   1020
   End
End
Attribute VB_Name = "frmChangePassword1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
   ' frmMMain.fraCmdButton.Enabled = True
    Unload Me
End Sub

Private Sub cmdSave_Click()
   Dim Conn As New ADODB.Connection
   Dim rst As New ADODB.Recordset

'   Dim Conn As New RDO.rdoConnection
'   Dim Rst As rdoResultset
   Dim SQLStr As String
   'connect to the database
'   Conn.Connect = getConnectionString
'   Conn.CursorDriver = rdUseIfNeeded
'   Conn.EstablishConnection rdDriverNoPrompt
   Conn.Open getConnectionString

   'get the current password from the usernames table
   SQLStr = "SELECT Password FROM UserNames WHERE UserName = '" & Text1.text & "'"
'   Set Rst = Conn.OpenResultset(SQLStr, rdOpenDynamic, rdConcurRowVer)
   rst.Open SQLStr, Conn, adOpenDynamic, adLockOptimistic

   'If the password entered is not the same as the password from the database tell
   'user password entered incorrectly and close resulset and connection before exiting procedure
   If LCase(rst!Password) <> LCase(Text2.text) Then
       MsgBox "The password entered is incorrect.", vbOKOnly + vbCritical, "Incorrect Password"
       rst.Close
       Conn.Close
       Exit Sub
   End If
   'If the user did not specify a new password tell user
   If Text3.text = "" Then
       MsgBox "You must specify a password!", vbOKOnly, "Password Required"
       Exit Sub
   End If
   'If new password entered twice do not match tell user
   If LCase(Text3.text) <> LCase(Text4.text) Then
       MsgBox "The re-entered new password is incorrect.", vbOKOnly + vbCritical, "Incorrect Password"
       rst.Close
       Conn.Close
       Exit Sub
   End If
   'save the new password to the database as lowercase

   rst!Password = LCase(Text3.text)
   rst.Update
   rst.Close
   Conn.Close
   MsgBox "Password has been changed successfully.", vbInformation + vbOKOnly, "Password"
   'return to frmMain
   Unload Me
End Sub

Private Sub Form_Load()
    Me.Top = 50
    Me.Left = 50
    Me.BackColor = MODULEBACKCOLOR
    Frame1.BackColor = MODULEBACKCOLOR

'    Me.Move (Screen.Width - Width) / 2, (Screen.Height - Height) / 2
    
    Text1.text = User

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   frmMMain.MousePointer = vbArrow
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'frmMMain.fraCmdButton.Enabled = True
    Unload Me
End Sub
