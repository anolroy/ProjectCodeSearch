VERSION 5.00
Begin VB.Form frmlogin 
   BackColor       =   &H00E5E5E5&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Log in - Prestige Property Management Program"
   ClientHeight    =   2730
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5145
   Icon            =   "login.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2730
   ScaleWidth      =   5145
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3960
      Picture         =   "login.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1920
      Width           =   1095
   End
   Begin VB.CommandButton cmdLogin 
      Caption         =   "&Login"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2280
      Picture         =   "login.frx":0B2C
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1920
      Width           =   1095
   End
   Begin VB.ComboBox cboShopCentre 
      Height          =   315
      Left            =   2280
      Style           =   2  'Dropdown List
      TabIndex        =   0
      ToolTipText     =   "Click to choose shopping Centre Name"
      Top             =   240
      Width           =   2775
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2280
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   3
      Text            =   "xx"
      Top             =   1320
      Width           =   2775
   End
   Begin VB.TextBox txtUserName 
      Height          =   285
      Left            =   2280
      MaxLength       =   10
      TabIndex        =   1
      Text            =   "manager"
      Top             =   840
      Width           =   2775
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00E5E5E5&
      Caption         =   "Select Property Company:"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   30
      TabIndex        =   5
      Top             =   240
      Width           =   2145
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00E5E5E5&
      Caption         =   "Enter Password:"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   855
      TabIndex        =   4
      Top             =   1320
      Width           =   1320
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00E5E5E5&
      Caption         =   "Enter User Name:"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   705
      TabIndex        =   2
      Top             =   840
      Width           =   1470
   End
End
Attribute VB_Name = "frmlogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Conn As New RDO.rdoConnection
Dim Env As rdoEnvironment
Dim Envs As rdoEnvironments
Dim Rst As rdoResultset
Dim SQLStr As String

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdLogin_Click()
    ' Pop up a box to remind them about support agreement each year

   If cboShopCentre.text = "" Then Exit Sub

   Dim howMany As Integer
   Dim support, strDecimal As String
   support = "01/01/2003"
   Dim lDecimalPos As Double
   Dim lDifference As Double

   lDifference = CDbl(CDbl(Now - CDate(support)) Mod 365)

   If lDifference <= 30 And howMany < 6 Then
      'this message will appear for a week
      MsgBox "Support for your program is now due to renewal. Please contact the help desk at PCM Consulting if you have not yet received the renewal. "
      howMany = howMany + 1
   Else
      If lDifference > 300 Then howMany = 0
   End If

   'Start of real program
   On Error GoTo ErrH

   Dim a As Integer

   For a = 2 To 4
      If Mid(cboShopCentre.text, a, 3) = " / " Then
         SCID = Left(cboShopCentre.text, a - 1)
         SCName = Right(cboShopCentre.text, Len(cboShopCentre.text) - a - 2)
      End If
   Next a

   Conn.Connect = "DSN=PrestigeBMControl;UID=;PWD="
   Conn.CursorDriver = rdUseOdbc
   Conn.EstablishConnection rdDriverNoPrompt

   SQLStr = "SELECT * FROM Databases WHERE ID = " & SCID
   Set Rst = Conn.OpenResultset(SQLStr, rdOpenStatic, rdConcurReadOnly)

   Adsn = Rst!AccessDSN
   gCurrentShopCentreName = SCName
   gCurrentShopCentreCode = SCID
   CompanyDatapath = "Company Datapath" & CStr(SCID)

   szPictureDBPath = App.Path + "\database\PBMc" & CStr(Format(Val(SCID), "000")) & ".mdb" 'db name

   szSageDSN = "SageLine50v12c" & CStr(SCID)

   Rst.Close
   Conn.Close

   If CheckLogin(Me) Then
      'Login
      User = txtUserName.text
      frmMMain.Show
      'clear password box so when user logs out their password is not still there
      Unload frmlogin
   Else
      MsgBox "Password and UserName are incorrect.  Please try again.", vbOKOnly + vbCritical, "Incorrect Login"
      Me.txtPassword.SetFocus
      'select the text
      Me.txtPassword.SelStart = 0
      Me.txtPassword.SelLength = Len(Me.txtPassword.text)
   End If

   Exit Sub

ErrH:
   If ERR.Number <> 0 Then
      If ERR.Number = 40002 Then
         If MsgBox("DSN Invalid. Please check with your system administrator.", vbRetryCancel + vbCritical, "DSN Set Up Error") = vbRetry Then Resume
      Else
         MsgBox ERR.Number & " -(pcm_004) " & ERR.description
      End If
   End If
End Sub

Private Sub Form_Load()
   Dim latest As Integer
   Dim Name As String
   Dim Name2 As String

   latest = 0
   Me.Move (Screen.Width - Width) / 2, (Screen.Height - Height) / 2
   'add all the shopping centre names from control database to cboshopcentre

   Conn.Connect = "DSN=PrestigeBMControl;UID=;PWD="
   Conn.CursorDriver = rdUseIfNeeded
   Conn.EstablishConnection rdDriverNoPrompt

   SQLStr = "SELECT * FROM Databases"
   Set Rst = Conn.OpenResultset(SQLStr, rdOpenStatic, rdConcurReadOnly)

   If Rst.EOF = False Then
       While Rst.EOF = False
           cboShopCentre.AddItem Rst!ID & " / " & Rst!SCName
           If Rst!ID > latest Then latest = Rst!ID
           Rst.MoveNext
       Wend
   End If
   Rst.Close

'   If latest = 1 Then cboShopCentre.text = cboShopCentre.List(0)
'
'   'search for any new databases(using fileexits function)
'   latest = latest + 1
'   If latest < 10 Then
'       Name = "PrestigeBMc0" & latest
'   Else
'       Name = "PrestigeBMc" & latest
'   End If
'   Name2 = App.Path & "\" & Name & ".mdb"
'   If FileExists(Name2) = True Then
'       Set Rst = Conn.OpenResultset("SELECT * FROM Databases", rdOpenDynamic, rdConcurRowVer)
'       Rst.AddNew
'       Rst!ID = latest
'       Rst!DBName = Name
'       Rst!AccessDSN = Name
'       Rst.Update
'       Rst.Close
'       Conn.Close
'       cboShopCentre.AddItem latest & " / " & Name
'   Else
'       Conn.Close
'   End If
   Me.Show
   cboShopCentre.SetFocus
End Sub
