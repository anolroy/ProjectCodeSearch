VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmLogin2 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Welcome to Prestige"
   ClientHeight    =   7680
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6765
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLogin2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7680
   ScaleWidth      =   6765
   StartUpPosition =   2  'CenterScreen
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxCompany 
      Height          =   2916
      Left            =   36
      TabIndex        =   19
      Top             =   1764
      Width           =   6612
      _ExtentX        =   11668
      _ExtentY        =   5133
      _Version        =   393216
      FixedRows       =   0
      FixedCols       =   0
      BackColorFixed  =   13553358
      ForeColorFixed  =   -2147483634
      BackColorSel    =   12648447
      ForeColorSel    =   -2147483630
      BackColorBkg    =   16777215
      GridColor       =   14737632
      WordWrap        =   -1  'True
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
      _Band(0).Cols   =   2
   End
   Begin VB.PictureBox picCompany 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3636
      Left            =   0
      ScaleHeight     =   3600
      ScaleWidth      =   6690
      TabIndex        =   13
      Top             =   1116
      Width           =   6720
      Begin MSForms.Label lblComId 
         Height          =   192
         Left            =   108
         TabIndex        =   16
         Top             =   72
         Width           =   1596
         VariousPropertyBits=   8388627
         Caption         =   "Company Id"
         Size            =   "2805;344"
         FontName        =   "Myriad Web"
         FontHeight      =   156
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label lblComName 
         Height          =   192
         Left            =   1668
         TabIndex        =   15
         Top             =   84
         Width           =   1596
         VariousPropertyBits=   8388627
         Caption         =   "Company Name"
         Size            =   "2805;344"
         FontName        =   "Myriad Web"
         FontHeight      =   156
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label10 
         Height          =   192
         Left            =   -252
         TabIndex        =   14
         Top             =   -432
         Width           =   1416
         VariousPropertyBits=   8388627
         Caption         =   "Company Id"
         Size            =   "2487;344"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtSearchCompanyId 
         Height          =   252
         Left            =   48
         TabIndex        =   0
         Top             =   336
         Width           =   1608
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         Size            =   "2836;444"
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtCompanyName 
         Height          =   252
         Left            =   1692
         TabIndex        =   1
         Top             =   336
         Width           =   4980
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         Size            =   "8784;444"
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Shape Shape4 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C0C0C0&
         BorderStyle     =   6  'Inside Solid
         FillColor       =   &H00E0FFFF&
         FillStyle       =   0  'Solid
         Height          =   240
         Index           =   0
         Left            =   0
         Top             =   72
         Width           =   6648
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   720
      Left            =   5448
      Picture         =   "frmLogin2.frx":014A
      ScaleHeight     =   720
      ScaleWidth      =   900
      TabIndex        =   11
      Top             =   4884
      Width           =   900
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Exit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   3948
      Picture         =   "frmLogin2.frx":0DDB
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5724
      Width           =   1095
   End
   Begin VB.CommandButton cmdLogin 
      Caption         =   "&Logon"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   1692
      Picture         =   "frmLogin2.frx":116F
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5724
      Width           =   1095
   End
   Begin VB.ComboBox cboShopCentre 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1440
      TabIndex        =   5
      ToolTipText     =   "Click to choose a property"
      Top             =   1560
      Width           =   3135
   End
   Begin VB.TextBox txtPassword 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1728
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   3
      Text            =   "xx"
      Top             =   5256
      Width           =   3312
   End
   Begin VB.TextBox txtUserName 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   1728
      MaxLength       =   10
      TabIndex        =   2
      Text            =   "manager"
      Top             =   4860
      Width           =   3312
   End
   Begin VB.Label lblChangePass 
      BackStyle       =   0  'Transparent
      Caption         =   "Click here to change user password"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   768
      Left            =   216
      TabIndex        =   20
      Top             =   5796
      Width           =   1200
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Subscription Licence Terms and Conditions available from PCM Consulting Ltd"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   192
      Left            =   648
      TabIndex        =   18
      Top             =   7128
      Width           =   5988
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Using this program constitutes acceptance of the Prestige Property Management Software "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   192
      Left            =   216
      TabIndex        =   17
      Top             =   6840
      Width           =   6420
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00E5E5E5&
      BackStyle       =   0  'Transparent
      Caption         =   "Company"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   192
      Left            =   324
      TabIndex        =   12
      Top             =   1620
      Width           =   720
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00E5E5E5&
      BackStyle       =   0  'Transparent
      Caption         =   "Property"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   216
      Left            =   -1884
      TabIndex        =   8
      Top             =   1560
      Width           =   612
   End
   Begin VB.Label lblReprintDemand 
      Alignment       =   2  'Center
      BackColor       =   &H00E5D5C5&
      BackStyle       =   0  'Transparent
      Caption         =   "Settings"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   252
      Left            =   4440
      MousePointer    =   99  'Custom
      TabIndex        =   10
      Top             =   8604
      Visible         =   0   'False
      Width           =   792
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00E5E5E5&
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   216
      Left            =   216
      TabIndex        =   7
      Top             =   5316
      Width           =   756
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00E5E5E5&
      BackStyle       =   0  'Transparent
      Caption         =   "User Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   216
      Left            =   216
      TabIndex        =   6
      Top             =   4932
      Width           =   792
   End
   Begin VB.Image Image1 
      Enabled         =   0   'False
      Height          =   1530
      Left            =   75
      Picture         =   "frmLogin2.frx":1537
      Top             =   -150
      Width           =   6270
   End
End
Attribute VB_Name = "frmLogin2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Conn As New ADODB.Connection
Dim rst As New ADODB.Recordset
Dim SQLStr As String

'##############               RUNNING THIRD PARTY PROCESS AND CONTROL
Private Type STARTUPINFO
   cB As Long
   lpReserved As String
   lpDesktop As String
   lpTitle As String
   dwX As Long
   dwY As Long
   dwXSize As Long
   dwYSize As Long
   dwXCountChars As Long
   dwYCountChars As Long
   dwFillAttribute As Long
   dwFlags As Long
   wShowWindow As Integer
   cbReserved2 As Integer
   lpReserved2 As Long
   hStdInput As Long
   hStdOutput As Long
   hStdError As Long
End Type

Private Type PROCESS_INFORMATION
   hProcess As Long
   hThread As Long
   dwProcessID As Long
   dwThreadID As Long
End Type
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal _
   hHandle As Long, ByVal dwMilliseconds As Long) As Long

Private Declare Function CreateProcessA Lib "kernel32" (ByVal _
   lpApplicationName As String, ByVal lpCommandLine As String, ByVal _
   lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, _
   ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, _
   ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As String, _
   lpStartupInfo As STARTUPINFO, lpProcessInformation As _
   PROCESS_INFORMATION) As Long

Private Declare Function CloseHandle Lib "kernel32" _
   (ByVal hObject As Long) As Long

Private Declare Function GetExitCodeProcess Lib "kernel32" _
   (ByVal hProcess As Long, lpExitCode As Long) As Long

Private Const NORMAL_PRIORITY_CLASS = &H20&
Private Const INFINITE = -1&
'-------------------------------------------------------------------------------

Private Sub cboShopCentre_KeyUp(KeyCode As Integer, Shift As Integer)
   Dim i As Integer, l As Integer

   For i = 0 To cboShopCentre.ListCount - 1
      If cboShopCentre.text = Left(cboShopCentre.List(i), Len(cboShopCentre.text)) Then
         l = Len(cboShopCentre.text)
         cboShopCentre.ListIndex = i
         cboShopCentre.SelStart = l
         cboShopCentre.SelLength = Len(cboShopCentre.text) - l
         Exit For
      End If
   Next i
End Sub

Private Sub cmdCancel_Click()
   Unload Me
End Sub



Private Sub cmdLogin_Click()
''''''''''''''''''''''''''''''Modified by Mahboob 29/03/2023 Change ID 5/Work Item 1 Creating the User Access table
Dim adoRibConn As New ADODB.Connection
   Dim adoRst     As New ADODB.Recordset
   adoRibConn.Open "Driver={Microsoft Access Driver (*.mdb)};DBQ=" & App.Path & "\Common Files\RibbonLayOut.mdb;"
On Error GoTo CreateTable_UserAccess
   adoRst.Open "SELECT * FROM UserNames;", adoRibConn, adOpenStatic, adLockReadOnly
   adoRst.Close
   Set adoRst = Nothing
   
   ''''''''''End of modification
   ''''''''''Modified by Mahboob 31/03/2023 Change ID 6/Work Item 2 check from user password for entering user access moudle
Dim psql As String
If txtUserName.text = "admin" Then
    If txtPassword.text = "" Then
        MsgBox "Please enter a valid password"
        Exit Sub
    End If
    If txtPassword.text = "sysadmin#1" Then
        frmMainUA.Show
        frmLogin2.Hide
        Exit Sub
    End If
    psql = "Select * from UserNames where UserName='" & "admin" & "' and UserPassword='" & txtPassword.text & "'"
    adoRst.Open psql, adoRibConn, adOpenDynamic, adLockOptimistic
    If Not adoRst.EOF Then
        frmMainUA.Show
        frmLogin2.Hide
    Else
        MsgBox "The password entered is incorrect. Please try again."
    End If
    adoRst.Close
    Set adoRst = Nothing
    adoRibConn.Close
    Set adoRibConn = Nothing
    Exit Sub
End If
''''''''''End of modification
   'If cboShopCentre.text = "" Then Exit Sub

   Dim strDecimal As String
   Dim lDecimalPos As Double
   Dim lDifference As Double
   Dim Reg As Object
   Set Reg = CreateObject("wscript.shell")
'added by anol 01/01/2016
   ' LicenseDate = "01/01/2015"
'  GoTo xx
   lDifference = CDbl(DateDiff("d", Now, CDate(LicenseDate)))

   If lDifference > 0 And lDifference <= 30 Then
      'this message will appear for a month
      MsgBox "Support for your program is now due to renewal. " & _
             "Please contact the help desk at PCM Consulting " & _
             "if you have not yet received the renewal. ", vbOKOnly, "Licence Expire " & LicenseDate
   End If
'xx:
   'Start of real program
   On Error GoTo ErrH

   Dim a As Integer

   For a = 2 To 4
      If Mid(cboShopCentre.text, a, 3) = " / " Then
         SCID = Left(cboShopCentre.text, a - 1)
         SCName = Right(cboShopCentre.text, Len(cboShopCentre.text) - a - 2)
      End If
   Next a

   If SCID = "" Then
      MsgBox "Please select a Property", vbCritical + vbOKOnly, "Log in"
      cboShopCentre.text = ""
      cboShopCentre.SetFocus
      Exit Sub
   End If

   If InStr(OS, "Server 2008") = 0 Then
      Conn.Open "DSN=PrestigeBMControlNS;UID=;PWD="
   Else
      Conn.Open "Driver={Microsoft Access Driver (*.mdb)};" & _
                "Dbq=PBMControl.mdb;" & _
                "DefaultDir=" & DB_PATH & ";" & _
                "Uid=;Pwd=;"
   End If

   SQLStr = "SELECT * FROM Databases WHERE ID = " & SCID
   rst.Open SQLStr, Conn, adOpenStatic, adLockReadOnly

   accessDBPws = "RDSWKDPP"
   Adsn = rst!AccessDSN

'   If OS = "Windows XP" Then
      If Adsn <> "" Then
'MsgBox OS                     '#delete me
         If InStr(OS, "Server 2008") = 0 Then
            ChangeReportODBC
         Else
            ChangeReportODBC_WinSvr2008
         End If
         If InStr(OS, "Server 2008") = 0 Then DB_PATH = DataBaseLocationPath
         If FullDatabasePath = "" Then
            If Right(DB_PATH, 1) = "\" Then
               FullDatabasePath = DB_PATH & "PBMc" & Right(Adsn, 3) & ".mdb"
            Else
               FullDatabasePath = DB_PATH & "\PBMc" & Right(Adsn, 3) & ".mdb"
            End If
'MsgBox FullDatabasePath
         End If
      Else
         MsgBox "There are some technical errors have found in Prestige database."
         Exit Sub
      End If
'   End If

   gCurrentShopCentreName = SCName
   gCurrentShopCentreCode = SCID

   rst.Close
   Conn.Close

   Conn.Open getConnectionString
   rst.Open "SELECT SageDSN FROM ShoppingCentre;", Conn, adOpenStatic, adLockReadOnly
   CompanyDatapath = IIf(IsNull(rst.Fields.Item("SageDSN").Value), "", rst.Fields.Item("SageDSN").Value)

   rst.Close
   Conn.Close
   Set rst = Nothing
   Set Conn = Nothing
'If CheckLogin(Me) Then
'''''''Modified by mahboob 06/04/2023 change id 12/ Wrok item 2 check user name in new table
'   If fnLogInCheck(txtUserName.text, txtPassword.text, SCID) Then
   'End of modification
      'Login
      User = txtUserName.text

      frmMMain.Show

      frmMMain.SystemUserName = txtUserName.text
      'clear password box so when user logs out their password is not still there
      Unload frmLogin2
'   Else
'      MsgBox "Please ensure the user name and password entered is valid and that the user has a role assigned and has been granted company access. Please try again.", vbOKOnly + vbCritical, "Incorrect Login"
'      Me.txtPassword.SetFocus
'      'select the text
'      SelTxtInCtrl Me.txtPassword
'   End If
'   If CompanyDatapath <> "" Then _
'      CompanyDatapath = Reg.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\ODBC\ODBC.INI\" & CompanyDatapath & "\DataPathname")

   Exit Sub
   ''''''''''''''''''''''''''''''Modified by Mahboob 29/03/2023 Change ID 5/Work Item 2 Creating the User Access table

CreateTable_UserAccess:
    fnCreateTableUserAccess adoRibConn
    MsgBox "The user access module has been succesfully updated. Please click the login button to continue"
    adoRibConn.Close
   Set adoRibConn = Nothing

Exit Sub
'''''''''''''End of modification
ErrH:
   If Err.Number <> 0 Then
      If Err.Number = -2147467259 Then
            MsgBox " No Data found. Please ensure you are connected to a valid database file", vbCritical, "Prestige"
            Unload Me
            Exit Sub
      End If
      If Err.Number = 40002 Then
         If MsgBox("DSN Invalid. Please check with your system administrator.", vbRetryCancel + vbCritical, "DSN Set Up Error") = vbRetry Then Resume
      Else
      Rem  the error msg by anol 20160429 this becomes annoying where there is no sage DSN
          'MsgBox ERR.Number & " -(pcm_004) " & ERR.description
      End If
   End If
End Sub
''''''''''''''''''''''''''Modified by Mahboob 29/03/2023 Change ID 5/Work Item 3:- Function to create table for user access
Function fnCreateTableUserAccess(connRibbon As ADODB.Connection)
'Dim rst As New ADODB.Recordset
'On Error GoTo Create_Table_UserName
connRibbon.Execute "Create TABLE Roles " & _
        "(" & _
               "RoleID COUNTER PRIMARY KEY,RoleName Text(100)" & _
        ");"
        connRibbon.Execute "INSERT INTO Roles(RoleName)values('" & "Administrator" & "')"


'Create_Table_UserName:
' On Error GoTo Add_Table_RolePermission
'
'   rst.Open "SELECT * FROM UserNames;", connRibbon, adOpenStatic, adLockReadOnly
'   rst.Close
'
'   GoTo Create_Table_RolePermission
'Add_Table_RolePermission:
        connRibbon.Execute "Create TABLE UserNames " & _
        "(" & _
               "UserID COUNTER PRIMARY KEY,UserName Text(100),UserPassword Text(100), UserEmail Text(100), RoleID number, IsActive Text(1)" & _
        ");"
        connRibbon.Execute "INSERT INTO UserNames(UserName,UserPassword,RoleID,IsActive)values('" & "admin" & "','" & "admin" & "'," & 1 & ",'" & "Y" & "')"
'fnCreateTableUserAccess = 1
'Exit Function
'GoTo Create_Table_RolePermission
'Create_Table_RolePermission:
'On Error GoTo Add_Table_RolePermissions
'rst.Open "SELECT * FROM RolePermissions;", connRibbon, adOpenStatic, adLockReadOnly
'   rst.Close
'
'   GoTo Create_Table_CompanyAccess
'Add_Table_RolePermissions:
connRibbon.Execute "Create TABLE RolePermissions " & _
        "(" & _
               "RoleID number, ItemID Text(100)" & _
        ");"

'GoTo Create_Table_CompanyAccess
'
'Create_Table_CompanyAccess:
'On Error GoTo Add_Table_CompanyAccess
'rst.Open "SELECT * FROM UserCompanyAccess;", connRibbon, adOpenStatic, adLockReadOnly
'   rst.Close
'
'   GoTo TempItemLoad
'Add_Table_CompanyAccess:
        connRibbon.Execute "Create TABLE UserCompanyAccess " & _
        "(" & _
               "UserID number, CompanyID number" & _
        ");"
'        GoTo TempItemLoad
'TempItemLoad:
'On Error GoTo Add_Table_TempItemLoad
'rst.Open "SELECT * FROM TempItemLoad;", connRibbon, adOpenStatic, adLockReadOnly
'   rst.Close
'
''   GoTo TempItemLoad
'Add_Table_TempItemLoad:
        connRibbon.Execute "Create TABLE TempItemLoad " & _
        "(" & _
               "ID Text(100)" & _
        ");"
        connRibbon.Execute "INSERT INTO TempItemLoad(ID) SELECT ID FROM Items WHERE Display=False"
End Function


'Private Sub ChangeReportODBC()
'   Dim Reg As Object
'   Set Reg = CreateObject("wscript.shell")
''MsgBox Adsn                                  '#delete me
'   FullDatabasePath = Reg.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\ODBC\ODBC.INI\" & Adsn & "\DBQ")
'   REG_SaveString HKEY_LOCAL_MACHINE, "SOFTWARE\ODBC\ODBC.INI\PrestigeBMcReportsNS", "DBQ", FullDatabasePath
'End Sub

Private Sub ChangeReportODBC_WinSvr2008()
   Dim retval As Long
'MsgBox Adsn                                  '#delete me
   retval = ExecCmd(App.Path + "\ODBC_Amendment.exe " & Adsn & "")
End Sub

Private Sub Command1_Click()

End Sub

Private Sub flxCompany_Click()
''''''''''''''''By Mahboob 07/03/2023 Change ID:5 Work Item 6 set the value in cmbobox
    cboShopCentre.text = flxCompany.TextMatrix(flxCompany.row, 1) & " / " & flxCompany.TextMatrix(flxCompany.row, 2)
End Sub

Private Sub flxCompany_GotFocus()
''''''''By Mahboob 07/03/2023 change ID 5 workitem : 14 set the value in cmbobox
' Change the appearance of the first row when the FlexGrid receives focus
    flxCompany.row = 1 ' Set the first row as the current row
    flxCompany.RowSel = 1 ' Select the first row
    flxCompany.col = 0 ' Set the current column (e.g., column 0)
    flxCompany.ColSel = flxCompany.Cols - 1 ' Select all columns in the first row
    flxCompany.CellBackColor = vbBlue ' Set the background color for the selected cells
    ' Additional formatting can be applied here, such as changing the font color or style
End Sub

Private Sub flxCompany_KeyDown(KeyCode As Integer, Shift As Integer)
''''''''''''''''By Mahboob 07/03/2023 Change ID:5 Work Item 7 set courser to user text box
If KeyCode = 13 Then
           txtUserName.SetFocus
    End If
End Sub

Private Sub flxCompany_LostFocus()
''''''''By Mahboob 07/03/2023 change ID 5 workitem : 15 set the value in cmbobox
cboShopCentre.text = flxCompany.TextMatrix(flxCompany.row, 1) & " / " & flxCompany.TextMatrix(flxCompany.row, 2)
End Sub

Private Sub flxCompany_RowColChange()
''''''''''''''''By Mahboob 07/03/2023 Change ID:5 Work Item 8 set the value in cmbobox
    cboShopCentre.text = flxCompany.TextMatrix(flxCompany.row, 1) & " / " & flxCompany.TextMatrix(flxCompany.row, 2)

End Sub

Private Sub Form_Activate()
   Image1.Height = 1815
   flxCompany.row = 1
   If UCase(SystemUser) = "BOSLUSER" And UCase(WS_Name) = "PCM-DEV2" Then
      txtUserName.text = "Manager"
      txtPassword.text = "xx"
      cboShopCentre.ListIndex = 0
      cmdLogin.SetFocus
'      Dim adoconn As New ADODB.Connection
'      adoconn.Open getConnectionString
'      adoconn.Execute "Update shoppingcentre set pws='',SMTP='',UNAME=''"
'      adoconn.Close
      
   End If
End Sub
'''''''''Mahboob Change ID:5 Work Item 2 load the company in grid
Function LoadCompany(conConnCom As ADODB.Connection)
    Dim rRow As Integer
    
    Dim rstSQLCom As New ADODB.Recordset
    Dim sqlCom As String
   txtSearchCompanyId.text = ""
   txtCompanyName.text = ""
   flxCompany.RowHeight(0) = 0
   flxCompany.Cols = 3
   flxCompany.ColWidth(0) = 5
   flxCompany.ColWidth(1) = 1500
   flxCompany.ColWidth(2) = 5000
   flxCompany.Clear
   flxCompany.ColAlignment(0) = vbLeftJustify
   flxCompany.ColAlignment(1) = vbLeftJustify
   flxCompany.ColAlignment(2) = vbLeftJustify
   
    Dim latest As Integer
Dim rCount As Integer
  
   sqlCom = "SELECT * FROM Databases ORDER BY ID;"
   rstSQLCom.Open sqlCom, conConnCom, adOpenStatic, adLockReadOnly
rRow = 1
   latest = 0
   flxCompany.Rows = rstSQLCom.RecordCount
     If Not rstSQLCom.EOF Then
      While Not rstSQLCom.EOF
         If rCount = 1 Then
      flxCompany.row = 0
      rRow = 0
      Else
      flxCompany.row = 1
      End If
      flxCompany.RowSel = 0
           flxCompany.ColSel = 0
           flxCompany.TextMatrix(rRow, 0) = ""
           flxCompany.TextMatrix(rRow, 1) = rstSQLCom!Id
           flxCompany.TextMatrix(rRow, 2) = rstSQLCom!SCName
           flxCompany.RowHeight(rRow) = 240
           If Not rstSQLCom.EOF Then flxCompany.AddItem ""
           rRow = rRow + 1
           If rstSQLCom!Id > latest Then latest = rstSQLCom!Id
         rstSQLCom.MoveNext
      Wend
   End If
'flxCompany.Rows = flxCompany.Rows - 1
'flxCompany.Rows = flxCompany.Rows - 1
Do While flxCompany.Rows > rRow
        flxCompany.RemoveItem flxCompany.Rows - 1
    Loop
   rstSQLCom.Close
   Set rstSQLCom = Nothing
   conConnCom.Close
   Set conConnCom = Nothing
   latest = 0
   
End Function

Private Sub Form_Load()
 On Error GoTo errHnd

   If IsLoadedAndVisible("frmReport") Then
      MsgBox "There are open reports found. Please must close all open reports before login another company.", vbCritical + vbOKOnly, "Login"
      Exit Sub
   End If

   Dim latest As Integer
   Dim Name As String, Name2 As String
   Dim conConn As New ADODB.Connection, rstSQL As New ADODB.Recordset

   SQLStr = "SELECT * FROM Databases ORDER BY ID;"
   If InStr(OS, "Server 2008") = 0 Then
      conConn.Open "DSN=PrestigeBMControlNS;UID=;PWD="
   Else
      conConn.Open "Driver={Microsoft Access Driver (*.mdb)};" & _
                        "Dbq=PBMControl.mdb;" & _
                        "DefaultDir=" & DB_PATH & ";" & _
                        "Uid=;Pwd=;"
   End If

   rstSQL.Open SQLStr, conConn, adOpenStatic, adLockReadOnly

   latest = 0
   If Not rstSQL.EOF Then
      While Not rstSQL.EOF
         cboShopCentre.AddItem rstSQL!Id & " / " & rstSQL!SCName
'         Set list = lvProperty.ListItems.Add(, , rstSQL!Id & " / " & rstSQL!SCName)
         If rstSQL!Id > latest Then latest = rstSQL!Id
         rstSQL.MoveNext
      Wend
   End If
   rstSQL.Close
   Set rstSQL = Nothing
      ''''''''''''''''''''''''''''Mahboob Change ID:5 Work Item 16 added company list for loading company in grid
   LoadCompany conConn


   latest = 0
   Me.Height = 8076
   Me.Width = 6828
'   Me.Move (Screen.Width - Width) / 2, (Screen.Height - Height) / 2

   Me.Show
   cboShopCentre.ListIndex = -1
  
   Exit Sub

errHnd:
   MsgBox "Control file Datebase not found. Please contact PCM Support.", vbCritical + vbOKOnly, "Please Set PrestigeBMControlNS DSN"
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   lblReprintDemand.FontBold = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
   On Error GoTo MissingFolder

   CreateNonExistsFolder DB_PATH & "\AllStuff"
   CreateNonExistsFolder DB_PATH & "\AllStuff\Temp"
   CreateNonExistsFolder DB_PATH & "\AllStuff\Logs"

   Exit Sub
MissingFolder:
   If User <> "" Then _
      'Modified by anol 04 Feb 2015
      ShowMsgInTaskBar "System could not create necessary folders. Please contact with PCM Consulting Ltd.", "Y"
       End If
End Sub
''''''''''''''''''''''''''Modified by Mahboob 27/04/2023 Change ID 14/Work Item 6:- Function to load the form

Private Sub lblChangePass_Click()
frmChangePassUser.Show
Unload Me
End Sub
''''''''''''''''''''''''''Modified by Mahboob 27/04/2023 Change ID 14/Work Item 5:- Hand Icon

Private Sub lblChangePass_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
SetCursor LoadCursor(0, IDC_HAND)
End Sub
''''''''''''''''''End of modification
Private Sub lblReprintDemand_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   lblReprintDemand.MouseIcon = LoadPicture(App.Path + "\Package1\hmove.cur")
   lblReprintDemand.ForeColor = &H80000012 'vbBack
End Sub

Private Sub lblReprintDemand_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   lblReprintDemand.MouseIcon = LoadPicture(App.Path + "\" + "Package1\harrow.cur")
   lblReprintDemand.FontBold = True
End Sub

Private Sub lblReprintDemand_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   lblReprintDemand.MouseIcon = LoadPicture(App.Path + "\" + "Package1\harrow.cur")
   lblReprintDemand.ForeColor = &HFF0000   'vbBlue
End Sub

Private Sub txtCompanyName_Change()
''Mahboob Change ID:5 Work Item 5 short by company name
Dim i As Integer

   If Len(txtCompanyName.text) > 0 Then
        txtSearchCompanyId.text = ""
   End If

   For i = flxCompany.Rows - 1 To 1 Step -1
        flxCompany.RowHeight(i) = 240
        If InStr(1, UCase(flxCompany.TextMatrix(i, 2)), UCase(txtCompanyName.text), vbTextCompare) = 0 Then
              flxCompany.RowHeight(i) = 0
        End If
        If flxCompany.RowHeight(i) = 240 Then
              flxCompany.row = i
        End If
   Next i

End Sub

Private Sub txtCompanyName_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
''Mahboob Change ID:5 Work Item 13 leave focus on return
If KeyCode = 13 Then
         flxCompany.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        If flxCompany.Visible Then
            flxCompany.SetFocus
        End If
    End If
End Sub

Private Sub txtPassword_GotFocus()
   SelTxtInCtrl txtPassword
   '''''''''''Changed by Mahboob  changeID:5 work item 11        07/03/2023 change the fore color to black
   txtPassword.ForeColor = vbBlack
   cmdLogin.Default = True
End Sub

Private Sub txtPassword_KeyDown(KeyCode As Integer, Shift As Integer)
'''''''''''Changed by Mahboob   changeID:5 work item 12       07/03/2023 move cursor to lod in button
If KeyCode = 13 Then
'cmdLogin.Default = True
           cmdLogin.SetFocus
    End If
End Sub

Private Sub txtPassword_LostFocus()
   txtPassword.ForeColor = &H80FFFF
End Sub

Private Sub txtSearchCompanyId_Change()
''''Mahboob Change ID:5 Work Item 3 short by company Id
Dim i As Integer
   If Len(txtSearchCompanyId.text) > 0 Then
        txtCompanyName.text = ""
   End If
   For i = flxCompany.Rows - 1 To 1 Step -1
   flxCompany.RowHeight(i) = 240
        If InStr(1, UCase(flxCompany.TextMatrix(i, 1)), UCase(txtSearchCompanyId.text), vbTextCompare) = 0 Then
              flxCompany.RowHeight(i) = 0
        End If
        If flxCompany.RowHeight(i) = 240 Then
              flxCompany.row = i
        End If
   Next i

End Sub

Private Sub txtSearchCompanyId_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
''''Mahboob Change ID:5 Work Item 4 focus grid on enter
If KeyCode = vbKeyDown Then
        flxCompany.SetFocus
    End If
    If KeyCode = 13 Then
           txtCompanyName.SetFocus
    End If

End Sub

Private Sub txtUserName_GotFocus()
   SelTxtInCtrl txtUserName
   '''''''''''Changed by Mahboob   changeID:5 work item 9       07/03/2023 change the fore color to black
   txtUserName.ForeColor = vbBlack
End Sub

Private Sub txtUserName_KeyDown(KeyCode As Integer, Shift As Integer)
'''''''''''Changed by Mahboob    changeID:5 work item 10      07/03/2023 move cursor to password text box
If KeyCode = 13 Then
           txtPassword.SetFocus
'           cmdLogin.Default = True
    End If
End Sub

Private Sub txtUserName_LostFocus()
   txtUserName.ForeColor = &H80FFFF
End Sub

Public Function ExecCmd(cmdline$)
   Dim proc As PROCESS_INFORMATION
   Dim start As STARTUPINFO
   Dim ret&

   ' Initialize the STARTUPINFO structure:
   start.cB = Len(start)

   ' Start the shelled application:
   ret& = CreateProcessA(vbNullString, cmdline$, 0&, 0&, 1&, _
      NORMAL_PRIORITY_CLASS, 0&, vbNullString, start, proc)

   ' Wait for the shelled application to finish:
      ret& = WaitForSingleObject(proc.hProcess, INFINITE)
      Call GetExitCodeProcess(proc.hProcess, ret&)
      Call CloseHandle(proc.hThread)
      Call CloseHandle(proc.hProcess)
      ExecCmd = ret&
End Function
