VERSION 5.00
Begin VB.Form frmAbout 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About "
   ClientHeight    =   5670
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   6915
   ClipControls    =   0   'False
   FillColor       =   &H00D3E3F0&
   BeginProperty Font 
      Name            =   "Myriad Web"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3916.094
   ScaleMode       =   0  'User
   ScaleWidth      =   6488.21
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdLicenceMenu 
      BackColor       =   &H000017B8&
      Caption         =   "Renew Licence"
      Height          =   345
      Index           =   0
      Left            =   4752
      TabIndex        =   31
      Top             =   2544
      Visible         =   0   'False
      Width           =   1597
   End
   Begin VB.CommandButton cmdLicenceMenu 
      Caption         =   "Batch Payment"
      Height          =   345
      Index           =   1
      Left            =   4752
      TabIndex        =   30
      Top             =   2904
      Visible         =   0   'False
      Width           =   1597
   End
   Begin VB.CommandButton cmdLicenceMenu 
      Caption         =   "Export Licence"
      Height          =   345
      Index           =   2
      Left            =   4752
      TabIndex        =   29
      Top             =   3288
      Visible         =   0   'False
      Width           =   1597
   End
   Begin VB.PictureBox PictureLogo 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFDF&
      ForeColor       =   &H80000008&
      Height          =   1860
      Left            =   2340
      Picture         =   "frmAbout.frx":F172
      ScaleHeight     =   1830
      ScaleWidth      =   2310
      TabIndex        =   20
      Top             =   90
      Width           =   2340
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Height          =   2295
      Left            =   11520
      TabIndex        =   7
      Top             =   2463
      Visible         =   0   'False
      Width           =   3255
      Begin VB.Label lblDescription 
         BackStyle       =   0  'Transparent
         Caption         =   "London"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   9
         Left            =   240
         MousePointer    =   4  'Icon
         TabIndex        =   16
         Top             =   992
         Width           =   2805
      End
      Begin VB.Label lblDescription 
         BackStyle       =   0  'Transparent
         Caption         =   "SW9 6DE"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   8
         Left            =   240
         MousePointer    =   4  'Icon
         TabIndex        =   15
         Top             =   1240
         Width           =   2805
      End
      Begin VB.Label lblDescription 
         BackStyle       =   0  'Transparent
         Caption         =   "+44(0)  870  922  2158"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   7
         Left            =   240
         MousePointer    =   4  'Icon
         TabIndex        =   14
         Top             =   1980
         Width           =   1965
      End
      Begin VB.Label lblDescription 
         BackStyle       =   0  'Transparent
         Caption         =   "www.pcmconsulting.com"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   6
         Left            =   240
         MousePointer    =   4  'Icon
         TabIndex        =   13
         Top             =   1736
         Width           =   2805
      End
      Begin VB.Label lblDescription 
         BackStyle       =   0  'Transparent
         Caption         =   "e-Mail: support@pcmuk.net"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   5
         Left            =   240
         MousePointer    =   4  'Icon
         TabIndex        =   12
         Top             =   1488
         Width           =   2805
      End
      Begin VB.Label lblDescription 
         BackStyle       =   0  'Transparent
         Caption         =   "1-3 Brixton Road"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   4
         Left            =   240
         MousePointer    =   4  'Icon
         TabIndex        =   11
         Top             =   720
         Width           =   2805
      End
      Begin VB.Label lblDescription 
         BackStyle       =   0  'Transparent
         Caption         =   "Kennington Park Business Centre"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   3
         Left            =   240
         MousePointer    =   4  'Icon
         TabIndex        =   10
         Top             =   496
         Width           =   2805
      End
      Begin VB.Label lblDescription 
         BackStyle       =   0  'Transparent
         Caption         =   "CH.219, Chester House"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   2
         Left            =   240
         MousePointer    =   4  'Icon
         TabIndex        =   9
         Top             =   248
         Width           =   2805
      End
      Begin VB.Label lblDescription 
         BackStyle       =   0  'Transparent
         Caption         =   "PCM CONSULTING LTD."
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   1
         Left            =   240
         MousePointer    =   4  'Icon
         TabIndex        =   8
         Top             =   0
         Width           =   2805
      End
   End
   Begin VB.CommandButton cmdDropDownMenu 
      Caption         =   "V"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   319
      Left            =   6300
      TabIndex        =   6
      Top             =   2232
      Width           =   256
   End
   Begin VB.CommandButton cmdLicence 
      Caption         =   "Licence"
      Height          =   345
      Left            =   4752
      TabIndex        =   1
      Top             =   2208
      Width           =   1608
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   465
      Left            =   11640
      Picture         =   "frmAbout.frx":11A8F
      ScaleHeight     =   326.585
      ScaleMode       =   0  'User
      ScaleWidth      =   389.795
      TabIndex        =   5
      Top             =   360
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Height          =   345
      Left            =   144
      TabIndex        =   0
      Top             =   5004
      Width           =   1860
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFDF&
      Caption         =   "PCM Consulting Ltd"
      Height          =   192
      Left            =   144
      TabIndex        =   32
      Top             =   2232
      Width           =   3468
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFFFDF&
      Caption         =   "This product is supplied as part of a subscription service"
      Height          =   192
      Left            =   144
      TabIndex        =   28
      Top             =   4500
      Width           =   4044
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFFDF&
      Caption         =   "Copyright © 2011 -2023 PCM Consulting Ltd "
      Height          =   192
      Left            =   144
      TabIndex        =   27
      Top             =   4213
      Width           =   3468
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFDF&
      Caption         =   "Website: www.pcmuk.net"
      Height          =   192
      Left            =   144
      TabIndex        =   26
      Top             =   3930
      Width           =   3468
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFDF&
      Caption         =   "Tel: 0203 474 0574"
      Height          =   192
      Left            =   144
      TabIndex        =   25
      Top             =   3647
      Width           =   3468
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFDF&
      Caption         =   "Email: Support@pcmuk.net"
      Height          =   192
      Left            =   144
      TabIndex        =   24
      Top             =   3364
      Width           =   3468
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFDF&
      Caption         =   "London SW6 4LZ"
      Height          =   192
      Left            =   144
      TabIndex        =   23
      Top             =   3081
      Width           =   3468
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFDF&
      Caption         =   "136-144 New Kings Road"
      Height          =   192
      Left            =   144
      TabIndex        =   22
      Top             =   2798
      Width           =   3468
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFDF&
      Caption         =   "New Kings House"
      Height          =   192
      Left            =   144
      TabIndex        =   21
      Top             =   2515
      Width           =   3468
   End
   Begin VB.Label lblUnits 
      Alignment       =   1  'Right Justify
      Height          =   240
      Left            =   4752
      TabIndex        =   19
      Top             =   4560
      Width           =   1632
   End
   Begin VB.Label lblPropertyCount 
      Alignment       =   1  'Right Justify
      Height          =   240
      Left            =   4752
      TabIndex        =   18
      Top             =   4200
      Width           =   1632
   End
   Begin VB.Label lblClientCount 
      Alignment       =   1  'Right Justify
      Height          =   240
      Left            =   4752
      TabIndex        =   17
      Top             =   3840
      Width           =   1632
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      Visible         =   0   'False
      X1              =   10747.07
      X2              =   16677.94
      Y1              =   993.873
      Y2              =   993.873
   End
   Begin VB.Label lblDescription 
      BackStyle       =   0  'Transparent
      Caption         =   "Property Block Management System"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Index           =   0
      Left            =   12600
      TabIndex        =   2
      Top             =   720
      Visible         =   0   'False
      Width           =   4605
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "PRESTIGE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   600
      Left            =   12600
      TabIndex        =   3
      Top             =   120
      Visible         =   0   'False
      Width           =   3285
   End
   Begin VB.Label lblVersion 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Version"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   228
      Left            =   1368
      TabIndex        =   4
      Top             =   1944
      Width           =   3888
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      Visible         =   0   'False
      X1              =   10747.07
      X2              =   16677.94
      Y1              =   1004.233
      Y2              =   1004.233
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Reg Key Security Options...
Const READ_CONTROL = &H20000
Const KEY_QUERY_VALUE = &H1
Const KEY_SET_VALUE = &H2
Const KEY_CREATE_SUB_KEY = &H4
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const KEY_NOTIFY = &H10
Const KEY_CREATE_LINK = &H20
Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + _
                       KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + _
                       KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
                     
' Reg Key ROOT Types...
Const HKEY_LOCAL_MACHINE = &H80000002
Const ERROR_SUCCESS = 0
Const REG_SZ = 1                         ' Unicode nul terminated string
Const REG_DWORD = 4                      ' 32-bit number

Const gREGKEYSYSINFOLOC = "SOFTWARE\Microsoft\Shared Tools Location"
Const gREGVALSYSINFOLOC = "MSINFO"
Const gREGKEYSYSINFO = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
Const gREGVALSYSINFO = "PATH"

Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal Hkey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal Hkey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal Hkey As Long) As Long

Private Sub cmdSysInfo_Click()
  Call StartSysInfo
End Sub

Private Sub cmdDropDownMenu_Click()
   If cmdDropDownMenu.Caption = "V" Then
      cmdLicenceMenu(0).Visible = True
      cmdLicenceMenu(1).Visible = True
      cmdLicenceMenu(2).Visible = True

      cmdDropDownMenu.Caption = "^"
   Else
      cmdLicenceMenu(0).Visible = False
      cmdLicenceMenu(1).Visible = False
      cmdLicenceMenu(2).Visible = False

      cmdDropDownMenu.Caption = "V"
   End If
End Sub

Private Sub cmdLicence_Click()
   cmdDropDownMenu_Click
End Sub

Private Sub cmdLicence__Click()
MsgBox ";"
End Sub

Private Sub cmdLicenceMenu_Click(Index As Integer)
   Dim szKey As String

   cmdLicenceMenu(0).Visible = False
   cmdLicenceMenu(1).Visible = False
   cmdLicenceMenu(2).Visible = False

   If Index = 0 Then
      Load frmKey
      frmKey.Show 1
   End If
   If Index = 1 Then
      MsgBox "Please close all opened modules.", vbInformation + vbOKOnly, "Module Activation"
      szKey = InputBox("Input the Key:", "Batch Payment Module Licence")

      If szKey = "" Then Exit Sub

      If szKey <> "667515655492843-7670497728507320-4808507653489898570779573695787840" Then
         MsgBox "Key is not valid.", vbInformation + vbOKOnly, "Batch Payment Licence"
      Else
         SaveSetting App.ProductName, "BT", "Module", szKey
         'frmMMain.mnuBatchPayment_.Enabled = True
         MsgBox "Batch Payment module has been activated.", vbInformation + vbOKOnly, "Thank You"
      End If
   End If
   If Index = 2 Then
      MsgBox "Please close all opened modules.", vbInformation + vbOKOnly, "Module Activation"
      szKey = InputBox("Input the Key:", "Export Module Licence")

      If szKey = "" Then Exit Sub

      If szKey <> "807515675492773-7320497678507790-4788507833499858480769483845737780715324762846686" Then
         MsgBox "Key is not valid.", vbInformation + vbOKOnly, "Export Moudle Licence"
      Else
         SaveSetting App.ProductName, "ExpMMS", "Module", szKey
         MsgBox "Export module has been activated.", vbInformation + vbOKOnly, "Thank You"
      End If
   End If
End Sub

Private Sub cmdOK_Click()
   'frmMMain.fraCmdButton.Enabled = True
   Unload Me
End Sub

Private Sub Command1_Click()
   'PopupMenu frmMenu.mnuFile, 0, Me.ScaleLeft, Me.ScaleHeight / 2
End Sub

Private Sub Form_Load()
   Me.Height = 6072
   Me.Width = 6984
   Me.Top = 100 '(frmMMain.Height / 2) - (Me.Height / 2) - 400
   Me.Left = 100 '(frmMMain.Width / 2) - (Me.Width / 2) - 400
   Me.Caption = "About " & App.Title
   Me.BackColor = MODULEBACKCOLOR
   Frame1.BackColor = Me.BackColor
   lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
   Dim adoConn As New ADODB.Connection
   Dim rsClient As New ADODB.Recordset
   Dim rsProperty As New ADODB.Recordset
   Dim rsUnits As New ADODB.Recordset
   adoConn.Open getConnectionString
   rsClient.Open "Select Count(*) as cnt  From Client", adoConn, adOpenForwardOnly, adLockReadOnly
   lblClientCount.Caption = "No. of Clients:  " & rsClient("cnt").Value
   rsClient.Close
   rsProperty.Open "Select Count(*) as cnt  From Property", adoConn, adOpenForwardOnly, adLockReadOnly
   lblPropertyCount = "No. of Properties:  " & rsProperty("cnt").Value
   rsProperty.Close
   rsUnits.Open "Select Count(*) as cnt  From Units", adoConn, adOpenForwardOnly, adLockReadOnly
   lblUnits = "No. of Units: " & rsUnits("cnt").Value
   rsUnits.Close
   
   adoConn.Close
   Set adoConn = Nothing
End Sub

Public Sub StartSysInfo()
    On Error GoTo SysInfoErr
  
    Dim rc As Long
    Dim SysInfoPath As String
    
    ' Try To Get System Info Program Path\Name From Registry...
    If GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFO, gREGVALSYSINFO, SysInfoPath) Then
    ' Try To Get System Info Program Path Only From Registry...
    ElseIf GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFOLOC, gREGVALSYSINFOLOC, SysInfoPath) Then
        ' Validate Existance Of Known 32 Bit File Version
        If (Dir(SysInfoPath & "\MSINFO32.EXE") <> "") Then
            SysInfoPath = SysInfoPath & "\MSINFO32.EXE"
            
        ' Error - File Can Not Be Found...
        Else
            GoTo SysInfoErr
        End If
    ' Error - Registry Entry Can Not Be Found...
    Else
        GoTo SysInfoErr
    End If
    
    Call Shell(SysInfoPath, vbNormalFocus)
    
    Exit Sub
SysInfoErr:
    MsgBox "System Information Is Unavailable At This Time", vbOKOnly
End Sub

Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, ByRef KeyVal As String) As Boolean
    Dim i As Long                                           ' Loop Counter
    Dim rc As Long                                          ' Return Code
    Dim Hkey As Long                                        ' Handle To An Open Registry Key
    Dim hDepth As Long                                      '
    Dim KeyValType As Long                                  ' Data Type Of A Registry Key
    Dim tmpVal As String                                    ' Tempory Storage For A Registry Key Value
    Dim KeyValSize As Long                                  ' Size Of Registry Key Variable
    '------------------------------------------------------------
    ' Open RegKey Under KeyRoot {HKEY_LOCAL_MACHINE...}
    '------------------------------------------------------------
    rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, Hkey) ' Open Registry Key
    
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Error...
    
    tmpVal = String$(1024, 0)                             ' Allocate Variable Space
    KeyValSize = 1024                                       ' Mark Variable Size
    
    '------------------------------------------------------------
    ' Retrieve Registry Key Value...
    '------------------------------------------------------------
    rc = RegQueryValueEx(Hkey, SubKeyRef, 0, _
                         KeyValType, tmpVal, KeyValSize)    ' Get/Create Key Value
                        
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Errors
    
    If (Asc(Mid(tmpVal, KeyValSize, 1)) = 0) Then           ' Win95 Adds Null Terminated String...
        tmpVal = Left(tmpVal, KeyValSize - 1)               ' Null Found, Extract From String
    Else                                                    ' WinNT Does NOT Null Terminate String...
        tmpVal = Left(tmpVal, KeyValSize)                   ' Null Not Found, Extract String Only
    End If
    '------------------------------------------------------------
    ' Determine Key Value Type For Conversion...
    '------------------------------------------------------------
    Select Case KeyValType                                  ' Search Data Types...
    Case REG_SZ                                             ' String Registry Key Data Type
        KeyVal = tmpVal                                     ' Copy String Value
    Case REG_DWORD                                          ' Double Word Registry Key Data Type
        For i = Len(tmpVal) To 1 Step -1                    ' Convert Each Bit
            KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, i, 1)))   ' Build Value Char. By Char.
        Next
        KeyVal = Format$("&h" + KeyVal)                     ' Convert Double Word To String
    End Select
    
    GetKeyValue = True                                      ' Return Success
    rc = RegCloseKey(Hkey)                                  ' Close Registry Key
    Exit Function                                           ' Exit
    
GetKeyError:      ' Cleanup After An Error Has Occured...
    KeyVal = ""                                             ' Set Return Val To Empty String
    GetKeyValue = False                                     ' Return Failure
    rc = RegCloseKey(Hkey)                                  ' Close Registry Key
End Function

Private Sub Form_Unload(Cancel As Integer)
   'frmMMain.fraCmdButton.Enabled = True
   Unload Me
End Sub

Private Sub lblDescription_Click(Index As Integer)
'   Call Shell("C:\Program Files\Internet Explorer\IEXPLORE.EXE http://www.pcmuk.net/", 1)

   Dim ie As Object
   Set ie = CreateObject("INTERNETEXPLORER.APPLICATION")
   ie.Navigate "http://www.pcmuk.net"
   ie.Visible = True

   While ie.busy
   DoEvents
   Wend
End Sub

Private Sub lblDescription_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'   If Index = 1 Then
'      MsgBox App.Path & "\Package1\POINT05.ICO"
'      lblDescription(1).MouseIcon = LoadPicture(App.Path & "\Package1\POINT05.ICO")
'   End If
End Sub

Private Sub picIcon_Click()
   Dim ie As Object
   Set ie = CreateObject("INTERNETEXPLORER.APPLICATION")
   ie.Navigate "http://www.pcmuk.net"
   ie.Visible = True

   While ie.busy
   DoEvents
   Wend
End Sub

Private Sub picIcon_Mouseup(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Button = 1 And Shift = 2 Then _
      MsgBox szDataBaseUpdateStatus, vbOKOnly, "Update Database Status"
End Sub

Private Sub Picture1_Click()
   Dim ie As Object
   Set ie = CreateObject("INTERNETEXPLORER.APPLICATION")
   ie.Navigate "http://www.pcmuk.net"
   ie.Visible = True

   While ie.busy
   DoEvents
   Wend
End Sub
