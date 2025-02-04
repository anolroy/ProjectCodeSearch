VERSION 5.00
Begin VB.Form frmKey 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Renew License"
   ClientHeight    =   1980
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   9015
   BeginProperty Font 
      Name            =   "Myriad Web"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmKey.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1980
   ScaleWidth      =   9015
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtCurExpDate 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   7125
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   1080
      Width           =   1740
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   330
      Left            =   7440
      TabIndex        =   3
      Top             =   1560
      Width           =   1485
   End
   Begin VB.TextBox txtDate 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFEFEF&
      Enabled         =   0   'False
      Height          =   285
      Left            =   1245
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   1080
      Width           =   1740
   End
   Begin VB.CommandButton cmdGenerate 
      Caption         =   "&Set Key"
      Height          =   330
      Left            =   4320
      TabIndex        =   1
      Top             =   1560
      Width           =   1485
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "&Reset"
      Height          =   315
      Left            =   1200
      TabIndex        =   2
      Top             =   1560
      Width           =   1485
   End
   Begin VB.TextBox txtKey 
      Height          =   285
      Left            =   1245
      TabIndex        =   0
      Top             =   165
      Width           =   7620
   End
   Begin VB.TextBox txtName 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFEFEF&
      Enabled         =   0   'False
      Height          =   285
      Left            =   1245
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   622
      Width           =   7620
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Current Licence Expire Date:"
      Height          =   195
      Left            =   4920
      TabIndex        =   10
      Top             =   1080
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Expire Date"
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   1080
      Width           =   1230
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Licence Key"
      Height          =   240
      Left            =   120
      TabIndex        =   6
      Top             =   165
      Width           =   825
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Company Name"
      Height          =   225
      Left            =   120
      TabIndex        =   4
      Top             =   615
      Width           =   1125
   End
End
Attribute VB_Name = "frmKey"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
   Unload Me
End Sub

Private Sub cmdGenerate_Click()
   If MsgBox("Please make sure that it is a valid key. Otherwise you may lost your current registration." & Chr(10) & _
               "Do you want to proceed?", vbExclamation + vbYesNo, "Licence Renewal") = vbNo Then _
      Exit Sub

   Dim key As String
   key = txtKey.text

   Dim prekey As String
   Dim i As Integer
   Dim valu As String

   For i = 1 To Len(key)
      valu = Mid(key, i, 2)
      If Left(valu, 1) = "-" Then
         prekey = prekey & "/"
         i = i - 1
      Else
         prekey = prekey & Chr(CLng(valu))
      End If
      i = i + 2
   Next

   'We've got the prekey, now split
   LicenseName = ""
   LicenseDate = ""

   Dim DAT As String

   For i = 1 To Len(prekey)
      LicenseName = LicenseName & Mid(prekey, i, 1)
      DAT = Mid(prekey, i + 1, 1)
      If IsNumeric(DAT) Then
         LicenseDate = LicenseDate & DAT
      Else
         If DAT = "/" Then
            LicenseDate = LicenseDate & DAT
         Else
            LicenseName = LicenseName & DAT
         End If
      End If
      i = i + 1
   Next

   If IsDate(LicenseDate) Then
      'Key seems ok
      setKey txtKey.text
      cmdGenerate.Enabled = False
      txtName.text = LicenseName
      txtDate.text = LicenseDate
   Else
      MsgBox "You have entered a wrong key", vbCritical, "Wrong Key"
   End If
End Sub

Private Sub setKey(key As String)
    SaveSetting App.ProductName, "License", "Key", key
    SaveSetting App.ProductName, "License", "Diff", DateDiff("d", Now, LicenseDate)
End Sub

Private Sub cmdReset_Click()
    txtDate.text = ""
    txtKey.text = ""
    txtName.text = ""
    cmdGenerate.Enabled = True
    setKey ""
End Sub

Private Sub Form_Load()
   Me.BackColor = MODULEBACKCOLOR
   txtCurExpDate.text = Format(LicenseDate, "dd/mm/yyyy")
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'added by anol 20180117
    If txtKey.text = "" Then
        End
    End If
End Sub

Private Sub txtKey_KeyPress(KeyAscii As Integer)
   LicenceTextKeyPress txtKey, KeyAscii
End Sub

Private Sub txtName_KeyUp(KeyCode As Integer, Shift As Integer)
    txtName.text = UCase(txtName.text)
    txtName.SelStart = Len(txtName.text)
End Sub
