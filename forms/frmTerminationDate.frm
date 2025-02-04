VERSION 5.00
Begin VB.Form frmTerminationDate 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1635
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   4140
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmTerminationDate.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1635
   ScaleWidth      =   4140
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd_OK 
      Caption         =   "&OK"
      Height          =   365
      Left            =   780
      TabIndex        =   1
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton cmd_Cancel 
      Caption         =   "&Cancel"
      Height          =   365
      Left            =   2100
      TabIndex        =   2
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton cmdGridUnitLookup 
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   120
      Width           =   300
   End
   Begin VB.TextBox txtTerminationDate 
      Appearance      =   0  'Flat
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
      Height          =   285
      Left            =   720
      TabIndex        =   0
      Top             =   540
      Width           =   2025
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFDFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Please enter termination Date"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   42
      Left            =   720
      TabIndex        =   4
      Top             =   240
      Width           =   2115
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00008000&
      Height          =   1440
      Index           =   0
      Left            =   120
      Top             =   120
      Width           =   3960
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   2
      Height          =   1455
      Index           =   1
      Left            =   120
      Top             =   120
      Width           =   3975
   End
End
Attribute VB_Name = "frmTerminationDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public szCallingForm As String
Public szClientID    As String
Public SourceOfCalling As String
Public LeaseEndDate As String
Public LeaseOverRide As Boolean

Private Sub cmd_Cancel_Click()
    Unload Me
    If Left(UCase(Trim(frmLease4.Label16(12).Caption)), 6) = "CURRENT" Then
        frmLease4.cmdTerminate.Visible = True
    Else
        frmLease4.cmdTerminate.Visible = False
    End If
End Sub

Private Sub cmd_OK_Click()
    If IsDate(txtTerminationDate.text) Then
        frmLease4.var = txtTerminationDate.text
        If LeaseEndDate <> "" Then
            If DateDiff("d", LeaseEndDate, txtTerminationDate.text) > 0 And LeaseOverRide = False Then
                    MsgBox "Lease termination date cannot be greater than lease end date."
                    txtTerminationDate.text = LeaseEndDate
                    Exit Sub
            End If
        End If
        If LeaseOverRide Then
            frmLease4.txtLeaseEndDate.text = txtTerminationDate.text
            frmLease4.chkOLED.Value = 0
        End If
        frmLease4.Label16(12).Caption = " Terminated (" & txtTerminationDate.text & ")"
        frmLease4.Label16(12).Tag = txtTerminationDate.text
   End If
   Unload Me
   If SourceOfCalling = "" Then
        frmLease4.terminate_lease_ERR2
   End If
   
End Sub

Private Sub cmdGridUnitLookup_Click(Index As Integer)
    Unload Me
    frmLease4.cmdTerminate.Visible = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If szCallingForm = "frmLease4" Then frmLease4.Enabled = True
   
End Sub

Private Sub txtTerminationDate_Change()
   TextBoxChangeDate txtTerminationDate
End Sub

Private Sub txtTerminationDate_GotFocus()
   SelTxtInCtrl txtTerminationDate
End Sub

Private Sub txtTerminationDate_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then Unload Me                                'Escape
   Dim iRet As Integer
   If KeyAscii = 13 Then                                 'Enter
        If IsDate(txtTerminationDate.text) = False Then
            MsgBox "Please Enter a Valid termination Date."
            txtTerminationDate.text = ""
            FocusControl txtTerminationDate
            Exit Sub
        End If
        FocusControl cmd_OK
   End If
   
   TextBoxKeyPrsDate txtTerminationDate, KeyAscii

End Sub



Private Sub txtTerminationDate_LostFocus()
     If txtTerminationDate.text <> "" Then TextBoxFormatDate txtTerminationDate
End Sub
