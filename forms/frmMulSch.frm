VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmMulSch 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Multiple Schedules"
   ClientHeight    =   4110
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6990
   Icon            =   "frmMulSch.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4110
   ScaleWidth      =   6990
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtPercentage 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   1200
      TabIndex        =   10
      Top             =   3360
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5760
      TabIndex        =   4
      Top             =   3120
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Done"
      Height          =   375
      Left            =   5640
      TabIndex        =   3
      Top             =   3600
      Width           =   1215
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxSchedule 
      Height          =   2370
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   4180
      _Version        =   393216
      FixedCols       =   0
      BackColorFixed  =   13553358
      ForeColorFixed  =   12632256
      BackColorSel    =   14737632
      ForeColorSel    =   -2147483640
      BackColorBkg    =   16777215
      GridColor       =   -2147483638
      WordWrap        =   -1  'True
      GridLinesFixed  =   1
      SelectionMode   =   1
      Appearance      =   0
      BandDisplay     =   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Total:"
      Height          =   255
      Left            =   4680
      TabIndex        =   9
      Top             =   3120
      Width           =   975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000016&
      Caption         =   "Value"
      Height          =   195
      Index           =   4
      Left            =   5520
      TabIndex        =   8
      Top             =   480
      Width           =   1365
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000016&
      Caption         =   "Percentage"
      Height          =   195
      Index           =   3
      Left            =   4425
      TabIndex        =   5
      Top             =   480
      Width           =   1305
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000016&
      Caption         =   "Name"
      Height          =   195
      Index           =   2
      Left            =   1095
      TabIndex        =   6
      Top             =   480
      Width           =   3645
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000016&
      Caption         =   "No."
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   7
      Top             =   480
      Width           =   1320
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Lessee:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   735
   End
   Begin VB.Label lblLessee 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   960
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmMulSch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
   Me.Left = frmLease4.tabLease.Left + frmLease4.Frame1(2).Left + frmLease4.txtSCFundCode.Left
   Me.Top = frmLease4.tabLease.Top + frmLease4.Frame1(2).Top + frmLease4.cboSchedule.Top + frmLease4.cboSchedule.Height
   Me.BackColor = MODULEBACKCOLOR

   Dim adoConn As New ADODB.Connection
   Dim szSQL As String

   adoConn.Open getConnectionString
   szSQL = "SELECT ScheduleID, ScheduleName FROM Schedule;"

   ConfigureFlxSchedule
   populateGridDefinedHeader adoConn, szSQL, flxSchedule

   adoConn.Close
   Set adoConn = Nothing

   Call WheelHook(Me.hWnd)
End Sub

Private Sub ConfigureFlxSchedule()
   Dim szHeader As String, iCol As Integer

   flxSchedule.Clear
   flxSchedule.Cols = 4
   flxSchedule.Rows = 2
   flxSchedule.RowHeight(0) = 0
   szHeader$ = "<ScheduleID|<ScheduleName|>Percentage|>Value"
   flxSchedule.FormatString = szHeader$

   flxSchedule.ColWidth(0) = Label1(2).Left - Label1(1).Left
   flxSchedule.ColWidth(1) = Label1(3).Left - Label1(2).Left
   flxSchedule.ColWidth(2) = Label1(4).Left - Label1(3).Left
   flxSchedule.ColWidth(3) = flxSchedule.Width + flxSchedule.Left - Label1(4).Left - 160
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call WheelUnHook(Me.hWnd)
   frmLease4.Enabled = True
End Sub

' Here you can add scrolling support to controls that don't normally respond.
' This Sub could always be moved to a module to make scrollwheel behaviour
' generic across forms.
' ===========================================================================
Public Sub MouseWheel(ByVal MouseKeys As Long, ByVal Rotation As Long, ByVal Xpos As Long, ByVal Ypos As Long)
  Dim ctl As Control
  Dim bHandled As Boolean
  Dim bOver As Boolean

  For Each ctl In Controls
    ' Is the mouse over the control
    On Error Resume Next
    bOver = (ctl.Visible And IsOver(ctl.hWnd, Xpos, Ypos))
    On Error GoTo 0

    If bOver Then
      ' If so, respond accordingly
      bHandled = True
      Select Case True

        Case TypeOf ctl Is MSHFlexGrid
          FlexGridScroll ctl, MouseKeys, Rotation, Xpos, Ypos

        Case TypeOf ctl Is PictureBox
          PictureBoxZoom ctl, MouseKeys, Rotation, Xpos, Ypos

        Case TypeOf ctl Is ListBox, TypeOf ctl Is TextBox, TypeOf ctl Is ComboBox
          ' These controls already handle the mousewheel themselves, so allow them to:
          If ctl.Enabled Then ctl.SetFocus

        Case Else
          bHandled = False

      End Select
      If bHandled Then Exit Sub
    End If
    bOver = False
  Next ctl
'
'  ' Scroll was not handled by any controls, so treat as a general message send to the form
'  Me.Caption = "Form Scroll " & IIf(Rotation < 0, "Down", "Up")
End Sub
