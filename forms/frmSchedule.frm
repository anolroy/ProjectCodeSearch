VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmSchedule 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Schedule"
   ClientHeight    =   6810
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9345
   BeginProperty Font 
      Name            =   "Myriad Web"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSchedule.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6810
   ScaleWidth      =   9345
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3136
      TabIndex        =   5
      Top             =   4680
      Width           =   1000
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Cl&ose"
      Height          =   375
      Left            =   6035
      TabIndex        =   7
      Top             =   4680
      Width           =   1000
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4584
      TabIndex        =   6
      Top             =   4680
      Width           =   1000
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "&Edit"
      Height          =   375
      Left            =   1688
      TabIndex        =   4
      Top             =   4680
      Width           =   1000
   End
   Begin VB.CommandButton cmdAddNew 
      Caption         =   "Add &New"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   4680
      Width           =   1000
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxSchedule 
      Height          =   3015
      Left            =   255
      TabIndex        =   0
      Top             =   480
      Width           =   6795
      _ExtentX        =   11986
      _ExtentY        =   5318
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
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Schedule Code:"
      Height          =   195
      Index           =   5
      Left            =   240
      TabIndex        =   14
      Top             =   3840
      Width           =   1095
   End
   Begin MSForms.TextBox txtScheduleCode 
      Height          =   315
      Left            =   1620
      TabIndex        =   1
      Top             =   3840
      Width           =   5415
      VariousPropertyBits=   746604571
      MaxLength       =   15
      BorderStyle     =   1
      Size            =   "9551;556"
      SpecialEffect   =   0
      FontName        =   "Myriad Web"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000016&
      Caption         =   "Date"
      Height          =   195
      Index           =   4
      Left            =   5760
      TabIndex        =   13
      Top             =   240
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000016&
      Caption         =   "Name"
      Height          =   195
      Index           =   3
      Left            =   1920
      TabIndex        =   12
      Top             =   240
      Width           =   5115
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFC0C0&
      BorderWidth     =   2
      Height          =   1455
      Index           =   3
      Left            =   120
      Top             =   3720
      Width           =   7095
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   3
      Left            =   1455
      TabIndex        =   11
      Top             =   4200
      Visible         =   0   'False
      Width           =   30
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF8080&
      BorderWidth     =   3
      Height          =   1455
      Index           =   2
      Left            =   120
      Top             =   3720
      Width           =   7095
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000004&
      Height          =   3495
      Index           =   1
      Left            =   120
      Top             =   120
      Width           =   7095
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   3
      Height          =   3495
      Index           =   0
      Left            =   120
      Top             =   120
      Width           =   7095
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000016&
      Caption         =   "Schedule Code"
      Height          =   195
      Index           =   2
      Left            =   750
      TabIndex        =   10
      Top             =   240
      Width           =   3345
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000016&
      Caption         =   "No."
      Height          =   195
      Index           =   1
      Left            =   260
      TabIndex        =   9
      Top             =   240
      Width           =   1320
   End
   Begin MSForms.TextBox txtName 
      Height          =   315
      Left            =   1620
      TabIndex        =   2
      Top             =   4200
      Width           =   5415
      VariousPropertyBits=   746604571
      MaxLength       =   100
      BorderStyle     =   1
      Size            =   "9551;556"
      SpecialEffect   =   0
      FontName        =   "Myriad Web"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Schedule Name:"
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   8
      Top             =   4200
      Width           =   1125
   End
End
Attribute VB_Name = "frmSchedule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private iNewEdit As Byte

Private Sub cmdAddNew_Click()
   cmdAddNew.Enabled = False
   cmdEdit.Enabled = False
   cmdSave.Enabled = True
   cmdCancel.Enabled = True
   flxSchedule.Enabled = False

   iNewEdit = 1
   Label2(3).Caption = ""
   txtName.text = ""
   txtScheduleCode.text = ""
   txtScheduleCode.SetFocus
End Sub

Private Sub cmdCancel_Click()
   If MsgBox("Do you like to discard the changes?", vbQuestion + vbYesNo, "Schedule") = vbNo Then Exit Sub

   cmdAddNew.Enabled = True
   cmdEdit.Enabled = True
   cmdSave.Enabled = False
   cmdCancel.Enabled = False
   flxSchedule.Enabled = True
   txtName.text = ""
   txtScheduleCode.text = ""
   Label2(3).Caption = ""

   iNewEdit = 0
End Sub

Private Sub cmdClose_Click()
   If iNewEdit <> 0 Then
      If MsgBox("Do you want to close without saving the changes?", vbQuestion + vbYesNo, "Schedule") = vbNo Then Exit Sub
   End If

   Unload Me
End Sub

Private Sub cmdEdit_Click()
   If flxSchedule.row = 0 Then
      ShowMsgInTaskBar "Please select a Schedule from the grid."
      Exit Sub
   End If

   cmdAddNew.Enabled = False
   cmdEdit.Enabled = False
   cmdSave.Enabled = True
   cmdCancel.Enabled = True
   flxSchedule.Enabled = False

   iNewEdit = 2

   Label2(3).Caption = flxSchedule.TextMatrix(flxSchedule.row, 0)
   txtScheduleCode.text = flxSchedule.TextMatrix(flxSchedule.row, 1)
   txtName.text = flxSchedule.TextMatrix(flxSchedule.row, 2)
   SelTxtInCtrl txtScheduleCode
   txtScheduleCode.SetFocus
End Sub

Private Sub cmdSave_Click()
   If iNewEdit = 0 Then Exit Sub
   If Trim(txtName.text) = "" Then
      ShowMsgInTaskBar "Please type the Name of the Schedule."
      txtName.text = ""
      txtScheduleCode.text = ""
      Label2(3).Caption = ""
      txtName.SetFocus
      Exit Sub
   End If

   On Error GoTo ErrorHandler

   Dim adoConn As New ADODB.Connection
   Dim szSQL As String

   adoConn.Open getConnectionString

   If iNewEdit = 1 Then
      szSQL = "INSERT INTO Schedule (ScheduleCode, ScheduleName, InsertDate) " & _
              "VALUES ('" & UCase(txtScheduleCode.text) & "', '" & txtName.text & "', #" & Format(Date, "dd mmmm yyyy") & "#);"
'Debug.Print szSQL
      adoConn.Execute szSQL
      ShowMsgInTaskBar "Schedule has been added successfully."
   End If
   If iNewEdit = 2 Then
      szSQL = "UPDATE Schedule " & _
              "SET ScheduleCode = '" & txtScheduleCode.text & "', " & _
                  "ScheduleName = '" & txtName.text & "', " & _
                  "InsertDate = #" & Format(Date, "dd mmmm yyyy") & "# " & _
              "WHERE ScheduleID = " & Val(flxSchedule.TextMatrix(flxSchedule.row, 0)) & ";"
      adoConn.Execute szSQL
      ShowMsgInTaskBar "Schedule has been edited successfully."
   End If

   cmdAddNew.Enabled = True
   cmdEdit.Enabled = True
   cmdSave.Enabled = False
   cmdCancel.Enabled = False
   flxSchedule.Enabled = True
   Label2(3).Caption = ""
   txtName.text = ""
   txtScheduleCode.text = ""

   szSQL = "SELECT ScheduleID, ScheduleCode, ScheduleName, InsertDate FROM Schedule;"
   populateGridDefinedHeader adoConn, szSQL, flxSchedule

   adoConn.Close
   Set adoConn = Nothing

   iNewEdit = 0

   Exit Sub
ErrorHandler:
   If iNewEdit = 1 Then ShowMsgInTaskBar "System could not add new Schedule." & Chr(13) & Err.Number & " " & Err.description, , "Y"
   If iNewEdit = 2 Then ShowMsgInTaskBar "System could not edit Schedule." & Chr(13) & Err.Number & " " & Err.description, , "Y"
End Sub

Private Sub Form_Load()
   frmMMain.Arrange vbCascade
   Me.ZOrder 0
   Me.Width = 7410
   Me.Height = 5745
   Me.BackColor = MODULEBACKCOLOR

   Dim adoConn As New ADODB.Connection
   Dim szSQL As String

   adoConn.Open getConnectionString
   szSQL = "SELECT ScheduleID, ScheduleCode, ScheduleName, InsertDate FROM Schedule;"

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
   szHeader$ = "<ScheduleID|<ScheduleCode|<ScheduleName|<Date"
   flxSchedule.FormatString = szHeader$

   flxSchedule.ColWidth(0) = Label1(2).Left - Label1(1).Left
   flxSchedule.ColWidth(1) = Label1(3).Left - Label1(2).Left
   flxSchedule.ColWidth(2) = (Label1(4).Left - Label1(3).Left) + flxSchedule.Width + flxSchedule.Left - Label1(4).Left - 120 'Label1(4).Left - Label1(3).Left
   flxSchedule.ColWidth(3) = 0 'flxSchedule.Width + flxSchedule.Left - Label1(4).Left - 120
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call WheelUnHook(Me.hWnd)
   'frmMMain.fraCmdButton.Enabled = True
   Unload Me
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
