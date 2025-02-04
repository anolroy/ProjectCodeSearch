VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmControlCode 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Control Code"
   ClientHeight    =   10890
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12705
   BeginProperty Font 
      Name            =   "Myriad Web"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmControlCode.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   10890
   ScaleWidth      =   12705
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2720
      TabIndex        =   8
      Top             =   6480
      Width           =   1485
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Cl&ose"
      Height          =   375
      Left            =   7920
      TabIndex        =   7
      Top             =   6480
      Width           =   1485
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5320
      TabIndex        =   6
      Top             =   6480
      Width           =   1485
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "&Edit"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   6480
      Width           =   1485
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxTransactionTypes 
      Height          =   5655
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   9285
      _ExtentX        =   16378
      _ExtentY        =   9975
      _Version        =   393216
      Cols            =   4
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
      _Band(0).Cols   =   4
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000016&
      Caption         =   "Control Code"
      Height          =   195
      Index           =   3
      Left            =   4920
      TabIndex        =   2
      Top             =   120
      Width           =   4425
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000016&
      Caption         =   "Type"
      Height          =   195
      Index           =   2
      Left            =   1200
      TabIndex        =   3
      Top             =   120
      Width           =   3945
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000016&
      Caption         =   "ID"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   1485
   End
   Begin MSForms.ComboBox cboCC 
      Height          =   315
      Left            =   4905
      TabIndex        =   1
      Top             =   360
      Width           =   4500
      VariousPropertyBits=   1753237529
      DisplayStyle    =   3
      Size            =   "7937;556"
      BoundColumn     =   0
      TextColumn      =   2
      ColumnCount     =   2
      ListRows        =   20
      cColumnInfo     =   1
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      SpecialEffect   =   6
      FontName        =   "Myriad Web"
      FontEffects     =   1073750016
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      Object.Width           =   "3527"
   End
End
Attribute VB_Name = "frmControlCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private iNewEdit As Byte

Private Sub ConfigureFlxTransctionTypes()
   Dim szHeader As String, iCol As Integer

   flxTransactionTypes.Clear
   flxTransactionTypes.Cols = 3
   flxTransactionTypes.Rows = 2
   flxTransactionTypes.RowHeight(0) = 0
   szHeader$ = "<TYPE_ID|<DESCRIPTION|<CtrlCode"
   flxTransactionTypes.FormatString = szHeader$

   flxTransactionTypes.ColWidth(0) = Label1(2).Left - Label1(1).Left
   flxTransactionTypes.ColWidth(1) = Label1(3).Left - Label1(2).Left
   flxTransactionTypes.ColWidth(2) = flxTransactionTypes.Width + flxTransactionTypes.Left - Label1(3).Left - 300
End Sub

Private Sub cmdCancel_Click()
   If MsgBox("Do you like to discard the changes?", vbQuestion + vbYesNo, "Control Code") = vbNo Then Exit Sub

   ControlHanlding DefaultMode
End Sub

Private Sub cmdClose_Click()
   If iNewEdit <> 0 Then
      If MsgBox("Do you want to close without saving the changes?", vbQuestion + vbYesNo, "Control Code") = vbNo Then Exit Sub
   End If

   Unload Me
End Sub

Private Sub cmdEdit_Click()
   If flxTransactionTypes.row = 0 Then
      MsgBox "Please select a Nominal code from the grid.", vbCritical + vbOKOnly, "Control Code"
      Exit Sub
   End If

   ControlHanlding EditMode
End Sub

Private Sub cmdSave_Click()
   If iNewEdit = 0 Then Exit Sub
   
   If Trim(cboCC.text) = "" Then Exit Sub

   On Error GoTo ErrorHandler

   Dim adoConn As New ADODB.Connection
   Dim szSQL As String

   adoConn.Open getConnectionString

   If iNewEdit = 2 Then
      szSQL = "UPDATE tlbTransactionTypes " & _
              "SET CtrlCode = '" & cboCC.Value & "' " & _
              "WHERE TYPE_ID = " & flxTransactionTypes.TextMatrix(flxTransactionTypes.row, 0) & ";"
      adoConn.Execute szSQL
      MsgBox "Control Code has been edited successfully.", vbInformation + vbOKOnly, "Control Code"
   End If

   szSQL = "SELECT TYPE_ID, DESCRIPTION, CtrlCode " & _
           "FROM tlbTransactionTypes AS T;"

   populateGridDefinedHeader adoConn, szSQL, flxTransactionTypes

   adoConn.Close
   Set adoConn = Nothing

   ControlHanlding DefaultMode

   Exit Sub
ErrorHandler:
   If iNewEdit = 2 Then MsgBox "System could not edit Control Code." & Chr(13) & ERR.Number & " " & ERR.description, vbCritical + vbOKOnly, "Error"
   ControlHanlding DefaultMode
End Sub

Private Sub Form_Load()
   Dim szSQL As String
   Dim adoConn As New ADODB.Connection

   adoConn.Open getConnectionString

   Me.Width = 9585
   Me.Height = 7425
   Me.BackColor = MODULEBACKCOLOR

   ConfigureFlxTransctionTypes

   szSQL = "SELECT TYPE_ID, DESCRIPTION, CtrlCode " & _
           "FROM tlbTransactionTypes AS T;"

   populateGridDefinedHeader adoConn, szSQL, flxTransactionTypes

   szSQL = "SELECT Code, Name " & _
           "FROM   NominalLedger;"

   populateCombo adoConn, szSQL, cboCC
   cboCC.ColumnWidths = "40pt;"
   
   adoConn.Close
   Set adoConn = Nothing
   
   Call WheelHook(Me.hWnd)
End Sub

Private Sub ControlHanlding(ByVal mode As ComponentMode)
   Select Case mode
      Case ComponentMode.DefaultMode
         cmdEdit.Enabled = True
         cmdSave.Enabled = False
         cmdCancel.Enabled = False
         cboCC.Enabled = False

         iNewEdit = 0
   
      Case ComponentMode.NewEntryMode
         cmdEdit.Enabled = False
         cmdSave.Enabled = True
         cmdCancel.Enabled = True
         cboCC.Enabled = True

         iNewEdit = 1

      Case ComponentMode.EditMode
         cmdEdit.Enabled = False
         cmdSave.Enabled = True
         cmdCancel.Enabled = True
         cboCC.Enabled = True

         iNewEdit = 2

      Case ComponentMode.SavedMode
         cmdEdit.Enabled = True
         cmdSave.Enabled = False
         cmdCancel.Enabled = False
         cboCC.Enabled = False

         iNewEdit = 0
   End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call WheelUnHook(Me.hWnd)
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
