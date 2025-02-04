VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmFund4RptPay 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Fund"
   ClientHeight    =   3930
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5805
   Icon            =   "frmFund4RptPay.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3930
   ScaleWidth      =   5805
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CancelButton 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4500
      TabIndex        =   1
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "&OK"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   3480
      Width           =   1215
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxFund 
      Height          =   2655
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   5595
      _ExtentX        =   9869
      _ExtentY        =   4683
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
      BackColor       =   &H80000016&
      Caption         =   "Fund Name"
      Height          =   195
      Index           =   3
      Left            =   2520
      TabIndex        =   6
      Top             =   480
      Width           =   3165
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Please select a fund:"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   210
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   1545
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000016&
      Caption         =   "Fund Code"
      Height          =   195
      Index           =   2
      Left            =   735
      TabIndex        =   3
      Top             =   480
      Width           =   1740
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000016&
      Caption         =   "No."
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   1320
   End
End
Attribute VB_Name = "frmFund4RptPay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public szCallerForm As String

Private Sub CancelButton_Click()
   If szCallerForm = "Demand" Then frmDemands3.lFund = -1

   If szCallerForm = "PI" Then frmPurchaseExpense.lFund = -1

   Unload Me
End Sub

Private Sub flxFund_DblClick()
   frmDemands3.lFund = -1
   OKButton_Click
End Sub

Private Sub Form_Activate()
    Dim adoConn As New adodb.Connection
   Dim szSQL As String

   adoConn.Open getConnectionString
   
   
   If szCallerForm = "Demand" And frmDemands3.chkApplyFund.Value = True Then
        szSQL = "SELECT F.FundID, FundCode, FundName FROM FUND F, BankFund C where C.FundId=F.FundID AND C.CLIENTID='" & frmDemands3.txtRptClientList.text & "' ;"
   Else
        szSQL = "SELECT FundID, FundCode, FundName FROM FUND;"
   End If

   ConfigureFlxFund
   populateGridDefinedHeader adoConn, szSQL, flxFund

   adoConn.Close
   Set adoConn = Nothing

   flxFund.row = 0
End Sub

Private Sub Form_Load()
   Me.BackColor = MODULEBACKCOLOR

   
   Call WheelHook(Me.hWnd)
End Sub

Private Sub ConfigureFlxFund()
   Dim szHeader As String, iCol As Integer

   flxFund.Clear
   flxFund.Cols = 3
   flxFund.Rows = 2
   flxFund.row = 1
   flxFund.RowHeight(0) = 0
   szHeader$ = "<FundID|<FundCode|<FundName"
   flxFund.FormatString = szHeader$

   flxFund.ColWidth(0) = Label1(2).Left - Label1(1).Left
   flxFund.ColWidth(1) = Label1(3).Left - Label1(2).Left
   flxFund.ColWidth(2) = flxFund.Width + flxFund.Left - Label1(3).Left - 320
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call WheelUnHook(Me.hWnd)
   If frmDemands3.lFund < 1 Then frmDemands3.lFund = -1
End Sub

Private Sub OKButton_Click()
   If flxFund.row < 1 Then
      MsgBox "Please select a fund.", vbOKOnly + vbExclamation, "Fund Selection"
   End If
   If flxFund.row > 0 And flxFund.TextMatrix(flxFund.row, 0) <> "" Then
      If szCallerForm = "Demand" Then
         frmDemands3.lFund = flxFund.TextMatrix(flxFund.row, 0)
      End If
      If szCallerForm = "PI" Then
         frmPurchaseExpense.lFund = flxFund.TextMatrix(flxFund.row, 0)
      End If

      Unload Me
   End If
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
End Sub
