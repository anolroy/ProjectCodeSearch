VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmLesseeStatement 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Lessee Statement"
   ClientHeight    =   10470
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11145
   BeginProperty Font 
      Name            =   "Myriad Web"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLesseeStatement.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   10470
   ScaleWidth      =   11145
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "Send &Email"
      Height          =   375
      Left            =   1800
      TabIndex        =   13
      Top             =   5640
      Width           =   1575
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   5280
      TabIndex        =   10
      Top             =   5640
      Width           =   1575
   End
   Begin VB.CommandButton cmdGenerate 
      Caption         =   "&Print"
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   5640
      Width           =   1575
   End
   Begin VB.TextBox txtTenantSearchID 
      Appearance      =   0  'Flat
      BackColor       =   &H80000014&
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   360
      TabIndex        =   2
      Top             =   1065
      Width           =   975
   End
   Begin VB.TextBox txtTenantSearchName 
      Appearance      =   0  'Flat
      BackColor       =   &H80000014&
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   1395
      TabIndex        =   3
      Top             =   1065
      Width           =   2500
   End
   Begin VB.TextBox txtTenantSearchUnitName 
      Appearance      =   0  'Flat
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   3915
      TabIndex        =   4
      Top             =   1065
      Width           =   2175
   End
   Begin VB.OptionButton optAllTran 
      Caption         =   "&All Transactions"
      Height          =   375
      Left            =   1920
      TabIndex        =   1
      Top             =   420
      Width           =   1695
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxLeaseList 
      Height          =   4140
      Left            =   120
      TabIndex        =   5
      Top             =   1440
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   7303
      _Version        =   393216
      Cols            =   5
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
      _Band(0).Cols   =   5
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.OptionButton optOSOnly 
      Caption         =   "&Outstanding Only"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   420
      Value           =   -1  'True
      Width           =   1695
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00008000&
      BorderWidth     =   5
      X1              =   0
      X2              =   6975
      Y1              =   0
      Y2              =   15
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Index           =   10
      Left            =   6120
      TabIndex        =   12
      Top             =   840
      Width           =   30
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00008000&
      Caption         =   "Produce Lessee Statement"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   375
      Left            =   0
      TabIndex        =   11
      Top             =   40
      Width           =   6975
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Unit Name"
      Height          =   195
      Index           =   2
      Left            =   3915
      TabIndex        =   8
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Lessee"
      Height          =   195
      Index           =   1
      Left            =   1395
      TabIndex        =   7
      Top             =   840
      Width           =   495
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ID"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   435
      TabIndex        =   6
      Top             =   840
      Width           =   165
   End
   Begin VB.Shape Shape4 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      BorderStyle     =   6  'Inside Solid
      FillColor       =   &H00C0FFFF&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   6
      Left            =   120
      Top             =   840
      Width           =   6480
   End
End
Attribute VB_Name = "frmLesseeStatement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private szLesseeList As String

Private Sub LoadFlxLeaseList(adoConn As ADODB.Connection)
   Me.MousePointer = vbHourglass

   Dim szSQL As String

   ConfigureFlxLeaseList
   
   szSQL = "SELECT Tenants.SageAccountNumber, Name, LeaseDetails.UnitNumber, " & _
               "IIF(Tenants.InvoiceTo = 'H', Tenants.Email1, Tenants.Email2) AS Email " & _
           "From   Tenants, LeaseDetails " & _
           "WHERE ((Tenants.Comments) IS NULL OR Tenants.Comments='') AND " & _
               "Tenants.SageAccountNumber = LeaseDetails.SageAccountNumber AND " & _
               "LeaseDetails.Status = True " & _
          "ORDER BY Tenants.SageAccountNumber;"

   PopulateTenantLookup adoConn, szSQL

'   UpdateBalance

   txtTenantSearchID.text = ""
   txtTenantSearchName.text = ""
   txtTenantSearchUnitName.text = ""

   Me.MousePointer = vbArrow
End Sub

Private Sub ConfigureFlxLeaseList()
   Dim szHeader As String

   flxLeaseList.Width = 6735
   Label20(0).Caption = "Lessee Id"
   Label20(1).Caption = "Lessee Name"
   Label20(0).Width = 690
   Label20(0).Left = 435
   Label20(1).Width = 930
   Label20(1).Left = 1560
   Label20(2).Visible = True
   txtTenantSearchUnitName.Visible = True
   Shape4(6).Width = 6480

   flxLeaseList.Clear
   flxLeaseList.Cols = 6
   flxLeaseList.RowHeight(0) = 0
   szHeader$ = "|<Lessee ID|<Lessee Name|<Lessee Name|>Ac Balance|Email"
   flxLeaseList.FormatString = szHeader$
   flxLeaseList.ColWidth(0) = Label20(0).Left - flxLeaseList.Left       '0        Solid column
   flxLeaseList.ColWidth(1) = Label20(1).Left - Label20(0).Left - 20    '1400    'Tenant ID
   flxLeaseList.ColWidth(2) = Label20(2).Left - Label20(1).Left - 20             'Tenant Name
   flxLeaseList.ColWidth(3) = Label20(10).Left - Label20(2).Left - 20            'Unit Name
   flxLeaseList.ColWidth(4) = 0 'Ac Balance
   flxLeaseList.ColWidth(5) = 0                                                  'Email
   flxLeaseList.Rows = 2
End Sub

Public Function PopulateTenantLookup(adoConn As ADODB.Connection, ByVal sSQLQuery_ As String)
   Dim adoRst As New ADODB.Recordset

   adoRst.Open sSQLQuery_, adoConn, adOpenStatic, adLockReadOnly

   Dim iRow As Integer
   iRow = 1

   While Not adoRst.EOF
      If Not IsNull(adoRst!EMail) Then
         flxLeaseList.TextMatrix(iRow, 1) = adoRst!SageAccountNumber
         flxLeaseList.TextMatrix(iRow, 2) = adoRst!Name
         flxLeaseList.TextMatrix(iRow, 3) = adoRst!UnitNumber
         flxLeaseList.TextMatrix(iRow, 5) = adoRst!EMail

         iRow = iRow + 1
         adoRst.MoveNext

         If Not adoRst.EOF Then flxLeaseList.AddItem ""
      Else
         adoRst.MoveNext
      End If
   Wend
   adoRst.Close
   Set adoRst = Nothing
End Function

Private Sub cmdClose_Click()
   Unload Me
End Sub

Private Sub flxLeaseList_Click()
   SelectFlxGridRow 0, flxLeaseList, flxLeaseList.row
End Sub

Private Sub CreateLesseeList()
   Dim i As Integer

   For i = 1 To flxLeaseList.Rows - 1
      If flxLeaseList.TextMatrix(i, 0) = "X" And flxLeaseList.RowHeight(i) > 0 Then
         szLesseeList = szLesseeList & "'" & flxLeaseList.TextMatrix(i, 1) & "'" & ", "
      End If
   Next i
   If Len(szLesseeList) > 0 Then szLesseeList = Left(szLesseeList, Len(szLesseeList) - 2)
End Sub

Private Sub cmdGenerate_Click()
   szLesseeList = ""
   CreateLesseeList
   
   If Len(szLesseeList) = 0 Then Exit Sub

   Dim adoConn As New ADODB.Connection
   Dim szSQL   As String
   Dim reportApp As New CRAXDRT.Application
   Dim Report As CRAXDRT.Report
   Dim rep As frmReport

   adoConn.Open getConnectionString
   
   szSQL = "UPDATE Tenants " & _
           "SET    spare2 = '' " & _
           "WHERE  spare2 = 'Y';"
   adoConn.Execute szSQL

   szSQL = "UPDATE Tenants " & _
           "SET    spare2 = 'Y' " & _
           "WHERE  SageAccountNumber IN (" & szLesseeList & ");"
   adoConn.Execute szSQL

   Set Report = reportApp.OpenReport(App.Path & szReportPath & "\LesseeStatement.rpt")
   Report.Database.Tables(1).ConnectionProperties.Item("Database Password") = accessDBPws

   Report.EnableParameterPrompting = False
   Report.DiscardSavedData

   If optOSOnly.Value Then
      Report.ParameterFields(1).AddCurrentValue False
   Else
      Report.ParameterFields(1).AddCurrentValue True
   End If
   Report.ParameterFields(2).AddCurrentValue False
   Report.ParameterFields(3).AddCurrentValue "1"

   Set rep = New frmReport
   Load rep
   rep.LoadReportViewer Report

   adoConn.Close
   Set adoConn = Nothing
End Sub

Private Sub Form_Load()
   Dim adoConn As New ADODB.Connection

   Me.Width = 7065
   Me.Height = 6585
   Me.BackColor = MODULEBACKCOLOR
   optOSOnly.BackColor = Me.BackColor
   optAllTran.BackColor = Me.BackColor

   adoConn.Open getConnectionString
   
   LoadFlxLeaseList adoConn

   adoConn.Close
   Set adoConn = Nothing

   Call WheelHook(Me.hWnd)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call WheelUnHook(Me.hWnd)
End Sub

Private Sub txtTenantSearchID_Change()
   Dim i As Integer

   If Len(txtTenantSearchID.text) > 0 Then
      txtTenantSearchName.text = ""
      txtTenantSearchUnitName.text = ""
   End If

   For i = 1 To flxLeaseList.Rows - 1
      flxLeaseList.RowHeight(i) = 240
      If UCase(Left(flxLeaseList.TextMatrix(i, 1), Len(txtTenantSearchID.text))) <> UCase(txtTenantSearchID.text) Then
         flxLeaseList.RowHeight(i) = 0
      End If
   Next i
End Sub

Private Sub txtTenantSearchName_Change()
   Dim i As Integer

   If Len(txtTenantSearchName.text) > 0 Then
      txtTenantSearchID.text = ""
      txtTenantSearchUnitName.text = ""
   End If

   For i = 1 To flxLeaseList.Rows - 1
      flxLeaseList.RowHeight(i) = 240
      If UCase(Left(flxLeaseList.TextMatrix(i, 2), Len(txtTenantSearchName.text))) <> UCase(txtTenantSearchName.text) Then
         flxLeaseList.RowHeight(i) = 0
      End If
   Next i
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
