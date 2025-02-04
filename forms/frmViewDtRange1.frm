VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmViewDtRange 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Rent Review Report"
   ClientHeight    =   3315
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   3840
   BeginProperty Font 
      Name            =   "Myriad Web"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmViewDtRange.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3315
   ScaleWidth      =   3840
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton GenerateReportButton 
      Caption         =   "Generate Report"
      Default         =   -1  'True
      Height          =   375
      Left            =   960
      TabIndex        =   2
      Top             =   2760
      Width           =   1695
   End
   Begin VB.Frame Frame2 
      Caption         =   "Specify Date Range"
      Height          =   1455
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   3615
      Begin VB.TextBox txtToDate 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   3
         EndProperty
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1200
         MaxLength       =   10
         TabIndex        =   1
         Top             =   960
         Width           =   1575
      End
      Begin VB.TextBox txtFromDate 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   3
         EndProperty
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1200
         MaxLength       =   10
         TabIndex        =   0
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "To:"
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   960
         Width           =   210
      End
      Begin VB.Label lblSpecifyDateRange 
         AutoSize        =   -1  'True
         Caption         =   "From:"
         Height          =   195
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   390
      End
   End
   Begin MSForms.ComboBox cboPropertyList 
      Height          =   315
      Left            =   810
      TabIndex        =   9
      Top             =   600
      Width           =   2925
      VariousPropertyBits=   1753237531
      DisplayStyle    =   3
      Size            =   "5159;556"
      BoundColumn     =   0
      TextColumn      =   1
      ColumnCount     =   3
      ListRows        =   20
      cColumnInfo     =   2
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      SpecialEffect   =   6
      FontName        =   "Myriad Web"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      Object.Width           =   "1411;4233"
   End
   Begin MSForms.ComboBox cboClientList 
      Height          =   315
      Left            =   810
      TabIndex        =   8
      Top             =   80
      Width           =   2925
      VariousPropertyBits=   1753237531
      DisplayStyle    =   3
      Size            =   "5159;556"
      BoundColumn     =   0
      TextColumn      =   1
      ColumnCount     =   8
      ListRows        =   20
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      SpecialEffect   =   6
      FontName        =   "Myriad Web"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label44 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Client:"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   7
      Top             =   80
      Width           =   465
   End
   Begin VB.Label Label44 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Property:"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   6
      Top             =   600
      Width           =   645
   End
End
Attribute VB_Name = "frmViewDtRange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
   Dim adoConn As New ADODB.Connection
'   connect to database
   adoConn.Open getConnectionString
   
   PrepareList adoConn, cboClientList, cboPropertyList
   
   txtToDate.text = Format(Now, "dd/mm/yyyy")
   txtFromDate.text = Format(Now, "dd/mm/yyyy")

   Me.Top = 0
   Me.Left = 0

   adoConn.Close
   Set adoConn = Nothing
End Sub

Private Sub PrepareList(adoConn As ADODB.Connection, cboClient As Control, cboProperty As Control)
   Dim adoRST As New ADODB.Recordset
   Dim szSQL As String

   On Error GoTo ErrorHandler

'*************************************** CLIENT COMBO ******************************************
   szSQL = "SELECT CLIENTID, CLIENTNAME, CLIENTPOSTCODE,  " & _
               "LandLordSageCustAC, LandLordSageSuppAC " & _
           "FROM CLIENT " & _
           "ORDER BY CLIENTNAME;"

   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If adoRST.EOF Then
      adoRST.Close
      Set adoRST = Nothing
      Exit Sub
   End If

   Dim TotalRow As Integer, TotalCol As Integer
   Dim Data() As String
   Dim i As Integer, j As Integer

   TotalRow = adoRST.RecordCount
   TotalCol = adoRST.Fields.count - 1

   ReDim Data(TotalCol, TotalRow) As String

   Data(0, 0) = "ALL"
   Data(1, 0) = "All Clients"
   For i = 1 To TotalRow
       For j = 0 To TotalCol - 1
           Data(j, i) = IIf(IsNull(adoRST.Fields(j).Value), "", adoRST.Fields(j).Value)
       Next j
       adoRST.MoveNext
       If adoRST.EOF Then Exit For
   Next i

   cboClient.Column() = Data()
   cboClient.ListIndex = 0
   adoRST.Close
'*************************************** PROPERTY ******************************************
   szSQL = "SELECT PropertyID, PropertyName, " & _
               "ProAddressLine1, ProPostCode " & _
           "FROM Property " & _
           "ORDER BY PropertyID;"

   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If adoRST.EOF Then
      adoRST.Close
      Set adoRST = Nothing
      Exit Sub
   End If

   TotalRow = adoRST.RecordCount
   TotalCol = adoRST.Fields.count - 1

   ReDim Data(TotalCol, TotalRow) As String

   Data(0, 0) = "ALL"
   Data(1, 0) = "All Properties"
   For i = 1 To TotalRow
      For j = 0 To TotalCol - 1
         Data(j, i) = IIf(IsNull(adoRST.Fields(j).Value), "", adoRST.Fields(j).Value)
      Next j
      adoRST.MoveNext
      If adoRST.EOF Then Exit For
   Next i
   cboProperty.Column() = Data()
   cboProperty.ListIndex = 0

   adoRST.Close
   Set adoRST = Nothing
   Exit Sub

ErrorHandler:
   MsgBox ERR.description & "::" & ERR.Number

   adoRST.Close
   Set adoRST = Nothing
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Me.MousePointer = vbArrow
End Sub

Private Sub Form_Unload(Cancel As Integer)
   frmMMain.fraCmdButton.Enabled = True
   Unload Me
End Sub

Private Sub Frame2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Frame2.MousePointer = vbArrow
End Sub

Private Sub GenerateReportButton_Click()
   Dim datetype As Integer

   If txtFromDate.text = "" Then
      MsgBox "Please enter the to date?", vbExclamation + vbOKOnly, "Enter Date"
      txtFromDate.SetFocus
      Exit Sub
   End If

   If txtToDate.text = "" Then
      MsgBox "Please enter the to date?", vbExclamation + vbOKOnly, "Enter Date"
      txtToDate.SetFocus
      Exit Sub
   End If

   Dim reportApp As New CRAXDRT.Application
   Dim Report As CRAXDRT.Report

   If cboClientList.ListIndex = 0 And cboPropertyList.ListIndex = 0 Then      'all CLIENTs and all PROPERTY
      Set Report = reportApp.OpenReport(App.Path & szReportPath & "\RentReviewReport.rpt")

      Report.Database.Tables(1).ConnectionProperties.Item("Database Password") = accessDBPws

      Report.EnableParameterPrompting = False
      Report.DiscardSavedData

      ' Passing the from and to date values to Crystal Reports
      Report.ParameterFields(1).AddCurrentValue CDate(txtFromDate.text)
      Report.ParameterFields(2).AddCurrentValue CDate(txtToDate.text)
   End If

   If cboClientList.ListIndex > 0 And cboPropertyList.ListIndex = 0 Then      'a CLIENT and all PROPERTY
      Set Report = reportApp.OpenReport(App.Path & szReportPath & "\RentReview_C.rpt")

      Report.Database.Tables(1).ConnectionProperties.Item("Database Password") = accessDBPws

      Report.EnableParameterPrompting = False
      Report.DiscardSavedData

      ' Passing the from and to date values to Crystal Reports
      Report.ParameterFields(1).AddCurrentValue CDate(txtFromDate.text)
      Report.ParameterFields(2).AddCurrentValue CDate(txtToDate.text)
      Report.ParameterFields(3).AddCurrentValue cboClientList.Column(0)
   End If

   If cboClientList.ListIndex = 0 And cboPropertyList.ListIndex > 0 Then      'all CLIENTs and a PROPERTY
      Set Report = reportApp.OpenReport(App.Path & szReportPath & "\RentReview_P.rpt")

      Report.Database.Tables(1).ConnectionProperties.Item("Database Password") = accessDBPws

      Report.EnableParameterPrompting = False
      Report.DiscardSavedData

      ' Passing the from and to date values to Crystal Reports
      Report.ParameterFields(1).AddCurrentValue CDate(txtFromDate.text)
      Report.ParameterFields(2).AddCurrentValue CDate(txtToDate.text)
      Report.ParameterFields(3).AddCurrentValue cboPropertyList.Column(0)
   End If

   If cboClientList.ListIndex > 0 And cboPropertyList.ListIndex > 0 Then      'a CLIENT and a PROPERTY
      Set Report = reportApp.OpenReport(App.Path & szReportPath & "\RentReview_CP.rpt")

      Report.Database.Tables(1).ConnectionProperties.Item("Database Password") = accessDBPws

      Report.EnableParameterPrompting = False
      Report.DiscardSavedData

      ' Passing the from and to date values to Crystal Reports
      Report.ParameterFields(1).AddCurrentValue CDate(txtFromDate.text)
      Report.ParameterFields(2).AddCurrentValue CDate(txtToDate.text)
      Report.ParameterFields(3).AddCurrentValue cboClientList.Column(0)
      Report.ParameterFields(4).AddCurrentValue cboPropertyList.Column(0)
   End If

   Load frmReport
   frmReport.LoadReportViewer Report
End Sub

Private Sub GenerateReportButton_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   GenerateReportButton.MousePointer = vbArrow
End Sub

Private Sub txtFromDate_Change()
   TextBoxChangeDate txtFromDate
End Sub

Private Sub txtFromDate_GotFocus()
   SelTxtInCtrl txtFromDate
End Sub

Private Sub txtFromDate_KeyPress(KeyAscii As Integer)
   TextBoxKeyPrsDate txtFromDate, KeyAscii
End Sub

Private Sub txtFromDate_LostFocus()
   TextBoxFormatDate txtFromDate
End Sub

Private Sub txtToDate_Change()
   TextBoxChangeDate txtToDate
End Sub

Private Sub txtToDate_GotFocus()
   SelTxtInCtrl txtToDate
End Sub

Private Sub txtToDate_KeyPress(KeyAscii As Integer)
   TextBoxKeyPrsDate txtToDate, KeyAscii
End Sub

Private Sub txtToDate_LostFocus()
   TextBoxFormatDate txtToDate
End Sub

