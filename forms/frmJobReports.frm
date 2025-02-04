VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmJobReports 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pending Jobs"
   ClientHeight    =   1590
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5610
   Icon            =   "frmJobReports.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1590
   ScaleMode       =   0  'User
   ScaleWidth      =   3000
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print Report"
      Height          =   375
      Left            =   2040
      TabIndex        =   1
      Top             =   1080
      Width           =   1575
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3720
      TabIndex        =   2
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label Label19 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Property"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   3
      Left            =   255
      TabIndex        =   3
      Top             =   420
      Width           =   720
   End
   Begin MSForms.ComboBox cboPropertyList 
      Height          =   315
      Left            =   1440
      TabIndex        =   0
      Top             =   360
      Width           =   3855
      VariousPropertyBits=   1753237531
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "6800;556"
      BoundColumn     =   0
      TextColumn      =   2
      ColumnCount     =   3
      ListRows        =   20
      cColumnInfo     =   1
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      SpecialEffect   =   0
      FontName        =   "Myriad Web"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      Object.Width           =   "1411"
   End
End
Attribute VB_Name = "frmJobReports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub PrepareList(adoConn As ADODB.Connection, cboProperty As Control)
   Dim adoRst     As New ADODB.Recordset
   Dim szSQL      As String
   Dim TotalRow   As Integer
   Dim TotalCol   As Integer
   Dim i          As Integer
   Dim j          As Integer

   On Error GoTo ErrorHandler

'*************************************** PROPERTY ******************************************
   szSQL = "SELECT PropertyID, PropertyName, " & _
               "ProAddressLine1, ProPostCode " & _
           "FROM Property " & _
           "ORDER BY PropertyID;"

   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If adoRst.EOF Then GoTo NoRes

   TotalRow = adoRst.RecordCount
   TotalCol = adoRst.Fields.count - 1

   ReDim Data(TotalCol, TotalRow) As String

   Data(0, 0) = "ALL"
   Data(1, 0) = "All Properties"
   For i = 1 To TotalRow
       For j = 0 To TotalCol - 1
           Data(j, i) = IIf(IsNull(adoRst.Fields(j).Value), "", adoRst.Fields(j).Value)
       Next j
       adoRst.MoveNext
       If adoRst.EOF Then Exit For
   Next i
   cboProperty.Column() = Data()
   cboProperty.ListIndex = 0

NoRes:
   adoRst.Close
   Set adoRst = Nothing
   Exit Sub

ErrorHandler:
   MsgBox ERR.description & "::" & ERR.Number

   adoRst.Close
   Set adoRst = Nothing
End Sub

Private Sub cmdCancel_Click()
   Unload Me
End Sub

Private Sub cmdPrint_Click()
   Dim adoConn As New ADODB.Connection
   Dim adoRst As New ADODB.Recordset, adoRSTR As New ADODB.Recordset
   Dim szSQL As String
   
   On Error GoTo ErrorHandler
   
   adoConn.Open getConnectionString
   
   If cboPropertyList.Column(0) = "ALL" Then
      szSQL = "SELECT * " & _
              "FROM PropertyMaintHistory " & _
              "WHERE ISNULL(DateCompleted);"
   Else
      szSQL = "SELECT * " & _
              "FROM PropertyMaintHistory " & _
              "WHERE PropertyID = '" & cboPropertyList.Column(0) & "' AND ISNULL(DateCompleted);"
   End If
'Debug.Print szSQL
   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If adoRst.EOF Then
      MsgBox "No Pending Jobs available for the selected Property", vbOKOnly, "No Pending Jobs"
      adoRst.Close
      adoConn.Close
      Exit Sub
   End If
   
   adoRst.Close
   adoConn.Close
   Set adoConn = Nothing
      
   Dim reportApp As New CRAXDRT.Application
   Dim Report As CRAXDRT.Report

   Set Report = reportApp.OpenReport(App.Path & szReportPath & "\PendingJob.rpt")
   Report.Database.Tables(1).ConnectionProperties.Item("Database Password") = accessDBPws

   Report.EnableParameterPrompting = False
   Report.DiscardSavedData
   
   Report.ParameterFields(1).AddCurrentValue "Job Name"
   
   Report.ParameterFields(2).AddCurrentValue cboPropertyList.Column(0)

   If cboPropertyList.Column(0) = "ALL" Then
      Report.ParameterFields(3).AddCurrentValue "For all properties"
   Else
      Report.ParameterFields(3).AddCurrentValue "For the Property: " & cboPropertyList.Column(0)
   End If
         
   Load frmReport
   frmReport.LoadReportViewer Report
   
   Exit Sub
   
ErrorHandler:
   MsgBox ERR.description & "::" & ERR.Number
   
End Sub

Private Sub Form_Load()
   Me.BackColor = MODULEBACKCOLOR

   Dim adoConn As New ADODB.Connection
   Dim szSQL As String

   adoConn.Open getConnectionString
   
   PrepareList adoConn, cboPropertyList
End Sub
