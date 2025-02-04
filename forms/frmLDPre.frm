VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmLDPre 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lessee Detail Report"
   ClientHeight    =   3255
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4635
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   4635
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   420
      Left            =   2370
      TabIndex        =   1
      Top             =   2250
      Width           =   1215
   End
   Begin VB.CommandButton cmdReport 
      Caption         =   "&Report"
      Height          =   420
      Left            =   810
      TabIndex        =   0
      Top             =   2250
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   1815
      Left            =   45
      TabIndex        =   2
      Top             =   45
      Width           =   4560
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Property "
         Height          =   375
         Index           =   2
         Left            =   315
         TabIndex        =   6
         Top             =   1065
         Width           =   735
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Client"
         Height          =   375
         Index           =   1
         Left            =   315
         TabIndex        =   5
         Top             =   585
         Width           =   735
      End
      Begin MSForms.ComboBox cboProperty 
         Height          =   345
         Left            =   990
         TabIndex        =   4
         Top             =   1005
         Width           =   2985
         VariousPropertyBits=   1753237531
         DisplayStyle    =   3
         Size            =   "5265;609"
         BoundColumn     =   0
         TextColumn      =   2
         ColumnCount     =   2
         ListRows        =   20
         cColumnInfo     =   1
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   6
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         Object.Width           =   "1411"
      End
      Begin MSForms.ComboBox cboClient 
         Height          =   345
         Left            =   990
         TabIndex        =   3
         Top             =   585
         Width           =   2985
         VariousPropertyBits=   1753237531
         DisplayStyle    =   3
         Size            =   "5265;609"
         BoundColumn     =   0
         TextColumn      =   2
         ColumnCount     =   2
         ListRows        =   20
         cColumnInfo     =   1
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   6
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         Object.Width           =   "1763"
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1365
      Left            =   45
      TabIndex        =   7
      Top             =   1845
      Width           =   4560
   End
End
Attribute VB_Name = "frmLDPre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dataClient() As String, dataProp() As String
'Lessee details report
'issue 483 note 997 added by anol 24 Mar 2015
'Not showing all lessee address details. Input needs to be added that allows user to select client and property when running this report.
Public strFrom As String

Private Sub cboClient_Click()
    Dim adoConn As New ADODB.Connection
    adoConn.Open getConnectionString
    LoadProperties adoConn, cboProperty, cboClient.Column(0)
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub
Private Sub PrepareList(adoConn As ADODB.Connection, cboC As Control, cboP As Control)
   Dim adoRst As New ADODB.Recordset
   Dim szSQL As String

   On Error GoTo ErrorHandler

'*************************************** CLIENT COMBO ******************************************
   szSQL = "SELECT CLIENTID, CLIENTNAME, CLIENTPOSTCODE,  " & _
               "LandLordSageCustAC, LandLordSageSuppAC " & _
           "FROM CLIENT " & _
           "ORDER BY CLIENTNAME;"
'Debug.Print szSQL
   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If adoRst.EOF Then GoTo NoRes

   Dim TotalRow As Integer, TotalCol As Integer
   Dim i As Integer, j As Integer

   TotalRow = adoRst.RecordCount
   TotalCol = adoRst.Fields.count - 1

   Dim Data() As String

   ReDim Data(TotalCol, TotalRow) As String

   'Resolved by BOSL
'issue 455
'Modified by Anol 21 Aug 2014
   Data(0, 0) = "ALL"
   Data(1, 0) = "All Clients"
   For i = 1 To TotalRow 'end of modification
       For j = 0 To TotalCol - 1
           Data(j, i) = IIf(IsNull(adoRst.Fields(j).Value), "", adoRst.Fields(j).Value)
       Next j
       adoRst.MoveNext
       If adoRst.EOF Then Exit For
   Next i
   cboC.Column() = Data()
   cboC.ListIndex = 0
   adoRst.Close
'*************************************** PROPERTY ******************************************
    'Resolved by BOSL
    'issue 455
    'Modified by Anol 21 Aug 2014
   If cboC.text = "All Clients" Or Trim(cboC.text) = "" Then
      szSQL = "SELECT PropertyID, PropertyName, " & _
                  "ProAddressLine1, ProPostCode " & _
              "FROM Property " & _
              "ORDER BY PropertyID;"
   Else
        szSQL = "SELECT PropertyID, PropertyName, " & _
                  "ProAddressLine1, ProPostCode " & _
              "FROM Property " & _
              "WHERE ClientID = '" & cboC.Column(0) & "' " & _
              "ORDER BY PropertyID;"
      
   End If
   'end of modification
'   Debug.Print szSQL
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
   
   cboP.Column() = Data()
   cboP.ListIndex = 0

NoRes:
   adoRst.Close
   Set adoRst = Nothing

   Exit Sub

ErrorHandler:
   MsgBox ERR.description & "::" & ERR.Number

   adoRst.Close
   Set adoRst = Nothing
End Sub
Private Sub LoadProperties(adoConn As ADODB.Connection, cboP As Control, szClientID As String)
   Dim adoRst As New ADODB.Recordset
   Dim szSQL As String, j As Integer
   Dim i As Integer, Data() As String
   Dim TotalRow As Integer, TotalCol As Integer

   On Error GoTo ErrorHandler

'***************************************  PROPERTY  ******************************************
   If cboClient.text = "All Clients" Or Trim(cboClient.text) = "" Then
      szSQL = "SELECT PropertyID, PropertyName, " & _
                  "ProAddressLine1, ProPostCode " & _
              "FROM Property " & _
              "ORDER BY PropertyID;"
   Else
        szSQL = "SELECT PropertyID, PropertyName, " & _
                  "ProAddressLine1, ProPostCode " & _
              "FROM Property " & _
              "WHERE ClientID = '" & cboClient.Column(0) & "' " & _
              "ORDER BY PropertyID;"
      
   End If
   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   cboP.Clear
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
   
   cboP.Column() = Data()
   cboP.ListIndex = 0

NoRes:
   adoRst.Close
   Set adoRst = Nothing

   Exit Sub

ErrorHandler:
   MsgBox ERR.description & "::" & ERR.Number

   adoRst.Close
   Set adoRst = Nothing
End Sub
Private Sub cmdReport_Click()
'  On Error GoTo ErrHanlder
' Declare the application object used to open the rpt file
   Dim crxApplication As New CRAXDRT.Application

' Declare the report object
   Dim Report As CRAXDRT.Report
   Dim i As Integer
   If strFrom = "TenantDetailsReport" Then
           Set Report = crxApplication.OpenReport(App.Path & szReportPath & "\TenantDetailsReport.rpt", 1)
        '   Report.Database.LogOnServer "C:\Samrat\PropertyManagementProgram\Non_Client_Server\Non_SAGE\BlockMng\Database\PBMc001.mdb;UID=;PWD=" & accessDBPws & ""
           Report.DiscardSavedData
           'Note bny anol
           ' I have found that Ac report Does not take client and property ID as parameter, but WPM does
           Report.Database.Tables(1).ConnectionProperties.Item("Database Password") = accessDBPws
           Report.ParameterFields(1).AddCurrentValue cboClient.Column(0)
           Report.ParameterFields(2).AddCurrentValue cboProperty.Column(0)
   ElseIf strFrom = "LeaseInformationReport" Then
         Set Report = crxApplication.OpenReport(App.Path & szReportPath & "\LeaseInformationReport.rpt", 1)
        '   Report.Database.LogOnServer "C:\Samrat\PropertyManagementProgram\Non_Client_Server\Non_SAGE\BlockMng\Database\PBMc001.mdb;UID=;PWD=" & accessDBPws & ""
           Report.DiscardSavedData
           'Note bny anol
           ' I have found that Ac report Does not take client and property ID as parameter, but WPM does
           Report.Database.Tables(1).ConnectionProperties.Item("Database Password") = accessDBPws
           Report.ParameterFields(1).AddCurrentValue cboClient.Column(0)
           Report.ParameterFields(2).AddCurrentValue cboProperty.Column(0)
   End If
    
   Dim strInvoiceAmt As String
   Dim rep As frmReport

   Set rep = New frmReport

   rep.LoadReportViewer Report
   Exit Sub

ErrHanlder:
   MsgBox ERR.description & ":::An error occured, please contact with PCM Consulting.", vbCritical + vbOKOnly, "Report Error"

End Sub

Private Sub Form_Load()
    Me.Left = 100
    Me.Top = 100
    Frame1.BackColor = MODULEBACKCOLOR
    Frame2.BackColor = MODULEBACKCOLOR
    Me.BackColor = MODULEBACKCOLOR
    Dim adoConn As New ADODB.Connection
    adoConn.Open getConnectionString
    PrepareList adoConn, cboClient, cboProperty
    adoConn.Close
    
End Sub
