VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmPJRptCriteria 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pending Jobs Report Criteria"
   ClientHeight    =   7530
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7800
   BeginProperty Font 
      Name            =   "Myriad Web"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPJRptCriteria.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7530
   ScaleWidth      =   7800
   Begin VB.TextBox txtDtReportedTo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   4680
      TabIndex        =   2
      Top             =   960
      Width           =   2895
   End
   Begin VB.TextBox txtDtReportedFrom 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1560
      TabIndex        =   1
      Text            =   "01/01/1980"
      Top             =   960
      Width           =   2895
   End
   Begin MSForms.ComboBox cmbUnitTo 
      Height          =   315
      Left            =   4680
      TabIndex        =   29
      Top             =   4680
      Width           =   2895
      VariousPropertyBits=   1753237531
      DisplayStyle    =   3
      Size            =   "5106;556"
      BoundColumn     =   0
      TextColumn      =   2
      ColumnCount     =   2
      ListRows        =   20
      cColumnInfo     =   2
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      SpecialEffect   =   6
      FontName        =   "Myriad Web"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      Object.Width           =   "1411;3527"
   End
   Begin MSForms.ComboBox cmbUnitFrom 
      Height          =   315
      Left            =   1560
      TabIndex        =   28
      Top             =   4680
      Width           =   2895
      VariousPropertyBits=   1753237531
      DisplayStyle    =   3
      Size            =   "5106;556"
      BoundColumn     =   0
      TextColumn      =   2
      ColumnCount     =   2
      ListRows        =   20
      cColumnInfo     =   2
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      SpecialEffect   =   6
      FontName        =   "Myriad Web"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      Object.Width           =   "1411;3527"
   End
   Begin MSForms.ComboBox cmbPropUnit 
      Height          =   285
      Left            =   240
      TabIndex        =   27
      Top             =   1460
      Width           =   1215
      VariousPropertyBits=   1753237531
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "2143;503"
      BoundColumn     =   0
      TextColumn      =   1
      ListRows        =   20
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      SpecialEffect   =   0
      FontName        =   "Myriad Web"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox txtJobAssignTo 
      Height          =   315
      Left            =   3240
      TabIndex        =   16
      Top             =   6000
      Visible         =   0   'False
      Width           =   2415
      VariousPropertyBits=   679495707
      BorderStyle     =   1
      Size            =   "4260;556"
      Value           =   "ZZZZZZ"
      BorderColor     =   8421504
      SpecialEffect   =   0
      FontName        =   "Myriad Web"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ComboBox cmbJobAssignTo 
      Height          =   315
      Left            =   4680
      TabIndex        =   10
      Top             =   2880
      Width           =   2895
      VariousPropertyBits=   1753237531
      DisplayStyle    =   3
      Size            =   "5106;556"
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
      Object.Width           =   "1058"
   End
   Begin MSForms.ComboBox cmbJobAssignFrom 
      Height          =   315
      Left            =   1560
      TabIndex        =   9
      Top             =   2880
      Width           =   2895
      VariousPropertyBits=   1753237531
      DisplayStyle    =   3
      Size            =   "5106;556"
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
      Object.Width           =   "1058"
   End
   Begin MSForms.ComboBox cmbJobOwnerFrom 
      Height          =   315
      Left            =   1560
      TabIndex        =   7
      Top             =   2400
      Width           =   2895
      VariousPropertyBits=   1753237531
      DisplayStyle    =   3
      Size            =   "5106;556"
      BoundColumn     =   0
      TextColumn      =   2
      ColumnCount     =   2
      ListRows        =   20
      cColumnInfo     =   2
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      SpecialEffect   =   6
      FontName        =   "Myriad Web"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      Object.Width           =   "1058;3527"
   End
   Begin MSForms.ComboBox cmbJobOwnerTo 
      Height          =   315
      Left            =   4680
      TabIndex        =   8
      Top             =   2400
      Width           =   2895
      VariousPropertyBits=   1753237531
      DisplayStyle    =   3
      Size            =   "5106;556"
      BoundColumn     =   0
      TextColumn      =   2
      ColumnCount     =   2
      ListRows        =   20
      cColumnInfo     =   2
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      SpecialEffect   =   6
      FontName        =   "Myriad Web"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      Object.Width           =   "1058;3527"
   End
   Begin MSForms.ComboBox cmbJobRefTo 
      Height          =   315
      Left            =   4680
      TabIndex        =   6
      Top             =   1920
      Width           =   2895
      VariousPropertyBits=   1753237531
      DisplayStyle    =   3
      Size            =   "5106;556"
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
      Object.Width           =   "1762"
   End
   Begin MSForms.ComboBox cmbJobRefFrom 
      Height          =   315
      Left            =   1560
      TabIndex        =   5
      Top             =   1920
      Width           =   2895
      VariousPropertyBits=   1753237531
      DisplayStyle    =   3
      Size            =   "5106;556"
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
      Object.Width           =   "1762"
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00808080&
      Height          =   3375
      Index           =   1
      Left            =   120
      Top             =   600
      Width           =   7575
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Height          =   3375
      Index           =   0
      Left            =   120
      Top             =   600
      Width           =   7575
   End
   Begin MSForms.ComboBox cmbClients 
      Height          =   315
      Left            =   2640
      TabIndex        =   0
      Top             =   120
      Width           =   3135
      VariousPropertyBits=   1753237531
      DisplayStyle    =   3
      Size            =   "5530;556"
      BoundColumn     =   0
      TextColumn      =   2
      ColumnCount     =   2
      ListRows        =   20
      cColumnInfo     =   2
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      SpecialEffect   =   6
      FontName        =   "Myriad Web"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      Object.Width           =   "1058;3527"
   End
   Begin MSForms.Label Label1 
      Height          =   255
      Index           =   7
      Left            =   2040
      TabIndex        =   26
      Top             =   120
      Width           =   735
      ForeColor       =   64
      VariousPropertyBits=   8388627
      Caption         =   "Client:"
      Size            =   "1296;450"
      FontName        =   "Myriad Web"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox txtJobAssignFrom 
      Height          =   315
      Left            =   360
      TabIndex        =   15
      Top             =   6000
      Visible         =   0   'False
      Width           =   2415
      VariousPropertyBits=   679495707
      BorderStyle     =   1
      Size            =   "4260;556"
      Value           =   "AAAAAA"
      BorderColor     =   8421504
      SpecialEffect   =   0
      FontName        =   "Myriad Web"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label1 
      Height          =   255
      Index           =   6
      Left            =   240
      TabIndex        =   25
      Top             =   1440
      Width           =   735
      ForeColor       =   64
      VariousPropertyBits=   8388627
      Caption         =   "Property"
      Size            =   "1296;450"
      FontName        =   "Myriad Web"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ComboBox cmbPropertyFrom 
      Height          =   315
      Left            =   1560
      TabIndex        =   3
      Top             =   1440
      Width           =   2895
      VariousPropertyBits=   1753237531
      DisplayStyle    =   3
      Size            =   "5106;556"
      BoundColumn     =   0
      TextColumn      =   2
      ColumnCount     =   2
      ListRows        =   20
      cColumnInfo     =   2
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      SpecialEffect   =   6
      FontName        =   "Myriad Web"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      Object.Width           =   "1058;3527"
   End
   Begin MSForms.ComboBox cmbPropertyTo 
      Height          =   315
      Left            =   4680
      TabIndex        =   4
      Top             =   1440
      Width           =   2895
      VariousPropertyBits=   1753237531
      DisplayStyle    =   3
      Size            =   "5106;556"
      BoundColumn     =   0
      TextColumn      =   2
      ColumnCount     =   2
      ListRows        =   20
      cColumnInfo     =   2
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      SpecialEffect   =   6
      FontName        =   "Myriad Web"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      Object.Width           =   "1058;3527"
   End
   Begin MSForms.CommandButton cmdRPCCancel 
      Height          =   375
      Left            =   6240
      TabIndex        =   12
      Top             =   3540
      Width           =   1335
      Caption         =   "Cancel"
      Size            =   "2355;661"
      FontName        =   "Myriad Web"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton cmdRPCOK 
      Default         =   -1  'True
      Height          =   375
      Left            =   4560
      TabIndex        =   11
      Top             =   3540
      Width           =   1335
      Caption         =   "OK"
      Size            =   "2355;661"
      FontName        =   "Myriad Web"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C00000&
      Index           =   0
      X1              =   120
      X2              =   7680
      Y1              =   3360
      Y2              =   3360
   End
   Begin MSForms.TextBox txtJobOwnerTo 
      Height          =   315
      Left            =   3240
      TabIndex        =   14
      Top             =   5520
      Visible         =   0   'False
      Width           =   2415
      VariousPropertyBits=   679495707
      BorderStyle     =   1
      Size            =   "4260;556"
      Value           =   "ZZZZZZ"
      BorderColor     =   8421504
      SpecialEffect   =   0
      FontName        =   "Myriad Web"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox txtJobOwnerFrom 
      Height          =   315
      Left            =   360
      TabIndex        =   13
      Top             =   5520
      Visible         =   0   'False
      Width           =   2415
      VariousPropertyBits=   679495707
      BorderStyle     =   1
      Size            =   "4260;556"
      Value           =   "AAAAAA"
      BorderColor     =   8421504
      SpecialEffect   =   0
      FontName        =   "Myriad Web"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox txtJobRefTo 
      Height          =   315
      Left            =   3240
      TabIndex        =   18
      Top             =   5040
      Visible         =   0   'False
      Width           =   2415
      VariousPropertyBits=   679495707
      BorderStyle     =   1
      Size            =   "4260;556"
      Value           =   "ZZZZZZ"
      BorderColor     =   8421504
      SpecialEffect   =   0
      FontName        =   "Myriad Web"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox txtJobRefFrom 
      Height          =   315
      Left            =   360
      TabIndex        =   17
      Top             =   5040
      Visible         =   0   'False
      Width           =   2415
      VariousPropertyBits=   679495707
      BorderStyle     =   1
      Size            =   "4260;556"
      Value           =   "AAAAAA"
      BorderColor     =   8421504
      SpecialEffect   =   0
      FontName        =   "Myriad Web"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label1 
      Height          =   255
      Index           =   5
      Left            =   1560
      TabIndex        =   24
      Top             =   720
      Width           =   495
      ForeColor       =   64
      VariousPropertyBits=   8388627
      Caption         =   "From"
      Size            =   "873;450"
      FontName        =   "Myriad Web"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.Label Label1 
      Height          =   255
      Index           =   4
      Left            =   4680
      TabIndex        =   23
      Top             =   720
      Width           =   255
      ForeColor       =   64
      VariousPropertyBits=   8388627
      Caption         =   "To"
      Size            =   "450;450"
      FontName        =   "Myriad Web"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.Label Label1 
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   22
      Top             =   2880
      Width           =   1095
      ForeColor       =   64
      VariousPropertyBits=   8388627
      Caption         =   "Job Assigned"
      Size            =   "1931;450"
      FontName        =   "Myriad Web"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label1 
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   21
      Top             =   2400
      Width           =   975
      ForeColor       =   64
      VariousPropertyBits=   8388627
      Caption         =   "Job Owner"
      Size            =   "1720;450"
      FontName        =   "Myriad Web"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label1 
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   20
      Top             =   1920
      Width           =   615
      ForeColor       =   64
      VariousPropertyBits=   8388627
      Caption         =   "Job No"
      Size            =   "1085;450"
      FontName        =   "Myriad Web"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label1 
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   19
      Top             =   960
      Width           =   1335
      ForeColor       =   64
      VariousPropertyBits=   8388627
      Caption         =   "Date Reported"
      Size            =   "2355;450"
      FontName        =   "Myriad Web"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   3
      Index           =   1
      X1              =   120
      X2              =   7680
      Y1              =   3360
      Y2              =   3360
   End
End
Attribute VB_Name = "frmPJRptCriteria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public CallingFrom As String

Private Sub cmbPropUnit_Click()
   If cmbPropUnit.text = "Unit" Then
      cmbPropertyFrom.Visible = False
      cmbPropertyTo.Visible = False
      cmbUnitFrom.Top = cmbPropertyFrom.Top
      cmbUnitTo.Top = cmbPropertyTo.Top
      cmbUnitFrom.Visible = True
      cmbUnitTo.Visible = True
   End If
   If cmbPropUnit.text = "Property" Then
      cmbUnitFrom.Visible = False
      cmbUnitTo.Visible = False
      cmbPropertyFrom.Visible = True
      cmbPropertyTo.Visible = True
   End If
End Sub

Private Sub cmdRPCCancel_Click()
   Unload Me
End Sub

Private Sub cmdRPCOK_Click()
   Dim reportApp As New CRAXDRT.Application
   Dim Report As CRAXDRT.Report

   If CallingFrom = "Budget" Then
      Set Report = reportApp.OpenReport(App.Path & szReportPath & "\JobsBudgetComp.rpt")
   End If

   If CallingFrom = "Pending" Then
      If cmbPropUnit.text = "Property" Then
         Set Report = reportApp.OpenReport(App.Path & szReportPath & "\PendingJobReport.rpt")
      Else
         Set Report = reportApp.OpenReport(App.Path & szReportPath & "\PendingJobReport_Unit.rpt")
      End If
   End If

   Report.EnableParameterPrompting = False
   Report.DiscardSavedData

   Report.ParameterFields(1).AddCurrentValue cmbClients.Column(1)
   If cmbPropUnit.text = "Property" Then
      Report.ParameterFields(2).AddCurrentValue cmbPropertyFrom.Column(1)
      Report.ParameterFields(3).AddCurrentValue cmbPropertyTo.Column(1)
   Else
      Report.ParameterFields(2).AddCurrentValue cmbUnitFrom.Column(1)
      Report.ParameterFields(3).AddCurrentValue cmbUnitTo.Column(1)
   End If
   Report.ParameterFields(4).AddCurrentValue CDate(IIf(txtDtReportedFrom.text = "", Format(Date, "dd mmmm yyyy"), txtDtReportedFrom.text))
   Report.ParameterFields(5).AddCurrentValue CDate(IIf(txtDtReportedTo.text = "", Format(Date, "dd mmmm yyyy"), txtDtReportedTo.text))
   Report.ParameterFields(6).AddCurrentValue IIf(cmbJobRefFrom.text = "", "AAAAAA", cmbJobRefFrom.Column(0))
   Report.ParameterFields(7).AddCurrentValue IIf(cmbJobRefTo.text = "", "ZZZZZZ", cmbJobRefTo.Column(0))
   Report.ParameterFields(8).AddCurrentValue IIf(cmbJobOwnerFrom.text = "", "AAAAAA", cmbJobOwnerFrom.text)
   Report.ParameterFields(9).AddCurrentValue IIf(cmbJobOwnerTo.text = "", "ZZZZZZ", cmbJobOwnerTo.text)
   Report.ParameterFields(10).AddCurrentValue IIf(cmbJobAssignFrom.text = "", "AAAAAA", cmbJobAssignFrom.text)
   Report.ParameterFields(11).AddCurrentValue IIf(cmbJobAssignTo.text = "", "ZZZZZZ", cmbJobAssignTo.text)

   Load frmReport
   frmReport.LoadReportViewer Report
   
   Exit Sub
'
'BUDGET_REPORT:
'
'
'
End Sub

Private Sub Form_Activate()
   If CallingFrom = "Budget" Then cmbPropUnit.Visible = False
End Sub

Private Sub Form_Load()
   If CallingFrom = "Pending" Then
      Me.Caption = "Pending Jobs Report Criteria"
   Else
      Me.Caption = "Jobs Budget Comparison Report"
   End If

   Dim adoConn As New ADODB.Connection

   adoConn.Open getConnectionString

   Me.Top = 50
   Me.Left = 50
   Me.Height = 4605
   Me.Width = 7890
   Me.BackColor = MODULEBACKCOLOR
   txtDtReportedTo.text = Format(Date, "DD/MM/YYYY")

   If Not LoadClients(adoConn) Then
      adoConn.Close
      Set adoConn = Nothing
      Unload Me
   End If

   adoConn.Close
   Set adoConn = Nothing
End Sub

Private Sub cmbClients_Click()
   Dim adoConn As New ADODB.Connection

   adoConn.Open getConnectionString

   LoadPropertyByClient adoConn
   LoadJobNo adoConn
   LoadOwner adoConn

   adoConn.Close
   Set adoConn = Nothing
End Sub

Private Sub LoadJobNo(adoConn As ADODB.Connection)
   Dim iRec As Integer
   Dim adoRST As New ADODB.Recordset
   Dim szSQL As String, szaData() As String

   On Error GoTo Error_Handler

   szSQL = "SELECT M.ID, M.Job_DiaryName " & _
           "FROM PropertyMaintHistory AS M, Property AS P " & _
           "WHERE P.ClientID = '" & cmbClients.Column(0) & "' AND " & _
               "P.PropertyID = M.PropertyID AND " & _
               "LEFT(M.ID, 1) = 'J' " & _
           "UNION " & _
           "SELECT M.ID, M.Job_DiaryName " & _
           "FROM PropertyMaintHistory AS M, Units AS U, Property AS P  " & _
           "WHERE P.ClientID = '" & cmbClients.Column(0) & "' AND " & _
               "U.UnitNumber = M.PropertyID AND " & _
               "U.PropertyID = P.PropertyID AND " & _
               "LEFT(M.ID, 1) = 'J' AND " & _
               "P.PropertyID = U.PropertyID " & _
           "ORDER BY M.ID;"
'Debug.Print szSQL
   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   ReDim szaData(1, adoRST.RecordCount - 1) As String

   While Not adoRST.EOF
      szaData(0, iRec) = adoRST.Fields.Item(0).Value
      szaData(1, iRec) = adoRST.Fields.Item(1).Value
      iRec = iRec + 1
      adoRST.MoveNext
   Wend

   cmbJobRefFrom.Clear
   cmbJobRefFrom.Column() = szaData()
   cmbJobRefFrom.ListIndex = 0
   cmbJobRefTo.Clear
   cmbJobRefTo.Column() = szaData()
   cmbJobRefTo.ListIndex = cmbJobRefTo.ListCount - 1
   
   ' Destroy Objects
   Set adoRST = Nothing

   Exit Sub

   ' Error Handling Code
Error_Handler:
   ' Destroy Objects
   Set adoRST = Nothing
End Sub

Private Sub LoadOwner(adoConn As ADODB.Connection)
   Dim iRec       As Integer
   Dim adoRST     As New ADODB.Recordset
   Dim adoRstSup  As New ADODB.Recordset
   Dim szSQL      As String
   Dim szaData()  As String
   Dim iData      As Integer
   Dim szTemp1    As String
   Dim szTemp2    As String
   Dim i          As Integer
   Dim j          As Integer

   On Error GoTo Error_Handler

'   szSQL = "SELECT CODE, VALUE " & _
'           "FROM SECONDARYCODE " & _
'           "WHERE PRIMARYCODE = 'MNTJOB' " & _
'           "ORDER BY Value;"
   szSQL = "SELECT DISTINCT S.CODE, S.VALUE " & _
           "FROM PropertyMaintHistory AS P, SECONDARYCODE AS S " & _
           "WHERE S.PRIMARYCODE = 'MNTJOB' AND " & _
                 "P.TaskOwner = S.Value " & _
           "ORDER BY S.Value;"

'Debug.Print szSQL
   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   ReDim szaData(1, adoRST.RecordCount - 1) As String

   While Not adoRST.EOF
      szaData(0, iRec) = adoRST.Fields.Item(0).Value
      szaData(1, iRec) = adoRST.Fields.Item(1).Value
      iRec = iRec + 1
      adoRST.MoveNext
   Wend
   adoRST.Close

   cmbJobOwnerFrom.Clear
   cmbJobOwnerFrom.Column() = szaData()
   cmbJobOwnerFrom.ListIndex = 0
   cmbJobOwnerTo.Clear
   cmbJobOwnerTo.Column() = szaData()
   cmbJobOwnerTo.ListIndex = cmbJobOwnerTo.ListCount - 1

'-------------------------------------------------------------------------------------------
   szSQL = "SELECT DISTINCT S.CODE, S.VALUE " & _
           "FROM PropertyMaintHistory AS P, SECONDARYCODE AS S " & _
           "WHERE S.PRIMARYCODE = 'MNTJOB' AND " & _
                 "P.AssignedTo = S.Value " & _
           "ORDER BY S.Value;"
'Debug.Print szSQL
   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
   iData = adoRST.RecordCount - 1

   szSQL = "SELECT DISTINCT S.SupplierID, S.SupplierName " & _
           "FROM PropertyMaintHistory AS P, Supplier AS S " & _
           "WHERE P.AssignedTo = S.SupplierName " & _
           "ORDER BY S.SupplierName;"
'Debug.Print szSQL
   adoRstSup.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
   iData = iData + adoRstSup.RecordCount

   ReDim szaData(1, iData) As String

   iRec = 0
   While Not adoRST.EOF
      szaData(0, iRec) = adoRST.Fields.Item(0).Value
      szaData(1, iRec) = adoRST.Fields.Item(1).Value
      iRec = iRec + 1
      adoRST.MoveNext
   Wend
   While Not adoRstSup.EOF
      szaData(0, iRec) = adoRstSup.Fields.Item(0).Value
      szaData(1, iRec) = adoRstSup.Fields.Item(1).Value
      iRec = iRec + 1
      adoRstSup.MoveNext
   Wend

   adoRST.Close
   adoRstSup.Close

   For i = 0 To iRec - 2
      For j = i + 1 To iRec - 1
         If szaData(1, i) > szaData(1, j) Then
            szTemp1 = szaData(0, i)
            szTemp2 = szaData(1, i)
            szaData(0, i) = szaData(0, j)
            szaData(1, i) = szaData(1, j)
            szaData(0, j) = szTemp1
            szaData(1, j) = szTemp2
         End If
      Next j
   Next i
   
   cmbJobAssignFrom.Clear
   cmbJobAssignFrom.Column() = szaData()
   cmbJobAssignFrom.ListIndex = 0
   cmbJobAssignTo.Clear
   cmbJobAssignTo.Column() = szaData()
   cmbJobAssignTo.ListIndex = cmbJobAssignTo.ListCount - 1

   ' Destroy Objects
   Set adoRST = Nothing
   Set adoRstSup = Nothing

   Exit Sub

   ' Error Handling Code
Error_Handler:
   ' Destroy Objects
   Set adoRST = Nothing
   Set adoRstSup = Nothing
End Sub

Private Sub LoadPropertyByClient(adoConn As ADODB.Connection)
   cmbPropUnit.AddItem "Property"
   cmbPropUnit.AddItem "Unit"
   cmbPropUnit.ListIndex = 0

   Dim iRec As Integer
   Dim adoRST As New ADODB.Recordset
   Dim szSQL As String, szaData() As String

   On Error GoTo Error_Handler

   szSQL = "SELECT PropertyID, PropertyName " & _
           "FROM Property " & _
           "WHERE ClientID = '" & cmbClients.Column(0) & "' " & _
           "ORDER BY PropertyName;"
'Debug.Print szSQL
   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   ReDim szaData(1, adoRST.RecordCount - 1) As String

   While Not adoRST.EOF
      szaData(0, iRec) = adoRST.Fields.Item("PropertyID").Value
      szaData(1, iRec) = adoRST.Fields.Item("PropertyName").Value
      iRec = iRec + 1
      adoRST.MoveNext
   Wend
   adoRST.Close

   cmbPropertyFrom.Clear
   cmbPropertyFrom.Column() = szaData()
   cmbPropertyFrom.ListIndex = 0
   cmbPropertyTo.Clear
   cmbPropertyTo.Column() = szaData()
   cmbPropertyTo.ListIndex = cmbPropertyTo.ListCount - 1
'###############################################################     UNIT     ##############################
   
   szSQL = "SELECT   U.UnitNumber, U.UnitName " & _
           "FROM     Property AS P, Units AS U " & _
           "WHERE    P.ClientID = '" & cmbClients.Column(0) & "' AND " & _
                    "U.PropertyID = P.PropertyID " & _
           "ORDER BY U.UnitName;"
   
   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   ReDim szaData(1, adoRST.RecordCount - 1) As String
   iRec = 0

   While Not adoRST.EOF
      szaData(0, iRec) = adoRST.Fields.Item(0).Value
      szaData(1, iRec) = adoRST.Fields.Item(1).Value
      iRec = iRec + 1
      adoRST.MoveNext
   Wend
   adoRST.Close
   
   cmbUnitFrom.Clear
   cmbUnitFrom.Column() = szaData()
   cmbUnitFrom.ListIndex = 0
   cmbUnitTo.Clear
   cmbUnitTo.Column() = szaData()
   cmbUnitTo.ListIndex = cmbUnitTo.ListCount - 1

   ' Destroy Objects
   Set adoRST = Nothing

   Exit Sub

   ' Error Handling Code
Error_Handler:
   ' Destroy Objects
   Set adoRST = Nothing
End Sub

Private Sub LoadProperty(adoConn As ADODB.Connection)
   Dim iRec As Integer
   Dim adoRST As New ADODB.Recordset
   Dim szSQL As String, szaData() As String

   On Error GoTo Error_Handler

   szSQL = "SELECT PropertyID, PropertyName " & _
           "FROM Property " & _
           "ORDER BY PropertyName;"
'Debug.Print szSQL
   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   ReDim szaData(1, adoRST.RecordCount - 1) As String

   While Not adoRST.EOF
      szaData(0, iRec) = adoRST.Fields.Item("PropertyID").Value
      szaData(1, iRec) = adoRST.Fields.Item("PropertyName").Value
      iRec = iRec + 1
      adoRST.MoveNext
   Wend

   cmbPropertyFrom.Clear
   cmbPropertyFrom.Column() = szaData()
   cmbPropertyFrom.ListIndex = 0
   cmbPropertyTo.Clear
   cmbPropertyTo.Column() = szaData()
   cmbPropertyTo.ListIndex = cmbPropertyTo.ListCount - 1

   ' Destroy Objects
   Set adoRST = Nothing

   Exit Sub

   ' Error Handling Code
Error_Handler:
   ' Destroy Objects
   Set adoRST = Nothing
End Sub

Private Function LoadClients(adoConn As ADODB.Connection) As Boolean
   Dim iRec As Integer
   Dim adoRST As New ADODB.Recordset
   Dim szSQL As String, szaData() As String

   LoadClients = True

   On Error GoTo Error_Handler

   szSQL = "SELECT ClientID, ClientName " & _
           "FROM Client " & _
           "ORDER BY ClientName;"
'Debug.Print szSQL
   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If adoRST.EOF Then
      ShowMsgInTaskBar "Please input Client details into Prestige."
      LoadClients = False
   Else
      ReDim szaData(1, adoRST.RecordCount - 1) As String

      While Not adoRST.EOF
         szaData(0, iRec) = adoRST.Fields.Item("ClientID").Value
         szaData(1, iRec) = adoRST.Fields.Item("ClientName").Value
         iRec = iRec + 1
         adoRST.MoveNext
      Wend
   End If

   cmbClients.Clear
   cmbClients.Column() = szaData()
   cmbClients.ListIndex = 0

   ' Destroy Objects
   Set adoRST = Nothing

   Exit Function

   ' Error Handling Code
Error_Handler:
   ' Destroy Objects
   Set adoRST = Nothing
End Function

Private Sub Form_Unload(Cancel As Integer)
   'frmMMain.fraCmdButton.Enabled = True
End Sub

Private Sub TextBox7_Change()

End Sub

Private Sub txtCustRefFrom_Change()

End Sub

Private Sub txtDtReportedFrom_Change()
   TextBoxChangeDate txtDtReportedFrom
End Sub

Private Sub txtDtReportedFrom_GotFocus()
   SelTxtInCtrl txtDtReportedFrom
End Sub

Private Sub txtDtReportedFrom_KeyPress(KeyAscii As Integer)
   TextBoxKeyPrsDate txtDtReportedFrom, KeyAscii
End Sub

Private Sub txtDtReportedFrom_LostFocus()
   TextBoxFormatDate txtDtReportedFrom
End Sub

Private Sub txtDtReportedto_Change()
   TextBoxChangeDate txtDtReportedTo
End Sub

Private Sub txtDtReportedTo_GotFocus()
   SelTxtInCtrl txtDtReportedTo
End Sub

Private Sub txtDtReportedto_KeyPress(KeyAscii As Integer)
   TextBoxKeyPrsDate txtDtReportedTo, KeyAscii
End Sub

Private Sub txtDtReportedto_LostFocus()
   TextBoxFormatDate txtDtReportedTo
End Sub
