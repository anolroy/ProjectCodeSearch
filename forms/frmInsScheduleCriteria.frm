VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmInsScheduleCriteria 
   BackColor       =   &H00E5E5E5&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Insurance Schedule Report Criteria"
   ClientHeight    =   3750
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7680
   BeginProperty Font 
      Name            =   "Myriad Web"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmInsScheduleCriteria.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3750
   ScaleWidth      =   7680
   Begin VB.TextBox txtDateTo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   5040
      TabIndex        =   2
      Top             =   1200
      Width           =   2415
   End
   Begin VB.TextBox txtDateFrom 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2160
      TabIndex        =   1
      Text            =   "01/01/1980"
      Top             =   1200
      Width           =   2415
   End
   Begin MSForms.ComboBox cmbTenantTo 
      Height          =   315
      Left            =   5040
      TabIndex        =   14
      Top             =   2160
      Width           =   2415
      VariousPropertyBits=   1753237531
      DisplayStyle    =   3
      Size            =   "4260;556"
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
   Begin MSForms.ComboBox cmbTenantFrom 
      Height          =   315
      Left            =   2160
      TabIndex        =   13
      Top             =   2160
      Width           =   2415
      VariousPropertyBits=   1753237531
      DisplayStyle    =   3
      Size            =   "4260;556"
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
      Index           =   1
      Left            =   240
      TabIndex        =   12
      Top             =   2160
      Width           =   735
      ForeColor       =   4194368
      VariousPropertyBits=   8388627
      Caption         =   "Tenant"
      Size            =   "1296;450"
      FontName        =   "Myriad Web"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00808080&
      Height          =   3015
      Index           =   1
      Left            =   120
      Top             =   600
      Width           =   7455
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Height          =   3015
      Index           =   0
      Left            =   120
      Top             =   600
      Width           =   7455
   End
   Begin MSForms.ComboBox cmbClients 
      Height          =   315
      Left            =   2760
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
      Left            =   2160
      TabIndex        =   11
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
   Begin MSForms.Label Label1 
      Height          =   255
      Index           =   6
      Left            =   240
      TabIndex        =   10
      Top             =   1680
      Width           =   735
      ForeColor       =   4194368
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
      Left            =   2160
      TabIndex        =   3
      Top             =   1680
      Width           =   2415
      VariousPropertyBits=   1753237531
      DisplayStyle    =   3
      Size            =   "4260;556"
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
      Left            =   5040
      TabIndex        =   4
      Top             =   1680
      Width           =   2415
      VariousPropertyBits=   1753237531
      DisplayStyle    =   3
      Size            =   "4260;556"
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
      Left            =   6120
      TabIndex        =   6
      Top             =   3060
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
      Left            =   4440
      TabIndex        =   5
      Top             =   3060
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
      X2              =   7560
      Y1              =   2760
      Y2              =   2760
   End
   Begin MSForms.Label Label1 
      Height          =   255
      Index           =   5
      Left            =   2160
      TabIndex        =   9
      Top             =   720
      Width           =   495
      ForeColor       =   4194368
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
      Left            =   5040
      TabIndex        =   8
      Top             =   720
      Width           =   255
      ForeColor       =   4194368
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
      Index           =   0
      Left            =   240
      TabIndex        =   7
      Top             =   1200
      Width           =   1335
      ForeColor       =   4194368
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
      X2              =   7560
      Y1              =   2760
      Y2              =   2760
   End
End
Attribute VB_Name = "frmInsScheduleCriteria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmbClients_Click()
   Dim adoConn As New ADODB.Connection

   adoConn.Open getConnectionString

   LoadPropertyByClient adoConn
   LoadTenant adoConn

   adoConn.Close
   Set adoConn = Nothing
End Sub

Private Sub cmbPropertyFrom_Click()
   LoadTenantByProperty
End Sub

Private Sub cmbPropertyTo_Click()
   LoadTenantByProperty
End Sub

Private Sub cmdRPCCancel_Click()
   Unload Me
End Sub

Private Sub cmdRPCOK_Click()
'   Dim adoConn As New ADODB.Connection
'   Dim rstRst As New ADODB.Recordset
'   Dim szSQL As String, i As Integer, szaUnit() As String
   Dim reportApp As New CRAXDRT.Application
   Dim Report As CRAXDRT.Report

'   adoConn.Open getConnectionString
'
'   szSQL = "SELECT * " & _
'           "FROM tlbReceipt " & _
'           "WHERE RDate >= #" & Format(txtDateFrom.text, "dd mmmm yyyy") & "# AND " & _
'               "RDate <= #" & Format(txtDateto.text, "dd mmmm yyyy") & "# AND " & _
'               "TransactionID >= " & Val(txtjobnofrom.text) & " AND " & _
'               "TransactionID <= " & Val(txtjobnoto.text) & " AND " & _
'               "SageAccountNumber >= '" & IIf(txtCustRefFrom.text = "", "AAAAAA", txtCustRefFrom.text) & "' AND " & _
'               "SageAccountNumber <= '" & IIf(txtCustRefTo.text = "", "ZZZZZZ", txtCustRefTo.text) & "' AND " & _
'               "BankCode >= '" & Val(cmbBankAcFrom.Column(0)) & "' AND " & _
'               "BankCode <= '" & Val(cmbBankAcTo.Column(0)) & "';"
''Debug.Print szSQL
'   rstRst.Open szSQL, adoConn, adOpenStatic, adLockOptimistic

   Set Report = reportApp.OpenReport(App.Path & szReportPath & "\InsuranceSchedule.rpt")

   Report.EnableParameterPrompting = False
   Report.DiscardSavedData

'   While Not rstRst.EOF
'      Report.Database.SetDataSource rstRst
'      rstRst.MoveNext
'   Wend

   Report.ParameterFields(1).AddCurrentValue cmbClients.Column(1)
   Report.ParameterFields(2).AddCurrentValue CDate(IIf(txtDateFrom.text = "", Format(Date, "dd mmmm yyyy"), txtDateFrom.text))
   Report.ParameterFields(3).AddCurrentValue CDate(IIf(txtDateTo.text = "", Format(Date, "dd mmmm yyyy"), txtDateTo.text))
   Report.ParameterFields(4).AddCurrentValue cmbTenantFrom.Column(1)
   Report.ParameterFields(5).AddCurrentValue cmbTenantTo.Column(1)
   Report.ParameterFields(6).AddCurrentValue cmbPropertyFrom.Column(1)
   Report.ParameterFields(7).AddCurrentValue cmbPropertyTo.Column(1)

   Load frmReport
   frmReport.LoadReportViewer Report

'   rstRst.Close
'   Set rstRst = Nothing
'
'   adoConn.Close
'   Set adoConn = Nothing
End Sub

Private Sub Form_Load()
   Dim adoConn As New ADODB.Connection

   adoConn.Open getConnectionString

   Me.Top = 50
   Me.Left = 50
   Me.BackColor = MODULEBACKCOLOR
   txtDateTo.text = Format(Date, "DD/MM/YYYY")

   If Not LoadClients(adoConn) Then
      adoConn.Close
      Set adoConn = Nothing
      Unload Me
   End If
   LoadTenant adoConn

   adoConn.Close
   Set adoConn = Nothing
'
'   LoadTenantByProperty
End Sub

Private Sub LoadTenantByProperty()
   Dim iRec As Integer
   Dim adoRST As New ADODB.Recordset
   Dim szSQL As String, szaData() As String, szProperties As String
   Dim adoConn As New ADODB.Connection

   If cmbPropertyFrom.ListCount < 1 Then Exit Sub
   If cmbPropertyTo.ListIndex < 0 Then Exit Sub

   For iRec = IIf(cmbPropertyFrom.ListIndex < cmbPropertyTo.ListIndex, _
                  cmbPropertyFrom.ListIndex, cmbPropertyTo.ListIndex) To _
                  IIf(cmbPropertyFrom.ListIndex > cmbPropertyTo.ListIndex, _
                  cmbPropertyFrom.ListIndex, cmbPropertyTo.ListIndex)
      szProperties = szProperties + IIf(Len(szProperties) > 0, ",", "") + "'" + cmbPropertyFrom.Column(0, iRec) + "'"
   Next iRec
'Debug.Print szProperties

   adoConn.Open getConnectionString

   On Error GoTo Error_Handler

   szSQL = "SELECT Tenants.SageAccountNumber, Tenants.Name " & _
           "FROM Tenants, LeaseDetails, Units, Property " & _
           "WHERE Tenants.SageAccountNumber = LeaseDetails.SageAccountNumber AND " & _
               "LeaseDetails.UnitNumber = Units.UnitNumber AND " & _
               "LeaseDetails.Status = TRUE AND " & _
               "Units.PropertyID IN (" & szProperties & ") AND " & _
               "Units.PropertyID = Property.PropertyID " & _
           "ORDER BY Tenants.Name;"
'Debug.Print szSQL
   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   ReDim szaData(1, adoRST.RecordCount - 1) As String

   iRec = 0
   While Not adoRST.EOF
      szaData(0, iRec) = adoRST.Fields.Item("SageAccountNumber").Value
      szaData(1, iRec) = adoRST.Fields.Item("Name").Value
      iRec = iRec + 1
      adoRST.MoveNext
   Wend

   cmbTenantFrom.Clear
   cmbTenantFrom.Column() = szaData()
   cmbTenantFrom.ListIndex = 0
   cmbTenantTo.Clear
   cmbTenantTo.Column() = szaData()
   cmbTenantTo.ListIndex = cmbTenantTo.ListCount - 1

   ' Destroy Objects
   Set adoRST = Nothing
   adoConn.Close
   Set adoConn = Nothing

   Exit Sub

   ' Error Handling Code
Error_Handler:
   ' Destroy Objects
   Set adoRST = Nothing
   adoConn.Close
   Set adoConn = Nothing
End Sub

Private Sub LoadPropertyByClient(adoConn As ADODB.Connection)
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

Private Sub LoadTenant(adoConn As ADODB.Connection)
   Dim iRec As Integer
   Dim adoRST As New ADODB.Recordset
   Dim szSQL As String, szaData() As String

   On Error GoTo Error_Handler

   szSQL = "SELECT Tenants.SageAccountNumber, Tenants.Name " & _
           "FROM Tenants, LeaseDetails, Units, Property " & _
           "WHERE Tenants.SageAccountNumber = LeaseDetails.SageAccountNumber AND " & _
               "LeaseDetails.UnitNumber = Units.UnitNumber AND " & _
               "LeaseDetails.Status = TRUE AND " & _
               "Units.PropertyID = Property.PropertyID AND " & _
               "Property.ClientID = '" & cmbClients.Column(0) & "' " & _
           "ORDER BY Tenants.Name;"
'Debug.Print szSQL
   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   ReDim szaData(1, adoRST.RecordCount - 1) As String

   iRec = 0
   While Not adoRST.EOF
      szaData(0, iRec) = adoRST.Fields.Item("SageAccountNumber").Value
      szaData(1, iRec) = adoRST.Fields.Item("Name").Value
      iRec = iRec + 1
      adoRST.MoveNext
   Wend

   cmbTenantFrom.Clear
   cmbTenantFrom.Column() = szaData()
   cmbTenantFrom.ListIndex = 0
   cmbTenantTo.Clear
   cmbTenantTo.Column() = szaData()
   cmbTenantTo.ListIndex = cmbTenantTo.ListCount - 1

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
      MsgBox "Please input Client details into Prestige."
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

Private Sub txtDateFrom_Change()
   TextBoxChangeDate txtDateFrom
End Sub

Private Sub txtDateFrom_GotFocus()
   SelTxtInCtrl txtDateFrom
End Sub

Private Sub txtDateFrom_KeyPress(KeyAscii As Integer)
   TextBoxKeyPrsDate txtDateFrom, KeyAscii
End Sub

Private Sub txtDateFrom_LostFocus()
   TextBoxFormatDate txtDateFrom
End Sub

Private Sub txtDateTo_Change()
   TextBoxChangeDate txtDateTo
End Sub

Private Sub txtDateTo_KeyPress(KeyAscii As Integer)
   TextBoxKeyPrsDate txtDateTo, KeyAscii
End Sub

Private Sub txtDateTo_LostFocus()
   TextBoxFormatDate txtDateTo
End Sub

