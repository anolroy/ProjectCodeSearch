VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmRPCriteria 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Receipt Payment Criteria"
   ClientHeight    =   3480
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6585
   BeginProperty Font 
      Name            =   "Myriad Web"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRPCriteria.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3480
   ScaleWidth      =   6585
   Begin VB.Frame Frame2 
      Height          =   2580
      Left            =   45
      TabIndex        =   3
      Top             =   45
      Width           =   6495
      Begin VB.TextBox txtTranDateFrom 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2010
         TabIndex        =   5
         Text            =   "01/01/1980"
         Top             =   675
         Width           =   1695
      End
      Begin VB.TextBox txtTranDateTo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   4500
         TabIndex        =   4
         Top             =   675
         Width           =   1695
      End
      Begin MSForms.Label Label1 
         Height          =   210
         Index           =   8
         Left            =   90
         TabIndex        =   21
         Top             =   315
         Width           =   465
         VariousPropertyBits=   276824083
         Caption         =   "Client"
         Size            =   "820;370"
         FontName        =   "Myriad Web"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox cboClientID 
         Height          =   315
         Left            =   1980
         TabIndex        =   20
         Top             =   270
         Width           =   4260
         VariousPropertyBits=   1753237531
         DisplayStyle    =   3
         Size            =   "7514;556"
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
      Begin MSForms.Label Label1 
         Height          =   195
         Index           =   7
         Left            =   4095
         TabIndex        =   19
         Top             =   1980
         Width           =   195
         VariousPropertyBits=   276824083
         Caption         =   "To"
         Size            =   "344;344"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label1 
         Height          =   195
         Index           =   6
         Left            =   4095
         TabIndex        =   18
         Top             =   1530
         Width           =   195
         VariousPropertyBits=   276824083
         Caption         =   "To"
         Size            =   "344;344"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label1 
         Height          =   195
         Index           =   5
         Left            =   4095
         TabIndex        =   17
         Top             =   1125
         Width           =   195
         VariousPropertyBits=   276824083
         Caption         =   "To"
         Size            =   "344;344"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label1 
         Height          =   420
         Index           =   0
         Left            =   90
         TabIndex        =   16
         Top             =   675
         Width           =   1920
         VariousPropertyBits=   276824083
         Caption         =   "Transaction Date From"
         Size            =   "3387;741"
         FontName        =   "Myriad Web"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label1 
         Height          =   240
         Index           =   1
         Left            =   90
         TabIndex        =   15
         Top             =   1125
         Width           =   2040
         VariousPropertyBits=   276824083
         Caption         =   "Transaction ID from"
         Size            =   "3598;423"
         FontName        =   "Myriad Web"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label1 
         Height          =   195
         Index           =   2
         Left            =   90
         TabIndex        =   14
         Top             =   1530
         Width           =   1935
         VariousPropertyBits=   276824083
         Caption         =   "Customer Ref From"
         Size            =   "3413;344"
         FontName        =   "Myriad Web"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label1 
         Height          =   270
         Index           =   3
         Left            =   90
         TabIndex        =   13
         Top             =   1890
         Width           =   1785
         VariousPropertyBits=   276824083
         Caption         =   "Bank Code From"
         Size            =   "3149;476"
         FontName        =   "Myriad Web"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label1 
         Height          =   195
         Index           =   4
         Left            =   4095
         TabIndex        =   12
         Top             =   720
         Width           =   195
         VariousPropertyBits=   276824083
         Caption         =   "To"
         Size            =   "344;344"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtTranNoFrom 
         Height          =   315
         Left            =   2010
         TabIndex        =   11
         Top             =   1080
         Width           =   1695
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         Size            =   "2990;556"
         Value           =   "1"
         BorderColor     =   8421504
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   2
      End
      Begin MSForms.TextBox txtTranNoTo 
         Height          =   315
         Left            =   4500
         TabIndex        =   10
         Top             =   1080
         Width           =   1695
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         Size            =   "2990;556"
         Value           =   "999999"
         BorderColor     =   8421504
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtCustRefFrom 
         Height          =   315
         Left            =   2010
         TabIndex        =   9
         Top             =   1485
         Width           =   1695
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         Size            =   "2990;556"
         BorderColor     =   8421504
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtCustRefTo 
         Height          =   315
         Left            =   4500
         TabIndex        =   8
         Top             =   1485
         Width           =   1695
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         Size            =   "2990;556"
         BorderColor     =   8421504
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox cmbBankAcFrom 
         Height          =   315
         Left            =   2010
         TabIndex        =   7
         Top             =   1935
         Width           =   1695
         VariousPropertyBits=   1753237531
         DisplayStyle    =   3
         Size            =   "2990;556"
         BoundColumn     =   0
         TextColumn      =   1
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
      Begin MSForms.ComboBox cmbBankAcTo 
         Height          =   315
         Left            =   4500
         TabIndex        =   6
         Top             =   1935
         Width           =   1695
         VariousPropertyBits=   1753237531
         DisplayStyle    =   3
         Size            =   "2990;556"
         BoundColumn     =   0
         TextColumn      =   1
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
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   45
      TabIndex        =   0
      Top             =   2700
      Width           =   6495
      Begin MSForms.CommandButton cmdRPCOK 
         Default         =   -1  'True
         Height          =   375
         Left            =   1755
         TabIndex        =   2
         Top             =   225
         Width           =   1695
         Caption         =   "OK"
         Size            =   "2990;661"
         FontName        =   "Myriad Web"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton cmdRPCCancel 
         Height          =   375
         Left            =   3795
         TabIndex        =   1
         Top             =   225
         Width           =   1695
         Caption         =   "Cancel"
         Size            =   "2990;661"
         FontName        =   "Myriad Web"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
   End
End
Attribute VB_Name = "frmRPCriteria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdRPCCancel_Click()
   Unload Me
End Sub

Private Sub LoadClient(adoConn As ADODB.Connection)
'added by anol 05 Nov 2015
   Dim szSQL As String
  
   Dim adoRST As New ADODB.Recordset
   Dim TotalRow As Integer, TotalCol As Integer
   Dim i As Integer, j As Integer
   ' Clients
   szSQL = "SELECT CLIENTID, CLIENTNAME FROM CLIENT "
   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   TotalRow = adoRST.RecordCount
   TotalCol = adoRST.Fields.Count

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
   cboClientID.Column() = Data()
   cboClientID.ListIndex = 0
   adoRST.Close

   Set adoRST = Nothing
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
'           "WHERE RDate >= #" & Format(txtTranDateFrom.text, "dd mmmm yyyy") & "# AND " & _
'               "RDate <= #" & Format(txtTranDateTo.text, "dd mmmm yyyy") & "# AND " & _
'               "TransactionID >= " & Val(txtTranNoFrom.text) & " AND " & _
'               "TransactionID <= " & Val(txtTranNoTo.text) & " AND " & _
'               "SageAccountNumber >= '" & IIf(txtCustRefFrom.text = "", "AAAAAA", txtCustRefFrom.text) & "' AND " & _
'               "SageAccountNumber <= '" & IIf(txtCustRefTo.text = "", "ZZZZZZ", txtCustRefTo.text) & "' AND " & _
'               "BankCode >= '" & Val(cmbBankAcFrom.Column(0)) & "' AND " & _
'               "BankCode <= '" & Val(cmbBankAcTo.Column(0)) & "';"
''Debug.Print szSQL
'   rstRst.Open szSQL, adoConn, adOpenStatic, adLockOptimistic

   Set Report = reportApp.OpenReport(App.Path & szReportPath & "\ReceiptListing.rpt")

   Report.EnableParameterPrompting = False
   Report.DiscardSavedData

'   While Not rstRst.EOF
'      Report.Database.SetDataSource rstRst
'      rstRst.MoveNext
'   Wend

   Report.ParameterFields(1).AddCurrentValue CDate(txtTranDateFrom.text)
   Report.ParameterFields(2).AddCurrentValue Val(txtTranNoFrom.text)
   Report.ParameterFields(3).AddCurrentValue Val(txtTranNoTo.text)
   Report.ParameterFields(4).AddCurrentValue txtCustRefFrom.text
   Report.ParameterFields(5).AddCurrentValue txtCustRefTo.text
   Report.ParameterFields(6).AddCurrentValue CDate(txtTranDateTo.text)
   Report.ParameterFields(7).AddCurrentValue cmbBankAcFrom.Column(0)
   Report.ParameterFields(8).AddCurrentValue cmbBankAcTo.Column(0)
   Report.ParameterFields(9).AddCurrentValue cboClientID.Column(0)

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
   txtTranDateTo.text = Format(Date, "DD/MM/YYYY")
   LoadClient adoConn
   LoadBankCode adoConn

   adoConn.Close
   Set adoConn = Nothing
End Sub

Private Sub LoadBankCode(adoConn As ADODB.Connection)
   Dim iRec As Integer
   Dim adoRST As New ADODB.Recordset
   Dim szSQL As String, szaData() As String

   On Error GoTo Error_Handler

   szSQL = "SELECT tlbClientBanks.NominalCode AS BNC, NominalLedger.Name AS BNN " & _
           "FROM tlbClientBanks, NominalLedger " & _
           "WHERE tlbClientBanks.NominalCode = NominalLedger.Code " & _
           "GROUP BY tlbClientBanks.NominalCode, NominalLedger.Name, tlbClientBanks.CurrentBalance;"
'Debug.Print szSQL
   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If adoRST.EOF Then
      ShowMsgInTaskBar "Please setup bank account for the client."
   Else
      ReDim szaData(1, adoRST.RecordCount - 1) As String

      While Not adoRST.EOF
         szaData(0, iRec) = adoRST.Fields.Item("BNC").Value
         szaData(1, iRec) = adoRST.Fields.Item("BNN").Value
         iRec = iRec + 1
         adoRST.MoveNext
      Wend
   End If

   cmbBankAcFrom.Clear
   cmbBankAcFrom.Column() = szaData()
   cmbBankAcFrom.ListIndex = 0
   cmbBankAcTo.Clear
   cmbBankAcTo.Column() = szaData()
   cmbBankAcTo.ListIndex = cmbBankAcTo.ListCount - 1

   ' Destroy Objects
   Set adoRST = Nothing

   Exit Sub

   ' Error Handling Code
Error_Handler:
   ' Destroy Objects
   Set adoRST = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'frmMMain.fraCmdButton.Enabled = True
End Sub

Private Sub TextBox7_Change()

End Sub

Private Sub txtTranDateFrom_Change()
   TextBoxChangeDate txtTranDateFrom
End Sub

Private Sub txtTranDateFrom_GotFocus()
   SelTxtInCtrl txtTranDateFrom
End Sub

Private Sub txtTranDateFrom_KeyPress(KeyAscii As Integer)
   TextBoxKeyPrsDate txtTranDateFrom, KeyAscii
End Sub

Private Sub txtTranDateFrom_LostFocus()
   TextBoxFormatDate txtTranDateFrom
End Sub

Private Sub txtTranDateTo_Change()
   TextBoxChangeDate txtTranDateTo
End Sub

Private Sub txtTranDateTo_KeyPress(KeyAscii As Integer)
   TextBoxKeyPrsDate txtTranDateTo, KeyAscii
End Sub

Private Sub txtTranDateTo_LostFocus()
   TextBoxFormatDate txtTranDateTo
End Sub
