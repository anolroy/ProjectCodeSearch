VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmFundSumry 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Fund Summary Report"
   ClientHeight    =   9450
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6210
   BeginProperty Font 
      Name            =   "Myriad Web"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmFundSumry.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9450
   ScaleWidth      =   6210
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chkDetails 
      Caption         =   "Detailed Transaction"
      Height          =   255
      Left            =   1800
      TabIndex        =   18
      Top             =   6300
      Width           =   1935
   End
   Begin VB.PictureBox picClientList 
      Appearance      =   0  'Flat
      BackColor       =   &H00C5C5C5&
      BeginProperty Font 
         Name            =   "Myriad Condensed Web"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2480
      Left            =   120
      ScaleHeight     =   2445
      ScaleWidth      =   5520
      TabIndex        =   16
      Top             =   6600
      Visible         =   0   'False
      Width           =   5555
      Begin VB.CommandButton cmdGridUnitLookup 
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5230
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   10
         Width           =   295
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxClientList 
         Height          =   2355
         Left            =   40
         TabIndex        =   19
         Top             =   45
         Width           =   5445
         _ExtentX        =   9604
         _ExtentY        =   4154
         _Version        =   393216
         Cols            =   5
         FixedCols       =   0
         BackColorFixed  =   13553358
         ForeColorFixed  =   12632256
         BackColorSel    =   12648447
         ForeColorSel    =   -2147483640
         BackColorBkg    =   16777215
         GridColor       =   -2147483638
         WordWrap        =   -1  'True
         GridLinesFixed  =   1
         SelectionMode   =   1
         Appearance      =   0
         BandDisplay     =   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   5
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
   End
   Begin VB.TextBox txtLlName 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   120
      Width           =   3375
   End
   Begin VB.Frame fraFunds 
      Caption         =   "Funds:"
      Height          =   1935
      Left            =   570
      TabIndex        =   15
      Top             =   4200
      Width           =   5565
      Begin VB.CheckBox chkFunds 
         Caption         =   "All Funds"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   1095
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxFunds 
         Height          =   1335
         Left            =   135
         TabIndex        =   5
         Top             =   525
         Width           =   5325
         _ExtentX        =   9393
         _ExtentY        =   2355
         _Version        =   393216
         FixedCols       =   0
         BackColorFixed  =   12632256
         ForeColorFixed  =   -2147483640
         BackColorSel    =   15329508
         ForeColorSel    =   -2147483640
         BackColorBkg    =   16777215
         BackColorUnpopulated=   -2147483643
         GridColor       =   -2147483638
         GridColorFixed  =   8421504
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
   End
   Begin VB.Frame fraBankAccounts 
      Caption         =   "Bank Accounts:"
      Enabled         =   0   'False
      Height          =   1935
      Left            =   120
      TabIndex        =   14
      Top             =   480
      Width           =   5565
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxBankAccounts 
         Height          =   1575
         Left            =   120
         TabIndex        =   1
         Top             =   285
         Width           =   5325
         _ExtentX        =   9393
         _ExtentY        =   2778
         _Version        =   393216
         FixedCols       =   0
         BackColorFixed  =   12632256
         ForeColorFixed  =   -2147483640
         BackColorSel    =   15329508
         ForeColorSel    =   -2147483640
         BackColorBkg    =   16777215
         BackColorUnpopulated=   -2147483643
         GridColor       =   -2147483638
         GridColorFixed  =   8421504
         WordWrap        =   -1  'True
         GridLinesFixed  =   1
         ScrollBars      =   1
         Appearance      =   0
         BandDisplay     =   1
         RowSizingMode   =   1
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
   End
   Begin VB.Frame Frame2 
      Caption         =   "Statement Date:"
      Height          =   735
      Left            =   120
      TabIndex        =   11
      Top             =   3480
      Width           =   5565
      Begin MSForms.ComboBox cboLtStDt 
         Height          =   315
         Left            =   3480
         TabIndex        =   2
         Top             =   240
         Width           =   1815
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "3201;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Last: "
         Height          =   255
         Index           =   1
         Left            =   3000
         TabIndex        =   13
         Top             =   240
         Width           =   735
      End
      Begin MSForms.ComboBox cboCurStDt 
         Height          =   315
         Left            =   960
         TabIndex        =   3
         Top             =   240
         Width           =   1815
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "3201;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Current:"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   12
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.CommandButton cmdClient 
      Caption         =   "V"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      Width           =   285
   End
   Begin VB.TextBox txtClientID 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   120
      Width           =   1090
   End
   Begin VB.CommandButton cmdGenReport 
      Caption         =   "&Generate Report"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   6225
      Width           =   1575
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   375
      Left            =   4080
      TabIndex        =   7
      Top             =   6240
      Width           =   1575
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFDFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Client:"
      Height          =   195
      Index           =   3
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   465
   End
End
Attribute VB_Name = "frmFundSumry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Samrat Rahman: On 07/02/2012
'First this form was created for Landlord Summary Statement.
'We are changing this form for Client Summary Statement.
'This refer to the modification document on page 126: Client Summary Statement.
'fund summary statement
Option Explicit

Private szDemandTypes      As String
Private bCallingFromGrid   As Boolean
Private szBanks            As String
Private szFunds            As String
Private szFundList         As String
Private Const INI_HEIGHT   As Integer = 6210
Private cBBF               As Currency
Private ReConDates()       As String
Private dtLastStDate       As Date
'Public bProceed            As Boolean

Private Sub cboCurStDt_Click()
   Dim i As Integer

   If cboCurStDt.text = "" Then Exit Sub
   cboLtStDt.Clear
   cboLtStDt.Column() = ReConDates()

   On Error GoTo EndOfRec

   For i = 0 To cboLtStDt.ListCount - 1
      If CDate(cboCurStDt.text) <= CDate(cboLtStDt.List(i)) Then
         cboLtStDt.RemoveItem (i)
         i = i - 1
      End If
   Next i

EndOfRec:
   Exit Sub
End Sub

Private Sub cboCurStDt_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = 0
End Sub

Private Sub cboLtStDt_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = 0
End Sub
'
'Private Sub chkBankAccounts_Click()
'   If bCallingFromGrid Then
'      bCallingFromGrid = False
'      Exit Sub
'   End If
'
'   Dim iRow As Integer
'
'   For iRow = 1 To flxBankAccounts.Rows - 1
'      If flxBankAccounts.RowHeight(iRow) > 0 And flxBankAccounts.TextMatrix(iRow, 0) = "X" Then
'         SelectFlxGridRow 0, flxBankAccounts, iRow
'      End If
'   Next iRow
'
'   For iRow = 1 To flxBankAccounts.Rows - 1
'      If flxBankAccounts.RowHeight(iRow) > 0 And chkBankAccounts.Value Then
'         SelectFlxGridRow 0, flxBankAccounts, iRow
'      End If
'   Next iRow
'End Sub

Private Sub chkFunds_Click()
   If bCallingFromGrid Then
      bCallingFromGrid = False
      Exit Sub
   End If

   Dim irow As Integer

   For irow = 1 To flxFunds.Rows - 1
      If flxFunds.RowHeight(irow) > 0 And flxFunds.TextMatrix(irow, 0) = "X" Then
         SelectFlxGridRow 0, flxFunds, irow
      End If
   Next irow

   For irow = 1 To flxFunds.Rows - 1
      If flxFunds.RowHeight(irow) > 0 And chkFunds.Value Then
         SelectFlxGridRow 0, flxFunds, irow
      End If
   Next irow
End Sub

Private Sub cmdClient_Click()
   picClientList.Top = txtClientID.Top + txtClientID.Height + 5
   picClientList.Left = Label1(3).Left + 5
   picClientList.Visible = True
   picClientList.ZOrder 0
   Me.Height = 7935
End Sub

Private Sub cmdClose_Click()
   Unload Me
End Sub

Private Function UpdateDatabaseInvPropertyID(Conn1 As ADODB.Connection)
    Dim Rst1 As New ADODB.Recordset
   On Error GoTo CHANGE_ADD_invPropertyID_tlbBankReconcilation

   Rst1.Open "SELECT invPropertyID FROM tlbBankReconcilation;", Conn1, adOpenStatic, adLockReadOnly

   Rst1.Close

   Exit Function

CHANGE_ADD_invPropertyID_tlbBankReconcilation:
   Conn1.Execute "ALTER TABLE tlbBankReconcilation ADD COLUMN invPropertyID TEXT(4);"
'

    
End Function
Private Function MarkInvPropertyID(Conn1 As ADODB.Connection)
   Conn1.Execute "UPDATE (tlbPaymentSplit INNER JOIN (((tlbPayment AS tlbPayment_1 INNER JOIN PayTransactions ON tlbPayment_1.TransactionID = PayTransactions.ToTran) " & _
   "INNER JOIN tblPurInv ON tlbPayment_1.PI = tblPurInv.MY_ID) INNER JOIN ((tlbBankReconcilation INNER JOIN tlbPayment ON tlbBankReconcilation.RefID = tlbPayment.szTransactionID)" & _
   "INNER JOIN Fund ON tlbBankReconcilation.FundID = Fund.FundID) ON PayTransactions.FromTran = tlbPayment.TransactionID) ON tlbPaymentSplit.PayHeader = tlbPayment.TransactionID)" & _
   "INNER JOIN tlbClientBanks ON tlbPayment.NominalCode = tlbClientBanks.NominalCode SET tlbBankReconcilation.invPropertyID=tblPurInv.PropertyID "
   
'    Debug.Print "UPDATE (tlbPaymentSplit INNER JOIN (((tlbPayment AS tlbPayment_1 INNER JOIN PayTransactions ON tlbPayment_1.TransactionID = " & _
'    "PayTransactions.ToTran) INNER JOIN tblPurInv ON tlbPayment_1.PI = tblPurInv.MY_ID) INNER JOIN ((tlbBankReconcilation INNER JOIN tlbPayment ON " & _
'    "tlbBankReconcilation.RefID = tlbPayment.szTransactionID) INNER JOIN Fund ON tlbBankReconcilation.FundID = Fund.FundID) ON ayTransactions.FromTran" & _
'    "= tlbPayment.TransactionID) ON tlbPaymentSplit.PayHeader = tlbPayment.TransactionID) INNER JOIN tlbClientBanks ON tlbPayment.NominalCode = tlbClientBanks.NominalCode " & _
'    " SET tlbBankReconcilation.invPropertyID=tblPurInv.PropertyID WHERE (((tlbBankReconcilation.TransactionType)=8 Or (tlbBankReconcilation.TransactionType)=9 Or (tlbBankReconcilation.TransactionType)=24) AND  " & _
'    "((Fund.SelFund)='Y') AND ((tlbPayment.SageAccountNumber)<>'LITCHFIE') AND ((tlbBankReconcilation.ReconDate)=#8/1/2018#) AND " & _
'    "((tlbClientBanks.SelBanks)='Y'));"
'
'     Conn1.Execute "UPDATE (tlbPaymentSplit INNER JOIN (((tlbPayment AS tlbPayment_1 INNER JOIN PayTransactions ON tlbPayment_1.TransactionID = " & _
'    "PayTransactions.ToTran) INNER JOIN tblPurInv ON tlbPayment_1.PI = tblPurInv.MY_ID) INNER JOIN ((tlbBankReconcilation INNER JOIN tlbPayment ON " & _
'    "tlbBankReconcilation.RefID = tlbPayment.szTransactionID) INNER JOIN Fund ON tlbBankReconcilation.FundID = Fund.FundID) ON ayTransactions.FromTran" & _
'    "= tlbPayment.TransactionID) ON tlbPaymentSplit.PayHeader = tlbPayment.TransactionID) INNER JOIN tlbClientBanks ON tlbPayment.NominalCode = tlbClientBanks.NominalCode " & _
'    " SET tlbBankReconcilation.invPropertyID=tblPurInv.PropertyID WHERE (((tlbBankReconcilation.TransactionType)=8 Or (tlbBankReconcilation.TransactionType)=9 Or (tlbBankReconcilation.TransactionType)=24) AND  " & _
'    "((Fund.SelFund)='Y') AND ((tlbPayment.SageAccountNumber)<>'LITCHFIE') AND ((tlbBankReconcilation.ReconDate)=#8/1/2018#) AND " & _
'    "((tlbClientBanks.SelBanks)='Y'));"
End Function
Private Sub cmdGenReport_Click()
   If txtClientID.text = "" Then
      MsgBox "Please a Client.", vbInformation + vbOKOnly, "Client"
      cmdClient.SetFocus
      Exit Sub
   End If
   If cboCurStDt.text = "" Then
      MsgBox "Please select current statement date.", vbInformation + vbOKOnly, "From Date"
      cboCurStDt.SetFocus
      Exit Sub
   End If

   If cboLtStDt.text = "" Then
      If cboLtStDt.ListCount > 1 Then
         If MsgBox("Do you wish to leave blank last statement date?", vbInformation + vbYesNo, "To Date") = vbYes Then
            dtLastStDate = Format(#1/1/2000#, "dd/mm/yyyy")
         Else
            cboLtStDt.SetFocus
            Exit Sub
         End If
      Else
         dtLastStDate = Format(#1/1/2000#, "dd/mm/yyyy")
      End If
   Else
      dtLastStDate = CDate(cboLtStDt.text)
   End If

   szBanks = ""
   If Not IsBankSelected Then
      MsgBox "Please select a bank account.", vbInformation + vbOKOnly, "Bank Account"
      flxBankAccounts.SetFocus
'      chkBankAccounts.SetFocus
      Exit Sub
   End If
   
   szFunds = ""
   szFundList = ""
   If Not IsFundSelected Then
      MsgBox "Please select a fund.", vbInformation + vbOKOnly, "Funds"
      chkFunds.SetFocus
      Exit Sub
   End If
   Dim adoConn  As New ADODB.Connection
   adoConn.Open getConnectionString
   
   MarkBankFund adoConn
   UpdateDatabaseInvPropertyID adoConn
   MarkInvPropertyID adoConn
   
   adoConn.Close
   Set adoConn = Nothing
   
   Dim reportApp As New CRAXDRT.Application
   Dim Report As CRAXDRT.Report
   Dim fReport As frmReport

   Set Report = reportApp.OpenReport(App.Path & szReportPath & "\FSR.rpt")

   Report.EnableParameterPrompting = False
   Report.DiscardSavedData

   Report.ParameterFields(1).AddCurrentValue CDate(dtLastStDate)
   Report.ParameterFields(2).AddCurrentValue CDate(cboCurStDt.text)
   'Report.ParameterFields(1).AddCurrentValue Format(dtLastStDate, "yyyy-mm-dd")
   'Report.ParameterFields(2).AddCurrentValue Format(cboCurStDt.text, "yyyy-mm-dd")
   Report.ParameterFields(3).AddCurrentValue txtClientID.text
   Report.ParameterFields(4).AddCurrentValue Val(cBBF)

   Set fReport = New frmReport
   Load fReport
   fReport.LoadReportViewer Report

   If chkDetails.Value = 0 Then Exit Sub

   Set Report = reportApp.OpenReport(App.Path & szReportPath & "\FSR_Details.rpt")

   Report.EnableParameterPrompting = False
   Report.DiscardSavedData

   Report.ParameterFields(1).AddCurrentValue CDate(dtLastStDate)
   Report.ParameterFields(2).AddCurrentValue CDate(cboCurStDt.text)
   Report.ParameterFields(3).AddCurrentValue txtClientID.text
   Report.ParameterFields(4).AddCurrentValue Val(cBBF)

   Set fReport = New frmReport
   Load fReport
   fReport.LoadReportViewer Report
End Sub

Private Sub CalculateBBF(adoConn As ADODB.Connection)
   Dim szSQL   As String
   Dim cR      As Currency
   Dim cP      As Currency
   Dim cB      As Currency
   Dim adoRst  As New ADODB.Recordset

'----------------------------------------------  SR & RoA
   szSQL = "SELECT SUM(S.Amount) AS T " & _
           "FROM   tlbBankReconcilation AS RC, " & _
                  "tlbReceipt AS R, " & _
                  "tlbReceiptSplit AS S, " & _
                  "Fund AS F, tlbClientBanks AS CB " & _
           "WHERE  RC.RefID = R.szTransactionID AND " & _
                  "R.TransactionID = S.RptHeader AND " & _
                  "S.FundID = F.FundID AND " & _
                  "R.BankCode = CB.NominalCode AND " & _
                  "RC.ReconDate < #" & Format(dtLastStDate, "dd mmmm yyyy") & "# AND " & _
                  "F.SelFund = 'Y' AND CB.SelBanks = 'Y' AND " & _
                  "R.TYPE IN (3, 4);"
   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If Not adoRst.EOF Then
      cR = IIf(IsNull(adoRst.Fields.Item(0).Value), 0, adoRst.Fields.Item(0).Value)
   Else
      cR = 0
   End If
   adoRst.Close
'------------------------------------------------  SRR
   szSQL = "SELECT SUM(S.Amount) AS T " & _
           "FROM   tlbBankReconcilation AS RC, " & _
                  "tlbReceipt AS R, " & _
                  "tlbReceiptSplit AS S, " & _
                  "Fund AS F, tlbClientBanks AS CB " & _
           "WHERE  RC.RefID = R.szTransactionID AND " & _
                  "R.TransactionID = S.RptHeader AND " & _
                  "S.FundID = F.FundID AND " & _
                  "R.BankCode = CB.NominalCode AND " & _
                  "RC.ReconDate < #" & Format(dtLastStDate, "dd mmmm yyyy") & "# AND " & _
                  "F.SelFund = 'Y' AND CB.SelBanks = 'Y' AND " & _
                  "R.TYPE = 23;"
'Debug.Print szSQL
   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If Not adoRst.EOF Then cR = cR - IIf(IsNull(adoRst.Fields.Item(0).Value), 0, adoRst.Fields.Item(0).Value)
   adoRst.Close
'-------------------------------------------------------------  PP & PoA
   szSQL = "SELECT SUM(S.Amount) AS T " & _
           "FROM   tlbBankReconcilation AS RC, " & _
                  "tlbPayment AS P, " & _
                  "tlbPaymentSplit AS S, " & _
                  "Fund AS F, tlbClientBanks AS CB " & _
           "WHERE  RC.RefID = CSTR(P.TransactionID) AND " & _
                  "P.TransactionID = S.PayHeader AND " & _
                  "S.FundID = F.FundID AND " & _
                  "P.BankCode = CB.NominalCode AND " & _
                  "RC.ReconDate < #" & Format(dtLastStDate, "dd mmmm yyyy") & "# AND " & _
                  "F.SelFund = 'Y' AND CB.SelBanks = 'Y' AND " & _
                  "P.TYPE IN (8, 9);"
'Debug.Print szSQL
   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If Not adoRst.EOF Then
      cP = IIf(IsNull(adoRst.Fields.Item(0).Value), 0, adoRst.Fields.Item(0).Value)
   Else
      cP = 0
   End If
   adoRst.Close
'-------------------------------------------------------------  PPR
   szSQL = "SELECT SUM(S.Amount) AS T " & _
           "FROM   tlbBankReconcilation AS RC, " & _
                  "tlbPayment AS P, " & _
                  "tlbPaymentSplit AS S, " & _
                  "Fund AS F, tlbClientBanks AS CB " & _
           "WHERE  RC.RefID = CSTR(P.TransactionID) AND " & _
                  "P.TransactionID = S.PayHeader AND " & _
                  "S.FundID = F.FundID AND " & _
                  "P.BankCode = CB.NominalCode AND " & _
                  "RC.ReconDate < #" & Format(dtLastStDate, "dd mmmm yyyy") & "# AND " & _
                  "F.SelFund = 'Y' AND CB.SelBanks = 'Y' AND " & _
                  "P.TYPE = 24;"
   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If Not adoRst.EOF Then cP = cP - IIf(IsNull(adoRst.Fields.Item(0).Value), 0, adoRst.Fields.Item(0).Value)
   adoRst.Close
'-----------------------------------------------------------  BR
   szSQL = "SELECT SUM(B.NET_AMOUNT + B.VAT) AS T " & _
           "FROM   tlbBankReconcilation AS RC, " & _
                  "tlbBankPayment AS B, " & _
                  "Fund AS F, tlbClientBanks AS CB " & _
           "WHERE  RC.RefID = B.MY_ID AND " & _
                  "CLNG(B.DEPT_ID) = F.FundID AND " & _
                  "B.BANK_AC = CB.NominalCode AND " & _
                  "RC.ReconDate < #" & Format(dtLastStDate, "dd mmmm yyyy") & "# AND " & _
                  "F.SelFund = 'Y' AND CB.SelBanks = 'Y' AND " & _
                  "B.TransactionType = 12;"
   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If Not adoRst.EOF Then
      cB = IIf(IsNull(adoRst.Fields.Item(0).Value), 0, adoRst.Fields.Item(0).Value)
   Else
      cB = 0
   End If
   adoRst.Close
'-----------------------------------------------------------  BP
   szSQL = "SELECT SUM(B.NET_AMOUNT + B.VAT) AS T " & _
           "FROM   tlbBankReconcilation AS RC, " & _
                  "tlbBankPayment AS B, " & _
                  "Fund AS F, tlbClientBanks AS CB " & _
           "WHERE  RC.RefID = B.MY_ID AND " & _
                  "CLNG(B.DEPT_ID) = F.FundID AND " & _
                  "B.BANK_AC = CB.NominalCode AND " & _
                  "RC.ReconDate < #" & Format(dtLastStDate, "dd mmmm yyyy") & "# AND " & _
                  "F.SelFund = 'Y' AND CB.SelBanks = 'Y' AND " & _
                  "B.TransactionType = 11;"
   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If Not adoRst.EOF Then cB = cB - IIf(IsNull(adoRst.Fields.Item(0).Value), 0, adoRst.Fields.Item(0).Value)
   adoRst.Close
'######################################################################

   cBBF = cR - cP + cB
End Sub

Private Sub cmdGridUnitLookup_Click()
   If txtClientID.text = "" Then Exit Sub

   picClientList.Visible = False
   'Me.Height = INI_HEIGHT

   fraBankAccounts.Enabled = True
   fraFunds.Enabled = True
End Sub

Private Sub flxBankAccounts_Click()
   If flxBankAccounts.row = 0 Then Exit Sub

   SelectOnly1RowFlxGrid flxBankAccounts, flxBankAccounts.row, 0
   
  ' Me.Height = INI_HEIGHT
'
'   SelectFlxGridRow 0, flxBankAccounts, flxBankAccounts.row
'
'   Dim iRow As Integer
'
'   For iRow = 1 To flxBankAccounts.Rows - 1
'      If flxBankAccounts.TextMatrix(iRow, 0) <> "X" And chkBankAccounts.Value Then
'         bCallingFromGrid = True
'         chkBankAccounts.Value = 0
'         Exit For
'      End If
'   Next iRow
    Dim conClient As New ADODB.Connection
    conClient.Open getConnectionString

    LoadTransactionDates conClient
    conClient.Close
    Set conClient = Nothing
End Sub

Private Sub flxClientList_Click()
   If flxClientList.TextMatrix(flxClientList.row, 1) = "" Then Exit Sub

   txtClientID.text = flxClientList.TextMatrix(flxClientList.row, 1)
   txtLlName.text = flxClientList.TextMatrix(flxClientList.row, 2)

   picClientList.Visible = False
   Me.Height = 7935

   fraBankAccounts.Enabled = True
   fraFunds.Enabled = True

   FilteringBanks
End Sub

Private Sub FilteringDates()
   Dim iCmb As Integer
   Dim irow As Integer

   cboCurStDt.Clear
   cboCurStDt.Column() = ReConDates()
   cboLtStDt.Clear
   cboLtStDt.Column() = ReConDates()
   For irow = 1 To flxBankAccounts.Rows - 1
      If flxBankAccounts.RowHeight(irow) = 0 Then
         iCmb = 0
         Do While (iCmb <= cboCurStDt.ListCount - 1)
            If cboCurStDt.Column(1, iCmb) = flxBankAccounts.TextMatrix(irow, 2) Then
               cboCurStDt.RemoveItem (iCmb)
               iCmb = iCmb - 1
            End If
            iCmb = iCmb + 1
         Loop

         iCmb = 0
         Do While (iCmb <= cboLtStDt.ListCount - 1)
            If cboLtStDt.Column(1, iCmb) = flxBankAccounts.TextMatrix(irow, 2) Then
               cboLtStDt.RemoveItem (iCmb)
               iCmb = iCmb - 1
            End If
            iCmb = iCmb + 1
         Loop
      End If
   Next irow
End Sub

Private Sub FilteringBanks()
   Dim irow As Integer

   For irow = 1 To flxBankAccounts.Rows - 1
      flxBankAccounts.RowHeight(irow) = 240
   Next irow

   For irow = 1 To flxBankAccounts.Rows - 1
      If flxBankAccounts.TextMatrix(irow, 4) <> txtClientID.text Then
         flxBankAccounts.RowHeight(irow) = 0
      End If
   Next irow

   'FilteringDates
End Sub

Private Sub flxFunds_Click()
   If flxFunds.row = 0 Then Exit Sub

   SelectFlxGridRow 0, flxFunds, flxFunds.row

   Dim irow As Integer

   For irow = 1 To flxFunds.Rows - 1
      If flxFunds.TextMatrix(irow, 0) <> "X" And chkFunds.Value Then
         bCallingFromGrid = True
         chkFunds.Value = 0
         Exit For
      End If
   Next irow
End Sub

Private Sub Form_Load()
   Dim conClient As New ADODB.Connection
   frmMMain.Arrange vbCascade
   Me.ZOrder 0
  ' Me.Height = INI_HEIGHT
   Me.Width = 5895
   Me.BackColor = MODULEBACKCOLOR
   fraBankAccounts.BackColor = MODULEBACKCOLOR
   Frame2.BackColor = MODULEBACKCOLOR
   fraFunds.BackColor = MODULEBACKCOLOR
   chkDetails.BackColor = MODULEBACKCOLOR
   chkFunds.BackColor = fraFunds.BackColor

   'bProceed = True

   bCallingFromGrid = False

   conClient.Open getConnectionString

   PrepareList conClient

   conClient.Close
   Set conClient = Nothing

   'If bProceed Then
      cmdClient_Click
      chkFunds.Value = 1
   'End If
   Me.Height = 7935
   Call WheelHook(Me.hWnd)
End Sub

Private Sub PrepareList(conClient As ADODB.Connection)
'   LoadTransactionDates conClient
'   If Not bProceed Then Exit Sub

   ConfigflxClientList
   ConfigFlxBankAccounts
   ConfigFlxFunds

   LoadflxClientList conClient
   LoadFlxBankAccounts conClient
   LoadFlxFunds conClient
End Sub

Private Sub LoadFlxFunds(conClient As ADODB.Connection)
   Dim rstClient   As New ADODB.Recordset
   Dim szSQL       As String
   Dim irow As Integer

   On Error GoTo ErrorHandler

   szSQL = "SELECT F.FundID, F.FundName, S.Value " & _
           "FROM Fund AS F, SecondaryCode AS S " & _
           "WHERE F.CategoryCode = CBYTE(S.Code) AND S.PrimaryCode = 'DCTG' " & _
           "ORDER BY FundID;"

   rstClient.Open szSQL, conClient, adOpenStatic, adLockReadOnly

   irow = 1

   While Not rstClient.EOF
      flxFunds.TextMatrix(irow, 1) = rstClient!fundID
      flxFunds.TextMatrix(irow, 2) = rstClient!FundName
      flxFunds.TextMatrix(irow, 3) = rstClient!Value
      rstClient.MoveNext
      If Not rstClient.EOF Then flxFunds.AddItem ""
      irow = irow + 1
   Wend

NoRes:
   rstClient.Close
   Set rstClient = Nothing
   Exit Sub

ErrorHandler:
   MsgBox Err.description & "::" & Err.Number

   Set rstClient = Nothing
End Sub

Private Sub LoadFlxBankAccounts(conClient As ADODB.Connection)
   Dim rstClient   As New ADODB.Recordset
   Dim szSQL       As String
   Dim irow As Integer

   On Error GoTo ErrorHandler

   szSQL = "SELECT MY_ID, NominalCode, Bank_AC_Name, CLIENT_ID " & _
           "FROM tlbClientBanks " & _
           "ORDER BY NominalCode;"

   rstClient.Open szSQL, conClient, adOpenStatic, adLockReadOnly

   irow = 1

   While Not rstClient.EOF
      flxBankAccounts.TextMatrix(irow, 1) = rstClient!My_ID
      flxBankAccounts.TextMatrix(irow, 2) = rstClient!nominalCode
      flxBankAccounts.TextMatrix(irow, 3) = rstClient!Bank_AC_Name
      flxBankAccounts.TextMatrix(irow, 4) = rstClient!CLIENT_ID
      rstClient.MoveNext
      If Not rstClient.EOF Then flxBankAccounts.AddItem ""
      irow = irow + 1
   Wend

NoRes:
   rstClient.Close
   Set rstClient = Nothing
   Exit Sub

ErrorHandler:
   MsgBox Err.description & "::" & Err.Number

   Set rstClient = Nothing
End Sub

Private Sub LoadflxClientList(conClient As ADODB.Connection)
   Dim rstClient   As New ADODB.Recordset
   Dim szSQL       As String
   Dim irow As Integer

   On Error GoTo ErrorHandler
   szSQL = "SELECT ClientID, ClientName " & _
           "FROM Client " & _
           "ORDER BY ClientName;"

   rstClient.Open szSQL, conClient, adOpenStatic, adLockReadOnly

   irow = 1

   While Not rstClient.EOF
      flxClientList.TextMatrix(irow, 1) = rstClient!clientID
      flxClientList.TextMatrix(irow, 2) = rstClient!ClientName
      flxClientList.TextMatrix(irow, 3) = "Client"
      rstClient.MoveNext
      If Not rstClient.EOF Then flxClientList.AddItem ""
      irow = irow + 1
   Wend

NoRes:
   rstClient.Close
   Set rstClient = Nothing
   Exit Sub

ErrorHandler:
   MsgBox Err.description & "::" & Err.Number

   Set rstClient = Nothing
End Sub

Private Sub ConfigflxClientList()
   Dim szHeader As String

   flxClientList.Cols = 4
   flxClientList.Clear
   szHeader$ = "|<ID|<Name|<Type"
   flxClientList.FormatString = szHeader$
   flxClientList.ColWidth(0) = 300        'Solid column
   flxClientList.ColWidth(1) = 900        'Client ID
   flxClientList.ColWidth(2) = 3000       'Client Name
   flxClientList.ColWidth(3) = 800        'Post Code
   flxClientList.Rows = 2

   flxClientList.RowHeightMin = 240
End Sub

Private Sub ConfigFlxBankAccounts()
   Dim szHeader As String

   flxBankAccounts.Cols = 5
   flxBankAccounts.Clear
   szHeader$ = "|<ID|<Account|<Name|Client_ID"
   flxBankAccounts.FormatString = szHeader$
   flxBankAccounts.ColWidth(0) = 280        'Solid column
   flxBankAccounts.ColWidth(1) = 0          'ID
   flxBankAccounts.ColWidth(2) = 900        'Account
   flxBankAccounts.ColWidth(3) = 3800       'Name
   flxBankAccounts.ColWidth(4) = 0          'ClientID
   flxBankAccounts.Rows = 2
End Sub

Private Sub ConfigFlxFunds()
   Dim szHeader As String

   flxFunds.Cols = 4
   flxFunds.Clear
   szHeader$ = "|<ID|<Name|<Category"
   flxFunds.FormatString = szHeader$
   flxFunds.ColWidth(0) = 280        'Solid column
   flxFunds.ColWidth(1) = 400        'ID
   flxFunds.ColWidth(2) = 2800       'Name
   flxFunds.ColWidth(3) = 1500       'Post Code
   flxFunds.Rows = 2

   flxFunds.RowHeightMin = 255
End Sub

Private Sub LoadTransactionDates(adoConn As ADODB.Connection)
   Dim adoRst  As New ADODB.Recordset
   Dim szSQL   As String
   Dim r       As Integer

'   szSQL = "SELECT LEFT(R.ReconNow, 10) AS SDate, FORMAT(LEFT(R.ReconNow, 10), 'YYYYMMDD') AS X " & _
'           "FROM tlbReceipt AS R, tlbTransactionTypes AS TT " & _
'           "WHERE (Type = 3 OR Type = 4 OR Type = 23) AND TT.TYPE_ID = R.Type " & _
'           "Union " & _
'           "SELECT LEFT(BP.ReconNow, 10) AS SDate, FORMAT(LEFT(BP.ReconNow, 10), 'YYYYMMDD') AS X " & _
'           "FROM tlbBankPayment AS BP, tlbTransactionTypes AS TT " & _
'           "WHERE (BP.TransactionType = 11 OR BP.TransactionType = 12) AND BP.TransactionType = TT.TYPE_ID " & _
'           "Union " & _
'           "SELECT LEFT(P.ReconNow, 10) AS SDate, FORMAT(LEFT(P.ReconNow, 10), 'YYYYMMDD') AS X " & _
'           "FROM tlbPayment AS P, tlbTransactionTypes AS TT " & _
'           "WHERE (P.Type = 8 OR P.Type = 9 OR P.Type = 24) AND P.Type = TT.TYPE_ID " & _
'           "GROUP BY ReconNow " & _
'           "ORDER BY X DESC;"
'
'   szSQL = "SELECT ReconDate, BankCode " & _
'           "FROM tlbBankReconcilation " & _
'           "GROUP BY ReconDate, BankCode " & _
'           "ORDER BY ReconDate DESC;"
'
'   szSQL = "SELECT StatementDate, BankCode " & _
'           "FROM tlbBankReconClosingBal " & _
'           "GROUP BY StatementDate, BankCode " & _
'           "ORDER BY StatementDate DESC;"

'Modified by anol 20181010 issue 642
    Dim irow As Integer
    Dim szBanks As String
    For irow = 1 To flxBankAccounts.Rows - 1
       If flxBankAccounts.TextMatrix(irow, 0) = "X" And flxBankAccounts.RowHeight(irow) > 0 Then
          szBanks = flxBankAccounts.TextMatrix(irow, 2)
       End If
    Next irow
   

    szSQL = "SELECT StatementDate, BankCode " & _
           "FROM tlbBankReconClosingBal " & _
           "WHERE BankCode = '" & szBanks & "' " & _
           "AND ClientID = '" & txtClientID.text & "' GROUP BY StatementDate, BankCode " & _
           "ORDER BY StatementDate DESC;"
           
'
'Debug.Print szSQL
   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If adoRst.RecordCount = 0 Then
      'ShowMsgInTaskBar "No bank reconciliation has been done yet", "Y", "N"
      Set adoRst = Nothing

     ' bProceed = False
      Exit Sub
   End If

   Dim TotalRow As Integer, TotalCol As Integer
   Dim i As Integer, j As Integer

   TotalRow = adoRst.RecordCount
   TotalCol = adoRst.Fields.count - 1

   ReDim ReConDates(TotalCol, TotalRow) As String

   For i = 0 To TotalRow - 1
      For j = 0 To TotalCol
         If Not IsNull(adoRst.Fields(j).Value) And adoRst.Fields(j).Value <> "" Then
            ReConDates(j, i) = adoRst.Fields(j).Value
         End If
      Next j
      adoRst.MoveNext
      If adoRst.EOF Then Exit For
   Next i

   cboCurStDt.Column() = ReConDates()
   cboLtStDt.Column() = ReConDates()

   adoRst.Close

   ' Error Handling Code
Error_Handler:

   ' Destroy Objects
   Set adoRst = Nothing
End Sub

Private Function IsBankSelected() As Boolean
   Dim irow          As Integer
   
   For irow = 1 To flxBankAccounts.Rows - 1
      If flxBankAccounts.TextMatrix(irow, 0) = "X" And flxBankAccounts.RowHeight(irow) > 0 Then
         szBanks = flxBankAccounts.TextMatrix(irow, 1) & ", " & szBanks
      End If
   Next irow
   If Len(szBanks) > 2 Then
      szBanks = Left(szBanks, Len(szBanks) - 2)
      IsBankSelected = True
   Else
      IsBankSelected = False
   End If
End Function

Private Function IsFundSelected() As Boolean
   Dim irow                As Integer
   
   For irow = 1 To flxFunds.Rows - 1
      If flxFunds.TextMatrix(irow, 0) = "X" And flxFunds.RowHeight(irow) > 0 Then
         szFunds = flxFunds.TextMatrix(irow, 1) & ", " & szFunds
         szFundList = flxFunds.TextMatrix(irow, 2) & ", " & szFundList
      End If
   Next irow
   If Len(szFunds) > 2 Then
      szFunds = Left(szFunds, Len(szFunds) - 2)
      szFundList = Left(szFundList, Len(szFundList) - 2)
      IsFundSelected = True
   Else
      IsFundSelected = False
      Exit Function
   End If
End Function

Private Sub MarkBankFund(adoConn As ADODB.Connection)
   

   adoConn.Execute "UPDATE tlbClientBanks " & _
                   "SET    SelBanks = '';"
   adoConn.Execute "UPDATE Fund " & _
                   "SET    SelFund = '', FundList = '';"

   adoConn.Execute "UPDATE tlbClientBanks " & _
                   "SET    SelBanks = 'Y' " & _
                   "WHERE  MY_ID IN (" & szBanks & ")"

   adoConn.Execute "UPDATE Fund " & _
                   "SET    SelFund = 'Y', FundList = '" & szFundList & "' " & _
                   "WHERE  FundID IN (" & szFunds & ")"

'   Currently the system calculates balance brought forward from all tables (sales, purchase and bank).
'   From now, system will get the BBF from the 'ProjClBal' of 'tlbBankReconClosingBal' on 'Last Statement Date'.
'   If the last statement date is empty then BBF will be 0.
'
'   CalculateBBF adoConn
'  ---------------------
  'issue 523
'Modified by anol 20 Jan 2015
   cBBF = StatementClosingBalance(adoConn, dtLastStDate, flxBankAccounts.TextMatrix(flxBankAccounts.row, 2), txtClientID.text)
   
   
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
          'PictureBoxZoom ctl, MouseKeys, Rotation, Xpos, Ypos
           bHandled = False

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
