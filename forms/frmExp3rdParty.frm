VERSION 5.00
Begin VB.Form frmExp3rdParty 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Export to Third Party Software"
   ClientHeight    =   7440
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8055
   BeginProperty Font 
      Name            =   "Myriad Web"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmExp3rdParty.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7440
   ScaleWidth      =   8055
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Caption         =   "Advance setting"
      Height          =   1815
      Left            =   120
      TabIndex        =   14
      Top             =   2400
      Width           =   6135
      Begin VB.TextBox txtCompany 
         Height          =   285
         Left            =   840
         TabIndex        =   19
         Top             =   720
         Width           =   5175
      End
      Begin VB.TextBox txtPath 
         Height          =   285
         Left            =   840
         TabIndex        =   17
         Top             =   360
         Width           =   5175
      End
      Begin VB.CheckBox chkdelreg1 
         Caption         =   "delete registry value if 2014 found  in sage path"
         Height          =   195
         Left            =   480
         TabIndex        =   16
         Top             =   1080
         Width           =   5175
      End
      Begin VB.CheckBox chkDelreg 
         Caption         =   "delete registry value if 2016 found  in sage path"
         Height          =   195
         Left            =   480
         TabIndex        =   15
         Top             =   1440
         Width           =   5415
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Company:"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   20
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Path:"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   18
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   255
      Left            =   240
      TabIndex        =   13
      Top             =   1200
      Width           =   735
   End
   Begin VB.Frame fraConfirmation 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Please Confirm:"
      Height          =   2415
      Left            =   0
      TabIndex        =   8
      Top             =   4320
      Width           =   5295
      Begin VB.CheckBox chkPreview 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "You have previewed your data - "
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   960
         Width           =   2655
      End
      Begin VB.CheckBox chkBackup 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "You have backed up your data - "
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   2655
      End
      Begin VB.Label lblClose 
         BackColor       =   &H00FFDEDE&
         BackStyle       =   0  'Transparent
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   150
         Left            =   4080
         TabIndex        =   9
         Top             =   120
         Width           =   120
      End
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "&Clear"
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CommandButton cmdLER 
      Caption         =   "Last Export"
      Height          =   375
      Left            =   1290
      TabIndex        =   2
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CommandButton cmdPreview 
      Caption         =   "&Preview"
      Height          =   375
      Left            =   2460
      TabIndex        =   3
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "C&lose"
      Height          =   375
      Left            =   4800
      TabIndex        =   5
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CommandButton cmdExport 
      Caption         =   "E&xport"
      Height          =   375
      Left            =   3630
      TabIndex        =   4
      Top             =   1560
      Width           =   1095
   End
   Begin VB.ComboBox cboTransactions 
      Height          =   315
      ItemData        =   "frmExp3rdParty.frx":0442
      Left            =   2040
      List            =   "frmExp3rdParty.frx":0452
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   840
      Width           =   3855
   End
   Begin VB.ComboBox cboSystem 
      Height          =   315
      ItemData        =   "frmExp3rdParty.frx":04BA
      Left            =   2040
      List            =   "frmExp3rdParty.frx":04BC
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   240
      Width           =   3855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Select System:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   12
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Transactions:"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   11
      Top             =   840
      Width           =   1215
   End
End
Attribute VB_Name = "frmExp3rdParty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'
''##############               RUNNING THIRD PARTY PROCESS AND CONTROL
'Private Type STARTUPINFO
'   cB As Long
'   lpReserved As String
'   lpDesktop As String
'   lpTitle As String
'   dwX As Long
'   dwY As Long
'   dwXSize As Long
'   dwYSize As Long
'   dwXCountChars As Long
'   dwYCountChars As Long
'   dwFillAttribute As Long
'   dwFlags As Long
'   wShowWindow As Integer
'   cbReserved2 As Integer
'   lpReserved2 As Long
'   hStdInput As Long
'   hStdOutput As Long
'   hStdError As Long
'End Type
'
'Private Type PROCESS_INFORMATION
'   hProcess As Long
'   hThread As Long
'   dwProcessID As Long
'   dwThreadID As Long
'End Type
'
'Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal _
'   hHandle As Long, ByVal dwMilliseconds As Long) As Long
'
'Private Declare Function CreateProcessA Lib "kernel32" (ByVal _
'   lpApplicationName As String, ByVal lpCommandLine As String, ByVal _
'   lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, _
'   ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, _
'   ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As String, _
'   lpStartupInfo As STARTUPINFO, lpProcessInformation As _
'   PROCESS_INFORMATION) As Long
'
'Private Declare Function CloseHandle Lib "kernel32" _
'   (ByVal hObject As Long) As Long
'
'Private Declare Function GetExitCodeProcess Lib "kernel32" _
'   (ByVal hProcess As Long, lpExitCode As Long) As Long
'
'Private Const NORMAL_PRIORITY_CLASS = &H20&
'Private Const INFINITE = -1&
'Private iTotalTran As Integer
''-------------------------------------------------------------------------------
'
'Private Sub cboSystem_Click()
'   If Left(cboSystem.text, 1) = "-" Then
'      MsgBox "This export module has not been subscribed." & Chr(13) & "Please contact with PCM.", vbInformation + vbOKOnly, "Export to " & Right(cboSystem.text, Len(cboSystem.text) - 1)
'   End If
'End Sub
'
'Private Function GetDataSource(adoconn As ADODB.Connection) As String
'   Dim adoRst As New ADODB.Recordset
'   Dim szSQL As String
'
'   szSQL = "SELECT * FROM ShoppingCentre;"
'   adoRst.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
'
'   If adoRst.RecordCount = 0 Then
'      GetDataSource = ""
'   Else
'      GetDataSource = adoRst!SageDSN
'   End If
'
'   adoRst.Close
'   Set adoRst = Nothing
'End Function
'
'Private Sub chkBackup_Click()
'   If chkBackup.Value = 1 And chkPreview.Value = 1 Then
'      If MsgBox("Do you wish to export these transactions to " & cboSystem.text & "?", vbQuestion + vbYesNo, "Export to 3rd Party") = vbYes Then
'         If cboSystem.text = "SAGE MMS" And cboTransactions.text = "Sales Invoices & Credit Notes" Then
'            ExportTransactionsSageMMS
'         End If
'         If cboSystem.text = "SAGE 200" And cboTransactions.text = "Sales Invoices & Credit Notes" Then
'            ExportTrans2Sage200
'         End If
'         If cboSystem.text = "SAGE Line 50 v14" And cboTransactions.text = "Sales Invoices & Credit Notes" Then
'            ExportTransIC2Sage50v14 'Export SI and SC to SAGE Line 50 v14
'         End If
'      Else
'         chkBackup.Value = 0
'         chkPreview.Value = 0
'      End If
'   End If
'End Sub
'
'Private Sub chkDelreg_Click()
'    On Error Resume Next
'    If chkDelreg.Value = 1 Then
'            If InStr(sageDirPath, "2016") > 0 Then
'                 DeleteSetting "PropertyManagement", "SageCompany"
'                 DeleteSetting "PropertyManagement", "SagePath"
'                 txtPath.text = GetSetting("PropertyManagement", "SagePath", "SageFolder")
'            End If
'    End If
'    If chkdelreg1.Value = 1 Then
'            If InStr(sageDirPath, "2014") > 0 Then
'                 DeleteSetting "PropertyManagement", "SageCompany"
'                 DeleteSetting "PropertyManagement", "SagePath"
'                 txtPath.text = GetSetting("PropertyManagement", "SagePath", "SageFolder")
'            End If
'    End If
''Err:
'
'End Sub
'
'Private Sub chkPreview_Click()
'   If chkBackup.Value = 1 And chkPreview.Value = 1 Then
'      If MsgBox("Do you wish to export these transactions to " & cboSystem.text & "?", vbQuestion + vbYesNo, "Export to 3rd Party") = vbYes Then
'         If cboSystem.text = "SAGE MMS" And cboTransactions.text = "Sales Invoices & Credit Notes" Then
'            ExportTransactionsSageMMS
'         End If
'         If cboSystem.text = "SAGE 200" And cboTransactions.text = "Sales Invoices & Credit Notes" Then
'            ExportTrans2Sage200
'         End If
'         If cboSystem.text = "SAGE Line 50 v14" And cboTransactions.text = "Sales Invoices & Credit Notes" Then
'            ExportTransIC2Sage50v14 'Export SI and SC to SAGE Line 50 v14
'         End If
'         If cboSystem.text = "SAGE Line 50 v16" And cboTransactions.text = "Sales Invoices & Credit Notes" Then
'            ExportTransIC2Sage50v16 'Export SI and SC to SAGE Line 50 v16
'         End If
'         If cboSystem.text = "SAGE Line 50 v14" And cboTransactions.text = "Sales Receipts" Then
'            ExportSR2Sage50v14 'Export SR to SAGE Line 50 v14
'         End If
'         If cboSystem.text = "SAGE Line 50 v16" And cboTransactions.text = "Sales Receipts" Then
'            ExportSR2Sage50v16 'Export SR to SAGE Line 50 v16
'         End If
'      Else
'         chkBackup.Value = 0
'         chkPreview.Value = 0
'      End If
'   End If
'End Sub
'
'Private Sub ExportSR2Sage50v16()
'   Const sageUserName   As String = "Prestige"
'   Const sagePassword   As String = "prestige"
'   Dim iCount        As Integer
'   Dim szSQLStr      As String
'   Dim i             As Integer
'   Dim iSplit        As Integer
'   Dim j             As Integer
'   Dim laTransID()   As Long
'    Dim Hkey             As Long
''   Dim szaTenant()   As String
'
'   On Error GoTo Error_Handler
'
''   Declare Variables for Database connectivity
'   Dim adoconn       As New ADODB.Connection
'   Dim adoRstRpt     As New ADODB.Recordset
'   Dim adoRstRptCh   As New ADODB.Recordset
'   Dim adoPoA        As New ADODB.Recordset
'
''  Set the connection to the databases
'   adoconn.Open getConnectionString
''   szaTenant = Split(txtTenantID.text, " \ ")
'
''  Receipt needs to export
'   szSQLStr = "SELECT RPT.TransactionID AS RPT_ID, RT.TransactionID AS RT_ID, RPT.SageAccountNumber, " & _
'                  "RT.AllocDate, RPT.TYPE, RT.BankCode, RT.ReceiptAmount, RPT.ExtRef " & _
'              "FROM RptTransactions AS RT INNER JOIN tlbReceipt AS RPT ON RT.FromTran = RPT.TransactionID " & _
'              "WHERE RT.IsSageUpdate And NOT RT.UpdateSage;"
''Debug.Print szSQLStr
'   adoRstRpt.Open szSQLStr, adoconn, adOpenStatic, adLockReadOnly
'
''  PoA needs to export
'   szSQLStr = "SELECT RPT.* " & _
'              "FROM tlbReceipt AS RPT " & _
'              "WHERE IsSageUpdate And Type = 4 And NOT UpDateSage;"
'
'   adoPoA.Open szSQLStr, adoconn, adOpenStatic, adLockReadOnly
'
'   If adoRstRpt.EOF And adoPoA.EOF Then
'      adoRstRpt.Close
'      adoPoA.Close
'      adoconn.Close
'      Set adoRstRpt = Nothing
'      Set adoPoA = Nothing
'      Set adoconn = Nothing
'      MsgBox "No record found to be posted", vbInformation, "Information"
'      Exit Sub
'   End If
'
'   ' Declare Objects
'   'Dim oSDO                As New SageDataObject220.SDOEngine     ' Create the SDO Engine Object
'   'Dim oWS                 As SageDataObject220.WorkSpace
''   Dim oTransactionPost    As SageDataObject220.TransactionPost
''   Dim oHeaderData         As SageDataObject220.HeaderData
''   Dim oSplitData          As SageDataObject220.SplitData
'
'   ' Declare Variables
'   Dim szDataPath          As String
'   Dim iCtr                As Integer
'   Dim bFlag               As Boolean
'   Dim bBreak              As Boolean
'   Dim lInvoiceHeader      As Long
'   Dim lReceiptHeader      As Long
'   Dim lReceiptSplit       As Long
'   Dim lSAmount            As Double
'   Dim lAmountLeft         As Double
'
''   Create the Workspace
''   Set oWS = oSDO.Workspaces.Add("Prestige")
'
''   Select Company.  The SelectCompany method takes the program install folder as a parameter
'   szDataPath = CompanyDatapath
''   If szDataPath = "C:\ProgramData\Sage\Accounts\2014" Then
''        MsgBox "Sage dll 2014 is now oboslate.Please select newer version.", vbInformation, "Warning"
''        Exit Sub
''   End If
''if no path is set ask for select the company by anol 2019-06-20
'    sageDirPath = GetSetting("PropertyManagement", "SagePath", "SageFolder")
'      If sageDirPath = "" Then
'         sageDirPath = BrowseForFolder(Hkey, "Please select Sage path...")
'         If sageDirPath = "" Then
'             End
'         End If
'         SaveSetting "PropertyManagement", "SagePath", "SageFolder", sageDirPath
'      End If
''now ask
'       szDataPath = oSDO.SelectCompany(sageDirPath)
'       CompanyDatapath = sageDirPath
''     Save company name in the registry
'      SaveSetting "PropertyManagement", "SageCompany", CompanyDatapath, szDataPath
'
'
''   A U.I. for company selection is presented to the user. If a company is selected,
''   the path will be passed to the szDataPath variable.
''   If not, or the Cancel button is selected, the variable will be left empty.
'   If szDataPath <> "" Then
''     Try to Connect - Will throw an exception if it fails
'      If Not oWS.Connect(szDataPath, sageUserName, sagePassword, "Prestige") Then
'         MsgBox "The Prestige has failed to create a connection with SAGE. Please contact with PCM Consulting.", vbCritical + vbOKOnly, "Connection with SAGE Failed"
'
'         adoRstRpt.Close
'         adoPoA.Close
'         adoconn.Close
'
''        Destroy the Objects
'         Set oWS = Nothing
'         Set oSDO = Nothing
'
'         Set adoRstRpt = Nothing
'         Set adoPoA = Nothing
'         Set adoconn = Nothing
'
'         Exit Sub
'      End If
'   Else
'      MsgBox "There are some problems with the SAGE configuration in the system registry. Please contact with PCM Consulting.", vbCritical + vbOKOnly, "Registry Error"
'
'      adoRstRpt.Close
'      adoconn.Close
'
''     Destroy the Objects
'      Set oWS = Nothing
'      Set oSDO = Nothing
'
'      Set adoRstRpt = Nothing
'      Set adoconn = Nothing
'      Exit Sub
'   End If
'
'   If adoRstRpt.EOF Then
'      adoRstRpt.Close
'      Set adoRstRpt = Nothing
'      GoTo PoA
'   End If
'
'   JustifyOSAmtOfRptToExport adoRstRpt, adoconn, oWS
'   adoRstRpt.Close
'   ReDim Preserve laTransID(iTotalTran) As Long
'
''   Main code segment to export to SAGE
''   szSQLStr = "SELECT DISTINCT R.Type, R.SageAccountNumber AS SAN, R.RDate, R.BankCode, " & _
''                  "R.SageDepartment, R.Amount, R.TransactionID, DemandRef, ExtRef " & _
''              "FROM tlbReceipt AS R, RptTransactions AS T " & _
''              "WHERE T.FromTran = R.TransactionID AND " & _
''                  "R.SageAccountNumber ='" & szaTenant(0) & "' AND " & _
''                  "T.IsSageUpdate AND NOT T.UpdateSage;"
'   szSQLStr = "SELECT DISTINCT R.Type, R.SageAccountNumber AS SAN, R.RDate, R.BankCode, " & _
'                  "R.SageDepartment, R.Amount, R.TransactionID, DemandRef, ExtRef, R.PostingDate " & _
'              "FROM tlbReceipt AS R, RptTransactions AS T " & _
'              "WHERE T.FromTran = R.TransactionID AND " & _
'                    "T.IsSageUpdate AND NOT T.UpdateSage;"
''Debug.Print szSQLStr
'   adoRstRpt.Open szSQLStr, adoconn, adOpenStatic, adLockReadOnly
'
'   iCount = 0
'   While Not adoRstRpt.EOF
'      If Val(adoRstRpt.Fields.Item("Type").Value) = CByte(sdoSR) Then
''         Create an instance of TransactionPost for the Sales Receipt                  #SR#
'         Set oTransactionPost = oWS.CreateObject("TransactionPost")
'
''         Fill in the Header fields
'         oTransactionPost.Header("ACCOUNT_REF").Value = CStr(adoRstRpt!SAN)
'         oTransactionPost.Header("DATE").Value = CDate(adoRstRpt!RDate)
'         oTransactionPost.Header("POSTED_DATE").Value = CDate(adoRstRpt!postingDate)
'         oTransactionPost.Header("TYPE").Value = CByte(sdoSR)                    'Sales Receipt -> 3
'         oTransactionPost.Header("DETAILS").Value = "Sales Receipt" & " - " & adoRstRpt!ExtRef
'         oTransactionPost.Header("BANK_CODE").Value = CStr(adoRstRpt!BankCode)
'         oTransactionPost.Header("Inv_Ref").Value = CStr(adoRstRpt!TransactionID & "/" & adoRstRpt!DemandRef)         'Receipt could be allocated against more than one inv
'
''         Create a split item by adding an empty split to the Items
'         Set oSplitData = oTransactionPost.Items.Add()
'
''         Fill in the Split fields - note a Sales Receipt only has one split
'         oSplitData.Fields.Item("TYPE").Value = CByte(sdoSR)                   'Sales Receipt -> 3
'         oSplitData.Fields.Item("DEPT_NUMBER").Value = CStr(adoRstRpt!SageDepartment)
'         oSplitData.Fields.Item("NOMINAL_CODE").Value = CStr(adoRstRpt!BankCode)
'         oSplitData.Fields.Item("TAX_CODE").Value = CInt(9)
'         oSplitData.Fields.Item("NET_AMOUNT").Value = CDbl(adoRstRpt!amount)
'         oSplitData.Fields.Item("TAX_AMOUNT").Value = CDbl(0)
'         oSplitData.Fields.Item("DATE").Value = CDate(oTransactionPost.Header("DATE").Value)
'         oSplitData.Fields.Item("DETAILS").Value = CStr(oTransactionPost.Header("DETAILS").Value)
'
''         Update the TransactionPost Object
'         If oTransactionPost.Update Then
'            lReceiptHeader = oTransactionPost.PostingNumber
'
'            Set oHeaderData = oWS.CreateObject("HeaderData")
'            oHeaderData.Read (lReceiptHeader)
'            lReceiptSplit = oHeaderData.Fields.Item("FIRST_SPLIT").Value
'
'            Set oTransactionPost = Nothing
'            Set oSplitData = Nothing
'
'            szSQLStr = "SELECT  DR.SAGEPostingNumber AS SPN, " & _
'                           "RT.AllocDate AS RDate, RT.ReceiptAmount, RT.TransactionID, " & _
'                           "SUM(DSR.TotalAmount) AS TA, Rpt.Type " & _
'                       "FROM RptTransactions AS RT, tlbReceipt as Rpt, DemandRecords as DR, " & _
'                           "DemandSplitRecords AS DSR " & _
'                       "WHERE RT.FromTran = " & adoRstRpt!TransactionID & " AND " & _
'                           "Rpt.TransactionID = RT.ToTran AND Rpt.DemandRef = DR.DemandID AND " & _
'                           "DR.DemandID = DSR.DemandID AND NOT RT.UpdateSage " & _
'                       "GROUP BY DR.SAGEPostingNumber, RT.AllocDate, " & _
'                           "RT.ReceiptAmount, RPT.TransactionID,  RT.TransactionID, Rpt.Type;"
''   Debug.Print szSQLStr
'            adoRstRptCh.Open szSQLStr, adoconn, adOpenStatic, adLockReadOnly
'
'            While Not adoRstRptCh.EOF
'               If adoRstRptCh.Fields.Item("Type").Value <> CByte(sdoSA) Then
'                  lInvoiceHeader = CLng(adoRstRptCh!SPN)             'SPN -> SAGE Posting Number
'                  oHeaderData.Read (lInvoiceHeader)
'
'                  Set oTransactionPost = oWS.CreateObject("TransactionPost")
'
'                  If Round(Val(oHeaderData.Fields.Item("NET_AMOUNT").Value) - _
'                     Val(oHeaderData.Fields.Item("AMOUNT_PAID").Value), 2) >= Val(adoRstRptCh!receiptAmount) Then
'
'                     Set oSplitData = oHeaderData.Link
'
'                     oSplitData.MoveFirst
'
'                     lAmountLeft = Val(adoRstRptCh!receiptAmount)
'
'                     For j = 1 To oSplitData.Count
'                        lSAmount = Val(oSplitData.Fields.Item("Net_Amount").Value) + _
'                           Val(oSplitData.Fields.Item("Tax_Amount").Value) - _
'                           Val(oSplitData.Fields.Item("AMOUNT_PAID").Value)
'
'                        If Round(lSAmount, 2) <= lAmountLeft Then
'                           lAmountLeft = lAmountLeft - lSAmount
'
'                           If oTransactionPost.AllocatePayment(CLng(oSplitData.RecordNumber), CLng(lReceiptSplit), _
'                                 lSAmount, CDate(adoRstRptCh!RDate)) Then
'                              laTransID(iCount) = adoRstRptCh!TransactionID
'                              iCount = iCount + 1
'                           End If
'                        Else
'                           If oTransactionPost.AllocatePayment(CLng(oSplitData.RecordNumber), CLng(lReceiptSplit), _
'                                 lAmountLeft, CDate(adoRstRptCh!RDate)) Then
'                              laTransID(iCount) = adoRstRptCh!TransactionID
'                              iCount = iCount + 1
'                           End If
'                           lAmountLeft = 0
'                        End If
'                        oSplitData.MoveNext
'                     Next j
'                  Else
'                     MsgBox "Prestige data does not match with SAGE data. Please contact with PCM Consulting Ltd.", vbCritical + vbOKOnly, "Prestige <> SAGE"
'                     UpdateMarked laTransID, iCount, adoconn, "RptTransactions"
'                     GoTo CloseConnection
'                  End If
'               Else                       'PAYMENT ON ACCOUNT - DOES NOT ALLOCATE AGAINST ANY TRANSACTION
'                  laTransID(iCount) = adoRstRptCh!TransactionID
'                  iCount = iCount + 1
'               End If
'               adoRstRptCh.MoveNext
'            Wend
'            adoRstRptCh.Close
'         End If
'      End If
'      adoRstRpt.MoveNext
'   Wend
'   adoRstRpt.Close
'   UpdateMarked laTransID, iCount, adoconn, "RptTransactions"
''MsgBox "BEFORE ALLOCATION OF CR AGAINST INV"
'' ***********************************************************************************************************
'' *********************** ALLOCATION OF CREDIT AGAINST INVOICE **********************************************
'' ***********************************************************************************************************
''  Exporting Allocations of 'Credit Invoice' against Invoice
'   Dim lSPN As Long, lSPN_Inv As Long, bAllocated As Boolean
'   Dim oCreditHeaderData As SageDataObject220.HeaderData, oCreditSplitData As SageDataObject220.SplitData
'   Dim oInvHeaderData As SageDataObject220.HeaderData, oInvSplitData As SageDataObject220.SplitData
'
'   szSQLStr = "SELECT RT.TransactionID, RT.FromTran, RT.ToTran, RT.AllocDate AS ADate, " & _
'                  "RT.ReceiptAmount AS RAmt, RT.BankCode, RT.NominalCode, Rpt.Type, " & _
'                  "Rpt.PoA_SPN, Rpt.DemandRef " & _
'              "FROM RptTransactions AS RT, tlbReceipt AS Rpt " & _
'              "WHERE RT.IsSageUpdate AND NOT RT.UpdateSage And " & _
'                  "(isnull(RT.Exp2Sage) OR LEFT(RT.Exp2Sage,1) = 'S') And " & _
'                  "RT.FromTran = Rpt.TransactionID;"
''Debug.Print szSQLStr
'   adoRstRpt.Open szSQLStr, adoconn, adOpenStatic, adLockReadOnly
'
'   Set oCreditHeaderData = oWS.CreateObject("HeaderData")
'   Set oInvHeaderData = oWS.CreateObject("HeaderData")
'
'   iCount = 0
'   While Not adoRstRpt.EOF
'      If adoRstRpt.Fields.Item("Type").Value = 4 Then                'PoA
''MsgBox adoRstRpt.Fields.Item("PoA_SPN").Value
'         lSPN = adoRstRpt.Fields.Item("PoA_SPN").Value
'      End If
'      If adoRstRpt.Fields.Item("Type").Value = 2 Then                'Cr Note
'         lSPN = SagePostingNumberCr(adoRstRpt.Fields.Item("DemandRef").Value, adoconn)
'      End If
'      oCreditHeaderData.Read (lSPN)
'
'      lSPN_Inv = SagePostingNumberInv(adoRstRpt.Fields.Item("ToTran").Value, adoconn)
'      oInvHeaderData.Read (lSPN_Inv)
'
'      'Link to the Credit Splits
'      Set oCreditSplitData = oCreditHeaderData.Link
'      oCreditSplitData.MoveFirst
'
'      Set oInvSplitData = oInvHeaderData.Link
'      oInvSplitData.MoveFirst
'
'      For i = 1 To oInvSplitData.Count
'         If Round(Val(oInvSplitData.Fields.Item("NET_AMOUNT").Value) - _
'               Val(oInvSplitData.Fields.Item("AMOUNT_PAID").Value), 2) >= _
'               Val(adoRstRpt.Fields.Item("RAmt").Value) Then
'            Exit For
'         End If
'         oInvSplitData.MoveNext
'      Next i
'      If i <= oInvSplitData.Count Then
'         Set oTransactionPost = oWS.CreateObject("TransactionPost")
'         bAllocated = oTransactionPost.AllocatePayment(CInt(oInvSplitData.RecordNumber), _
'                        CInt(oCreditSplitData.RecordNumber), CDbl(adoRstRpt.Fields.Item("RAmt").Value), _
'                        CDate(adoRstRpt.Fields.Item("ADate").Value))
'         If bAllocated Then
'            laTransID(iCount) = adoRstRpt.Fields.Item("TransactionID").Value
'            iCount = iCount + 1
'         End If
'      End If
'      Set oCreditSplitData = Nothing
'      Set oInvSplitData = Nothing
'
'      adoRstRpt.MoveNext
'   Wend
'
''   MsgBox iCount & " Transactions (out of " & iTotalTran & ") Posted to SAGE successfully", vbOKOnly, "Posted to SAGE"
'   If iTotalTran - iCount > 0 Then
''    i have stopped this report on basis of bug report of 23rd April, 2007
''     Report of posting exceptions to be printed
''      ShowReport App.Path & szReportPath & "\ReceiptExceptionReport.rpt"
'   End If
'   UpdateMarked laTransID, iCount, adoconn, "RptTransactions"
'
'   If adoPoA.EOF Then
''      MsgBox "There is no Payment on Account to update to SAGE.", vbInformation + vbOKOnly, "Payment on Account "
'      adoPoA.Close
'      Set adoPoA = Nothing
'      GoTo CloseConnection
'   End If
'
'PoA:
''  Update Payment On Account
'   iCount = 0
'   ReDim laTransID(adoPoA.RecordCount) As Long              'Refreshing the array to save all the PoA transaction id
'
'   While Not adoPoA.EOF
''      Create an instance of TransactionPost for the Sales Receipt
'      Set oTransactionPost = oWS.CreateObject("TransactionPost")
'
''      Fill in the Header fields
'      oTransactionPost.Header("ACCOUNT_REF").Value = CStr(adoPoA!SageAccountNumber)
'      oTransactionPost.Header("DATE").Value = CDate(adoPoA!RDate)
'      oTransactionPost.Header("POSTED_DATE").Value = CDate(Date)
'      oTransactionPost.Header("TYPE").Value = CByte(sdoSA)
'      oTransactionPost.Header("DETAILS").Value = "Payment on Account"
'      oTransactionPost.Header("BANK_CODE").Value = CStr(adoPoA!BankCode)
'      oTransactionPost.Header("INV_REF").Value = adoPoA!TransactionID & " - " & IIf(IsNull(adoPoA!ExtRef), "", adoPoA!ExtRef)
'
''      Create a split item by adding an empty split to the Items
''      collection of the TransactionPost Object.
'      Set oSplitData = oTransactionPost.Items.Add()
'
''      Fill in the Split fields - note a Sales Receipt only has one split
'      oSplitData.Fields.Item("TYPE").Value = CByte(sdoSA)
'      oSplitData.Fields.Item("DEPT_NUMBER").Value = CStr(PropDeptNumTent(adoPoA!SageAccountNumber, adoconn))
'      oSplitData.Fields.Item("NOMINAL_CODE").Value = CStr(adoPoA!BankCode)
'      oSplitData.Fields.Item("TAX_CODE").Value = CInt(9)
'      oSplitData.Fields.Item("NET_AMOUNT").Value = CDbl(adoPoA!amount)
'      oSplitData.Fields.Item("TAX_AMOUNT").Value = CDbl(0)
'      oSplitData.Fields.Item("DATE").Value = CDate(oTransactionPost.Header("DATE").Value)
'      oSplitData.Fields.Item("DETAILS").Value = CStr(oTransactionPost.Header("DETAILS").Value)
'      oSplitData.Fields.Item("INTERNAL_REF").Value = CStr(adoPoA.Fields("Ref").Value)
'
'      If oTransactionPost.Update Then
'         SavePostingNumberPoA "tlbReceipt", "TransactionID", CLng(adoPoA.Fields("TransactionID").Value), oTransactionPost.PostingNumber, adoconn
'         laTransID(iCount) = adoPoA!TransactionID
'         iCount = iCount + 1
'      Else
'         MsgBox "Payment on Account has not been exported to SAGE.", vbCritical + vbOKOnly, "PoA - Exported"
'      End If
'
'      adoPoA.MoveNext
'   Wend
'
'   UpdateMarked laTransID, iCount, adoconn, "tlbReceipt"
'
'CloseConnection:
'
'   oWS.Disconnect
'
''   Destroy the Objects
'   Set oCreditHeaderData = Nothing
'   Set oInvHeaderData = Nothing
'   Set oTransactionPost = Nothing
'   Set oSplitData = Nothing
'   Set oHeaderData = Nothing
'   Set oWS = Nothing
'   Set oSDO = Nothing
'
'   Set adoRstRptCh = Nothing
'   Set adoRstRpt = Nothing
'   Set adoconn = Nothing
'
'   Exit Sub
'
'' Error Handling Code
'Error_Handler:
'
'   MsgBox "The SDO generated the following error: " & oSDO.LastError.text & ERR.Number & " -(pcm_SR_Posting) " & ERR.description, vbOKOnly, "Posted to SAGE"
'
'   Set oCreditHeaderData = Nothing
'   Set oInvHeaderData = Nothing
'   Set oTransactionPost = Nothing
'   Set oSplitData = Nothing
'   Set oHeaderData = Nothing
'   Set oWS = Nothing
'   Set oSDO = Nothing
'   If adoRstRpt.State = 1 Then
'        adoRstRpt.Close
'   End If
'   If adoconn.State = 1 Then
'        adoconn.Close
'   End If
'   Set adoRstRpt = Nothing
'   Set adoconn = Nothing
'End Sub
'Private Sub ExportSR2Sage50v14()
'   Const sageUserName   As String = "Prestige"
'   Const sagePassword   As String = "prestige"
'   Dim iCount        As Integer
'   Dim szSQLStr      As String
'   Dim i             As Integer
'   Dim iSplit        As Integer
'   Dim j             As Integer
'   Dim laTransID()   As Long
''   Dim szaTenant()   As String
'
'   On Error GoTo Error_Handler
'
''   Declare Variables for Database connectivity
'   Dim adoconn       As New ADODB.Connection
'   Dim adoRstRpt     As New ADODB.Recordset
'   Dim adoRstRptCh   As New ADODB.Recordset
'   Dim adoPoA        As New ADODB.Recordset
'
''  Set the connection to the databases
'   adoconn.Open getConnectionString
''   szaTenant = Split(txtTenantID.text, " \ ")
'
''  Receipt needs to export
'   szSQLStr = "SELECT RPT.TransactionID AS RPT_ID, RT.TransactionID AS RT_ID, RPT.SageAccountNumber, " & _
'                  "RT.AllocDate, RPT.TYPE, RT.BankCode, RT.ReceiptAmount, RPT.ExtRef " & _
'              "FROM RptTransactions AS RT INNER JOIN tlbReceipt AS RPT ON RT.FromTran = RPT.TransactionID " & _
'              "WHERE RT.IsSageUpdate And NOT RT.UpdateSage;"
''Debug.Print szSQLStr
'   adoRstRpt.Open szSQLStr, adoconn, adOpenStatic, adLockReadOnly
'
''  PoA needs to export
'   szSQLStr = "SELECT RPT.* " & _
'              "FROM tlbReceipt AS RPT " & _
'              "WHERE IsSageUpdate And Type = 4 And NOT UpDateSage;"
'
'   adoPoA.Open szSQLStr, adoconn, adOpenStatic, adLockReadOnly
'
'   If adoRstRpt.EOF And adoPoA.EOF Then
'      adoRstRpt.Close
'      adoPoA.Close
'      adoconn.Close
'      Set adoRstRpt = Nothing
'      Set adoPoA = Nothing
'      Set adoconn = Nothing
'      Exit Sub
'   End If
'
'   ' Declare Objects
'   Dim oSDO                As New SageDataObject220.SDOEngine     ' Create the SDO Engine Object
'   Dim oWS                 As SageDataObject220.WorkSpace
'   Dim oTransactionPost    As SageDataObject220.TransactionPost
'   Dim oHeaderData         As SageDataObject220.HeaderData
'   Dim oSplitData          As SageDataObject220.SplitData
'
'   ' Declare Variables
'   Dim szDataPath          As String
'   Dim iCtr                As Integer
'   Dim bFlag               As Boolean
'   Dim bBreak              As Boolean
'   Dim lInvoiceHeader      As Long
'   Dim lReceiptHeader      As Long
'   Dim lReceiptSplit       As Long
'   Dim lSAmount            As Double
'   Dim lAmountLeft         As Double
'
''   Create the Workspace
'   Set oWS = oSDO.Workspaces.Add("Prestige")
'
''   Select Company.  The SelectCompany method takes the program install folder as a parameter
'   szDataPath = CompanyDatapath
'
''   A U.I. for company selection is presented to the user. If a company is selected,
''   the path will be passed to the szDataPath variable.
''   If not, or the Cancel button is selected, the variable will be left empty.
'   If szDataPath <> "" Then
''     Try to Connect - Will throw an exception if it fails
'      If Not oWS.Connect(szDataPath, sageUserName, sagePassword, "Prestige") Then
'         MsgBox "The Prestige has failed to create a connection with SAGE. Please contact with PCM Consulting.", vbCritical + vbOKOnly, "Connection with SAGE Failed"
'
'         adoRstRpt.Close
'         adoPoA.Close
'         adoconn.Close
'
''        Destroy the Objects
'         Set oWS = Nothing
'         Set oSDO = Nothing
'
'         Set adoRstRpt = Nothing
'         Set adoPoA = Nothing
'         Set adoconn = Nothing
'
'         Exit Sub
'      End If
'   Else
'      MsgBox "There are some problems with the SAGE configuration in the system registry. Please contact with PCM Consulting.", vbCritical + vbOKOnly, "Registry Error"
'
'      adoRstRpt.Close
'      adoconn.Close
'
''     Destroy the Objects
'      Set oWS = Nothing
'      Set oSDO = Nothing
'
'      Set adoRstRpt = Nothing
'      Set adoconn = Nothing
'      Exit Sub
'   End If
'
'   If adoRstRpt.EOF Then
'      adoRstRpt.Close
'      Set adoRstRpt = Nothing
'      GoTo PoA
'   End If
'
'   JustifyOSAmtOfRptToExport adoRstRpt, adoconn, oWS
'   adoRstRpt.Close
'   ReDim Preserve laTransID(iTotalTran) As Long
'
''   Main code segment to export to SAGE
''   szSQLStr = "SELECT DISTINCT R.Type, R.SageAccountNumber AS SAN, R.RDate, R.BankCode, " & _
''                  "R.SageDepartment, R.Amount, R.TransactionID, DemandRef, ExtRef " & _
''              "FROM tlbReceipt AS R, RptTransactions AS T " & _
''              "WHERE T.FromTran = R.TransactionID AND " & _
''                  "R.SageAccountNumber ='" & szaTenant(0) & "' AND " & _
''                  "T.IsSageUpdate AND NOT T.UpdateSage;"
'   szSQLStr = "SELECT DISTINCT R.Type, R.SageAccountNumber AS SAN, R.RDate, R.BankCode, " & _
'                  "R.SageDepartment, R.Amount, R.TransactionID, DemandRef, ExtRef, R.PostingDate " & _
'              "FROM tlbReceipt AS R, RptTransactions AS T " & _
'              "WHERE T.FromTran = R.TransactionID AND " & _
'                    "T.IsSageUpdate AND NOT T.UpdateSage;"
''Debug.Print szSQLStr
'   adoRstRpt.Open szSQLStr, adoconn, adOpenStatic, adLockReadOnly
'
'   iCount = 0
'   While Not adoRstRpt.EOF
'      If Val(adoRstRpt.Fields.Item("Type").Value) = CByte(sdoSR) Then
''         Create an instance of TransactionPost for the Sales Receipt                  #SR#
'         Set oTransactionPost = oWS.CreateObject("TransactionPost")
'
''         Fill in the Header fields
'         oTransactionPost.Header("ACCOUNT_REF").Value = CStr(adoRstRpt!SAN)
'         oTransactionPost.Header("DATE").Value = CDate(adoRstRpt!RDate)
'         oTransactionPost.Header("POSTED_DATE").Value = CDate(adoRstRpt!postingDate)
'         oTransactionPost.Header("TYPE").Value = CByte(sdoSR)                    'Sales Receipt -> 3
'         oTransactionPost.Header("DETAILS").Value = "Sales Receipt" & " - " & adoRstRpt!ExtRef
'         oTransactionPost.Header("BANK_CODE").Value = CStr(adoRstRpt!BankCode)
'         oTransactionPost.Header("Inv_Ref").Value = CStr(adoRstRpt!TransactionID & "/" & adoRstRpt!DemandRef)         'Receipt could be allocated against more than one inv
'
''         Create a split item by adding an empty split to the Items
'         Set oSplitData = oTransactionPost.Items.Add()
'
''         Fill in the Split fields - note a Sales Receipt only has one split
'         oSplitData.Fields.Item("TYPE").Value = CByte(sdoSR)                   'Sales Receipt -> 3
'         oSplitData.Fields.Item("DEPT_NUMBER").Value = CStr(adoRstRpt!SageDepartment)
'         oSplitData.Fields.Item("NOMINAL_CODE").Value = CStr(adoRstRpt!BankCode)
'         oSplitData.Fields.Item("TAX_CODE").Value = CInt(9)
'         oSplitData.Fields.Item("NET_AMOUNT").Value = CDbl(adoRstRpt!amount)
'         oSplitData.Fields.Item("TAX_AMOUNT").Value = CDbl(0)
'         oSplitData.Fields.Item("DATE").Value = CDate(oTransactionPost.Header("DATE").Value)
'         oSplitData.Fields.Item("DETAILS").Value = CStr(oTransactionPost.Header("DETAILS").Value)
'
''         Update the TransactionPost Object
'         If oTransactionPost.Update Then
'            lReceiptHeader = oTransactionPost.PostingNumber
'
'            Set oHeaderData = oWS.CreateObject("HeaderData")
'            oHeaderData.Read (lReceiptHeader)
'            lReceiptSplit = oHeaderData.Fields.Item("FIRST_SPLIT").Value
'
'            Set oTransactionPost = Nothing
'            Set oSplitData = Nothing
'
'            szSQLStr = "SELECT  DR.SAGEPostingNumber AS SPN, " & _
'                           "RT.AllocDate AS RDate, RT.ReceiptAmount, RT.TransactionID, " & _
'                           "SUM(DSR.TotalAmount) AS TA, Rpt.Type " & _
'                       "FROM RptTransactions AS RT, tlbReceipt as Rpt, DemandRecords as DR, " & _
'                           "DemandSplitRecords AS DSR " & _
'                       "WHERE RT.FromTran = " & adoRstRpt!TransactionID & " AND " & _
'                           "Rpt.TransactionID = RT.ToTran AND Rpt.DemandRef = DR.DemandID AND " & _
'                           "DR.DemandID = DSR.DemandID AND NOT RT.UpdateSage " & _
'                       "GROUP BY DR.SAGEPostingNumber, RT.AllocDate, " & _
'                           "RT.ReceiptAmount, RPT.TransactionID,  RT.TransactionID, Rpt.Type;"
''   Debug.Print szSQLStr
'            adoRstRptCh.Open szSQLStr, adoconn, adOpenStatic, adLockReadOnly
'
'            While Not adoRstRptCh.EOF
'               If adoRstRptCh.Fields.Item("Type").Value <> CByte(sdoSA) Then
'                  lInvoiceHeader = CLng(adoRstRptCh!SPN)             'SPN -> SAGE Posting Number
'                  oHeaderData.Read (lInvoiceHeader)
'
'                  Set oTransactionPost = oWS.CreateObject("TransactionPost")
'
'                  If Round(Val(oHeaderData.Fields.Item("NET_AMOUNT").Value) - _
'                     Val(oHeaderData.Fields.Item("AMOUNT_PAID").Value), 2) >= Val(adoRstRptCh!receiptAmount) Then
'
'                     Set oSplitData = oHeaderData.Link
'
'                     oSplitData.MoveFirst
'
'                     lAmountLeft = Val(adoRstRptCh!receiptAmount)
'
'                     For j = 1 To oSplitData.Count
'                        lSAmount = Val(oSplitData.Fields.Item("Net_Amount").Value) + _
'                           Val(oSplitData.Fields.Item("Tax_Amount").Value) - _
'                           Val(oSplitData.Fields.Item("AMOUNT_PAID").Value)
'
'                        If Round(lSAmount, 2) <= lAmountLeft Then
'                           lAmountLeft = lAmountLeft - lSAmount
'
'                           If oTransactionPost.AllocatePayment(CLng(oSplitData.RecordNumber), CLng(lReceiptSplit), _
'                                 lSAmount, CDate(adoRstRptCh!RDate)) Then
'                              laTransID(iCount) = adoRstRptCh!TransactionID
'                              iCount = iCount + 1
'                           End If
'                        Else
'                           If oTransactionPost.AllocatePayment(CLng(oSplitData.RecordNumber), CLng(lReceiptSplit), _
'                                 lAmountLeft, CDate(adoRstRptCh!RDate)) Then
'                              laTransID(iCount) = adoRstRptCh!TransactionID
'                              iCount = iCount + 1
'                           End If
'                           lAmountLeft = 0
'                        End If
'                        oSplitData.MoveNext
'                     Next j
'                  Else
'                     MsgBox "Prestige data does not match with SAGE data. Please contact with PCM Consulting Ltd.", vbCritical + vbOKOnly, "Prestige <> SAGE"
'                     UpdateMarked laTransID, iCount, adoconn, "RptTransactions"
'                     GoTo CloseConnection
'                  End If
'               Else                       'PAYMENT ON ACCOUNT - DOES NOT ALLOCATE AGAINST ANY TRANSACTION
'                  laTransID(iCount) = adoRstRptCh!TransactionID
'                  iCount = iCount + 1
'               End If
'               adoRstRptCh.MoveNext
'            Wend
'            adoRstRptCh.Close
'         End If
'      End If
'      adoRstRpt.MoveNext
'   Wend
'   adoRstRpt.Close
'   UpdateMarked laTransID, iCount, adoconn, "RptTransactions"
''MsgBox "BEFORE ALLOCATION OF CR AGAINST INV"
'' ***********************************************************************************************************
'' *********************** ALLOCATION OF CREDIT AGAINST INVOICE **********************************************
'' ***********************************************************************************************************
''  Exporting Allocations of 'Credit Invoice' against Invoice
'   Dim lSPN As Long, lSPN_Inv As Long, bAllocated As Boolean
'   Dim oCreditHeaderData As SageDataObject220.HeaderData, oCreditSplitData As SageDataObject220.SplitData
'   Dim oInvHeaderData As SageDataObject220.HeaderData, oInvSplitData As SageDataObject220.SplitData
'
'   szSQLStr = "SELECT RT.TransactionID, RT.FromTran, RT.ToTran, RT.AllocDate AS ADate, " & _
'                  "RT.ReceiptAmount AS RAmt, RT.BankCode, RT.NominalCode, Rpt.Type, " & _
'                  "Rpt.PoA_SPN, Rpt.DemandRef " & _
'              "FROM RptTransactions AS RT, tlbReceipt AS Rpt " & _
'              "WHERE RT.IsSageUpdate AND NOT RT.UpdateSage And " & _
'                  "(isnull(RT.Exp2Sage) OR LEFT(RT.Exp2Sage,1) = 'S') And " & _
'                  "RT.FromTran = Rpt.TransactionID;"
''Debug.Print szSQLStr
'   adoRstRpt.Open szSQLStr, adoconn, adOpenStatic, adLockReadOnly
'
'   Set oCreditHeaderData = oWS.CreateObject("HeaderData")
'   Set oInvHeaderData = oWS.CreateObject("HeaderData")
'
'   iCount = 0
'   While Not adoRstRpt.EOF
'      If adoRstRpt.Fields.Item("Type").Value = 4 Then                'PoA
''MsgBox adoRstRpt.Fields.Item("PoA_SPN").Value
'         lSPN = adoRstRpt.Fields.Item("PoA_SPN").Value
'      End If
'      If adoRstRpt.Fields.Item("Type").Value = 2 Then                'Cr Note
'         lSPN = SagePostingNumberCr(adoRstRpt.Fields.Item("DemandRef").Value, adoconn)
'      End If
'      oCreditHeaderData.Read (lSPN)
'
'      lSPN_Inv = SagePostingNumberInv(adoRstRpt.Fields.Item("ToTran").Value, adoconn)
'      oInvHeaderData.Read (lSPN_Inv)
'
'      'Link to the Credit Splits
'      Set oCreditSplitData = oCreditHeaderData.Link
'      oCreditSplitData.MoveFirst
'
'      Set oInvSplitData = oInvHeaderData.Link
'      oInvSplitData.MoveFirst
'
'      For i = 1 To oInvSplitData.Count
'         If Round(Val(oInvSplitData.Fields.Item("NET_AMOUNT").Value) - _
'               Val(oInvSplitData.Fields.Item("AMOUNT_PAID").Value), 2) >= _
'               Val(adoRstRpt.Fields.Item("RAmt").Value) Then
'            Exit For
'         End If
'         oInvSplitData.MoveNext
'      Next i
'      If i <= oInvSplitData.Count Then
'         Set oTransactionPost = oWS.CreateObject("TransactionPost")
'         bAllocated = oTransactionPost.AllocatePayment(CInt(oInvSplitData.RecordNumber), _
'                        CInt(oCreditSplitData.RecordNumber), CDbl(adoRstRpt.Fields.Item("RAmt").Value), _
'                        CDate(adoRstRpt.Fields.Item("ADate").Value))
'         If bAllocated Then
'            laTransID(iCount) = adoRstRpt.Fields.Item("TransactionID").Value
'            iCount = iCount + 1
'         End If
'      End If
'      Set oCreditSplitData = Nothing
'      Set oInvSplitData = Nothing
'
'      adoRstRpt.MoveNext
'   Wend
'
''   MsgBox iCount & " Transactions (out of " & iTotalTran & ") Posted to SAGE successfully", vbOKOnly, "Posted to SAGE"
'   If iTotalTran - iCount > 0 Then
''    i have stopped this report on basis of bug report of 23rd April, 2007
''     Report of posting exceptions to be printed
''      ShowReport App.Path & szReportPath & "\ReceiptExceptionReport.rpt"
'   End If
'   UpdateMarked laTransID, iCount, adoconn, "RptTransactions"
'
'   If adoPoA.EOF Then
''      MsgBox "There is no Payment on Account to update to SAGE.", vbInformation + vbOKOnly, "Payment on Account "
'      adoPoA.Close
'      Set adoPoA = Nothing
'      GoTo CloseConnection
'   End If
'
'PoA:
''  Update Payment On Account
'   iCount = 0
'   ReDim laTransID(adoPoA.RecordCount) As Long              'Refreshing the array to save all the PoA transaction id
'
'   While Not adoPoA.EOF
''      Create an instance of TransactionPost for the Sales Receipt
'      Set oTransactionPost = oWS.CreateObject("TransactionPost")
'
''      Fill in the Header fields
'      oTransactionPost.Header("ACCOUNT_REF").Value = CStr(adoPoA!SageAccountNumber)
'      oTransactionPost.Header("DATE").Value = CDate(adoPoA!RDate)
'      oTransactionPost.Header("POSTED_DATE").Value = CDate(Date)
'      oTransactionPost.Header("TYPE").Value = CByte(sdoSA)
'      oTransactionPost.Header("DETAILS").Value = "Payment on Account"
'      oTransactionPost.Header("BANK_CODE").Value = CStr(adoPoA!BankCode)
'      oTransactionPost.Header("INV_REF").Value = adoPoA!TransactionID & " - " & IIf(IsNull(adoPoA!ExtRef), "", adoPoA!ExtRef)
'
''      Create a split item by adding an empty split to the Items
''      collection of the TransactionPost Object.
'      Set oSplitData = oTransactionPost.Items.Add()
'
''      Fill in the Split fields - note a Sales Receipt only has one split
'      oSplitData.Fields.Item("TYPE").Value = CByte(sdoSA)
'      oSplitData.Fields.Item("DEPT_NUMBER").Value = CStr(PropDeptNumTent(adoPoA!SageAccountNumber, adoconn))
'      oSplitData.Fields.Item("NOMINAL_CODE").Value = CStr(adoPoA!BankCode)
'      oSplitData.Fields.Item("TAX_CODE").Value = CInt(9)
'      oSplitData.Fields.Item("NET_AMOUNT").Value = CDbl(adoPoA!amount)
'      oSplitData.Fields.Item("TAX_AMOUNT").Value = CDbl(0)
'      oSplitData.Fields.Item("DATE").Value = CDate(oTransactionPost.Header("DATE").Value)
'      oSplitData.Fields.Item("DETAILS").Value = CStr(oTransactionPost.Header("DETAILS").Value)
'      oSplitData.Fields.Item("INTERNAL_REF").Value = CStr(adoPoA.Fields("Ref").Value)
'
'      If oTransactionPost.Update Then
'         SavePostingNumberPoA "tlbReceipt", "TransactionID", CLng(adoPoA.Fields("TransactionID").Value), oTransactionPost.PostingNumber, adoconn
'         laTransID(iCount) = adoPoA!TransactionID
'         iCount = iCount + 1
'      Else
'         MsgBox "Payment on Account has not been exported to SAGE.", vbCritical + vbOKOnly, "PoA - Exported"
'      End If
'
'      adoPoA.MoveNext
'   Wend
'
'   UpdateMarked laTransID, iCount, adoconn, "tlbReceipt"
'
'CloseConnection:
'
'   oWS.Disconnect
'
''   Destroy the Objects
'   Set oCreditHeaderData = Nothing
'   Set oInvHeaderData = Nothing
'   Set oTransactionPost = Nothing
'   Set oSplitData = Nothing
'   Set oHeaderData = Nothing
'   Set oWS = Nothing
'   Set oSDO = Nothing
'
'   Set adoRstRptCh = Nothing
'   Set adoRstRpt = Nothing
'   Set adoconn = Nothing
'
'   Exit Sub
'
'' Error Handling Code
'Error_Handler:
'
'   MsgBox "The SDO generated the following error: " & oSDO.LastError.text & ERR.Number & " -(pcm_SR_Posting) " & ERR.description, vbOKOnly, "Posted to SAGE"
'
'   Set oCreditHeaderData = Nothing
'   Set oInvHeaderData = Nothing
'   Set oTransactionPost = Nothing
'   Set oSplitData = Nothing
'   Set oHeaderData = Nothing
'   Set oWS = Nothing
'   Set oSDO = Nothing
'
'   adoRstRpt.Close
'   adoconn.Close
'   Set adoRstRpt = Nothing
'   Set adoconn = Nothing
'End Sub
'Private Sub ExportTransIC2Sage50v14()
'   Dim reportApp        As New CRAXDRT.Application
'   Dim Report           As CRAXDRT.Report
'   Dim Hkey             As Long
'   Const sageUserName   As String = "Prestige"
'   Const sagePassword   As String = "prestige"
'
'   sageDirPath = GetSetting("PropertyManagement", "SagePath", "SageFolder")
'   If sageDirPath = "" Then
'      sageDirPath = BrowseForFolder(Hkey, "Please select Sage path...")
'      If sageDirPath = "" Then
'          End
'      End If
'      SaveSetting "PropertyManagement", "SagePath", "SageFolder", sageDirPath
'   End If
'
'   On Error GoTo Error_Handler                           '  Error Handler
'
''  Declare SAGE Objects
'   Dim oSDO             As SageDataObject200.SDOEngine
'   Dim oWS              As SageDataObject200.WorkSpace
'   Dim oTransactionPost As SageDataObject200.TransactionPost
'   Dim oSplitData       As SageDataObject200.SplitData
'   Dim oSalesRecord     As SageDataObject200.SalesRecord
'   Dim oNomRec          As SageDataObject200.NominalRecord
'
''  Declare Variables for Database connectivity
'   Dim adoconn          As New ADODB.Connection
'   Dim adoRstDmdHd      As New ADODB.Recordset
'   Dim adoRstDmdSpt     As New ADODB.Recordset
'   Dim adoRst           As New ADODB.Recordset
'
''  Declare Variables
'   Dim szDataPath       As String
'   Dim szSQLStr         As String
'   Dim iCount           As Integer
'   Dim iKount           As Integer
'   Dim i                As Integer
'
''  Set the connection to the databases
'   adoconn.Open getConnectionString
'
''  Create the SDO Engine Object
'   Set oSDO = New SageDataObject200.SDOEngine
'
''  Create the Workspace
'   Set oWS = oSDO.Workspaces.Add("Prestige")
'
''  Select Company.  The SelectCompany method takes the program install folder as a parameter read datapath from registr
'   szDataPath = CompanyDatapath
'   If szDataPath = "" Then
''     Select Company. The SelectCompany method takes the program install folder as a parameter
'      szDataPath = oSDO.SelectCompany(sageDirPath)
'      'written by anol 23 nov 2015
'       CompanyDatapath = sageDirPath
''     Save company name in the registry
'      SaveSetting "PropertyManagement", "SageCompany", CompanyDatapath, szDataPath
'   End If
'
''  The system Connects SAGE; if it fail, throws an exception
'   If oWS.Connect(szDataPath, sageUserName, sagePassword, "Prestige") Then
'
''  Transactions are posting here **********************************************************************
''  Step 1:  Check customer's account in SAGE
''  Step 2:  Check transactions nominal code
''  Step 3:  Upload transacitons into SAGE MMS
''------------------------------------------------------------------------------------------------------
'
'   '  Clear previous marked CF and NF to FAILED
'      adoconn.Execute "UPDATE DemandRecords " & _
'                      "SET spare3 = 'FAILED' " & _
'                      "WHERE spare3 = 'CF' OR spare3 = 'NF' OR spare3 = 'NI' OR spare3 = 'NC';"
'
'   '  Update E -> Y (E: Recently updated transactions, Y: Previously exported)
'      adoconn.Execute "UPDATE DemandRecords " & _
'                      "SET spare3 = 'Y' " & _
'                      "WHERE spare3 = 'E';"
'
''     STEP 1
''     check Customer 's account in SAGE
'      szSQLStr = "SELECT DISTINCT DemandRecords.SageAccountNumber " & _
'                 "FROM DemandRecords INNER JOIN Tenants ON DemandRecords.SageAccountNumber = Tenants.SageAccountNumber " & _
'                 "WHERE (DemandRecords.spare3 = '' OR ISNULL(DemandRecords.spare3)) AND NOT DemandRecords.UPDATE_SAGE AND " & _
'                       "(Tenants.Comments = '' OR ISNULL(Tenants.Comments)) " & _
'                 "ORDER BY DemandRecords.SageAccountNumber;"
'
'      adoRst.Open szSQLStr, adoconn, adOpenStatic, adLockReadOnly
'
'      If Not adoRst.EOF Then
'         Set oSalesRecord = oWS.CreateObject("SalesRecord")
'
'         While Not adoRst.EOF
'            oSalesRecord.Fields.Item("ACCOUNT_REF").Value = UCase(adoRst.Fields.Item("SageAccountNumber").Value)
'            If Not oSalesRecord.Find(False) Then
'
''              Customer not found in SAGE = CF, mark demands not to be exported.
'               szSQLStr = "UPDATE DemandRecords " & _
'                          "SET spare3 = 'CF' " & _
'                          "WHERE SageAccountNumber = '" & adoRst.Fields.Item("SageAccountNumber").Value & "' AND " & _
'                                "spare3 = '';"
'               adoconn.Execute szSQLStr
'               iKount = iKount + 1
'            End If
'            adoRst.MoveNext
'         Wend
'      End If
'      adoRst.Close
'      Set oSalesRecord = Nothing
'
''     Step 2
''     Check transactions nominal code
''     All demand record must have only one split line. System can produce only Single demand.
'      szSQLStr = "SELECT DISTINCT S.NominalCodeforAmount " & _
'                 "FROM DemandRecords AS D, DemandSplitRecords AS S, Tenants AS T " & _
'                 "WHERE (D.spare3 = '' OR ISNULL(D.spare3)) AND D.DemandID = S.DemandID AND " & _
'                      "NOT D.UPDATE_SAGE AND " & _
'                      "D.SageAccountNumber = T.SageAccountNumber AND " & _
'                      "(T.Comments = '' OR ISNULL(T.Comments))"
'   'Debug.Print szSQL
'      adoRst.Open szSQLStr, adoconn, adOpenStatic, adLockReadOnly
'
'      If Not adoRst.EOF Then
'         Set oNomRec = oWS.CreateObject("NominalRecord")
'
'         While Not adoRst.EOF
'            oNomRec.Fields.Item("ACCOUNT_REF").Value = UCase(adoRst.Fields.Item("NominalCodeforAmount").Value)
'
'            If Not oNomRec.Find(False) Then
'               szSQLStr = "UPDATE DemandRecords AS D, DemandSplitRecords AS S " & _
'                          "SET spare3 = 'NF' " & _
'                          "WHERE S.NominalCodeforAmount = '" & adoRst.Fields.Item("NominalCodeforAmount").Value & "';"
'               adoconn.Execute szSQLStr
'               iKount = iKount + 1
'            End If
'            adoRst.MoveNext
'         Wend
'      End If
'      adoRst.Close
'      Set adoRst = Nothing
'      Set oNomRec = Nothing
'
''  Step 3
''  Upload transacitons into SAGE MMS
''  Select all records to be updated
'
''     Connect to Demands table to add new demands.
'      szSQLStr = "SELECT D.DemandID as D_ID, D.SageAccountNumber as S_AC, " & _
'                     "D.TransactionType as T_TYPE, D.IssueDate as I_DATE, " & _
'                     "SUM(S.Amount) as AMT, SUM(S.TotalAmount) as TAMT, " & _
'                     "D.PostingDate " & _
'                 "FROM DemandRecords AS D, DemandSplitRecords AS S " & _
'                 "WHERE (D.spare3 = '' OR ISNULL(D.spare3)) AND  " & _
'                     "D.UPDATE_SAGE = False AND D.DEMANDID =  S.DEMANDID " & _
'                 "GROUP BY D.DemandID, D.SageAccountNumber, " & _
'                     "D.TransactionType, D.IssueDate, D.PostingDate " & _
'                 "ORDER BY D.DEMANDID;"
''   Debug.Print szSQLStr
'      adoRstDmdHd.Open szSQLStr, adoconn, adOpenStatic, adLockReadOnly
'
'      If adoRstDmdHd.EOF Then
'         MsgBox "There no demand to be exported to SAGE.", vbOKOnly, "SAGE"
'
'         Set adoconn = Nothing
'         Set adoRstDmdHd = Nothing
'         Exit Sub
'      End If
'
'      iCount = 0           'Success counter
'      iKount = 0           'Fail counter
'
'      While Not adoRstDmdHd.EOF
'         If Val(adoRstDmdHd.Fields("AMT").Value) > 0 Then
''           Create Instances of Objects
'            Set oTransactionPost = oWS.CreateObject("TransactionPost")
'
''           Get the split lines of the header line
''           The Transactions have 1 or more splits
''   szSQLStr = "SELECT DSR.SplitID as S_ID, " & _
''                           "DSR.NominalCodeforAmount as NCA, " & _
''                           "DSR.Amount AS AMT, DSR.VATAmount AS VAMT, " & _
''                           "DSR.SageRef AS SAGEREF, DSR.DueDate AS D_DT, " & _
''                           "DSR.DateFrom AS F_DT, DSR.DateTo AS T_DT, " & _
''                           "DSR.Description AS DESCP, DSR.VAT_CODE AS V_CODE, " & _
''                           "DSR.SageDepartment AS S_DEPT " & _
''                       "FROM DemandSplitRecords AS DSR " & _
''                       "WHERE DSR.DemandID = " & adoRstDmdHd.Fields("D_ID").Value & " " & _
''                           "AND DSR.DEMANDSTATEMENT=TRUE;"
''Modified by anol 05 Jan 2016
'            szSQLStr = "SELECT DSR.SplitID as S_ID, " & _
'                           "DSR.NominalCodeforAmount as NCA, " & _
'                           "DSR.Amount AS AMT, DSR.VATAmount AS VAMT, " & _
'                           "DSR.SageRef AS SAGEREF, DSR.DueDate AS D_DT, " & _
'                           "DSR.DateFrom AS F_DT, DSR.DateTo AS T_DT, " & _
'                           "DSR.Description AS DESCP, DSR.VAT_CODE AS V_CODE, " & _
'                           "Fund.FundCode AS S_DEPT " & _
'                       "FROM DemandSplitRecords AS DSR,Fund " & _
'                       "WHERE Fund.FundID=DSR.SageDepartment AND DSR.DemandID = " & adoRstDmdHd.Fields("D_ID").Value & " " & _
'                           "AND DSR.DEMANDSTATEMENT=TRUE;"
''         Debug.Print szSQLStr
'            adoRstDmdSpt.Open szSQLStr, adoconn, adOpenStatic, adLockReadOnly
'
''********** Exporting the header -->
'            oTransactionPost.Header("ACCOUNT_REF").Value = CStr(adoRstDmdHd.Fields("S_AC").Value)
'            oTransactionPost.Header("BANK_CODE").Value = CStr(GlbBankID(adoRstDmdHd.Fields("D_ID").Value))
'            oTransactionPost.Header("Date").Value = CDate(adoRstDmdHd.Fields("I_DATE").Value)
'            oTransactionPost.Header("Date_Due").Value = CDate(adoRstDmdSpt.Fields("D_DT").Value)
'            oTransactionPost.Header("DETAILS").Value = CStr(Left(adoRstDmdSpt.Fields("DESCP").Value, 42) & " " & _
'                                                       Format(adoRstDmdSpt.Fields("F_DT").Value, "DD/MM/YY") & "-" & _
'                                                       Format(adoRstDmdSpt.Fields("T_DT").Value, "DD/MM/YY"))
'            oTransactionPost.Header("EURO_GROSS").Value = CDbl(adoRstDmdHd.Fields("TAMT").Value)
'            oTransactionPost.Header("EURO_RATE").Value = CDbl(1)
'            oTransactionPost.Header("FOREIGN_GROSS").Value = CDbl(adoRstDmdHd.Fields("TAMT").Value)
'            oTransactionPost.Header("FOREIGN_RATE").Value = CDbl(1)
'            oTransactionPost.Header("INTEREST_RATE").Value = CDbl(0)
'            oTransactionPost.Header("Inv_Ref").Value = CStr(adoRstDmdHd.Fields("D_ID").Value)
'            oTransactionPost.Header("NET_AMOUNT").Value = CDbl(adoRstDmdHd.Fields("AMT").Value)
'            oTransactionPost.Header("TYPE").Value = CByte(adoRstDmdHd.Fields("T_TYPE").Value)
'            oTransactionPost.Header("Posted_Date").Value = CDate(adoRstDmdHd.Fields("PostingDate").Value)
'
''********** Exporting the split lines -->
'            While Not adoRstDmdSpt.EOF
''              Add a split to the Header's Item collection
'               Set oSplitData = oTransactionPost.Items.Add
'
''              Populate Split Fields
'               oSplitData.Fields.Item("DATE").Value = CDate(oTransactionPost.Header("DATE").Value)
'               oSplitData.Fields.Item("DEPT_NUMBER").Value = CInt(adoRstDmdSpt.Fields("S_DEPT").Value)
'               oSplitData.Fields.Item("DETAILS").Value = CStr(Left(adoRstDmdSpt.Fields("DESCP").Value, 42) & " " & _
'                                                         Format(adoRstDmdSpt.Fields("F_DT").Value, "DD/MM/YY") & "-" & _
'                                                         Format(adoRstDmdSpt.Fields("T_DT").Value, "DD/MM/YY")) 'data inserting: 'Service charge description'
'               oSplitData.Fields.Item("NET_AMOUNT").Value = CDbl(adoRstDmdSpt.Fields("AMT").Value)
'               oSplitData.Fields.Item("NOMINAL_CODE").Value = CStr(adoRstDmdSpt.Fields("NCA").Value)
'               oSplitData.Fields.Item("POSTED_DATE").Value = CDate(oTransactionPost.Header("POSTED_DATE").Value)
'               oSplitData.Fields.Item("TAX_AMOUNT").Value = CDbl(adoRstDmdSpt.Fields("VAMT").Value)
'               oSplitData.Fields.Item("TAX_CODE").Value = CInt(adoRstDmdSpt.Fields("V_CODE").Value)
'               oSplitData.Fields.Item("TYPE").Value = CByte(oTransactionPost.Header("TYPE").Value)
'               oSplitData.Fields.Item("INTERNAL_REF").Value = CStr(adoRstDmdSpt.Fields("SageRef").Value)
'
'               adoRstDmdSpt.MoveNext
'            Wend
'
'            adoRstDmdSpt.Close
'            If oTransactionPost.Update Then
'               SavePostingNumber "DemandRecords", "DemandID", CLng(adoRstDmdHd.Fields("D_ID").Value), oTransactionPost.PostingNumber, adoconn
'               iCount = iCount + 1
'               adoconn.Execute "UPDATE DemandRecords " & _
'                               "SET UPDATE_SAGE = TRUE, spare3 = 'E' " & _
'                               "WHERE UPDATE_SAGE = FALSE AND " & _
'                                     "DemandRecords.DemandID = " & CLng(adoRstDmdHd.Fields("D_ID").Value) & ""
'            Else
'               iKount = iKount + 1
'            End If
'
'            Set oSplitData = Nothing
'            Set oTransactionPost = Nothing
'         Else                                         'NEGATIVE INVOICE -> NI
'            If Val(adoRstDmdHd.Fields("AMT").Value) < 0 Then
'               adoconn.Execute "UPDATE DemandRecords " & _
'                               "SET spare3 = 'NI' " & _
'                               "WHERE DemandID = " & adoRstDmdHd.Fields.Item("D_ID").Value
'               iKount = iKount + 1
'            End If
'         End If
'
'         adoRstDmdHd.MoveNext
'      Wend
'      adoRstDmdHd.Close
'
'      MsgBox iCount & " Transactions (out of " & iCount + iKount & ") Posted to SAGE successfully", vbOKOnly, "Posted to SAGE"
'      oWS.Disconnect
'   End If
'
'   Set oWS = Nothing
'   Set oSDO = Nothing
'
'   Set adoRstDmdSpt = Nothing
'   Set adoRstDmdHd = Nothing
'   Set adoconn = Nothing
''*************************************************************************************************************
''  Print Report
'
'   Dim rep As New frmReport
'
'   If iKount > 0 Then
'      Set Report = reportApp.OpenReport(App.Path & szReportPath & "\ExpDmdList_Not.rpt")
'      Report.Database.Tables(1).ConnectionProperties.Item("Database Password") = accessDBPws
'
'      Report.EnableParameterPrompting = False
'      Report.DiscardSavedData
'
'      Load rep
'      rep.LoadReportViewer Report
'   End If
'
'   If iCount > 0 Then
'      Set Report = reportApp.OpenReport(App.Path & szReportPath & "\ExpDmdList.rpt")
'
'      Report.Database.Tables(1).ConnectionProperties.Item("Database Password") = accessDBPws
'
'      Report.EnableParameterPrompting = False
'      Report.DiscardSavedData
'
'      Report.ParameterFields(1).AddCurrentValue "ExportedRpt"
'      Report.ParameterFields(2).AddCurrentValue "List of demands Exported Successfully"
'      Report.ParameterFields(3).AddCurrentValue "E"
'
'      Load rep
'      rep.LoadReportViewer Report
'   End If
'
'   Exit Sub
'
''    Error Handling Code
'Error_Handler:
'
'   Select Case oSDO.LastError.Code
'
'    ' Invalid Password
'    Case Is = sdoLogonNameInUse
'        MsgBox "Logon Failed Due To: " & oSDO.LastError.text, vbInformation, ERR.description
'
'    ' User is Already Logged On
'    Case Is = sdoLogonNameInUse
'        MsgBox "Logon Failed Due To: " & oSDO.LastError.text, vbInformation, ERR.description
'
'    ' Sage 50 Accounts is in Exclusive Mode i.e. File Maintenance
'    Case Is = sdoLogonExclusive
'        MsgBox "Logon Failed Due To: " & oSDO.LastError.text, vbInformation, ERR.description
'
'    ' Invalid Data Path on Connect
'    Case Is = sdoBadDataPath
'        MsgBox "Logon Failed Due To: " & oSDO.LastError.text, vbInformation, ERR.description
'
'    ' Sage 50 Accounts is not Same Version as SDO
'    Case Is = sdoWrongVersion
'        MsgBox "Logon Failed Due To: " & oSDO.LastError.text, vbInformation, ERR.description
'
'    ' SDO is Not Registered on this Machine
'    Case Is = sdoNotRegistered
'        MsgBox "Logon Failed Due To: " & oSDO.LastError.text, vbInformation, ERR.description
'
'   Case Else
'        MsgBox "Logon Failed Due To: " & oSDO.LastError.text, vbInformation, ERR.description
'   End Select
'
'   ' Destroy Objects
'   Set oTransactionPost = Nothing
'   Set oSplitData = Nothing
'   Set oWS = Nothing
'   Set oSDO = Nothing
'
'   Set adoconn = Nothing
'   Set adoRstDmdHd = Nothing
'End Sub
'Private Sub ExportTransIC2Sage50v16()
'   Dim reportApp        As New CRAXDRT.Application
'   Dim Report           As CRAXDRT.Report
'   Dim Hkey             As Long
'   Const sageUserName   As String = "Prestige"
'   Const sagePassword   As String = "prestige"
''   1.Read if there is any dirpath found in registry
''   2.if dirpath empty pop up browse folder
''   3.Save dir path in registry
''   4.after select ing a company dll shall give you DataPath
''   5. having the datapath you can connect to company
''   6.after connecting company u can create work space
''   7.You can create object for transaction post
''  8.post transaction
'   'SaveSetting "PropertyManagement", "SagePath", "SageFolder", "C:\ProgramData\Sage\Accounts\2014"
'   sageDirPath = GetSetting("PropertyManagement", "SagePath", "SageFolder")
'   If sageDirPath = "" Then
'      sageDirPath = BrowseForFolder(Hkey, "Please select Sage path...")
'      If sageDirPath = "" Then
'          End
'      End If
'      SaveSetting "PropertyManagement", "SagePath", "SageFolder", sageDirPath
'   End If
'
'   On Error GoTo Error_Handler                           '  Error Handler
'
''  Declare SAGE Objects
'   Dim oSDO             As SageDataObject220.SDOEngine
'   Dim oWS              As SageDataObject220.WorkSpace
'   Dim oTransactionPost As SageDataObject220.TransactionPost
'   Dim oSplitData       As SageDataObject220.SplitData
'   Dim oSalesRecord     As SageDataObject220.SalesRecord
'   Dim oNomRec          As SageDataObject220.NominalRecord
'
''  Declare Variables for Database connectivity
'   Dim adoconn          As New ADODB.Connection
'   Dim adoRstDmdHd      As New ADODB.Recordset
'   Dim adoRstDmdSpt     As New ADODB.Recordset
'   Dim adoRst           As New ADODB.Recordset
'
''  Declare Variables
'   Dim szDataPath       As String
'   Dim szSQLStr         As String
'   Dim iCount           As Integer
'   Dim iKount           As Integer
'   Dim i                As Integer
'
''  Set the connection to the databases
'   adoconn.Open getConnectionString
'
''  Create the SDO Engine Object
'   Set oSDO = New SageDataObject220.SDOEngine
'
''  Create the Workspace
'   Set oWS = oSDO.Workspaces.Add("Prestige")
''SaveSetting "PropertyManagement", "SageCompany", CompanyDatapath, "C:\ProgramData\Sage\Accounts\2016\COMPANY.001\ACCDATA"
''Exit Sub
''  CompanyDatapath = Reg.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\ODBC\ODBC.INI\" & CompanyDatapath & "\DataPathname")
''szDataPath = "C:\ProgramData\Sage\Accounts\2016\COMPANY.001\ACCDATA"R
''  Select Company.  The SelectCompany method takes the program install folder as a parameter read datapath from registr
'   szDataPath = CompanyDatapath
'   'If szDataPath = "" Then' you want always a popup for companies so rem this line salia
''     Select Company. The SelectCompany method takes the program install folder as a parameter
'      szDataPath = oSDO.SelectCompany(sageDirPath)
'      'written by anol 23 nov 2015
'       CompanyDatapath = sageDirPath
''     Save company name in the registry
'      SaveSetting "PropertyManagement", "SageCompany", CompanyDatapath, szDataPath
'  'End If
''szDataPath = "C:\ProgramData\Sage\Accounts\2016\COMPANY.001\ACCDATA"
'
''  The system Connects SAGE; if it fail, throws an exception
'   If oWS.Connect(szDataPath, sageUserName, sagePassword, "Prestige") Then
'
''  Transactions are posting here **********************************************************************
''  Step 1:  Check customer's account in SAGE
''  Step 2:  Check transactions nominal code
''  Step 3:  Upload transacitons into SAGE MMS
''------------------------------------------------------------------------------------------------------
'
'   '  Clear previous marked CF and NF to FAILED
'      adoconn.Execute "UPDATE DemandRecords " & _
'                      "SET spare3 = 'FAILED' " & _
'                      "WHERE spare3 = 'CF' OR spare3 = 'NF' OR spare3 = 'NI' OR spare3 = 'NC';"
'
'   '  Update E -> Y (E: Recently updated transactions, Y: Previously exported)
'      adoconn.Execute "UPDATE DemandRecords " & _
'                      "SET spare3 = 'Y' " & _
'                      "WHERE spare3 = 'E';"
'
''     STEP 1
''     check Customer 's account in SAGE
'      szSQLStr = "SELECT DISTINCT DemandRecords.SageAccountNumber " & _
'                 "FROM DemandRecords INNER JOIN Tenants ON DemandRecords.SageAccountNumber = Tenants.SageAccountNumber " & _
'                 "WHERE (DemandRecords.spare3 = '' OR ISNULL(DemandRecords.spare3)) AND NOT DemandRecords.UPDATE_SAGE AND " & _
'                       "(Tenants.Comments = '' OR ISNULL(Tenants.Comments)) " & _
'                 "ORDER BY DemandRecords.SageAccountNumber;"
'
'      adoRst.Open szSQLStr, adoconn, adOpenStatic, adLockReadOnly
'
'      If Not adoRst.EOF Then
'         Set oSalesRecord = oWS.CreateObject("SalesRecord")
'
'         While Not adoRst.EOF
'            oSalesRecord.Fields.Item("ACCOUNT_REF").Value = UCase(adoRst.Fields.Item("SageAccountNumber").Value)
'            If Not oSalesRecord.Find(False) Then
'
''              Customer not found in SAGE = CF, mark demands not to be exported.
'               szSQLStr = "UPDATE DemandRecords " & _
'                          "SET spare3 = 'CF' " & _
'                          "WHERE SageAccountNumber = '" & adoRst.Fields.Item("SageAccountNumber").Value & "' AND " & _
'                                "spare3 = '';"
'               adoconn.Execute szSQLStr
'               iKount = iKount + 1
'            End If
'            adoRst.MoveNext
'         Wend
'      End If
'      adoRst.Close
'      Set oSalesRecord = Nothing
'
''     Step 2
''     Check transactions nominal code
''     All demand record must have only one split line. System can produce only Single demand.
'      szSQLStr = "SELECT DISTINCT S.NominalCodeforAmount " & _
'                 "FROM DemandRecords AS D, DemandSplitRecords AS S, Tenants AS T " & _
'                 "WHERE (D.spare3 = '' OR ISNULL(D.spare3)) AND D.DemandID = S.DemandID AND " & _
'                      "NOT D.UPDATE_SAGE AND " & _
'                      "D.SageAccountNumber = T.SageAccountNumber AND " & _
'                      "(T.Comments = '' OR ISNULL(T.Comments))"
'   'Debug.Print szSQL
'      adoRst.Open szSQLStr, adoconn, adOpenStatic, adLockReadOnly
'
'      If Not adoRst.EOF Then
'         Set oNomRec = oWS.CreateObject("NominalRecord")
'
'         While Not adoRst.EOF
'            oNomRec.Fields.Item("ACCOUNT_REF").Value = UCase(adoRst.Fields.Item("NominalCodeforAmount").Value)
'
'            If Not oNomRec.Find(False) Then
'               szSQLStr = "UPDATE DemandRecords AS D, DemandSplitRecords AS S " & _
'                          "SET spare3 = 'NF' " & _
'                          "WHERE S.NominalCodeforAmount = '" & adoRst.Fields.Item("NominalCodeforAmount").Value & "';"
'               adoconn.Execute szSQLStr
'               iKount = iKount + 1
'            End If
'            adoRst.MoveNext
'         Wend
'      End If
'      adoRst.Close
'      Set adoRst = Nothing
'      Set oNomRec = Nothing
'
''  Step 3
''  Upload transacitons into SAGE MMS
''  Select all records to be updated
'
''     Connect to Demands table to add new demands.
'      szSQLStr = "SELECT D.DemandID as D_ID, D.SageAccountNumber as S_AC, " & _
'                     "D.TransactionType as T_TYPE, D.IssueDate as I_DATE, " & _
'                     "SUM(S.Amount) as AMT, SUM(S.TotalAmount) as TAMT, " & _
'                     "D.PostingDate " & _
'                 "FROM DemandRecords AS D, DemandSplitRecords AS S " & _
'                 "WHERE (D.spare3 = '' OR ISNULL(D.spare3)) AND  " & _
'                     "D.UPDATE_SAGE = False AND D.DEMANDID =  S.DEMANDID " & _
'                 "GROUP BY D.DemandID, D.SageAccountNumber, " & _
'                     "D.TransactionType, D.IssueDate, D.PostingDate " & _
'                 "ORDER BY D.DEMANDID;"
''   Debug.Print szSQLStr
'      adoRstDmdHd.Open szSQLStr, adoconn, adOpenStatic, adLockReadOnly
'
'      If adoRstDmdHd.EOF Then
'         MsgBox "There no demand to be exported to SAGE.", vbOKOnly, "SAGE"
'
'         Set adoconn = Nothing
'         Set adoRstDmdHd = Nothing
'         Exit Sub
'      End If
'
'      iCount = 0           'Success counter
'      iKount = 0           'Fail counter
'
'      While Not adoRstDmdHd.EOF
'         If Val(adoRstDmdHd.Fields("AMT").Value) > 0 Then
''           Create Instances of Objects
'            Set oTransactionPost = oWS.CreateObject("TransactionPost")
'
''           Get the split lines of the header line
''           The Transactions have 1 or more splits
''   szSQLStr = "SELECT DSR.SplitID as S_ID, " & _
''                           "DSR.NominalCodeforAmount as NCA, " & _
''                           "DSR.Amount AS AMT, DSR.VATAmount AS VAMT, " & _
''                           "DSR.SageRef AS SAGEREF, DSR.DueDate AS D_DT, " & _
''                           "DSR.DateFrom AS F_DT, DSR.DateTo AS T_DT, " & _
''                           "DSR.Description AS DESCP, DSR.VAT_CODE AS V_CODE, " & _
''                           "DSR.SageDepartment AS S_DEPT " & _
''                       "FROM DemandSplitRecords AS DSR " & _
''                       "WHERE DSR.DemandID = " & adoRstDmdHd.Fields("D_ID").Value & " " & _
''                           "AND DSR.DEMANDSTATEMENT=TRUE;"
''Modified by anol 05 Jan 2016
'            szSQLStr = "SELECT DSR.SplitID as S_ID, " & _
'                           "DSR.NominalCodeforAmount as NCA, " & _
'                           "DSR.Amount AS AMT, DSR.VATAmount AS VAMT, " & _
'                           "DSR.SageRef AS SAGEREF, DSR.DueDate AS D_DT, " & _
'                           "DSR.DateFrom AS F_DT, DSR.DateTo AS T_DT, " & _
'                           "DSR.Description AS DESCP, DSR.VAT_CODE AS V_CODE, " & _
'                           "Fund.FundCode AS S_DEPT " & _
'                       "FROM DemandSplitRecords AS DSR,Fund " & _
'                       "WHERE Fund.FundID=DSR.SageDepartment AND DSR.DemandID = " & adoRstDmdHd.Fields("D_ID").Value & " " & _
'                           "AND DSR.DEMANDSTATEMENT=TRUE;"
''         Debug.Print szSQLStr
'            adoRstDmdSpt.Open szSQLStr, adoconn, adOpenStatic, adLockReadOnly
'
''********** Exporting the header -->
'            oTransactionPost.Header("ACCOUNT_REF").Value = CStr(adoRstDmdHd.Fields("S_AC").Value)
'            oTransactionPost.Header("BANK_CODE").Value = CStr(GlbBankID(adoRstDmdHd.Fields("D_ID").Value))
'            oTransactionPost.Header("Date").Value = CDate(adoRstDmdHd.Fields("I_DATE").Value)
'            oTransactionPost.Header("Date_Due").Value = CDate(adoRstDmdSpt.Fields("D_DT").Value)
'            oTransactionPost.Header("DETAILS").Value = CStr(Left(adoRstDmdSpt.Fields("DESCP").Value, 42) & " " & _
'                                                       Format(adoRstDmdSpt.Fields("F_DT").Value, "DD/MM/YY") & "-" & _
'                                                       Format(adoRstDmdSpt.Fields("T_DT").Value, "DD/MM/YY"))
'            oTransactionPost.Header("EURO_GROSS").Value = CDbl(adoRstDmdHd.Fields("TAMT").Value)
'            oTransactionPost.Header("EURO_RATE").Value = CDbl(1)
'            oTransactionPost.Header("FOREIGN_GROSS").Value = CDbl(adoRstDmdHd.Fields("TAMT").Value)
'            oTransactionPost.Header("FOREIGN_RATE").Value = CDbl(1)
'            oTransactionPost.Header("INTEREST_RATE").Value = CDbl(0)
'            oTransactionPost.Header("Inv_Ref").Value = CStr(adoRstDmdHd.Fields("D_ID").Value)
'            oTransactionPost.Header("NET_AMOUNT").Value = CDbl(adoRstDmdHd.Fields("AMT").Value)
'            oTransactionPost.Header("TYPE").Value = CByte(adoRstDmdHd.Fields("T_TYPE").Value)
'            oTransactionPost.Header("Posted_Date").Value = CDate(adoRstDmdHd.Fields("PostingDate").Value)
'
''********** Exporting the split lines -->
'            While Not adoRstDmdSpt.EOF
''              Add a split to the Header's Item collection
'               Set oSplitData = oTransactionPost.Items.Add
'
''              Populate Split Fields
'               oSplitData.Fields.Item("DATE").Value = CDate(oTransactionPost.Header("DATE").Value)
'               oSplitData.Fields.Item("DEPT_NUMBER").Value = CInt(adoRstDmdSpt.Fields("S_DEPT").Value)
'               oSplitData.Fields.Item("DETAILS").Value = CStr(Left(adoRstDmdSpt.Fields("DESCP").Value, 42) & " " & _
'                                                         Format(adoRstDmdSpt.Fields("F_DT").Value, "DD/MM/YY") & "-" & _
'                                                         Format(adoRstDmdSpt.Fields("T_DT").Value, "DD/MM/YY")) 'data inserting: 'Service charge description'
'               oSplitData.Fields.Item("NET_AMOUNT").Value = CDbl(adoRstDmdSpt.Fields("AMT").Value)
'               oSplitData.Fields.Item("NOMINAL_CODE").Value = CStr(adoRstDmdSpt.Fields("NCA").Value)
'               oSplitData.Fields.Item("POSTED_DATE").Value = CDate(oTransactionPost.Header("POSTED_DATE").Value)
'               oSplitData.Fields.Item("TAX_AMOUNT").Value = CDbl(adoRstDmdSpt.Fields("VAMT").Value)
'               'oSplitData.Fields.Item("TAX_CODE").Value = CInt(adoRstDmdSpt.Fields("V_CODE").Value)
'               'null checking has been implemented by anol 2021-02-23
'               oSplitData.Fields.Item("TAX_CODE").Value = CInt(IIf(IsNull(adoRstDmdSpt.Fields("V_CODE").Value), 0, adoRstDmdSpt.Fields("V_CODE").Value))
'               oSplitData.Fields.Item("TYPE").Value = CByte(oTransactionPost.Header("TYPE").Value)
'               oSplitData.Fields.Item("INTERNAL_REF").Value = CStr(adoRstDmdSpt.Fields("SageRef").Value)
'
'               adoRstDmdSpt.MoveNext
'            Wend
'
'            adoRstDmdSpt.Close
'            If oTransactionPost.Update Then
'               SavePostingNumber "DemandRecords", "DemandID", CLng(adoRstDmdHd.Fields("D_ID").Value), oTransactionPost.PostingNumber, adoconn
'               iCount = iCount + 1
'               adoconn.Execute "UPDATE DemandRecords " & _
'                               "SET UPDATE_SAGE = TRUE, spare3 = 'E' " & _
'                               "WHERE UPDATE_SAGE = FALSE AND " & _
'                                     "DemandRecords.DemandID = " & CLng(adoRstDmdHd.Fields("D_ID").Value) & ""
'            Else
'               iKount = iKount + 1
'            End If
'
'            Set oSplitData = Nothing
'            Set oTransactionPost = Nothing
'         Else                                         'NEGATIVE INVOICE -> NI
'            If Val(adoRstDmdHd.Fields("AMT").Value) < 0 Then
'               adoconn.Execute "UPDATE DemandRecords " & _
'                               "SET spare3 = 'NI' " & _
'                               "WHERE DemandID = " & adoRstDmdHd.Fields.Item("D_ID").Value
'               iKount = iKount + 1
'            End If
'         End If
'
'         adoRstDmdHd.MoveNext
'      Wend
'      adoRstDmdHd.Close
'
'      MsgBox iCount & " Transactions (out of " & iCount + iKount & ") Posted to SAGE successfully", vbOKOnly, "Posted to SAGE"
'      oWS.Disconnect
'   End If
'
'   Set oWS = Nothing
'   Set oSDO = Nothing
'
'   Set adoRstDmdSpt = Nothing
'   Set adoRstDmdHd = Nothing
'   Set adoconn = Nothing
''*************************************************************************************************************
''  Print Report
'
'   Dim rep As New frmReport
'
'   If iKount > 0 Then
'      Set Report = reportApp.OpenReport(App.Path & szReportPath & "\ExpDmdList_Not.rpt")
'      Report.Database.Tables(1).ConnectionProperties.Item("Database Password") = accessDBPws
'
'      Report.EnableParameterPrompting = False
'      Report.DiscardSavedData
'
'      Load rep
'      rep.LoadReportViewer Report
'   End If
'
'   If iCount > 0 Then
'      Set Report = reportApp.OpenReport(App.Path & szReportPath & "\ExpDmdList.rpt")
'
'      Report.Database.Tables(1).ConnectionProperties.Item("Database Password") = accessDBPws
'
'      Report.EnableParameterPrompting = False
'      Report.DiscardSavedData
'
'      Report.ParameterFields(1).AddCurrentValue "ExportedRpt"
'      Report.ParameterFields(2).AddCurrentValue "List of demands Exported Successfully"
'      Report.ParameterFields(3).AddCurrentValue "E"
'
'      Load rep
'      rep.LoadReportViewer Report
'   End If
'
'   Exit Sub
'
''    Error Handling Code
'Error_Handler:
'
'   Select Case oSDO.LastError.Code
'
'    ' Invalid Password
'    Case Is = sdoLogonNameInUse
'        MsgBox "Logon Failed Due To: " & oSDO.LastError.text, vbInformation, ERR.description
'
'    ' User is Already Logged On
'    Case Is = sdoLogonNameInUse
'        MsgBox "Logon Failed Due To: " & oSDO.LastError.text, vbInformation, ERR.description
'
'    ' Sage 50 Accounts is in Exclusive Mode i.e. File Maintenance
'    Case Is = sdoLogonExclusive
'        MsgBox "Logon Failed Due To: " & oSDO.LastError.text, vbInformation, ERR.description
'
'    ' Invalid Data Path on Connect
'    Case Is = sdoBadDataPath
'        MsgBox "Logon Failed Due To: " & oSDO.LastError.text, vbInformation, ERR.description
'
'    ' Sage 50 Accounts is not Same Version as SDO
'    Case Is = sdoWrongVersion
'        MsgBox "Logon Failed Due To: " & oSDO.LastError.text, vbInformation, ERR.description
'
'    ' SDO is Not Registered on this Machine
'    Case Is = sdoNotRegistered
'        MsgBox "Logon Failed Due To: " & oSDO.LastError.text, vbInformation, ERR.description
'
'   Case Else
'        MsgBox "Logon Failed Due To: " & oSDO.LastError.text, vbInformation, ERR.description
'   End Select
'
'   ' Destroy Objects
'   Set oTransactionPost = Nothing
'   Set oSplitData = Nothing
'   Set oWS = Nothing
'   Set oSDO = Nothing
'
'   Set adoconn = Nothing
'   Set adoRstDmdHd = Nothing
'End Sub
'Private Sub ExportTrans2Sage200()
'   Dim retval As Long
'   Dim szDataSource As String, iTrans As Integer, iExpTran As Integer
'   Dim szSQL As String
'   Dim adoconn As New ADODB.Connection
'   Dim adoRst As New ADODB.Recordset
'   Dim reportApp As New CRAXDRT.Application
'   Dim Report As CRAXDRT.Report
'
'   On Error GoTo CatchErr
'
'   Me.MousePointer = vbHourglass
'   frmMMain.MousePointer = vbHourglass
'
'   adoconn.Open getConnectionString
'
'   szDataSource = GetDataSource(adoconn)
'   If szDataSource = "" Then
'      MsgBox "Please enter Sage 200 Company Name shown in Sage 200 SYSADMIN in the 3rd Party data source field in the Company Setup screen in Prestige.", vbCritical + vbOKOnly, "SAGE MMS"
'
'      adoconn.Close
'      Set adoconn = Nothing
'      Me.MousePointer = vbArrow
'      frmMMain.MousePointer = vbArrow
'      Exit Sub
'   End If
'
''  Transactions are posting here **********************************************************************
''  Step 1:  Check customer's account in SAGE
''  Step 2:  Check transactions nominal code
''  Step 3:  Upload transacitons into SAGE MMS
''------------------------------------------------------------------------------------------------------
'
''  Clear previous marked CF and NF to FAILED
'   adoconn.Execute "UPDATE DemandRecords " & _
'                   "SET spare3 = 'FAILED' " & _
'                   "WHERE spare3 = 'CF' OR spare3 = 'NF' OR spare3 = 'NI' OR spare3 = 'NC';"
'
''  Update E -> Y (E: Recently updated transactions, Y: Previously exported)
'   adoconn.Execute "UPDATE DemandRecords " & _
'                   "SET spare3 = 'Y' " & _
'                   "WHERE spare3 = 'E';"
'   'update DemandSplitRecords set SageRef = (Format(DateFrom,"dd/mm/yy") & "-"& Format(dateTo,"dd/mm/yy"))
''   adoConn.Execute "UPDATE DemandSplitRecords " & _
''                   "SET Sageref =(Format(DateFrom,'dd/mm/yy') & " - "& Format(dateTo,'dd/mm/yy'));"
'
'   szDataSource = GetDataSource(adoconn)
'    'Old code version2011
'   'retval = ExecCmd(App.Path + "\Export2Sage200v2011.exe " & Adsn & " """ & szDataSource & """")
'
''   'new code version 2017
''   retval = ExecCmd(App.Path + "\\Export2Sage200v2017\\postingdata2Sage.exe " & Adsn & " """ & szDataSource & """")
''new code version 2018
'   'retval = ExecCmd(App.Path + "\\Export2Sage200v2018\\postingdata2Sage.exe " & Adsn & " """ & szDataSource & """")
'   'Export2Sage200v2021R1
'   retval = ExecCmd(App.Path + "\\Export2Sage200v2022R2\\postingdata2Sage.exe " & Adsn & " """ & szDataSource & """")
'   MsgBox "The Export to Sage 200 is complete .", vbInformation + vbOKOnly, "SAGE 200"
'
'   szSQL = "SELECT COUNT(*) FROM DemandRecords WHERE  spare3 = 'E';"
'   adoRst.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
'   iExpTran = 0
'   iExpTran = adoRst.Fields.Item(0).Value
'   adoRst.Close
'
'   szSQL = "SELECT COUNT(*) FROM DemandRecords WHERE spare3 = 'CF' OR spare3 = 'NF' OR spare3 = 'NI' OR spare3 = 'NC';"
'   adoRst.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
'   iTrans = 0
'   iTrans = adoRst.Fields.Item(0).Value
'   adoRst.Close
'   Set adoRst = Nothing
'
'   adoconn.Close
'   Set adoconn = Nothing
''*************************************************************************************************************
'   Dim rep As New frmReport
'
'   If iTrans > 0 Then
'      MsgBox "There are " & iTrans & " transaction(s) that have not been posted.", vbCritical + vbOKOnly, "Exported"
''  Print Report
'      Set Report = reportApp.OpenReport(App.Path & szReportPath & "\ExpDmdList_Not.rpt")
'      Report.Database.Tables(1).ConnectionProperties.Item("Database Password") = accessDBPws
'
'      Report.EnableParameterPrompting = False
'      Report.DiscardSavedData
'
'      Load rep
'      rep.LoadReportViewer Report
'   End If
'
'   If iExpTran > 0 Then
'      MsgBox iExpTran & " Transactions have been posted successfully.", vbInformation + vbOKOnly, "Exported"
'
'      Set Report = reportApp.OpenReport(App.Path & szReportPath & "\ExpDmdList.rpt")
'
'      Report.Database.Tables(1).ConnectionProperties.Item("Database Password") = accessDBPws
'
'      Report.EnableParameterPrompting = False
'      Report.DiscardSavedData
'
'      Report.ParameterFields(1).AddCurrentValue "ExportedRpt"
'      Report.ParameterFields(2).AddCurrentValue "List of demands Exported Successfully"
'      Report.ParameterFields(3).AddCurrentValue "E"
'
'      Load rep
'      rep.LoadReportViewer Report
'   End If
'
'   fraConfirmation.Top = 4000
'
'   Me.MousePointer = vbArrow
'   frmMMain.MousePointer = vbArrow
'
'   Exit Sub
'CatchErr:
'   If ERR.Number = 94 Then
'      MsgBox "Please specify the name of your 3rd party software in company settings.", vbInformation + vbOKOnly, "Warning"
'   Else
'      MsgBox ERR.Number & ": " & ERR.description, vbCritical + vbOKOnly, "Warning.."
'   End If
'   adoconn.Close
'   Set adoconn = Nothing
'   Me.MousePointer = vbArrow
'   frmMMain.MousePointer = vbArrow
'End Sub
'
'Private Sub ExportTransactionsSageMMS()
'   Dim szDataSource As String, iTrans As Integer, iExpTran As Integer
'   Dim adoconn As New ADODB.Connection
'   Dim adoCustomer As New ADODB.Recordset, adoDemands As New ADODB.Recordset
'   Dim szSQL As String, iKompany As Integer
'   Dim oApplication As New SAGEACCOUNTINGLib.Application
'   Dim oCustomer As SAGEACCOUNTINGLib.Customer
'   Dim oCustomers As SAGEACCOUNTINGLib.Customers
'   Dim oSI_Instrument As SAGEACCOUNTINGLib.SalesInvoiceInstrument
'   Dim oSC_Instrument As SAGEACCOUNTINGLib.SalesCreditNoteInstrument
'   Dim oNominalCodes As SAGEACCOUNTINGLib.NominalCodes
'   Dim reportApp As New CRAXDRT.Application
'   Dim Report As CRAXDRT.Report
'
'   On Error GoTo CatchErr
'
'   Me.MousePointer = vbHourglass
'   frmMMain.MousePointer = vbHourglass
'
'   adoconn.Open getConnectionString
'
'   szDataSource = GetDataSource(adoconn)
'   If szDataSource = "" Then
'      MsgBox "Please set the data source name in the company setup form, which will be found in tools.", vbCritical + vbOKOnly, "SAGE MMS"
'
'      adoconn.Close
'      Set adoconn = Nothing
'      Me.MousePointer = vbArrow
'      frmMMain.MousePointer = vbArrow
'      Exit Sub
'   End If
'
''  Connecting SAGE Object
'   oApplication.Connect "PCM", "", "PrestigeMMS"
'
'   For iKompany = 0 To oApplication.Companies.Count - 1
'      If UCase(oApplication.Companies.Item(iKompany).Name) = UCase(szDataSource) Then Exit For
'   Next iKompany
'   If iKompany = oApplication.Companies.Count Then
'      MsgBox "Prestige could not found the company in SAGE.", vbCritical + vbOKOnly, "SAGE MMS"
'
'      adoconn.Close
'      Set adoconn = Nothing
'      Me.MousePointer = vbArrow
'      frmMMain.MousePointer = vbArrow
'      Exit Sub
'   End If
'
'   oApplication.ActiveCompany = oApplication.Companies.Item(iKompany)
'
''  Transactions are posting here **********************************************************************
''  Step 1:  Check customer's account in SAGE
''  Step 2:  Check transactions nominal code
''  Step 3:  Upload transacitons into SAGE MMS
''------------------------------------------------------------------------------------------------------
'
''  Clear previous marked CF and NF to FAILED
'   adoconn.Execute "UPDATE DemandRecords " & _
'                   "SET spare3 = 'FAILED' " & _
'                   "WHERE spare3 = 'CF' OR spare3 = 'NF' OR spare3 = 'NI' OR spare3 = 'NC';"
'
''  Update E -> Y (E: Recently updated transactions, Y: Previously exported)
'   adoconn.Execute "UPDATE DemandRecords " & _
'                   "SET spare3 = 'Y' " & _
'                   "WHERE spare3 = 'E';"
'
''  STEP 1
''  Check customer's account in SAGE
'   szSQL = "SELECT DISTINCT SageAccountNumber " & _
'           "FROM DemandRecords " & _
'           "WHERE spare3 = '' OR ISNULL(spare3) " & _
'           "ORDER BY SageAccountNumber;"
'
'   adoCustomer.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
'
'   If Not adoCustomer.EOF Then
'      While Not adoCustomer.EOF
'         Set oCustomers = oApplication.SalesModule.Customers
'
'         Set oCustomer = oCustomers(UCase(adoCustomer.Fields.Item("SageAccountNumber").Value))
'         If oCustomer Is Nothing Then
'
''           Customer not found in SAGE = CF, mark demands not to be exported.
'            szSQL = "UPDATE DemandRecords " & _
'                    "SET spare3 = 'CF' " & _
'                    "WHERE SageAccountNumber = '" & adoCustomer.Fields.Item("SageAccountNumber").Value & "' AND " & _
'                          "spare3 = '';"
'            adoconn.Execute szSQL
'            iTrans = iTrans + 1
'         End If
'         adoCustomer.MoveNext
'      Wend
'   End If
'   adoCustomer.Close
'
''  Step 2
''  Check transactions nominal code
''  All demand record must have only one split line. System can produce only Single demand.
'   szSQL = "SELECT DISTINCT S.NominalCodeforAmount " & _
'           "FROM DemandRecords AS D, DemandSplitRecords AS S " & _
'           "WHERE (D.spare3 = '' OR ISNULL(D.spare3)) AND D.DemandID = S.DemandID;"
''Debug.Print szSQL
'   adoCustomer.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
'
'   If Not adoCustomer.EOF Then
'      While Not adoCustomer.EOF
'         Set oNominalCodes = oApplication.NominalModule.NominalCodes
'
'         oNominalCodes.Filter = "ACCOUNT-NUMBER='" & adoCustomer.Fields.Item("NominalCodeforAmount").Value & "'"
'         oNominalCodes.Requery
'
'         If oNominalCodes.Count = 0 Then        'Nominal code not found in SAGE = NF
'
'            'mark demands not to be exported. NF = Not Found
'            szSQL = "UPDATE DemandRecords AS D, DemandSplitRecords AS S " & _
'                    "SET spare3 = 'NF' " & _
'                    "WHERE S.NominalCodeforAmount = '" & adoCustomer.Fields.Item("NominalCodeforAmount").Value & "';"
'            adoconn.Execute szSQL
'            iTrans = iTrans + 1
'         End If
'         adoCustomer.MoveNext
'      Wend
'   End If
'   adoCustomer.Close
'
''  Step 3
''  Upload transacitons into SAGE MMS
''  Select all records to be updated
'
'   szSQL = "SELECT S.Amount AS T_Amt, S.VATAmount AS T_VAT, S.SageRef, " & _
'               "D.SageAccountNumber, D.spare3, D.TransactionType, T.Prefix, " & _
'               "D.ISSUEDATE , D.DemandId, s.DueDate, D.DmdSlNo, s.NominalCodeforAmount, D.DmdSlNo " & _
'           "FROM DemandRecords AS D, DemandSplitRecords AS S, DemandTypes AS T " & _
'           "WHERE (D.spare3 = '' OR ISNULL(D.spare3)) AND D.DemandID = S.DemandID AND " & _
'               "S.TypeOfDemand = T.ID AND D.IsPrinted = TRUE;"
'
''Debug.Print szSQL
'   adoDemands.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
'
'   If adoDemands.EOF Then
'      MsgBox "There are no demands to be exported into Sage", vbOKOnly + vbInformation, "No Demands"
'   Else
'      While Not adoDemands.EOF
''          new sales transaction record
''          Trap exceptions thrown by Sage Objects
''          Extract first Customer from collection
'         Set oCustomer = oApplication.SalesModule.Customers(UCase(adoDemands.Fields.Item("SageAccountNumber").Value))
'
''         Check transaction iKount
'         Set oSI_Instrument = oApplication.SalesModule.CreateInstrument(CreateSalesInvoice)
'         Set oSC_Instrument = oApplication.SalesModule.CreateInstrument(CreateSalesCreditNote)
'
'         If adoDemands.Fields.Item("TransactionType").Value = 1 Then    'INVOICE
'            If CDbl(adoDemands.Fields.Item("T_Amt").Value) > 0 Then
'               With oSI_Instrument
'                  .Customer = oCustomer
'                  .InstrumentNo = adoDemands!prefix & Format(adoDemands!DmdSlNo, "000000")
'                  .SecondReferenceNo = CStr(adoDemands.Fields.Item("SageRef").Value)
'
'                  .InstrumentDate = adoDemands.Fields.Item("ISSUEDATE").Value
'                  .PostedDate = CDate(Date)
'                  .DueDate = adoDemands.Fields.Item("ISSUEDATE").Value
'                  .NetValue = Format(CDbl(adoDemands.Fields.Item("T_Amt").Value), "0.00")
'                  .AllocatedValue = 0
'                  .TaxValue = Format(CDbl(adoDemands.Fields.Item("T_VAT").Value), "0.00")
''                   Nominal Analysis Item 1
'
'                  .NominalAnalysisItems(0).amount = Format(CDbl(adoDemands.Fields.Item("T_Amt").Value), "0.00")
'                  .NominalAnalysisItems(0).Narrative = CStr(adoDemands.Fields.Item("SageRef").Value)
'                  .NominalAnalysisItems(0).NominalSpecification.Reference = _
'                                          CStr(adoDemands.Fields.Item("NominalCodeforAmount").Value)
'                  .Update
'
''                 Mark Prestige DB as Exported to SAGE
'                  adoconn.Execute "UPDATE DemandRecords " & _
'                                  "SET spare3 = 'E' " & _
'                                  "WHERE DemandID = " & adoDemands.Fields.Item("DemandID").Value
'               End With
'            Else                                         'NEGATIVE INVOICE -> NI
'               adoconn.Execute "UPDATE DemandRecords " & _
'                               "SET spare3 = 'NI' " & _
'                               "WHERE DemandID = " & adoDemands.Fields.Item("DemandID").Value
'               iTrans = iTrans + 1
'            End If
'         End If
'
'         If adoDemands.Fields.Item("TransactionType").Value = 2 Then    'CREDIT NOTE
'            If CDbl(adoDemands.Fields.Item("T_Amt").Value) > 0 Then
'               With oSC_Instrument
'                  .Customer = oCustomer
'                  .InstrumentNo = adoDemands!prefix & Format(adoDemands!DmdSlNo, "000000")  'CStr(adoDemands.Fields.Item("SageRef").Value)
'                  .SecondReferenceNo = CStr(adoDemands.Fields.Item("SageRef").Value)
'
'                  .InstrumentDate = adoDemands.Fields.Item("ISSUEDATE").Value
'                  .PostedDate = CDate(Date)
'                  .DueDate = adoDemands.Fields.Item("ISSUEDATE").Value
'                  .NetValue = Format(CDbl(adoDemands.Fields.Item("T_Amt").Value), "0.00")
'                  .AllocatedValue = 0
'                  .TaxValue = Format(CDbl(adoDemands.Fields.Item("T_VAT").Value), "0.00")
''                   Nominal Analysis Item 1
'                  .NominalAnalysisItems(0).amount = Format(CDbl(adoDemands.Fields.Item("T_Amt").Value), "0.00")
'                  .NominalAnalysisItems(0).Narrative = CStr(adoDemands.Fields.Item("SageRef").Value)
'                  .NominalAnalysisItems(0).NominalSpecification.Reference = _
'                                          CStr(adoDemands.Fields.Item("NominalCodeforAmount").Value)
'                  .Update
'
''                 Mark Prestige DB as Exported to SAGE
'                  adoconn.Execute "UPDATE DemandRecords " & _
'                                  "SET spare3 = 'E' " & _
'                                  "WHERE DemandID = " & adoDemands.Fields.Item("DemandID").Value
'               End With
'            Else                                      'NEGATIVE CREDIT NOTE -> NC
'               adoconn.Execute "UPDATE DemandRecords " & _
'                               "SET spare3 = 'NC' " & _
'                               "WHERE DemandID = " & adoDemands.Fields.Item("DemandID").Value
'               iTrans = iTrans + 1
'            End If
'         End If
'
'         Set oSI_Instrument = Nothing
'         Set oSC_Instrument = Nothing
'
'         adoDemands.MoveNext
'         iExpTran = iExpTran + 1
'      Wend
'   End If
''*************************************************************************************************************
'
'   Dim rep As New frmReport
'
'   If iTrans > 0 Then
'      MsgBox "There are " & iTrans & " transaction(s) have not been posted.", vbCritical + vbOKOnly, "Exported"
''  Print Report
'      Set Report = reportApp.OpenReport(App.Path & szReportPath & "\ExpDmdList_Not.rpt")
'      Report.Database.Tables(1).ConnectionProperties.Item("Database Password") = accessDBPws
'
'      Report.EnableParameterPrompting = False
'      Report.DiscardSavedData
'
'      Load rep
'      rep.LoadReportViewer Report
'   End If
'
'   If iExpTran > 0 Then
'      MsgBox iExpTran & " Transactions have been posted successfully.", vbInformation + vbOKOnly, "Exported"
'
'      Set Report = reportApp.OpenReport(App.Path & szReportPath & "\ExpDmdList.rpt")
'
'      Report.Database.Tables(1).ConnectionProperties.Item("Database Password") = accessDBPws
'
'      Report.EnableParameterPrompting = False
'      Report.DiscardSavedData
'
'      Report.ParameterFields(1).AddCurrentValue "ExportedRpt"
'      Report.ParameterFields(2).AddCurrentValue "List of demands Exported Successfully"
'      Report.ParameterFields(3).AddCurrentValue "E"
'
'      Load rep
'      rep.LoadReportViewer Report
'   End If
'
'   fraConfirmation.Top = 4000
'
'   adoconn.Close
'   Set adoconn = Nothing
'   Me.MousePointer = vbArrow
'   frmMMain.MousePointer = vbArrow
'
'   Exit Sub
'CatchErr:
'   If ERR.Number = 94 Then
'      MsgBox "Please specify the name of your 3rd party software in company settings.", vbInformation + vbOKOnly, "Warning"
'   Else
'      MsgBox ERR.Number & ": " & ERR.description, vbCritical + vbOKOnly, "Warning.."
'   End If
'   adoconn.Close
'   Set adoconn = Nothing
'   Me.MousePointer = vbArrow
'   frmMMain.MousePointer = vbArrow
'End Sub
'
'Private Sub cmdClose_Click()
'   Unload Me
'End Sub
'
'Private Sub SavePostingNumber(tlbName As String, tblFieldName As String, lDemandID, lSPN As Long, adoconn As ADODB.Connection)
'   Dim SQLStr As String
'
'   'get the current password from the usernames table
'   If VarType(lDemandID) = vbString Then
'      SQLStr = "UPDATE " & tlbName & " " & _
'               "SET SAGEPostingNumber = " & lSPN & " " & _
'               "WHERE UPDATE_SAGE = FALSE AND " & _
'                     "" & tblFieldName & " = '" & lDemandID & "';"
'   Else
'      SQLStr = "UPDATE " & tlbName & " " & _
'               "SET SAGEPostingNumber = " & lSPN & " " & _
'               "WHERE UPDATE_SAGE = FALSE AND " & _
'                     "" & tblFieldName & " = " & lDemandID & ";"
'   End If
''Debug.Print SQLStr
'   adoconn.Execute SQLStr
'End Sub
'
'Private Sub cmdExport_Click()
'   Dim szChoice As String, szaChoice(1) As String
'
'   If cboSystem.text = "" Then
'      cboSystem.SetFocus
'      Exit Sub
'   End If
'   If cboTransactions.text = "" Then
'      cboTransactions.SetFocus
'      Exit Sub
'   End If
'   If Left(cboSystem.text, 1) = "-" Then Exit Sub
'
'   szaChoice(0) = cboSystem.ListIndex
'   szaChoice(1) = cboTransactions.ListIndex
'
'   szChoice = Join(szaChoice, "#")
'
'   SaveSetting "PropertyManagement", "ChoosedOption", "ExpMMS-c" & CStr(SCID), szChoice
'
'   fraConfirmation.Top = 0
'   fraConfirmation.Left = 40
'   fraConfirmation.Width = Me.Width - 180
'   fraConfirmation.Height = Me.Height - 520
'   lblClose.Left = fraConfirmation.Left + fraConfirmation.Width - 220
'
'   chkBackup.Value = 0
'   chkPreview.Value = 0
'   chkBackup.SetFocus
'End Sub
'
'Private Sub cmdLER_Click()
'   Dim reportApp As New CRAXDRT.Application
'   Dim Report As CRAXDRT.Report
'   Dim rep As New frmReport
'
'   On Error GoTo ErrHandler
'
'   Set Report = reportApp.OpenReport(App.Path & szReportPath & "\ExpDmdList.rpt")
'
'   Report.Database.Tables(1).ConnectionProperties.Item("Database Password") = accessDBPws
'
'   Report.EnableParameterPrompting = False
'   Report.DiscardSavedData
'
'   Report.ParameterFields(1).AddCurrentValue "ExportedRpt"
'   Report.ParameterFields(2).AddCurrentValue "List of demands Exported Successfully"
'   Report.ParameterFields(3).AddCurrentValue "E"
'
'   Load rep
'   rep.LoadReportViewer Report
'   Exit Sub
'
'ErrHandler:
'   ShowMsgInTaskBar "Please check the report path in the tools option", "Y", "N"
'End Sub
'
'Private Sub cmdPreview_Click()
'   If cboSystem.text = "" Then
'      cboSystem.SetFocus
'      Exit Sub
'   End If
'   If cboTransactions.text = "" Then
'      cboTransactions.SetFocus
'      Exit Sub
'   End If
'   If Left(cboSystem.text, 1) = "-" Then Exit Sub
'
'   If cboSystem.text = "SAGE 200" And cboTransactions.text = "Sales Invoices & Credit Notes" Then
'      Dim reportApp As New CRAXDRT.Application
'      Dim Report As CRAXDRT.Report
'
'      Set Report = reportApp.OpenReport(App.Path & szReportPath & "\PreVuExpDmdList.rpt")
'
'      Report.Database.Tables(1).ConnectionProperties.Item("Database Password") = accessDBPws
'
'      Report.EnableParameterPrompting = False
'      Report.DiscardSavedData
'
'      Dim rep As New frmReport
'
'      Report.ParameterFields(1).AddCurrentValue "XXX"
'      Report.ParameterFields(2).AddCurrentValue "List of demands for Export"
'
'      Load rep
'      rep.LoadReportViewer Report
'   End If
'
'   If (cboSystem.text = "SAGE Line 50 v14" Or cboSystem.text = "SAGE Line 50 v16") And cboTransactions.text = "Sales Invoices & Credit Notes" Then
'      ShowReport App.Path & szReportPath & "\PreViewExportedSAGE.rpt"
'   End If
'
'   Dim szChoice As String, szaChoice(1) As String
'
'   If cboSystem.text = "" Then
'      cboSystem.SetFocus
'      Exit Sub
'   End If
'   If cboTransactions.text = "" Then
'      cboTransactions.SetFocus
'      Exit Sub
'   End If
'   If Left(cboSystem.text, 1) = "-" Then Exit Sub
'
'   szaChoice(0) = cboSystem.ListIndex
'   szaChoice(1) = cboTransactions.ListIndex
'
'   szChoice = Join(szaChoice, "#")
'
'   SaveSetting "PropertyManagement", "ChoosedOption", "ExpMMS-c" & CStr(SCID), szChoice
'End Sub
'
'Private Sub cmdClear_Click()
'   cboSystem.ListIndex = -1
'   cboTransactions.ListIndex = -1
'
'   SaveSetting "PropertyManagement", "ChoosedOption", "ExpMMS-c" & CStr(SCID), ""
'End Sub
'
'Private Sub Form_Load()
'    On Error GoTo ERR
'   Dim szChoice As String, szaChoice() As String
'
'   Me.Height = 2610
'   Me.Width = 6165
'   Me.BackColor = MODULEBACKCOLOR
'   Frame1.BackColor = MODULEBACKCOLOR
'   fraConfirmation.BackColor = Me.BackColor
'   chkBackup.BackColor = Me.BackColor
'   chkPreview.BackColor = Me.BackColor
'   iTotalTran = 0
'
'   cboSystem.Clear
'
'   If Not ExpTrans3rdParty Then
'      cboSystem.AddItem "-SAGE MMS"
'      cboSystem.AddItem "-SAGE 200"
'      cboSystem.AddItem "-SAGE Line 50 v14"
'      cboSystem.AddItem "-SAGE Line 50 v16"
'   Else
'     ' cboSystem.AddItem "SAGE MMS"
'      cboSystem.AddItem "SAGE 200"
'      'cboSystem.AddItem "SAGE Line 50 v14"
'      cboSystem.AddItem "SAGE Line 50 v16"
'
'      szChoice = GetSetting("PropertyManagement", "ChoosedOption", "ExpMMS-c" & CStr(SCID))
'      szaChoice = Split(szChoice, "#")
'
'      If UBound(szaChoice) > 0 Then
'         cboSystem.ListIndex = szaChoice(0)
'         cboTransactions.ListIndex = szaChoice(1)
'      End If
'   End If
'   'delete all 2014 registry entry
'   'Or sageDirPath =sageDirPath, "C:\ProgramData\Sage\Accounts\2014"
'   'sample DeleteSetting "MyApp", "Startup"
''   sageDirPath = GetSetting("PropertyManagement", "SagePath", "SageFolder")
''   If InStr(sageDirPath, "2014") > 0 Then
''        DeleteSetting "PropertyManagement", "SageCompany"
''        DeleteSetting "PropertyManagement", "SagePath"
''   End If
'Exit Sub
'ERR:
'End Sub
'
'Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Me.MousePointer = vbArrow
'End Sub
'
'Private Sub Frame1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    Dim var As String
'    If Button = 2 And Shift = 1 Then
'        var = InputBox("Value")
'        If var = "xx" Then
'            Me.Height = 4755
'            txtPath.text = GetSetting("PropertyManagement", "SagePath", "SageFolder")
'            txtCompany.text = GetSetting("PropertyManagement", "Sagecompany", "Sagecompany")
'        Else
'             Me.Height = 2490
'        End If
'    End If
'End Sub
'
'Private Sub lblClose_Click()
'   fraConfirmation.Top = 4000
'End Sub
'
'Public Function ExecCmd(cmdline$)
'   Dim proc As PROCESS_INFORMATION
'   Dim start As STARTUPINFO
'   Dim ret&
'
'   ' Initialize the STARTUPINFO structure:
'   start.cB = Len(start)
'
'   ' Start the shelled application:
'   ret& = CreateProcessA(vbNullString, cmdline$, 0&, 0&, 1&, _
'      NORMAL_PRIORITY_CLASS, 0&, vbNullString, start, proc)
'
'   ' Wait for the shelled application to finish:
'      ret& = WaitForSingleObject(proc.hProcess, INFINITE)
'      Call GetExitCodeProcess(proc.hProcess, ret&)
'      Call CloseHandle(proc.hThread)
'      Call CloseHandle(proc.hProcess)
'      ExecCmd = ret&
'End Function
'
'' This method will check all invoices' outstanding amount in SAGE with the receipt amount in Prestige.
'' If receipt amount (Prestige) > outstanding amount (SAGE) then the receipt transaction will be marked as
'' not to exprot to SAGE. because sage does not allow it.
''
'' If the receipt transaction not prossible to update then it will be marked at "spare2" field,
'' in tlbReceipt table. If its possible to update then it will mark "S" (means success) into "spare2" field.
'' Otherwise it will mark by "F" (failed) + outstanding amount in SAGE of the invoice (ex. F100).
'
''RE-WRITE THE DEFINATION OF THE PROCESS OF THIS METHOD
'
'Private Sub JustifyOSAmtOfRptToExport(ByVal adoRpt As ADODB.Recordset, adoconn As ADODB.Connection, ByVal oWS As SageDataObject200.WorkSpace)
'   Dim szSQLStr As String, lInvoiceHeader As Long, bSucc As Boolean, j As Integer
'   Dim curOS  As Currency, bRef As Boolean
'   Dim oHeaderData As SageDataObject200.HeaderData
'   Dim oSplitData As SageDataObject200.SplitData
'   Dim adoRptCh As New ADODB.Recordset
'
'   adoRpt.MoveFirst
'
''   While Not adoRpt.EOF
'   szSQLStr = "SELECT DR.SageAccountNumber AS SAN, DR.SAGEPostingNumber AS SPN, " & _
'                  "T.TransactionID AS RT_TID, T.ToTran AS RT_INV, " & _
'                  "Rpt_1.TransactionID AS Rpt_TID, T.ReceiptAmount, " & _
'                  "T.AllocDate, DR.Details, Rpt_1.Ref, Rpt_1.Type " & _
'              "FROM ((RptTransactions AS T INNER JOIN tlbReceipt AS Rpt_1 ON T.FromTran = " & _
'                  "Rpt_1.TransactionID) INNER JOIN tlbReceipt AS Rpt_2 ON T.ToTran = " & _
'                  "Rpt_2.TransactionID) INNER JOIN DemandRecords AS DR ON Rpt_2.DemandRef = DR.DemandID " & _
'              "WHERE NOT T.UpdateSage And T.IsSageUpdate " & _
'              "GROUP BY DR.SageAccountNumber, DR.SAGEPostingNumber, T.TransactionID, " & _
'                  "T.ToTran, Rpt_1.TransactionID, T.ReceiptAmount, " & _
'                  "T.AllocDate, DR.Details, Rpt_1.Ref, Rpt_1.Type;"
'
''Debug.Print szSQLStr
'   adoRptCh.Open szSQLStr, adoconn, adOpenStatic, adLockReadOnly
'
'   While Not adoRptCh.EOF
'      If adoRptCh!Type <> CByte(sdoSA) Then
'         Set oHeaderData = oWS.CreateObject("HeaderData")
'         lInvoiceHeader = CLng(adoRptCh!SPN)             'SPN -> SGAE POSTING NUMBER
'         oHeaderData.Read (lInvoiceHeader)
'         bSucc = False
'         curOS = -1
'
'         curOS = CCur(oHeaderData.Fields.Item("NET_AMOUNT").Value - _
'                                         oHeaderData.Fields.Item("AMOUNT_PAID").Value)
'
'         If UCase(oHeaderData.Fields.Item("DETAILS").Value) <> "OPENING BALANCE" Then
'            bRef = True
'            bSucc = IIf(CCur(adoRptCh!receiptAmount) <= curOS, True, False)
'         Else
'            bRef = False
'            bSucc = False
'         End If
'
'         Call SetMark(bRef, bSucc, curOS, adoRptCh!RT_TID, adoconn)
'      Else
'         Call SetMark(True, True, 0, adoRptCh!RT_TID, adoconn)
'      End If
'      adoRptCh.MoveNext
'   Wend
'
'  ' Destroy the Objects
'   adoRptCh.Close
'   Set adoRptCh = Nothing
'   Set oSplitData = Nothing
'   Set oHeaderData = Nothing
'End Sub
'
'Private Sub SetMark(ByVal bRef As Boolean, ByVal bSucc As Boolean, ByVal curOS As Currency, ByVal lTranID As Long, adoconn As ADODB.Connection)
'   Dim szSQLStr As String, szMark As String
'   Dim adoRptCh As New ADODB.Recordset
'
'   If bSucc Then
'      szMark = "S"                  'S -> SUCCESS
'   Else
'      If bRef Then
'         szMark = "F" & CStr(curOS) 'F -> FAIL, eg.  szMark = F100
'      Else
'         szMark = "F-1"             'Not possible to export and not found the OS AMT
'      End If
'   End If
'
'   szSQLStr = "SELECT UpdateSage, IsSageUpdate, Exp2Sage  " & _
'              "FROM RptTransactions AS RT " & _
'              "WHERE RT.TransactionID = " & lTranID & ";"
'
'   adoRptCh.Open szSQLStr, adoconn, adOpenDynamic, adLockPessimistic
'
'   adoRptCh.Fields.Item("UpdateSage").Value = IIf(szMark = "S", False, True)
'   adoRptCh.Fields.Item("Exp2Sage").Value = szMark
'   adoRptCh.Update
'
'   adoRptCh.Close
'   Set adoRptCh = Nothing
'
'   iTotalTran = iTotalTran + 1
'End Sub
'
''THIS METHOD UPDATE THE SELECTED RECORDS OF RptReceipt
'Private Sub UpdateMarked(lTransID() As Long, iCount As Integer, adoconn As ADODB.Connection, szTable As String)
'   If iCount = 0 Then Exit Sub
'
'   Dim i As Integer
'   Dim SQLStr As String
'   Dim adoRpt As New ADODB.Recordset
'
'   SQLStr = "TransactionID = " & lTransID(0)
'
'   For i = 1 To iCount - 1
'      SQLStr = SQLStr & " or " & "TransactionID = " & lTransID(i)
'   Next i
'
'   SQLStr = "SELECT UpdateSage, IsSageUpdate " & _
'            "FROM " & szTable & " " & _
'            "WHERE " & SQLStr
''Debug.Print SQLStr
'   adoRpt.Open SQLStr, adoconn, adOpenDynamic, adLockPessimistic
'
'   While Not adoRpt.EOF
'      adoRpt.Fields.Item("UpdateSage").Value = True
'      adoRpt.Fields.Item("IsSageUpdate").Value = False
'      adoRpt.Update
'      adoRpt.MoveNext
'   Wend
'
'   adoRpt.Close
'   Set adoRpt = Nothing
'End Sub
'
'Private Function SagePostingNumberCr(szTransactionID As String, adoconn As ADODB.Connection) As Long
'   Dim szStr As String
'   Dim adoRst As New ADODB.Recordset
'
'   szStr = "SELECT SAGEPostingNumber FROM DemandRecords " & _
'           "WHERE DemandID = " & CLng(szTransactionID) & ";"
'
'   adoRst.Open szStr, adoconn, adOpenStatic, adLockReadOnly
'
'   SagePostingNumberCr = adoRst.Fields.Item("SAGEPostingNumber").Value
'
'   adoRst.Close
'   Set adoRst = Nothing
'End Function
'
'Private Function SagePostingNumberInv(szTransactionID As String, adoconn As ADODB.Connection) As Long
'   Dim szStr As String
'   Dim adoRst As New ADODB.Recordset
'
'   szStr = "SELECT DemandRecords.SAGEPostingNumber " & _
'           "FROM DemandRecords, tlbReceipt " & _
'           "WHERE tlbReceipt.TransactionID = " & CLng(szTransactionID) & " AND " & _
'               "tlbReceipt.DemandRef = DemandRecords.DemandID;"
''Debug.Print szStr
'   adoRst.Open szStr, adoconn, adOpenStatic, adLockReadOnly
'
'   SagePostingNumberInv = adoRst.Fields.Item("SAGEPostingNumber").Value
'
'   adoRst.Close
'   Set adoRst = Nothing
'End Function
'
'Private Sub SavePostingNumberPoA(tlbName As String, tblFieldName As String, lDemandID, lSPN As Long, adoconn As ADODB.Connection)
'   Dim SQLStr As String
'
'   SQLStr = "UPDATE " & tlbName & " " & _
'            "SET PoA_SPN = " & lSPN & " " & _
'            "WHERE NOT UpDateSage AND " & tblFieldName & " = " & lDemandID & ";"
''Debug.Print SQLStr
'   adoconn.Execute SQLStr
'End Sub
