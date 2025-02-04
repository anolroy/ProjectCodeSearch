VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmCopyTransaction 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Copy Transaction"
   ClientHeight    =   9135
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9360
   BeginProperty Font 
      Name            =   "Myriad Web"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCopyTransaction.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9135
   ScaleWidth      =   9360
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdSetAmtType 
      Caption         =   "+"
      Height          =   315
      Left            =   3840
      TabIndex        =   2
      Top             =   1493
      Width           =   315
   End
   Begin VB.TextBox txtAmount 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1155
      TabIndex        =   5
      Top             =   2640
      Width           =   1575
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00F0F0F0&
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   315
      Left            =   3090
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3000
      Width           =   1080
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00F0F0F0&
      Caption         =   "&Save"
      Height          =   315
      Left            =   1155
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3000
      Width           =   1080
   End
   Begin VB.TextBox txtSPDate 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000014&
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   1155
      MaxLength       =   10
      TabIndex        =   4
      Top             =   2277
      Width           =   1575
   End
   Begin VB.TextBox txtReceiptReference 
      Appearance      =   0  'Flat
      BackColor       =   &H80000014&
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   1155
      MaxLength       =   20
      TabIndex        =   0
      Top             =   1131
      Width           =   3015
   End
   Begin MSForms.ComboBox cmbFund 
      Height          =   315
      Left            =   1155
      TabIndex        =   3
      Top             =   1890
      Width           =   3045
      VariousPropertyBits=   1753237531
      DisplayStyle    =   3
      Size            =   "5371;556"
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
      Object.Width           =   "703"
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fund:"
      Height          =   195
      Index           =   0
      Left            =   75
      TabIndex        =   18
      Top             =   1890
      Width           =   390
   End
   Begin MSForms.TextBox txtProperty 
      Height          =   285
      Left            =   1155
      TabIndex        =   17
      Top             =   407
      Width           =   3015
      VariousPropertyBits=   679495711
      BackColor       =   15858158
      BorderStyle     =   1
      Size            =   "5318;503"
      SpecialEffect   =   0
      FontName        =   "Myriad Web"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox txtBankAC 
      Height          =   285
      Left            =   1155
      TabIndex        =   16
      Top             =   769
      Width           =   3015
      VariousPropertyBits=   679495711
      BackColor       =   15858158
      BorderStyle     =   1
      Size            =   "5318;503"
      SpecialEffect   =   0
      FontName        =   "Myriad Web"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bank A/C:"
      Height          =   195
      Index           =   6
      Left            =   75
      TabIndex        =   14
      Top             =   769
      Width           =   660
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Lessee:"
      Height          =   195
      Index           =   0
      Left            =   75
      TabIndex        =   13
      Top             =   45
      Width           =   525
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Amount:"
      Height          =   195
      Index           =   6
      Left            =   75
      TabIndex        =   12
      Top             =   2640
      Width           =   585
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date:"
      Height          =   195
      Index           =   7
      Left            =   75
      TabIndex        =   11
      Top             =   2277
      Width           =   375
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Reference:"
      Height          =   195
      Index           =   9
      Left            =   75
      TabIndex        =   10
      Top             =   1131
      Width           =   750
   End
   Begin VB.Label Label19 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Property:"
      Height          =   195
      Index           =   1
      Left            =   75
      TabIndex        =   9
      Top             =   407
      Width           =   645
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Amount Type:"
      Height          =   195
      Index           =   11
      Left            =   75
      TabIndex        =   8
      Top             =   1493
      Width           =   960
   End
   Begin MSForms.ComboBox cmbRptAmtType 
      Height          =   315
      Left            =   1155
      TabIndex        =   1
      Top             =   1493
      Width           =   2685
      VariousPropertyBits=   1753237531
      DisplayStyle    =   3
      Size            =   "4736;556"
      TextColumn      =   2
      ColumnCount     =   3
      ListRows        =   20
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      SpecialEffect   =   6
      FontName        =   "Myriad Web"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox txtTenantID 
      Height          =   285
      Left            =   1155
      TabIndex        =   15
      Top             =   45
      Width           =   3015
      VariousPropertyBits=   679495711
      BackColor       =   15858158
      BorderStyle     =   1
      Size            =   "5318;503"
      SpecialEffect   =   0
      FontName        =   "Myriad Web"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
End
Attribute VB_Name = "frmCopyTransaction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public CALLING_FORM  As String

Private Sub cmdClose_Click()
   Unload Me
End Sub

Private Sub cmdSave_Click()
   If IsNull(cmbFund.Value) Then
      MsgBox "Please select a fund.", vbInformation + vbOKOnly, "Copy Transaction"
      cmbFund.SetFocus
      Exit Sub
   End If

   If Right(CALLING_FORM, 16) = "_RECEIPT_REVERSE" Or _
         Right(CALLING_FORM, 11) = "_SA_REVERSE" Or _
         Right(CALLING_FORM, 4) = "_SRR" Then
      If MsgBox("Do you wish to create Sales Receipt Refund?", vbQuestion + vbYesNo, "Copy/Copy Reverse") = vbNo Then
         Exit Sub
      End If
   End If
   If CALLING_FORM = "DEMANDRECEIPT" Or _
         Right(CALLING_FORM, 3) = "_SR" Or _
         Right(CALLING_FORM, 3) = "_SA" Or _
         Right(CALLING_FORM, 12) = "_SRR_REVERSE" Then
      If MsgBox("Do you wish to create Sales Receipt on Account?", vbQuestion + vbYesNo, "Copy/Copy Reverse") = vbNo Then
         Exit Sub
      End If
   End If
   If CALLING_FORM = "PURCHASE_PAYMENT" Or _
         CALLING_FORM = "PURCHASE_PAYMENT_ACCOUNT" Or _
         CALLING_FORM = "PAYMENT_REFUND_REVERSE" Then
      If MsgBox("Do you wish to create Purchase Payment on Account?", vbQuestion + vbYesNo, "Copy Transaction") = vbNo Then
         Exit Sub
      End If
   End If
   If CALLING_FORM = "PURCHASE_PAYMENT_REFUND" Or _
         CALLING_FORM = "PAYMENT_ACCOUNT_REVERSE" Or _
         CALLING_FORM = "PURCHASE_PAYMENT_REVERSE" Then
      If MsgBox("Do you wish to create Purchase Payment Refund?", vbQuestion + vbYesNo, "Copy Reverse") = vbNo Then
         Exit Sub
      End If
   End If

   Dim adoConn    As New ADODB.Connection
   Dim adoRst     As New ADODB.Recordset
   Dim szSQL      As String
   Dim lRpt_ID    As Long
   Dim lSlNumber  As Long

   adoConn.Open getConnectionString

   If InStr(CALLING_FORM, "DEMAND") > 0 Then
      szSQL = "SELECT MAX(TransactionID) AS TID FROM tlbReceipt;"
      adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
      lRpt_ID = CLng(IIf(IsNull(adoRst!TID), 0, adoRst!TID)) + 1
      adoRst.Close

      szSQL = "SELECT * FROM tlbReceipt;"
      adoRst.Open szSQL, adoConn, adOpenDynamic, adLockOptimistic

      If CALLING_FORM = "DEMANDRECEIPT" Or _
            Right(CALLING_FORM, 3) = "_SR" Or _
            Right(CALLING_FORM, 3) = "_SA" Or _
            Right(CALLING_FORM, 12) = "_SRR_REVERSE" Then
         lSlNumber = SlNumber("SA", "tlbReceipt", adoConn)

         With adoRst
            .AddNew
            !TransactionID = lRpt_ID
            !Type = 4                                             'CByte(sdoSA)   Sales Receipt on Account
            !SageAccountNumber = txtTenantID.text
            !unitid = GetUnitIDbyTenantID(txtTenantID.text, adoConn)
            !RDate = Format(txtSPDate.text, "dd mmmm yyyy")
            !dDate = Format(txtSPDate.text, "dd mmmm yyyy")
            !ref = "SA" & Format(Now, "yymmddhhmmss")
            !Details = "Receipt on Account"
            !amount = txtAmount.text
            !OSAmount = !amount           'amount to be allocated
            !ReceiptView = True
            !BankCode = txtBankAC.text
            !nominalCode = !BankCode
            !ExtRef = txtReceiptReference.text
            !RptAmtType = cmbRptAmtType.Value
            !SlNumber = lSlNumber
            !fundID = cmbFund.Value

            .Update
            .Close
         End With
      End If

      If Right(CALLING_FORM, 16) = "_RECEIPT_REVERSE" Or _
            Right(CALLING_FORM, 11) = "_SA_REVERSE" Or _
            Right(CALLING_FORM, 4) = "_SRR" Then
         lSlNumber = SlNumber("SRR", "tlbReceipt", adoConn)

         With adoRst
            .AddNew
            !TransactionID = lRpt_ID
            !Type = 23
            !SageAccountNumber = txtTenantID.text
            !unitid = GetUnitIDbyTenantID(txtTenantID.text, adoConn)
            !RDate = Format(txtSPDate.text, "dd mmmm yyyy")
            !dDate = Format(txtSPDate.text, "dd mmmm yyyy")
            !ref = "SRR" & Format(Now, "yymmddhhmmss")
            !Details = "Sales Receipt Refund"
            !amount = txtAmount.text
            !OSAmount = !amount
            !ReceiptView = True
            !BankCode = txtBankAC.text
            !nominalCode = !BankCode
            !ExtRef = txtReceiptReference.text
            !RptAmtType = cmbRptAmtType.Value
            !SlNumber = lSlNumber
            !fundID = cmbFund.Value

            .Update
            .Close
         End With

   '     Saving the split(s) of the header
         szSQL = "SELECT * FROM tlbReceiptSplit;"
         adoRst.Open szSQL, adoConn, adOpenDynamic, adLockPessimistic

         With adoRst
            .AddNew
            .Fields.Item("TransactionID").Value = UniqueID()
            .Fields.Item("RptHeader").Value = lRpt_ID
            .Fields.Item("FundID").Value = cmbFund.Value
            .Fields.Item("Amount").Value = !amount
            .Fields.Item("SplitID").Value = 1
            .Fields.Item("DueDate").Value = Format(txtSPDate.text, "dd mmmm yyyy")
            .Fields.Item("Description").Value = "Sales Receipt Refund"
            .Update
            .Close
         End With
      End If
   End If

   If InStr(CALLING_FORM, "PURCHASE") > 0 Then
      szSQL = "SELECT MAX(TransactionID) AS TID FROM tlbPayment;"
      adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
      lRpt_ID = CLng(IIf(IsNull(adoRst!TID), 0, adoRst!TID)) + 1
      adoRst.Close

      szSQL = "SELECT * FROM tlbPayment;"
      adoRst.Open szSQL, adoConn, adOpenDynamic, adLockOptimistic

      If CALLING_FORM = "PURCHASE_PAYMENT" Or _
            CALLING_FORM = "PAYMENT_REFUND_REVERSE" Or _
            CALLING_FORM = "PURCHASE_PAYMENT_ACCOUNT" Then
         lSlNumber = SlNumber("PA", "tlbPayment", adoConn)

         With adoRst
            .AddNew
            !TransactionID = lRpt_ID
            !Type = 9                                 'CByte(sdoPA)  Purchase Payment on Account
            !SageAccountNumber = txtTenantID.text
            !PDate = Format(txtSPDate.text, "dd mmmm yyyy")
            !ref = "PA" & Format(Now, "yymmddhhmmss")
            !Details = "PAYMENT ON ACCOUNT"
            !amount = txtAmount.text
            !OSAmount = !amount           'amount to be allocated
            !PaymentView = True
            !BankCode = txtBankAC.text
            !nominalCode = !BankCode
            !ExtRef = txtReceiptReference.text
            !PayAmtType = cmbRptAmtType.Value
            !SlNumber = lSlNumber
            !fundID = cmbFund.Value
            .Update
            .Close
         End With
      End If

      If CALLING_FORM = "PURCHASE_PAYMENT_REFUND" Or _
            CALLING_FORM = "PAYMENT_ACCOUNT_REVERSE" Or _
            CALLING_FORM = "PURCHASE_PAYMENT_REVERSE" Then
         lSlNumber = SlNumber("PPR", "tlbPayment", adoConn)

         With adoRst
            .AddNew
            !TransactionID = lRpt_ID
            !Type = 24                                 'CByte(sdoPA)  Purchase Payment on Account
            !SageAccountNumber = txtTenantID.text
            !PDate = Format(txtSPDate.text, "dd mmmm yyyy")
            !ref = "PPR" & Format(Now, "yymmddhhmmss")
            !Details = "Purchase Payment Refund"
            !amount = txtAmount.text
            !OSAmount = !amount           'amount to be allocated
            !PaymentView = True
            !BankCode = txtBankAC.text
            !nominalCode = !BankCode
            !ExtRef = txtReceiptReference.text
            !PayAmtType = cmbRptAmtType.Value
            !SlNumber = lSlNumber
            !fundID = cmbFund.Value
            .Update
            .Close
         End With
      End If
   End If

   Set adoRst = Nothing
   adoConn.Close
   Set adoConn = Nothing

   Unload Me
   ShowMsgInTaskBar "Transaction has been created sucessfully.", "Y", "P"
End Sub

Private Sub Form_Activate()
   If InStr(CALLING_FORM, "DEMAND") > 0 Or _
         InStr(CALLING_FORM, "LESSEE") > 0 Or _
         Right(CALLING_FORM, 3) = "_SR" Or _
         Right(CALLING_FORM, 3) = "_SA" Then
      If CALLING_FORM = "DEMANDRECEIPT" Or _
         Right(CALLING_FORM, 3) = "_SR" Then Me.Caption = "Copy Transaction - Sales Receipt"
      If Right(CALLING_FORM, 4) = "_SRR" Then Me.Caption = "Copy Transaction - Sales Refund"
      If Right(CALLING_FORM, 3) = "_SA" Then Me.Caption = "Copy - Receipt on Account"

      If Right(CALLING_FORM, 16) = "_RECEIPT_REVERSE" Then Me.Caption = "Copy Reverse - Sales Receipt"
      If Right(CALLING_FORM, 11) = "_SA_REVERSE" Then Me.Caption = "Copy Reverse - Receipt on Account"
      If Right(CALLING_FORM, 12) = "_SRR_REVERSE" Then Me.Caption = "Copy Reverse - Sales Refund"

      LoadReceipt
   End If
   If InStr(CALLING_FORM, "PAYMENT") > 0 Then
      If CALLING_FORM = "PURCHASE_PAYMENT" Then Me.Caption = "Copy Transaction - Purchase Payment"
      If CALLING_FORM = "PURCHASE_PAYMENT_PPR" Then Me.Caption = "Copy Transaction - Payment Refund"
      If CALLING_FORM = "PURCHASE_PAYMENT_ACCOUNT" Then Me.Caption = "Copy - Payment on Account"

      If CALLING_FORM = "PURCHASE_PAYMENT_REVERSE" Then Me.Caption = "Copy Reverse - Purchase Payment"
      If CALLING_FORM = "PAYMENT_ACCOUNT_REVERSE" Then Me.Caption = "Copy Reverse - Payment on Account"
      If CALLING_FORM = "PAYMENT_REFUND_REVERSE" Then Me.Caption = "Copy Reverse - Payment Refund"

      LoadPayment
   End If
End Sub

Private Sub Form_Load()
   Me.Height = 3870
   Me.Width = 4395
   Me.BackColor = MODULEBACKCOLOR
End Sub

Private Sub LoadFund(adoConn As ADODB.Connection)
   Dim adoRst  As New ADODB.Recordset
   Dim szSQL   As String
   Dim i       As Integer

   szSQL = "SELECT FundID, FundName FROM FUND;"

   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   ReDim szaData(1, adoRst.RecordCount - 1) As String

   i = 0
   While Not adoRst.EOF
      szaData(0, i) = adoRst.Fields.Item("FundID").Value
      szaData(1, i) = adoRst.Fields.Item("FundName").Value
      i = i + 1
      adoRst.MoveNext
   Wend

   cmbFund.Clear
   cmbFund.Column() = szaData()

   adoRst.Close
   Set adoRst = Nothing
End Sub

Private Sub LoadReceipt()
   Dim adoConn As New ADODB.Connection
   Dim adoRst  As New ADODB.Recordset
   Dim szSQL   As String
   Dim btTT    As Byte

   If CALLING_FORM = "DEMANDRECEIPT" Or _
      Right(CALLING_FORM, 16) = "_RECEIPT_REVERSE" Or _
      Right(CALLING_FORM, 3) = "_SR" Then btTT = 3

   If Right(CALLING_FORM, 4) = "_SRR" Or _
      Right(CALLING_FORM, 12) = "_SRR_REVERSE" Then btTT = 23

   If Right(CALLING_FORM, 11) = "_SA_REVERSE" Or _
      Right(CALLING_FORM, 3) = "_SA" Then btTT = 4

   adoConn.Open getConnectionString

   If Left(CALLING_FORM, 6) = "DEMAND" Then
      szSQL = "SELECT * FROM tlbReceipt " & _
              "WHERE Type = " & btTT & " AND " & _
                  "SlNumber = " & StrDigitVal(frmDemands3.flxReceiptHistory.TextMatrix(frmDemands3.flxReceiptHistory.row, 0)) & ";"
   End If
   If Left(CALLING_FORM, 6) = "LESSEE" Then
      szSQL = "SELECT * FROM tlbReceipt " & _
              "WHERE Type = " & btTT & " AND " & _
                  "SlNumber = " & StrDigitVal(frmLeasee1.flxACHistory.TextMatrix(frmLeasee1.flxACHistory.row, 1)) & ";"
   End If
   If Left(CALLING_FORM, 3) = "CB_" Then
      szSQL = "SELECT * FROM tlbReceipt " & _
              "WHERE Type = " & btTT & " AND " & _
                  "SlNumber = " & StrDigitVal(frmCashbook.flxCashBook.TextMatrix(frmCashbook.flxCashBook.row, 2)) & ";"
   End If

   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If Not adoRst.EOF Then
      With adoRst.Fields
         txtSPDate.text = .Item("RDate").Value
         txtAmount.text = Format(.Item("Amount").Value, "0.00")
         txtReceiptReference.text = .Item("ExtRef").Value
         txtTenantID.text = .Item("SageAccountNumber").Value
         txtBankAC.text = .Item("BankCode").Value
         txtProperty.text = GetPropByLessee(txtTenantID.text, adoConn)
         LoadRptAmtType "RECEIPT AMOUNT TYPE", adoConn
         LoadFund adoConn
         cmbRptAmtType.Value = .Item("RptAmtType").Value
      End With
   End If

   adoRst.Close
   Set adoRst = Nothing
   adoConn.Close
   Set adoConn = Nothing
End Sub

Private Sub LoadPayment()
   Dim adoConn As New ADODB.Connection
   Dim adoRst  As New ADODB.Recordset
   Dim szSQL   As String
   Dim btTT    As Byte

   If CALLING_FORM = "PURCHASE_PAYMENT" Or CALLING_FORM = "PURCHASE_PAYMENT_REVERSE" Then btTT = 8
   If CALLING_FORM = "PURCHASE_PAYMENT_PPR" Or CALLING_FORM = "PAYMENT_REFUND_REVERSE" Then btTT = 24
   If CALLING_FORM = "PURCHASE_PAYMENT_ACCOUNT" Or CALLING_FORM = "PAYMENT_ACCOUNT_REVERSE" Then btTT = 9

   adoConn.Open getConnectionString

   szSQL = "SELECT * FROM tlbPayment " & _
           "WHERE Type = " & btTT & " AND " & _
               "TransactionID = " & frmPurchaseExpense.flxPurchPPHistory.TextMatrix(frmPurchaseExpense.flxPurchPPHistory.row, 0) & ";"

   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If Not adoRst.EOF Then
      With adoRst.Fields
         txtSPDate.text = .Item("PDate").Value
         txtAmount.text = Format(.Item("Amount").Value, "0.00")
         txtReceiptReference.text = .Item("ExtRef").Value
         txtTenantID.text = .Item("SageAccountNumber").Value
         txtBankAC.text = .Item("BankCode").Value
         txtProperty.text = IIf(IsNull(.Item("UnitID").Value), "", .Item("UnitID").Value)
         LoadRptAmtType "PAYMENT AMOUNT TYPE", adoConn
         LoadFund adoConn
         cmbRptAmtType.Value = .Item("PayAmtType").Value
      End With
   End If

   adoRst.Close
   Set adoRst = Nothing
   adoConn.Close
   Set adoConn = Nothing
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Me.MousePointer = vbArrow
End Sub

Private Sub txtSPDate_Change()
   TextBoxChangeDate txtSPDate
End Sub

Private Sub txtSPDate_GotFocus()
   SelTxtInCtrl txtSPDate
End Sub

Private Sub txtSPDate_KeyPress(KeyAscii As Integer)
   TextBoxKeyPrsDate txtSPDate, KeyAscii
End Sub

Private Sub txtSPDate_LostFocus()
   TextBoxFormatDate txtSPDate
End Sub

Private Sub txtAmount_GotFocus()
   SelTxtInCtrl txtAmount
End Sub

Private Sub txtAmount_KeyPress(KeyAscii As Integer)
   DigitTextKeyPress txtAmount, KeyAscii
End Sub

Private Sub txtAmount_LostFocus()
   txtAmount.text = Format(Val(txtAmount.text), "0.00")
End Sub

Private Sub LoadRptAmtType(szValue As String, adoConn As ADODB.Connection)
   Dim SQLStr1 As String, szaData() As String, i As Integer
   Dim adoRst As New ADODB.Recordset

   SQLStr1 = "SELECT SecondaryCode.Code as C, SecondaryCode.Value as V " & _
             "FROM PrimaryCode, SecondaryCode " & _
             "WHERE PrimaryCode.Value = '" & szValue & "' AND " & _
                  "PrimaryCode.CODE = SecondaryCode.PrimaryCode " & _
             "ORDER BY SecondaryCode.Value;"

   adoRst.Open SQLStr1, adoConn, adOpenStatic, adLockReadOnly

   If adoRst.EOF Then
      adoRst.Close
      Set adoRst = Nothing
      Exit Sub
   End If

   ReDim szaData(1, adoRst.RecordCount - 1) As String

   cmbRptAmtType.Clear
   i = 0
   While Not adoRst.EOF
      szaData(0, i) = adoRst!c
      szaData(1, i) = adoRst!V
      adoRst.MoveNext
      i = i + 1
   Wend
   adoRst.Close
   Set adoRst = Nothing

   cmbRptAmtType.Column() = szaData()
End Sub
