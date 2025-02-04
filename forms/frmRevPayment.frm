VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmRevPayment 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reverse Payment"
   ClientHeight    =   8010
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15195
   BeginProperty Font 
      Name            =   "Myriad Web"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRevPayment.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8010
   ScaleWidth      =   15195
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtPayDt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00DAEADA&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   10845
      TabIndex        =   29
      Top             =   4770
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdFix 
      Caption         =   "&Fix"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   13365
      TabIndex        =   27
      Top             =   7560
      Visible         =   0   'False
      Width           =   1700
   End
   Begin VB.CommandButton cmdAmendAllocationDate 
      Caption         =   "&Apply New  Allocation Date"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   9135
      TabIndex        =   1
      Top             =   6885
      Width           =   2460
   End
   Begin VB.CheckBox chkCredits 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   135
      TabIndex        =   26
      Top             =   315
      Width           =   195
   End
   Begin VB.CommandButton cmdSPClose 
      Caption         =   "C&lose"
      Height          =   400
      Left            =   13635
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6930
      Width           =   1400
   End
   Begin VB.CommandButton cmeRevereseAllocation 
      Caption         =   "&Reverse Allocation"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   165
      TabIndex        =   0
      Top             =   7020
      Width           =   1700
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxSPayment 
      Height          =   2415
      Left            =   120
      TabIndex        =   3
      Top             =   4380
      Width           =   15000
      _ExtentX        =   26458
      _ExtentY        =   4260
      _Version        =   393216
      FixedCols       =   0
      BackColorFixed  =   12632256
      ForeColorFixed  =   -2147483640
      BackColorSel    =   15329508
      ForeColorSel    =   -2147483640
      BackColorBkg    =   16777215
      GridColor       =   -2147483638
      GridColorFixed  =   -2147483630
      WordWrap        =   -1  'True
      GridLinesFixed  =   1
      Appearance      =   0
      BandDisplay     =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxCrPoA 
      Height          =   2985
      Left            =   120
      TabIndex        =   4
      Top             =   540
      Width           =   15000
      _ExtentX        =   26458
      _ExtentY        =   5265
      _Version        =   393216
      FixedCols       =   0
      BackColorFixed  =   12632256
      ForeColorFixed  =   -2147483640
      BackColorSel    =   15329508
      ForeColorSel    =   -2147483640
      BackColorBkg    =   16777215
      GridColor       =   -2147483638
      GridColorFixed  =   -2147483630
      WordWrap        =   -1  'True
      GridLinesFixed  =   1
      SelectionMode   =   1
      Appearance      =   0
      BandDisplay     =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Allocation Date"
      Height          =   195
      Index           =   0
      Left            =   10350
      TabIndex        =   28
      Top             =   4095
      Width           =   1095
   End
   Begin VB.Line Line2 
      X1              =   45
      X2              =   14760
      Y1              =   4005
      Y2              =   4005
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00008000&
      Caption         =   "  Reverse Allocation View  "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   330
      Index           =   2
      Left            =   4815
      TabIndex        =   6
      Top             =   3600
      Width           =   3300
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Credit:"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400040&
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   25
      Top             =   3765
      Width           =   510
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   2
      X1              =   675
      X2              =   13060
      Y1              =   105
      Y2              =   105
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Debit:"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400040&
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   24
      Top             =   30
      Width           =   465
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "allocating row no"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400040&
      Height          =   195
      Index           =   3
      Left            =   12015
      TabIndex        =   23
      Top             =   3690
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "credit row no"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400040&
      Height          =   195
      Index           =   4
      Left            =   12120
      TabIndex        =   22
      Top             =   90
      Visible         =   0   'False
      Width           =   990
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "No."
      Height          =   195
      Index           =   10
      Left            =   300
      TabIndex        =   21
      Top             =   4050
      Width           =   240
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Type"
      Height          =   195
      Index           =   11
      Left            =   1080
      TabIndex        =   20
      Top             =   4050
      Width           =   345
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Unit ID"
      Height          =   195
      Index           =   12
      Left            =   3420
      TabIndex        =   19
      Top             =   4050
      Width           =   495
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Due Date"
      Height          =   195
      Index           =   13
      Left            =   4320
      TabIndex        =   18
      Top             =   4050
      Width           =   675
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ref"
      Height          =   195
      Index           =   14
      Left            =   5460
      TabIndex        =   17
      Top             =   4050
      Width           =   225
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Details"
      Height          =   195
      Index           =   15
      Left            =   6600
      TabIndex        =   16
      Top             =   4050
      Width           =   510
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Amount £"
      Height          =   195
      Index           =   16
      Left            =   11820
      TabIndex        =   15
      Top             =   4095
      Width           =   675
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "O/S Amt. £"
      Height          =   195
      Index           =   17
      Left            =   12780
      TabIndex        =   14
      Top             =   4095
      Width           =   735
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Amount £"
      Height          =   195
      Index           =   26
      Left            =   13995
      TabIndex        =   13
      Top             =   315
      Width           =   675
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Details"
      Height          =   195
      Index           =   25
      Left            =   8160
      TabIndex        =   12
      Top             =   330
      Width           =   510
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ref"
      Height          =   195
      Index           =   24
      Left            =   6360
      TabIndex        =   11
      Top             =   330
      Width           =   225
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      Height          =   195
      Index           =   23
      Left            =   4800
      TabIndex        =   10
      Top             =   330
      Width           =   345
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Unit ID"
      Height          =   195
      Index           =   22
      Left            =   3180
      TabIndex        =   9
      Top             =   330
      Width           =   495
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Type"
      Height          =   195
      Index           =   21
      Left            =   1320
      TabIndex        =   8
      Top             =   330
      Width           =   345
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "No."
      Height          =   195
      Index           =   20
      Left            =   390
      TabIndex        =   7
      Top             =   330
      Width           =   240
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Receipt Amt £"
      Height          =   195
      Index           =   18
      Left            =   13980
      TabIndex        =   5
      Top             =   4095
      Width           =   1005
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0FFC0&
      BorderWidth     =   2
      Index           =   3
      X1              =   675
      X2              =   13060
      Y1              =   120
      Y2              =   120
   End
End
Attribute VB_Name = "frmRevPayment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private szSupplierID As String
Dim iCurRow As Integer
Private Sub chkCredits_Click()
        Dim iCount As Integer
        If chkCredits = 1 Then
            For iCount = 1 To flxCrPoA.Rows - 1
                  flxCrPoA.TextMatrix(iCount, 1) = "X"
             Next
        Else
             For iCount = 1 To flxCrPoA.Rows - 1
                   flxCrPoA.TextMatrix(iCount, 1) = ""
              Next
        End If
End Sub
Private Function checkOnlyOneSelected() As Boolean
     Dim iCount As Integer
     Dim iSel As Integer
     For iCount = 1 To flxCrPoA.Rows - 1
        If flxCrPoA.TextMatrix(iCount, 1) = "X" Then
            iSel = iSel + 1
        End If
     Next
     If iSel = 1 Then
        checkOnlyOneSelected = True
     Else
        checkOnlyOneSelected = False
     End If
        
End Function
Private Sub cmdAmendAllocationDate_Click()
    On Error GoTo ERR
    Dim szDate As String
    Dim iCount As Integer
    Dim adoconn As New ADODB.Connection
    Dim rsAllocation As New ADODB.Recordset
    Dim TransactionID As Long
    Dim szAllocationTransactionID As Long
    Dim szAllocdate As Date
    If checkOnlyOneSelected = False Then
        MsgBox "Please Select one Transaction from the list", vbOKOnly, "Please Select"
        FocusControl flxCrPoA
        Exit Sub
    End If
    If MsgBox("Do you wish to amend Allocation Date?", vbYesNo, "Please confirm") = vbNo Then Exit Sub
    For iCount = 1 To flxCrPoA.Rows - 1
        If flxCrPoA.TextMatrix(iCount, 1) = "X" Then
            TransactionID = flxCrPoA.TextMatrix(iCount, 0)
            szDate = flxCrPoA.TextMatrix(iCount, 5)
            Exit For
        End If
    Next
    FocusControl cmdSPClose
    adoconn.Open getConnectionString
'    rsAllocation.Open "Select allocdate,T.TransactionID from paytransactions T,tlbPayment P where P.transactionId=T.FromTran AND " & _
'            "P.transactionID=" & TransactionID & "", adoconn, adOpenKeyset, adLockReadOnly
'    If Not rsAllocation.EOF Then
'         szAllocdate = rsAllocation("allocdate").Value
'         frmPostingDate.szAllocationTransactionID = rsAllocation("TransactionID").Value
'    Else
'        MsgBox "Allocation date not found for this transaction", vbInformation, "Warning"
'        Exit Sub
'    End If
'    rsAllocation.Close

   For iCount = 1 To flxSPayment.Rows - 1
            szAllocationTransactionID = flxSPayment.TextMatrix(iCount, 0)
            szDate = flxSPayment.TextMatrix(iCount, 8)
            adoconn.Execute "Update PayTransactions T set allocdate=#" & Format(szDate, "dd MMM yyyy") & "# where T.FromTran=" & TransactionID & "" & _
           "and T.ToTran=" & szAllocationTransactionID & " and T.DeleteFlag=False"
           adoconn.Execute "Update PayTransactionsSplit T set allocdate=#" & Format(szDate, "dd MMM yyyy") & "# where T.FromTran=" & TransactionID & "" & _
           "and T.ToTran=" & szAllocationTransactionID & " and T.DeleteFlag=False"
            
    Next
    
    
    adoconn.Close
    Set adoconn = Nothing
    MsgBox "Record has been updated"
'    frmPostingDate.szCallingForm = Me.Name
'    frmPostingDate.szTransactionDate = szDate
'    Load frmPostingDate
'    frmPostingDate.txtPostingDate.text = szAllocdate
'
'    frmPostingDate.Top = Me.Top + Me.Height / 2 - frmPostingDate.Height / 2
'    frmPostingDate.Left = Me.Left + Me.Width / 2 - frmPostingDate.Width / 2
'
'    Me.Enabled = False
'    frmPostingDate.Show
    
    Exit Sub
ERR:
    MsgBox ERR.description
End Sub

Private Sub cmdFix_Click()
    Dim adoconn As New ADODB.Connection
    If MsgBox("Do you wish to run the fix?", vbYesNo, "Yes/No") = vbNo Then Exit Sub
    adoconn.Open getConnectionString
'    adoconn.Execute "Update tlbpaymentSplit set AllocTranID=2110191437090151242 where payheader=106"
    adoconn.Execute "Update tlbpaymentsplit set alloctranID='21123010480500081378' where Payheader=827"
    adoconn.Execute "Update tlbpaymentsplit set alloctranID='2112301045130005489' where Payheader=826"
    adoconn.Execute "Update tlbpaymentsplit set alloctranID='21120215254200098625' where Payheader=587"

    adoconn.Close
End Sub

Private Sub cmdSPClose_Click()
   Unload Me
   
End Sub
Private Function Exists(ByRef oCol As Collection, ByVal vKey As Variant) As Boolean
    On Error Resume Next
    oCol.Item vKey
    Exists = (ERR.Number = 0)
    ERR.Clear
End Function
Private Sub addAmount(ByRef oCol As Collection, Index As String, amt As Double)
    Dim tmpValue  As Double
    tmpValue = oCol(Index)
    tmpValue = tmpValue + amt
    oCol.Remove Index
    oCol.Add tmpValue, Index
End Sub
Private Function UnAllocatePayment2(adoconn As ADODB.Connection, PaymentTransactionID As Long) As Boolean
    Dim PayTransactions As New ADODB.Recordset
    Dim dicPayHeader As New Dictionary
    Dim dicPayHeaderAmount As New Dictionary
    Dim dblPerInvPayAmount As Double
    Dim rsPaymentSplit As New ADODB.Recordset
    Dim adoRst As New ADODB.Recordset
    Dim iTotalTransaction As Integer
    ReDim szaAllocTransID(20) As String
    ReDim szaAllocTransactionID(20) As String
    ReDim daSIRptID(20, 1) As Double
    Dim j As Integer
    Dim strMethod As String
    Dim i As Integer
    Dim K As Integer
    Dim strSupplierID As String
    Dim szSQL As String
    Dim rsAlloc As New ADODB.Recordset
    Dim AllocTransactionID  As String
    Dim PaymentType  As String
    Dim strPI As String
    'On Error GoTo Err
    
    strMethod = "2"
    'Payment side unallocate *****************************************************
       szSQL = "UPDATE tlbPayment AS P " & _
           "SET P.OSAmount = P.Amount, P.PaymentView = TRUE " & _
           "WHERE P.TransactionID = " & PaymentTransactionID & ";"

   adoconn.Execute szSQL

'  Update the OSAmt = Amt of all Payment splits
   szSQL = "SELECT * " & _
           "FROM tlbPaymentSplit " & _
           "WHERE PayHeader = " & PaymentTransactionID & " Order by amount DESC;"
   adoRst.Open szSQL, adoconn, adOpenDynamic, adLockOptimistic
   'Updating Payment os amount
   'Collectng the payment splits from payments
   While Not adoRst.EOF
      With adoRst
        If Not IsNull(.Fields.Item("ClientStatementID").Value) Then
                MsgBox "This transaction cannot be unallocated because it is included in a Client statement", vbInformation, "Warning"
                Exit Function
        End If
         .Fields.Item("OSAmount").Value = .Fields.Item("Amount").Value
         .Update
         szaAllocTransID(i) = CStr(.Fields.Item("Amount").Value)
         Debug.Print "From tlbPaymentSplit Amount is :" & CStr(.Fields.Item("Amount").Value)
         If IsNull(.Fields.Item("AllocTranID").Value) Then
                'PaymentTransactionID
                rsAlloc.Open "Select Type from  tlbPayment  where transactionID =" & PaymentTransactionID & " ", adoconn, adOpenStatic, adLockReadOnly
                If RecordCount(rsAlloc) = 1 Then
                    PaymentType = rsAlloc("Type").Value
                End If
                rsAlloc.Close
                If PaymentType = 7 Then
                        rsAlloc.Open "Select ToTran from PayTransactions where FromTran=" & PaymentTransactionID & "", adoconn, adOpenStatic, adLockReadOnly
                        If RecordCount(rsAlloc) = 1 Then
                            AllocTransactionID = rsAlloc("ToTran").Value
                            adoconn.Execute "Update tlbPayment SET OSAmount=Amount,paymentView=true where transactionID =" & PaymentTransactionID & ""
                            adoconn.Execute "Update tlbPaymentSplit SET Amount=OSAmount where PayHeader =" & PaymentTransactionID & ""
                        End If
                        rsAlloc.Close
                        adoconn.Execute "Update tlbPayment SET OSAmount=Amount,paymentView=true where transactionID =" & AllocTransactionID & ""
                        adoconn.Execute "Update tlbPaymentSplit SET Amount=OSAmount where PayHeader =" & AllocTransactionID & ""
                        adoconn.Execute "DELETE FROM PayTransactions where FromTran =" & PaymentTransactionID & ""
                        adoconn.Execute "DELETE FROM PayTransactionsSplit where FromTran =" & PaymentTransactionID & ""
'                        rsAlloc.Open "Select * from  tlbPayment where transactionID =" & AllocTransactionID & " ", adoconn, adOpenStatic, adLockReadOnly
'                        If RecordCount(rsAlloc) = 1 Then
'                            AllocTransactionID = rsAlloc("PI").Value
'                        End If
'                        rsAlloc.Close
                       ' adoconn.Execute "Update tblPurInv SET Amount=OSAmount where transactionID =" & AllocTransactionID & ""
                End If
                UnAllocatePayment2 = True
                Exit Function
         End If
         szaAllocTransactionID(i) = CStr(.Fields.Item("AllocTranID").Value)
         Debug.Print "From tlbPaymentSplit AllocTranID is :" & szaAllocTransactionID(i)
         
         i = i + 1
         .MoveNext
      End With
   Wend
   adoRst.Close
   '****************************************************************************************************
    'Going to find PI from allocation table
    'Collect payheader/payheaders that are related this current payment and storing them  into strPayTransactions variable
        Dim strPayTransactions As String
        Dim strPayTransactionsheaders As String
        PayTransactions.Open "Select ToTran from PayTransactions where FromTran=" & PaymentTransactionID & " AND DeleteFlag=False", adoconn, adOpenKeyset, adLockReadOnly
        While Not PayTransactions.EOF 'loop because there can two invoice in payment'Here loop is going through invoice
            If strPayTransactions = "" Then
                strPayTransactions = PayTransactions("ToTran").Value
            Else
                strPayTransactions = strPayTransactions & "," & PayTransactions("ToTran").Value
            End If
            PayTransactions.MoveNext
        Wend
         Debug.Print "Collect payheader/payheaders from  PayTransactions(these are invoices tran ID) :" & strPayTransactions
       'now strPayTransactions  variable  shall hold all the PI Invoices that are relates to this payment
       'Now search in the PI that matches with this payment amount and unallocate only that line
       'Going to the PI and unallocate
        Dim dblPayAmount As Double
        Dim dblPIAmount As Double
        '***********PI split unallocation 'SQL shall select the line to unallocate
        'Now there is other method by using AllocTranID column to find the perfect line for unallocation
        'I cant remember a long time ago I had found this relationsship was not maintained I did not go with that method
        'Now I am trying to have a version using that method
        If strMethod = "1" Then
                For K = 0 To i - 1 'I hold payment line count
                    Debug.Print K
                    rsPaymentSplit.Open "SELECT * FROM tlbPaymentSplit where PayHeader in (" & _
                        strPayTransactions & ") AND  amount >=" & szaAllocTransID(K) & " AND Amount>=OSamount+" & szaAllocTransID(K) & " Order by amount DESC", adoconn, adOpenDynamic, adLockOptimistic
                        'Amount<>OSamount explanation : OSamount is zero when nothing has been paid for this invoice
                        'if OSamount=Amount that means invoice is fully unpaid
                        'if OSamount<>Amount that means some or full part of the invoice had been paid
                        
                    While Not rsPaymentSplit.EOF 'loop because there can two split in an PI invoice,loop is going through PI invoice
                        If rsPaymentSplit.Fields.Item("OSAmount").Value = 0 Then
                            rsPaymentSplit.Fields.Item("OSAmount").Value = CDbl(szaAllocTransID(K))
                        Else
                            rsPaymentSplit.Fields.Item("OSAmount").Value = rsPaymentSplit.Fields.Item("OSAmount").Value + CDbl(szaAllocTransID(K))
                        End If
                        dblPayAmount = dblPayAmount + CDbl(szaAllocTransID(K)) 'saving the amount for future checksome
                        rsPaymentSplit.Update
                        If rsPaymentSplit.Fields.Item("OSAmount").Value > rsPaymentSplit.Fields.Item("Amount").Value Then
                            rsPaymentSplit.Close
                            UnAllocatePayment2 = False
                            Exit Function
                        End If
                        If strPayTransactionsheaders = "" Then
                            strPayTransactionsheaders = rsPaymentSplit("PayHeader").Value
                        Else
                            strPayTransactionsheaders = strPayTransactions & "," & rsPaymentSplit("PayHeader").Value
                        End If
                      
                        rsPaymentSplit.MoveNext
                    Wend
                    rsPaymentSplit.Close
                Next K
     Else
            For K = 0 To i - 1 'i is for payment line count
                    Debug.Print K
                    If strPayTransactions = "" Then Exit Function
                    rsPaymentSplit.Open "SELECT * FROM tlbPaymentSplit where PayHeader in (" & _
                        strPayTransactions & ") AND  TransactionID ='" & szaAllocTransactionID(K) & "' AND Amount>=OSamount+" & szaAllocTransID(K) & " Order by amount DESC", adoconn, adOpenDynamic, adLockOptimistic
                        'Amount<>OSamount explanation : OSamount is zero when nothing has been paid for this invoice
                        'if OSamount=Amount that means invoice is fully unpaid
                        'if OSamount<>Amount that means some or full part of the invoice had been paid
'                      Debug.Print "These are the PI for redeem"
'                      Debug.Print "SELECT * FROM tlbPaymentSplit where PayHeader in (" & _
'                        strPayTransactions & ") AND  TransactionID ='" & szaAllocTransactionID(k) & "' AND Amount>=OSamount+" & szaAllocTransID(k) & " Order by amount DESC"
                        
                    While Not rsPaymentSplit.EOF 'loop because there can two split in an PI invoice,loop is going through PI invoice
                        If rsPaymentSplit.Fields.Item("OSAmount").Value = 0 Then
                            rsPaymentSplit.Fields.Item("OSAmount").Value = CDbl(szaAllocTransID(K))
                        Else
                            rsPaymentSplit.Fields.Item("OSAmount").Value = rsPaymentSplit.Fields.Item("OSAmount").Value + CDbl(szaAllocTransID(K))
                        End If
                        If dicPayHeader.Exists(rsPaymentSplit("PayHeader").Value) Then
                                 'rsPaymentSplit("PayHeader").Value is the index of the collection
                                 'MsgBox "This item Already exists ,cannot add again.But I am adding amount again "
                                 'addAmount function add amount to an existing transation to the collection
                             dicPayHeader(rsPaymentSplit("PayHeader").Value) = dicPayHeader(rsPaymentSplit("PayHeader").Value) + CDbl(szaAllocTransID(K))
                        Else
                             dicPayHeader.Add rsPaymentSplit("PayHeader").Value, CDbl(szaAllocTransID(K))
                        End If
                        dblPayAmount = dblPayAmount + CDbl(szaAllocTransID(K)) 'saving the amount for future checksome
                        rsPaymentSplit.Update
                        If rsPaymentSplit.Fields.Item("OSAmount").Value > rsPaymentSplit.Fields.Item("Amount").Value Then
                            rsPaymentSplit.Close
                            UnAllocatePayment2 = False
                            Exit Function
                        End If
                        If strPayTransactionsheaders = "" Then
                            strPayTransactionsheaders = rsPaymentSplit("PayHeader").Value
                        Else
                            strPayTransactionsheaders = strPayTransactions & "," & rsPaymentSplit("PayHeader").Value
                        End If
                      
                        rsPaymentSplit.MoveNext
                    Wend
                    rsPaymentSplit.Close
                Next K
     End If
     'dblPerInvPayAmount
     '******************* I am trying new code for each  invoice
     Dim key
     For Each key In dicPayHeader.Keys
             adoRst.Open "SELECT * FROM tlbPayment where TransactionID in (" & key & ") Order by OSAmount ", adoconn, adOpenDynamic, adLockOptimistic
                dblPerInvPayAmount = dicPayHeader(key)
                'There will be only one record on each invoice
             If Not adoRst.EOF Then
                 adoRst.Fields.Item("OSAmount").Value = adoRst.Fields.Item("OSAmount").Value + dblPerInvPayAmount
                 dblPIAmount = dblPIAmount + dblPerInvPayAmount
                     
                 If strSupplierID = "" Then
                     strSupplierID = " '" & adoRst.Fields.Item("SageAccountNumber").Value & " '"
                 Else
                     strSupplierID = strSupplierID & ",'" & adoRst.Fields.Item("SageAccountNumber").Value & "'"
                 End If
                 If adoRst.Fields.Item("OSAmount").Value > adoRst.Fields.Item("Amount").Value Then
                             'adoRst.Close
                             UnAllocatePayment2 = False
                             Exit Function
                 End If
                 adoRst.Fields.Item("PaymentView").Value = True
                 adoRst.Update
                 
            End If
            adoRst.Close
    Next key

     '*********
    'Update the header of all PIs in the payment table.
'    adoRst.Open "SELECT * FROM tlbPayment where TransactionID in (" & strPayTransactionsheaders & ") Order by OSAmount ", adoConn, adOpenDynamic, adLockOptimistic
'
'    While Not adoRst.EOF
'        'dblPIAmount = 0
'        'dblPIAmount = dblPIAmount + adoRst.Fields.Item("Amount").Value
'        'Now this part is clearly failing when payment is less than the PI invoice amount So I am adding a if condition here
'        'Need a separate section for that 2020-05-12 by anol
'        If dblPayAmount >= adoRst.Fields.Item("Amount").Value Then
'            adoRst.Fields.Item("OSAmount").Value = adoRst.Fields.Item("Amount").Value
'            dblPIAmount = dblPIAmount + adoRst.Fields.Item("Amount").Value
'        Else
'            If adoRst.Fields.Item("OSAmount").Value = 0 Then
'                'for first time un allocate/withdraw  payment to invoice
'                adoRst.Fields.Item("OSAmount").Value = dblPayAmount
'                dblPIAmount = dblPIAmount + dblPayAmount
'             Else
'                'for second timeun allocate/withdraw payment and so on  to invoice
'                adoRst.Fields.Item("OSAmount").Value = adoRst.Fields.Item("OSAmount").Value + dblPayAmount
'                dblPIAmount = dblPIAmount + dblPayAmount
'             End If
'        End If
'        If strSupplierID = "" Then
'            strSupplierID = " '" & adoRst.Fields.Item("SageAccountNumber").Value & " '"
'        Else
'            strSupplierID = strSupplierID & ",'" & adoRst.Fields.Item("SageAccountNumber").Value & "'"
'        End If
'        If adoRst.Fields.Item("OSAmount").Value > adoRst.Fields.Item("Amount").Value Then
'                    'adoRst.Close
'                    UnAllocatePayment2 = False
'                    Exit Function
'        End If
'        adoRst.Fields.Item("PaymentView").Value = True
'        adoRst.Update
'        adoRst.MoveNext
'   Wend
'   adoRst.Close
   
   
    If Round(dblPayAmount, 2) = Round(dblPIAmount, 2) Then
    Else
        UnAllocatePayment2 = False
        MsgBox "Amount mismatch at payment and PI while un allocation . UnAllocation failed."
        Exit Function
    End If
    'Here deleting by payment from allocation table
   szSQL = "Update PayTransactions SET DeleteFlag=true " & _
           "WHERE PayTransactions.FromTran = " & PaymentTransactionID & ";"

'Now delete for split allocation table
   adoconn.Execute szSQL
   szSQL = "Update PayTransactionsSplit SET DeleteFlag=true " & _
           "WHERE PayTransactionsSplit.FromTran = " & PaymentTransactionID & ";"

   adoconn.Execute szSQL
   'this function is written by anol 2019 05 23 when found that (issue 673 )Updated OS amount  extra incorrectly
    'This function shall prevent saving the data if when outstading amount on payment is not updated.
    'This functionshall compare allocation with payment amount and outstanding amount
    Dim rsChecksum As New ADODB.Recordset
    Dim strWhere As String
    Dim szTran2Fix As String
    If strSupplierID <> "" Then
        strWhere = " AND R.sageaccountnumber in (" & strSupplierID & ")"
    Else
        MsgBox "Amount mismatch at payment and PI while un allocation . UnAllocation failed."
        Exit Function
    End If
    
    '        rsChecksum.Open "Select SlNumber,amount,osamount,amt,R.sageaccountnumber from tlbReceipt R,(select Sum(ReceiptAmount) as amt,ToTran from RptTransactions  group By ToTran ) as A " & _
    '                    "Where a.ToTran = r.TransactionID  " & StrWhere & " and  Round((amount - amt), 2) <> Round(OSAmount, 2)", adoConn, adOpenStatic, adLockReadOnly
        
    rsChecksum.Open "Select  R.transactionID,R.SlNumber,'',R.sageaccountnumber,amount,osamount,amt from tlbPayment R,(select Sum(PaymentAmount) as amt," & _
      " ToTran from PayTransactions  where DeleteFlag=False group By ToTran ) as A where A.ToTran=R.transactionID " & strWhere & " AND round((amount-amt),2)<>round(osamount,2)", adoconn, adOpenStatic, adLockReadOnly
    
    While Not rsChecksum.EOF
        szTran2Fix = szTran2Fix + IIf(szTran2Fix = "", "PI", ",PI") + CStr(rsChecksum("SlNumber").Value) + " (" + rsChecksum("sageaccountnumber").Value + ") "
        rsChecksum.MoveNext
    Wend
    
    rsChecksum.Close
    Set rsChecksum = Nothing
    
    'Dim adoRst As New ADODB.Recordset
    szSQL = "SELECT  P.TransactionID " & _
               "FROM tlbPayment AS P, (" & _
                     "SELECT PayHeader, ROUND(Sum(Amount) - Sum(OSAmount), 2) AS T " & _
                     "From tlbPaymentSplit " & _
                     "Group by PayHeader " & _
                     ") AS Q " & _
               "WHERE P.TransactionID = Q.PayHeader AND P.Amount <> P.OSAmount AND " & _
                     "ROUND(P.Amount - P.OSAmount, 2) <> Q.T;"

      rsChecksum.Open szSQL, adoconn, adOpenStatic, adLockReadOnly

      While Not rsChecksum.EOF
         szTran2Fix = szTran2Fix + ", " + CStr(rsChecksum.Fields.Item("TransactionID").Value)

         rsChecksum.MoveNext
      Wend

      rsChecksum.Close
      Set rsChecksum = Nothing
    
    If szTran2Fix = "" Then
         UnAllocatePayment2 = True
    Else
        MsgBox "A problem occurred while unallocating: " & _
             Chr(13) & szTran2Fix & "." & _
             "This transaction has not been saved.", _
             vbInformation + vbOKOnly, "PI unAllocation was not successful!"
             UnAllocatePayment2 = False
             Exit Function
    End If
   UnAllocatePayment2 = True
   Exit Function
ERR:
   UnAllocatePayment2 = False
     MsgBox "A problem occurred while unallocating: " & ERR.description
End Function
'Private Function unAllocatePayment(adoConn As ADODB.Connection, PaymentTransactionID As Long) As Boolean
'   Dim szaAllocTransID()   As String
'   Dim daSIRptID()         As Double
'   Dim szaTemp()           As String
'   Dim i                   As Integer
'   Dim j                   As Integer
'   On Error GoTo Err
'   Dim szSQL      As String, iRow As Integer
'   Dim adoRst     As New ADODB.Recordset
'
'   szSQL = "UPDATE tlbPayment AS P " & _
'           "SET P.OSAmount = P.Amount, P.PaymentView = TRUE " & _
'           "WHERE P.TransactionID = " & PaymentTransactionID & ";"
'
'   adoConn.Execute szSQL
'
''  Update the OSAmt = Amt of all Payment splits
''  At this stage save the AllocTranID in an array @szaAllocTransID
'   szSQL = "SELECT * " & _
'           "FROM tlbPaymentSplit " & _
'           "WHERE PayHeader = " & PaymentTransactionID & ";"
'   adoRst.Open szSQL, adoConn, adOpenDynamic, adLockOptimistic
''Debug.Print szSQL
'   ReDim szaAllocTransID(RecordCount(adoRst) - 1) As String
'   ReDim daSIRptID(RecordCount(adoRst) - 1, 1) As Double
'   'Updating Payment os amount
'   While Not adoRst.EOF
'      With adoRst
'         .Fields.Item("OSAmount").Value = .Fields.Item("Amount").Value
'         .Update
'
'         szaAllocTransID(i) = .Fields.Item("AllocTranID").Value & _
'                              "-#-" & CStr(.Fields.Item("Amount").Value)
'         i = i + 1
'         .MoveNext
'      End With
'   Wend
'   adoRst.Close
'
'   'Here the program trying to update PI split amount
' For j = 0 To i - 1 'if a payment has two lines this loop shall run twice  but now it has one line payment tran ID 60
'      szaTemp = Split(szaAllocTransID(j), "-#-")
'    'szaTemp(0) contains the AllocTranID from tlbPaymentsplit table
'     'szaTemp(1) contains the Amount from tlbReceiptsplit table
'     'This tlbpaymentsplit  is  PI paymentsplit which has a allocation ralationship
'     'I have a case in 2020-03-05 where PI split ID has been deleted or edited or changed i dont know how is that done which is after allocation.
'      If Len(szaTemp(0)) > 10 Then
'          'modified by anol 2019-06-11
'         adoRst.Open "SELECT * FROM tlbPaymentSplit where TransactionID = '" & szaTemp(0) & "' AND round(Amount,2)>=" & szaTemp(1) & " AND Amount<>OSamount;", adoConn, adOpenDynamic, adLockOptimistic
'      Else
'        'Modified by anol 2019-06-11
'         adoRst.Open "SELECT * FROM tlbPaymentSplit where PayHeader = " & szaTemp(0) & " AND round(Amount,2)>=" & szaTemp(1) & " AND Amount<>OSamount;", adoConn, adOpenDynamic, adLockOptimistic 'herer szaTemp(0) =15 which have two split
'      End If
'      If RecordCount(adoRst) = 1 Then
'            adoRst.Fields.Item("OSAmount").Value = CDbl(adoRst.Fields.Item("OSAmount").Value) + _
'                                                      CDbl(szaTemp(1))
'            adoRst.Update
'            If adoRst.Fields.Item("OSAmount").Value > adoRst.Fields.Item("Amount").Value Then
'                adoRst.Close
'                unAllocatePayment = False
'                Exit Function
'            End If
'            daSIRptID(j, 0) = CDbl(adoRst.Fields.Item("PayHeader").Value)
'            daSIRptID(j, 1) = CDbl(szaTemp(1))
'      Else
'            'No such record found as per allocation ID in receipt so exit this function and say failed
'            adoRst.Close
'            unAllocatePayment = False
'            Exit Function
'      End If
'      adoRst.Close
'   Next j
'
''original code start
''     Update all PI splits in the payment split table
'''   adoRst.Open "SELECT * FROM tlbPaymentSplit;", adoConn, adOpenDynamic, adLockOptimistic
'''
'''   For j = 0 To i - 1
'''      szaTemp = Split(szaAllocTransID(j), "-#-")
'''
'''      If Len(szaTemp(0)) > 10 Then
'''         adoRst.Find ("TransactionID = '" & szaTemp(0) & "'"), , , 1
'''      Else
'''         adoRst.Find ("RptHeader = " & CLng(szaTemp(0)) & ""), , , 1
'''      End If
'''      adoRst.Fields.Item("OSAmount").Value = CDbl(adoRst.Fields.Item("OSAmount").Value) + _
'''                                                CDbl(szaTemp(1))
'''      adoRst.Update
'''      daSIRptID(j, 0) = CDbl(adoRst.Fields.Item("PayHeader").Value)
'''      daSIRptID(j, 1) = CDbl(szaTemp(1))
'''   Next j
'''   adoRst.Close
''original code end
'
'
'
''   adoconn.BeginTrans'rem by anol 20170802 Becaue this was causing two simultanous transa locking
'
''  Update the header of all PIs in the payment table.
'   adoRst.Open "SELECT * FROM tlbPayment;", adoConn, adOpenDynamic, adLockOptimistic
'
'   For j = 0 To i - 1
'      adoRst.Find "TransactionID = " & daSIRptID(j, 0) & "", , , 1
''Debug.Print daSIRptID(j, 1)
'      adoRst.Fields.Item("OSAmount").Value = CDbl(adoRst.Fields.Item("OSAmount").Value) + _
'                                                daSIRptID(j, 1)
''Debug.Print adoRst.Fields.Item("OSAmount").Value
'      adoRst.Fields.Item("PaymentView").Value = True
'      adoRst.Update
'   Next j
'   adoRst.Close
'Rem out by anol 20170822
''   adoconn.Execute _
''           "DELETE S.* " & _
''           "FROM   tlbPaymentSplit AS S, tlbPayment AS R " & _
''           "WHERE  S.PayHeader = R.TransactionID AND " & _
''               "R.TransactionID = " & PaymentTransactionID & " AND " & _
''               "R.TYPE = 9;"
'
''   adoconn.CommitTrans
'
'   szSQL = "DELETE * " & _
'           "FROM PayTransactions " & _
'           "WHERE PayTransactions.FromTran = " & PaymentTransactionID & ";"
'
'   adoConn.Execute szSQL
'   unAllocatePayment = True
'   Exit Function
'Err:
'   unAllocatePayment = False
'End Function
Private Function CountSelectedItem() As Long
   Dim rCount As Integer
   For rCount = 1 To flxCrPoA.Rows - 1
        If flxCrPoA.TextMatrix(rCount, 1) = "X" Then
            CountSelectedItem = CountSelectedItem + 1
        End If
   Next
End Function
Private Sub cmeRevereseAllocation_Click()

   Dim adoconn    As New ADODB.Connection
   Dim rCount As Long
   Dim iCount  As Long
   Dim jCount  As Long
   Dim bResult As Boolean
'   If flxCrPoA.TextMatrix(flxCrPoA.row, 0) = "" Then
'      ShowMsgInTaskBar "Please select a credit transaction first to reverse allocate.", , "N"
'      Exit Sub
'   End If
    If CountSelectedItem = 0 Then
      ShowMsgInTaskBar "Please select a credit transaction first to reverse allocate.", , "N"
      Exit Sub
   End If
   If cmdSPClose.Enabled Then
        FocusControl cmdSPClose
   End If
   
   If MsgBox("Do you wish to reverse the allocation of selected credit transaction?", vbQuestion + vbYesNo, _
         "Reverse Allocation") = vbNo Then Exit Sub

   '   connect to database
   
   For rCount = 1 To flxCrPoA.Rows - 1
        If flxCrPoA.TextMatrix(rCount, 1) = "X" Then
                adoconn.Open getConnectionString
                adoconn.BeginTrans
                'research by anol 2020-03-05 So the rule is one payment can serve two invoices
                   ' and also one inovoice can be servered by two payments so its a complex relationship
                   If flxCrPoA.TextMatrix(rCount, 0) = "" Then
                        adoconn.RollbackTrans
                        adoconn.Close
                        Set adoconn = Nothing
                        GoTo NextLoop
                   End If
                     bResult = UnAllocatePayment2(adoconn, flxCrPoA.TextMatrix(rCount, 0)) 'flxCrPoA is loading TransactionID at 0 index FROM tlbPayment by anol
'                     If bResult = False Then
'                         bResult = unAllocatePayment(adoConn, flxCrPoA.TextMatrix(rCount, 0))
'                     End If
'                    bResult = unAllocatePayment(adoConn, flxCrPoA.TextMatrix(rCount, 0)) 'flxCrPoA is loading TransactionID at 0 index FROM tlbPayment by anol
                If bResult = True Then
                    adoconn.CommitTrans
                    iCount = iCount + 1
                    flxCrPoA.RowHeight(rCount) = 0
                    Debug.Print "Tag" & rCount
                Else
                    flxCrPoA.RowHeight(rCount) = 240
                    adoconn.RollbackTrans
                    Debug.Print "Tag2" & rCount
                    jCount = jCount + 1
                End If
                adoconn.Close
                Set adoconn = Nothing
               
                
                flxCrPoA.Refresh
        Else
            
        End If
NextLoop:
    Next
   'ConfigFlxCrPoA flxCrPoA, 9, 7, 20
   ConfigFlxSPayment
   flxSPayment.Cols = 12
   flxSPayment.ColWidth(11) = 0
   If adoconn.State = 1 Then
         'adoconn.Close
         Exit Sub
   End If
   adoconn.Open getConnectionString
   'LoadFlxCrPoA adoConn

   frmPurchaseExpense.LoadFlxSPayment adoconn
   frmPurchaseExpense.LoadFlxSCrPoA adoconn
   adoconn.Close
   Set adoconn = Nothing
   chkCredits.Value = 0
'   If jCount = 0 Then
'        MsgBox iCount & " Transaction has been reversed successfully.", vbInformation, "Information"
'   Else
'        MsgBox iCount & " Transaction has been reversed successfully. Failed : " & jCount, vbInformation, "Information"
'   End If
   flxCrPoA.row = 0
   
End Sub

Private Sub cmeRevereseAllocation_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmeRevereseAllocation.Refresh
End Sub

Private Sub flxCrPoA_Click()
    Dim szSQL As String, iRow As Integer
    Dim adoconn As New ADODB.Connection
    Dim adoRst As New ADODB.Recordset
    adoconn.Open getConnectionString
     If flxCrPoA.TextMatrix(flxCrPoA.row, 1) = "" And flxCrPoA.TextMatrix(flxCrPoA.row, 0) <> "" Then
           flxCrPoA.TextMatrix(flxCrPoA.row, 1) = "X"
           
'loading all the credits
    Dim rsCheck As New ADODB.Recordset
        rsCheck.Open "Select SplitIDofPI  from PayTransactions AS PT where PT.FromTran = " & flxCrPoA.TextMatrix(flxCrPoA.row, 0) & " AND DeleteFlag=false", adoconn, adOpenStatic, adLockReadOnly
        If rsCheck.EOF Then
                MsgBox "Allocation entry not found , Please contact PCM Consulting Support", vbInformation, "Warning"
                  Exit Sub
                rsCheck.Close
        End If
        If IsNull(rsCheck("SplitIDofPI").Value) Then
           szSQL = "SELECT T.DESCRIPTION, P.SlNumber, P.UnitID, P.PDate, P.PI, " & _
                             "P.Details, P.Amount, P.OSAmount, PT.PaymentAmount, P.Ref, P.TransactionID,PT.AllocDate, " & _
                             "MID(T.CONSTANT, 4, LEN(T.CONSTANT)-3) AS PF " & _
                      "FROM PayTransactions AS PT, tlbTransactionTypes AS T, tlbPayment AS P INNER JOIN " & _
                          "(SELECT PT.ToTran " & _
                           "FROM PayTransactions AS PT " & _
                           "Where DeleteFlag=False and PT.FromTran = " & flxCrPoA.TextMatrix(flxCrPoA.row, 0) & " " & _
                          ") AS X ON P.TransactionID = X.ToTran  " & _
                      "Where PT.DeleteFlag=False and  PT.ToTran = P.TransactionID And " & _
                         "PT.FromTran = " & flxCrPoA.TextMatrix(flxCrPoA.row, 0) & " And " & _
                         "P.Type = T.TYPE_ID;"
        Else
                   szSQL = "SELECT T.DESCRIPTION, P.SlNumber, P.UnitID, P.PDate, P.PI, " & _
                             "P.Details, P.Amount, P.OSAmount, PT.PaymentAmount, P.Ref, P.TransactionID,PT.AllocDate, " & _
                             "MID(T.CONSTANT, 4, LEN(T.CONSTANT)-3) AS PF " & _
                      "FROM PayTransactions AS PT, tlbTransactionTypes AS T,tlbPaymentSplit AS PS ,tlbPayment AS P " & _
                      "Where PT.DeleteFlag=False  AND P.TransactionID=PS.payheader and ps.SplitID=PT.SplitIDofPI  AND PT.ToTran = P.TransactionID And " & _
                         "PT.FromTran = " & flxCrPoA.TextMatrix(flxCrPoA.row, 0) & " And " & _
                         "P.Type = T.TYPE_ID;"
        End If
                         
'
    rsCheck.Close

        'ToTran is PI
           adoRst.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
        
           ConfigFlxSPayment
        
           iRow = 1
           While Not adoRst.EOF
              flxSPayment.TextMatrix(iRow, 0) = adoRst.Fields.Item("TransactionID").Value
              flxSPayment.TextMatrix(iRow, 2) = adoRst.Fields.Item("PF").Value & adoRst.Fields.Item("SlNumber").Value
              flxSPayment.TextMatrix(iRow, 3) = IIf(IsNull(adoRst.Fields.Item("DESCRIPTION").Value), "", adoRst.Fields.Item("DESCRIPTION").Value)
              flxSPayment.TextMatrix(iRow, 4) = IIf(IsNull(adoRst.Fields.Item("UnitID").Value), "", adoRst.Fields.Item("UnitID").Value)
              flxSPayment.TextMatrix(iRow, 5) = adoRst.Fields.Item("PDate").Value
              flxSPayment.TextMatrix(iRow, 6) = IIf(IsNull(adoRst.Fields.Item("Ref").Value), "", adoRst.Fields.Item("Ref").Value)
              flxSPayment.TextMatrix(iRow, 7) = IIf(IsNull(adoRst.Fields.Item("Details").Value), "", adoRst.Fields.Item("Details").Value)
              flxSPayment.TextMatrix(iRow, 8) = Format(adoRst.Fields.Item("AllocDate").Value, "dd/MM/yyyy")
              flxSPayment.TextMatrix(iRow, 9) = Format(adoRst.Fields.Item("Amount").Value, "0.00")
              flxSPayment.TextMatrix(iRow, 10) = Format(adoRst.Fields.Item("OSAmount").Value, "0.00")
              flxSPayment.TextMatrix(iRow, 11) = Format(adoRst.Fields.Item("PaymentAmount").Value, "0.00")
        
              adoRst.MoveNext
              If Not adoRst.EOF Then
                 iRow = iRow + 1
                 flxSPayment.AddItem ""
              End If
           Wend
        
           adoRst.Close
           Set adoRst = Nothing
           
     Else
           flxCrPoA.TextMatrix(flxCrPoA.row, 1) = ""
           ConfigFlxSPayment
     End If
     adoconn.Close
     Set adoconn = Nothing
End Sub

Private Sub Form_Activate()
     cmeRevereseAllocation.Refresh
End Sub

Private Sub Form_Load()
   Me.Width = 15285
   Me.Height = 8490
'   frmMMain.Arrange vbCascade
'   Me.ZOrder 0
   Me.BackColor = MODULEBACKCOLOR

   ConfigFlxCrPoA flxCrPoA, 9, 7, 20
   ConfigFlxSPayment

   Me.Caption = "Reverse Allocation of Payment - " & frmPurchaseExpense.txtSPSupplier.text

  

   szSupplierID = frmPurchaseExpense.txtSPSupplier.Tag

   LoadFlxCrPoA

   flxCrPoA.row = 0
   flxCrPoA.col = 0
'   adoConn.Close
'   Set adoConn = Nothing
   ' cmd button was not refreshing 20170219 anol
  
   Call WheelHook(Me.hWnd)
End Sub

Public Sub LoadFlxCrPoA()
   Dim szSQL As String, i As Integer
   Dim adoRst As New ADODB.Recordset
    Dim adoconn As New ADODB.Connection


   adoconn.Open getConnectionString
' AND P.ClientID='"& frmPurchaseExpense.cboClient.Column(0) &"' has been added by anol
'Date 25 Aug 2015 issue 571 Note 1148
'
' szSQL = "SELECT P.TransactionID, P.SlNumber, P.Amount, P.PDate, P.UnitID, P.Details, " & _
'               "T.DESCRIPTION, P.ExtRef, " & _
'                  "MID(T.CONSTANT, 4, LEN(T.CONSTANT)-3) AS PF " & _
'           "FROM tlbPayment AS P, tlbTransactionTypes AS T " & _
'           "WHERE " & _
'               "(P.Type = 7 OR P.Type = 8 OR P.Type = 9) AND " & _
'               "P.Type = T.TYPE_ID AND P.Amount > P.OSAmount " & _
'           "ORDER BY P.TransactionID, P.Type;"

   szSQL = "SELECT P.TransactionID, P.SlNumber, P.Amount, P.PDate, P.UnitID, P.Details, " & _
               "T.DESCRIPTION, P.ExtRef, " & _
                  "MID(T.CONSTANT, 4, LEN(T.CONSTANT)-3) AS PF " & _
           "FROM tlbPayment AS P, tlbTransactionTypes AS T " & _
           "WHERE " & _
               "P.SageAccountNumber = '" & szSupplierID & "' AND P.ClientID='" & frmPurchaseExpense.txtClientIDPurPay.text & "' AND " & _
               "(P.Type = 7 OR P.Type = 8 OR P.Type = 9) AND " & _
               "P.Type = T.TYPE_ID AND P.Amount > P.OSAmount " & _
           "ORDER BY P.TransactionID, P.Type;"

   adoRst.Open szSQL, adoconn, adOpenStatic, adLockReadOnly

   i = 1
   While Not adoRst.EOF
      flxCrPoA.TextMatrix(i, 0) = adoRst.Fields.Item("TransactionID").Value
      flxCrPoA.TextMatrix(i, 2) = adoRst.Fields.Item("PF").Value & adoRst.Fields.Item("SlNumber").Value

      flxCrPoA.TextMatrix(i, 3) = IIf(IsNull(adoRst.Fields.Item("DESCRIPTION").Value), "", adoRst.Fields.Item("DESCRIPTION").Value)
      If InStr(flxCrPoA.TextMatrix(i, 3), "Payment") > 0 And InStr(flxCrPoA.TextMatrix(i, 3), "Account") = 0 Then flxCrPoA.TextMatrix(i, 3) = "Payment"
      If InStr(flxCrPoA.TextMatrix(i, 3), "Account") > 0 Then flxCrPoA.TextMatrix(i, 3) = "Payment on A/C"

      flxCrPoA.TextMatrix(i, 4) = IIf(IsNull(adoRst.Fields.Item("UnitID").Value), "", adoRst.Fields.Item("UnitID").Value)
      flxCrPoA.TextMatrix(i, 5) = adoRst.Fields.Item("PDate").Value
      flxCrPoA.TextMatrix(i, 6) = IIf(IsNull(adoRst.Fields.Item("ExtRef").Value), "", adoRst.Fields.Item("ExtRef").Value)
      flxCrPoA.TextMatrix(i, 7) = IIf(IsNull(adoRst.Fields.Item("Details").Value), "", adoRst.Fields.Item("Details").Value)
      flxCrPoA.TextMatrix(i, 8) = Format(adoRst.Fields.Item("Amount").Value, "0.00")
      adoRst.MoveNext

      'If Not adoRst.EOF Then
         i = i + 1
         flxCrPoA.AddItem ""
     ' End If
   Wend

   adoRst.Close
   Set adoRst = Nothing
   adoconn.Close
End Sub

Private Sub ConfigFlxCrPoA(ctrFlxGrid As MSHFlexGrid, iCols As Integer, iLabels As Integer, iLblFstIdx As Integer, Optional szHeader As String)
   Dim iCol As Integer

   ctrFlxGrid.Clear
   ctrFlxGrid.Cols = iCols
   ctrFlxGrid.Rows = 2
   ctrFlxGrid.RowHeight(0) = 0

   If szHeader <> "" Then _
      ctrFlxGrid.FormatString = szHeader$

   ctrFlxGrid.ColWidth(0) = 0                                           'ID of the transactions
   ctrFlxGrid.ColWidth(1) = Label19(iLblFstIdx).Left - ctrFlxGrid.Left  'Sign -> X, +, -

   For iCol = 2 To ctrFlxGrid.Cols - 2
      ctrFlxGrid.ColWidth(iCol) = Label19(iCol + iLblFstIdx - 1).Left - Label19(iCol - 1 + iLblFstIdx - 1).Left - 50
   Next iCol
   ctrFlxGrid.ColWidth(iCol) = ctrFlxGrid.Width + ctrFlxGrid.Left - Label19(iCol - 2 + iLblFstIdx).Left
End Sub
Private Sub ConfigFlxSPayment()
    Dim iCol As Integer
    
    flxSPayment.Clear
    flxSPayment.Cols = 12
    flxSPayment.Rows = 2
    flxSPayment.RowHeight(0) = 0
    flxSPayment.FormatString = "|<No|<Type|<UnitID||<Date|<Ref|Details|>Amt"
    flxSPayment.ColWidth(0) = 0                                           'ID of the transactions
    flxSPayment.ColWidth(1) = 250 'Sign -> X
    'flxSPayment.ColWidth(2) = 250   'Sign ->, +, -
    flxSPayment.ColWidth(2) = Label19(11).Left - Label19(10).Left
    flxSPayment.ColWidth(3) = Label19(12).Left - Label19(11).Left
    flxSPayment.ColWidth(4) = Label19(13).Left - Label19(12).Left
    flxSPayment.ColWidth(5) = Label19(14).Left - Label19(13).Left
    flxSPayment.ColWidth(6) = Label19(15).Left - Label19(14).Left
    flxSPayment.ColAlignment(6) = vbLeftJustify
    flxSPayment.ColWidth(7) = 3500 ' Label19(16).Left - Label19(15).Left
    flxSPayment.ColAlignment(7) = vbLeftJustify
    flxSPayment.ColWidth(8) = 1200 'Label19(27).Left - Label19(26).Left
     flxSPayment.ColWidth(9) = 1200
    flxSPayment.ColWidth(10) = 1200
    flxSPayment.ColWidth(11) = 1200
'    FlxSPayment.ColWidth(2) = 250
'    FlxSPayment.ColWidth(iCol) = 3500
'    FlxSPayment.ColWidth(iCol + 1) = 1300
End Sub
'Private Sub ConfigFlxSPayment(ctrFlxGrid As MSHFlexGrid, iCols As Integer, iLabels As Integer, iLblFstIdx As Integer, Optional szHeader As String)
'   Dim iCol As Integer
'
'   ctrFlxGrid.Clear
'   ctrFlxGrid.Cols = iCols
'   ctrFlxGrid.Rows = 2
'   ctrFlxGrid.RowHeight(0) = 0
'
'   If szHeader <> "" Then _
'      ctrFlxGrid.FormatString = szHeader$
'
'   ctrFlxGrid.ColWidth(0) = 0                                           'ID of the transactions
'   ctrFlxGrid.ColWidth(1) = Label19(iLblFstIdx).Left - ctrFlxGrid.Left  'Sign -> X, +, -
'
'   For iCol = 2 To ctrFlxGrid.Cols - 2
'      ctrFlxGrid.ColWidth(iCol) = Label19(iCol + iLblFstIdx - 1).Left - Label19(iCol - 1 + iLblFstIdx - 1).Left
'   Next iCol
'   ctrFlxGrid.ColWidth(iCol) = ctrFlxGrid.Width + ctrFlxGrid.Left - Label19(iCol - 2 + iLblFstIdx).Left - 340
'End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Me.MousePointer = vbArrow
End Sub

Private Sub Form_Unload(Cancel As Integer)
'   Dim adoconn As New ADODB.Connection
'   Dim adoRst As New ADODB.Recordset
'
   frmPurchaseExpense.Show
'
'   adoconn.Open getConnectionString
'
'   frmPurchaseExpense.LoadFlxSPayment adoconn
'   frmPurchaseExpense.LoadFlxSCrPoA adoconn
'   frmPurchaseExpense.RedeclareArray
'
'   adoconn.Close
'   Set adoconn = Nothing
'
'   Call WheelUnHook(Me.hWnd)
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
            bHandled = False
'          PictureBoxZoom ctl, MouseKeys, Rotation, Xpos, Ypos

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
Private Sub txtPayDt_Change()
   TextBoxChangeDate txtPayDt
End Sub

Private Sub txtPayDt_GotFocus()
   'iCurRow = StarFound

   If flxSPayment.col = 8 Then
      If flxSPayment.TextMatrix(flxSPayment.row, 8) = "" Then
         txtPayDt.text = Format(Now, "dd/mm/yyyy")
      Else
         txtPayDt.text = Format(flxSPayment.TextMatrix(flxSPayment.row, 8), "dd/mm/yyyy")
      End If
   End If
   SelTxtInCtrl txtPayDt
End Sub

Private Sub txtPayDt_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
      flxSPayment.TextMatrix(iCurRow, 8) = txtPayDt.text
      If IsDate(txtPayDt.text) = False Then
            ShowMsgInTaskBar "Date format is not correct!"
            Exit Sub 'added by anol 20160911
      End If
      txtPayDt.Visible = False
      flxSPayment.ScrollBars = flexScrollBarVertical
   End If
   TextBoxKeyPrsDate txtPayDt, KeyAscii
End Sub

Private Sub txtPayDt_LostFocus()
        If txtPayDt.text <> "" Then TextBoxFormatDate txtPayDt
        If txtPayDt.text = "" Then
            flxSPayment.TextMatrix(iCurRow, 8) = Format(Date, "dd/MM/yyyy")
        Else
            flxSPayment.TextMatrix(iCurRow, 8) = txtPayDt.text
        End If
        txtPayDt.Visible = False
'        If flxSPayment.TextMatrix(iCurRow, 8) <> "" Then
'            flxSPayment.TextMatrix(iCurRow, 8) = flxSPayment.TextMatrix(flxSPayment.row, 8)
'        End If
End Sub
Private Sub flxSPayment_dblClick()
    If flxSPayment.col = 8 Then
      iCurRow = flxSPayment.row
      'flxSPayment.col = iFlxSPayCol
      txtPayDt.Top = flxSPayment.CellTop + flxSPayment.Top
     ' iTop = txtSPayment.Top
      txtPayDt.Left = flxSPayment.CellLeft + flxSPayment.Left
      'iLeft = flxSPayment.CellLeft + flxSPayment.Left
      txtPayDt.Width = 1200 'flxSPayment.ColWidth(iFlxSPayCol)
      txtPayDt.Height = flxSPayment.RowHeight(flxSPayment.row) - 15
      txtPayDt.text = flxSPayment.TextMatrix(flxSPayment.row, 8)
      txtPayDt.Visible = True
      'txtPayDt.ScrollBars = flexScrollBarNone
      txtPayDt.SetFocus
   End If
End Sub

