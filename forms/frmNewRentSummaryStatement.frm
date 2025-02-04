VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmNewRentSummaryStatement 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "New Rent Summary Statement"
   ClientHeight    =   11235
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14130
   FillColor       =   &H00FFFFFF&
   BeginProperty Font 
      Name            =   "Myriad Web"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmNewRentSummaryStatement.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   11235
   ScaleWidth      =   14130
   Begin VB.Frame Frame1 
      Caption         =   "Produce Rent Summary Statement"
      Height          =   10770
      Index           =   6
      Left            =   0
      TabIndex        =   0
      Top             =   135
      Width           =   14055
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxRetensionDetails 
         Height          =   2100
         Left            =   2655
         TabIndex        =   29
         Top             =   6705
         Visible         =   0   'False
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   3704
         _Version        =   393216
         Cols            =   3
         FixedCols       =   0
         BackColorFixed  =   12632256
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
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   3
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin VB.Frame Frame1 
         Caption         =   "Properties:"
         Height          =   4455
         Index           =   9
         Left            =   135
         TabIndex        =   23
         Top             =   4950
         Width           =   6690
         Begin VB.CheckBox chkAllProperties 
            Caption         =   "All Properties"
            Height          =   255
            Left            =   135
            TabIndex        =   24
            Top             =   240
            Width           =   2025
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxProperties 
            Height          =   3810
            Left            =   135
            TabIndex        =   25
            Top             =   495
            Width           =   6405
            _ExtentX        =   11298
            _ExtentY        =   6720
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
      Begin VB.Frame Frame1 
         Caption         =   "Retention"
         Height          =   870
         Index           =   7
         Left            =   135
         TabIndex        =   20
         Top             =   9405
         Width           =   4785
         Begin VB.TextBox txtRetention 
            Alignment       =   1  'Right Justify
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
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   330
            Left            =   2610
            Locked          =   -1  'True
            MaxLength       =   10
            TabIndex        =   22
            Text            =   "0.00"
            Top             =   225
            Width           =   1485
         End
         Begin VB.CommandButton cmdAddRetention 
            Caption         =   "Add Retention"
            Height          =   375
            Left            =   90
            TabIndex        =   21
            Top             =   225
            Width           =   2310
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Funds:"
         Height          =   4455
         Index           =   10
         Left            =   6930
         TabIndex        =   17
         Top             =   4950
         Width           =   6780
         Begin VB.CheckBox chkInFunds 
            Caption         =   "All Funds"
            Height          =   255
            Left            =   180
            TabIndex        =   18
            Top             =   270
            Width           =   1095
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxInFunds 
            Height          =   3765
            Left            =   120
            TabIndex        =   19
            Top             =   570
            Width           =   6495
            _ExtentX        =   11456
            _ExtentY        =   6641
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
      Begin VB.Frame Frame1 
         Caption         =   "Bank Accounts:"
         Height          =   4185
         Index           =   8
         Left            =   6930
         TabIndex        =   15
         Top             =   765
         Width           =   6780
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxBankAccounts 
            Height          =   3810
            Left            =   90
            TabIndex        =   16
            Top             =   300
            Width           =   6585
            _ExtentX        =   11615
            _ExtentY        =   6720
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
      Begin VB.CommandButton cmdOKInouts 
         Caption         =   "&Preview"
         Height          =   465
         Left            =   11205
         TabIndex        =   14
         Top             =   9495
         Width           =   1170
      End
      Begin VB.TextBox txtClientSearch 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1485
         MaxLength       =   10
         TabIndex        =   13
         Top             =   810
         Width           =   1575
      End
      Begin VB.TextBox txtLastStatementDate1 
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   4905
         MaxLength       =   10
         TabIndex        =   12
         Text            =   "01/01/2000"
         Top             =   450
         Width           =   1575
      End
      Begin VB.TextBox txtStatementDate1 
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1485
         MaxLength       =   10
         TabIndex        =   11
         Text            =   "01/01/2000"
         Top             =   450
         Width           =   1575
      End
      Begin VB.Frame Frame1 
         Caption         =   "Clients:"
         Height          =   3780
         Index           =   12
         Left            =   135
         TabIndex        =   9
         Top             =   1170
         Width           =   6690
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxClients 
            Height          =   3495
            Left            =   135
            TabIndex        =   10
            Top             =   225
            Width           =   6375
            _ExtentX        =   11245
            _ExtentY        =   6165
            _Version        =   393216
            Cols            =   3
            FixedCols       =   0
            BackColorFixed  =   12632256
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
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            _NumberOfBands  =   1
            _Band(0).Cols   =   3
            _Band(0).TextStyleBand=   0
            _Band(0).TextStyleHeader=   0
         End
      End
      Begin VB.CommandButton cmdTestReport 
         Caption         =   "Report For Test Purpose"
         Height          =   420
         Left            =   8730
         TabIndex        =   8
         Top             =   10125
         Visible         =   0   'False
         Width           =   3705
      End
      Begin VB.Frame Frame1 
         Caption         =   "Available  Fund"
         Height          =   870
         Index           =   13
         Left            =   5220
         TabIndex        =   6
         Top             =   9405
         Width           =   1635
         Begin VB.TextBox txtAvailableFunds 
            Alignment       =   1  'Right Justify
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
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   45
            MaxLength       =   10
            TabIndex        =   7
            Text            =   "0.00"
            Top             =   225
            Width           =   1485
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Rent Payable"
         Height          =   870
         Index           =   14
         Left            =   6885
         TabIndex        =   4
         Top             =   9405
         Width           =   1770
         Begin VB.TextBox txtRentPayable 
            Alignment       =   1  'Right Justify
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
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   45
            MaxLength       =   10
            TabIndex        =   5
            Text            =   "0.00"
            Top             =   225
            Width           =   1485
         End
      End
      Begin VB.CommandButton cmdCalculateAvailableFund 
         Caption         =   "Calculate Available Fund"
         Height          =   465
         Left            =   8730
         TabIndex        =   3
         Top             =   9495
         Width           =   2385
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Height          =   465
         Left            =   12600
         TabIndex        =   2
         Top             =   9495
         Width           =   1170
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Close"
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   0
         Left            =   12600
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   10125
         Width           =   1200
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Client (Search)"
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   225
         TabIndex        =   28
         Top             =   810
         Width           =   1035
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Last Statement Date"
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   3285
         TabIndex        =   27
         Top             =   450
         Width           =   1425
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Statement Date"
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   225
         TabIndex        =   26
         Top             =   450
         Width           =   1110
      End
   End
End
Attribute VB_Name = "frmNewRentSummaryStatement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public NO_AGENT_INFO As Boolean

Private szProertyID As String
Private bData As Boolean
Private szSupplierAccount As String
'Private lLastID As Long
Private szaFreq() As String
Private szaTrans(1) As Integer
Private szAgentVATCode As String
Dim szPayableTypes As String
Dim szSelectedClient As String
Dim szSelectedBankAccount As String
Dim bPreviewMode As Boolean
Dim szSelectedStatement As String
Dim szSelectedFund As String
Dim szAvailableFund1 As Double
Dim whichFieldToCheck As String
Dim hasSelProperty As Boolean
Dim hasSelBankAccounts As Boolean
Public szCurrentStatementID As String
Public bEditMode As Boolean
Dim boolConsolidatedStatement As Integer
'Dim szSelectedPayableTypeID As String
'Dim szCurrentRentsummarySTID As String
Private Sub MarkAllTransactionsWithSS(strSSID As String)
    Dim adoconn As New adodb.Connection
    adoconn.Open getConnectionString
    Dim szSQL As String
    'This will update all supplier,M agent,Client,landlord transaction
    szSQL = "Update tlbPayment P,tlbPaymentSplit S,Fund F,Supplier SP SET " & whichFieldToCheck & "='" & strSSID & "'  where " & _
            "P.PDate >=#" & Format(txtLastStatementDate1.text, "dd/mmm/yyyy") & "# AND P.PDate <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "# AND " & _
            "SP.SupplierID=P.SageaccountNumber AND  P.TransactionID=S.PayHeader AND P.TYPE IN(7,8,9) AND S.FundID=F.FundID AND   F.FundCode in (" & _
             ListOfFunds & ") AND (P.RentSumStatement='' OR isnull(P.RentSumStatement)) AND P.ClientID ='" & szSelectedClient & "'"
    adoconn.Execute szSQL
    
    szSQL = "Update tlbReceipt R,tlbReceiptSplit S,Fund F SET " & whichFieldToCheck & "='" & strSSID & "'  where R.TransactionID= S.RptHeader AND TYPE IN(3,4,23) " & _
            "AND S.FundID=F.FundID AND R.BankCODE='" & szSelectedBankAccount & "' and F.FundCode<>'TENANTDEPOSIT'  AND (R.RentSumStatement='' OR isnull(R.RentSumStatement)) " & _
            "AND R.RDate >=#" & Format(txtLastStatementDate1.text, "dd/mmm/yyyy") & "# AND R.RDate <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "# " & _
            "AND ClientID ='" & szSelectedClient & "'"
    adoconn.Execute szSQL
     
    szSQL = "Update tlbBankPayment B, Fund F  SET " & whichFieldToCheck & "='" & strSSID & "'  where B.DEPT_ID=cstr(F.FundID) " & _
            "AND B.TRAN_DATE >=#" & Format(txtLastStatementDate1.text, "dd/mmm/yyyy") & "# AND B.TRAN_DATE <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "# " & _
            "and BANK_AC='" & szSelectedBankAccount & "' AND F.FundCode in (" & ListOfFunds & ") AND (B.RentSumStatement='' OR isnull(B.RentSumStatement)) and clientID='" & szSelectedClient & "' "
            
    adoconn.Execute szSQL
    
    adoconn.Close
    Set adoconn = Nothing
    
    
End Sub
Private Function getClosingBalance(dblLasClosingBalance As Double) As Double
    'Pass propery as parameter for selected property
    'No property spec:
    Dim adoconn As New adodb.Connection
    Dim rsReceipt As New adodb.Recordset
    Dim rsPayment As New adodb.Recordset
    Dim rsBankPaymentAndRcpt As New adodb.Recordset
    Dim dblAmt As Double
    adoconn.Open getConnectionString
    Dim szSQL As String
    Dim AccrualsControl As String
    'we are not using property filter here
    'B )***********************  Sum of Rent received Paid/Refunded ***********************************


    szSQL = "Select  SUM(switch(type=3,S.Amount,type=4,S.Amount,type=23,-S.Amount)) as DR from tlbReceipt R,tlbReceiptSplit S,Fund F where R.TransactionID= S.RptHeader AND TYPE IN(3,4,23) " & _
    "AND S.FundID=F.FundID AND R.BankCODE='" & szSelectedBankAccount & "' and F.FundCode<>'TENANTDEPOSIT'  AND (R.RentSumStatement='' or isnull(R.RentSumStatement)) AND ClientID ='" & szSelectedClient & "' " & _
    "AND R.RDate >=#" & Format(txtLastStatementDate1.text, "dd/mmm/yyyy") & "# AND R.RDate <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "#"
    rsReceipt.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
    If Not rsReceipt.EOF Then
            dblAmt = IIf(IsNull(rsReceipt.Fields.Item("Dr").Value), 0, rsReceipt.Fields.Item("Dr").Value)
             'result is 255
    End If
    rsReceipt.Close
    Set rsReceipt = Nothing
    getClosingBalance = dblLasClosingBalance + dblAmt
 
   'c   (-): Sum of Supplier amounts Paid/Refunded (Both allocated and unallocated)
 
    szSQL = "Select  SUM(SWITCH(P.TYPE=24,S.Amount,P.TYPE=8,-S.Amount,P.TYPE=9,-S.Amount)) as AMT from tlbPayment P,tlbPaymentSplit S,Fund F,Supplier SP where  " & _
            "SP.SupplierID=P.SageAccountNumber AND P.TransactionID=S.PayHeader AND P.TYPE IN(8,9,24) AND (P.RentSumStatement='' or isnull(P.RentSumStatement)) AND " & _
            "P.PDate >=#" & Format(txtLastStatementDate1.text, "dd/mmm/yyyy") & "# AND P.PDate <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "# AND " & _
            "S.FundID=F.FundID and  P.BankCODE='" & szSelectedBankAccount & "' and  F.FundCode in (" & ListOfFunds & ") AND ClientID ='" & szSelectedClient & "' "
    rsPayment.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
    If Not rsPayment.EOF Then
            dblAmt = IIf(IsNull(rsPayment.Fields.Item("AMT").Value), 0, rsPayment.Fields.Item("AMT").Value)
        'result is -50
    End If
    rsPayment.Close
    Set rsPayment = Nothing
    getClosingBalance = getClosingBalance + dblAmt

    'd)  Add (+): Sum of Bank payments and receipts
   
     szSQL = "Select  SUM(SWITCH(TransactionType=11 ,B.NET_AMOUNT,TransactionType=12 ,-B.NET_AMOUNT)) as AMT from tlbBankPayment B, Fund F  where B.DEPT_ID=cstr(F.FundID) " & _
            "and BANK_AC='" & szSelectedBankAccount & "' AND F.FundCode in (" & ListOfFunds & ") AND (B.RentSumStatement='' OR isnull(B.RentSumStatement)) and clientID='" & szSelectedClient & "' " & _
            "AND B.TRAN_DATE >=#" & Format(txtLastStatementDate1.text, "dd/mmm/yyyy") & "# AND B.TRAN_DATE <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "#"
    rsBankPaymentAndRcpt.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
    If Not rsBankPaymentAndRcpt.EOF Then
        dblAmt = IIf(IsNull(rsBankPaymentAndRcpt.Fields.Item("AMT").Value), 0, rsBankPaymentAndRcpt.Fields.Item("AMT").Value)
           'result is 0
    End If
    rsBankPaymentAndRcpt.Close
    Set rsBankPaymentAndRcpt = Nothing
 End Function
Private Function getAvailablefunds(dblLasClosingBalance As Double) As Double
    'Pass propery as parameter for selected property
    'No property spec:
    'Exit Function
    Dim adoconn As New adodb.Connection
    Dim rsReceipt As New adodb.Recordset
    Dim rsPayment As New adodb.Recordset
    Dim rsBankPaymentAndRcpt As New adodb.Recordset
    Dim dblAmt As Double
    adoconn.Open getConnectionString
    Dim szSQL As String
    Dim AccrualsControl As String
    Dim whereProperty As String
    'we are not using property filter here
    'B )***********************  Sum of Rent received Paid/Refunded ***********************************
    'AND tlbReceipt.RDate >=#" & Format(txtLastStatementDate1.text, "dd/mmm/yyyy") & "# AND tlbReceipt.RDate <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "#"
    Dim rsGlobalData1 As New adodb.Recordset
                 Dim bolVatOptionEnabled As Boolean
                 Dim bolOptedTotax As String
                 Dim bolisAgentToSubmit As Boolean
                 Dim strManagingAgentID As String
                 rsGlobalData1.Open "Select vatOptionEnabled,isAgentToSubmit from Globaldata G,Property P where P.PropertyID=G.PropertyID AND P.PropertyID='" & _
                                    szSQL & "' ", adoconn, adOpenStatic, adLockReadOnly
                                    
                If Not rsGlobalData1.EOF Then
                        bolVatOptionEnabled = rsGlobalData1("vatOptionEnabled").Value
                        bolisAgentToSubmit = rsGlobalData1("isAgentToSubmit").Value
                End If
                rsGlobalData1.Close
'                rsGlobalData1.Open "SELECT optedTotax,* FROM Supplier where supplierID='" & strManagingAgentID & "'", adoConn, adOpenStatic, adLockReadOnly
'                If Not rsGlobalData1.EOF Then
'                        bolOptedTotax = rsGlobalData1("optedTotax").Value
'                End If
'                rsGlobalData1.Close
                
    If boolConsolidatedStatement = 1 Then
            whereProperty = "(U.PropertyID in (" & ListOfProperties & ")OR isnull(U.PropertyID) OR U.PropertyID='' ) AND "
    Else
            whereProperty = "U.PropertyID in (" & ListOfProperties & ") AND "
    End If
    'take the  VAT amount from the allocation table
    szSQL = "Select  SUM(switch(type=3,S.Amount,type=4,S.Amount,type=23,-S.Amount)) as DR from tlbReceipt R,tlbReceiptSplit S,Fund F, Units U where R.TransactionID= S.RptHeader AND TYPE IN(3,4,23) " & _
    "AND S.FundID=F.FundID AND R.BankCODE='" & szSelectedBankAccount & "' and F.FundCode<>'TENANTDEPOSIT'  AND (R.RentSumStatement='' or isnull(R.RentSumStatement)) AND ClientID ='" & szSelectedClient & "' " & _
    "AND R.UnitID=U.UnitNumber and  F.FundCode in (" & ListOfFunds & ") AND " & whereProperty & " R.RDate >=#" & Format(txtLastStatementDate1.text, "dd/mmm/yyyy") & "# AND R.RDate <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "#"
    rsReceipt.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
    If Not rsReceipt.EOF Then
            dblAmt = IIf(IsNull(rsReceipt.Fields.Item("Dr").Value), 0, rsReceipt.Fields.Item("Dr").Value)
             'result is 175
    End If
    rsReceipt.Close
    Set rsReceipt = Nothing
    getAvailablefunds = dblLasClosingBalance + dblAmt
    'Vat calculation of B)
'    If bolisAgentToSubmit = True Then
            szSQL = "Select  SUM(AL.VatAmount) as DR from tlbReceipt R,rptTransactions AL,tlbReceiptSplit S,Fund F, Units U,GLobalData G where G.PropertyID=U.PropertyID AND G.isAgentToSubmit=true AND R.TransactionID= S.RptHeader AND TYPE IN(3,4,23) " & _
            "AND S.FundID=F.FundID AND R.BankCODE='" & szSelectedBankAccount & "' and F.FundCode<>'TENANTDEPOSIT'  AND (R.RentSumStatement='' or isnull(R.RentSumStatement)) AND ClientID ='" & szSelectedClient & "' " & _
            "AND R.UnitID=U.UnitNumber and AL.Deleteflag=false and AL.FromTran=R.TransactionID AND  F.FundCode in (" & ListOfFunds & ") AND " & whereProperty & " R.RDate >=#" & Format(txtLastStatementDate1.text, "dd/mmm/yyyy") & "# AND R.RDate <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "#"
            rsReceipt.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
            If Not rsReceipt.EOF Then
                    dblAmt = IIf(IsNull(rsReceipt.Fields.Item("Dr").Value), 0, rsReceipt.Fields.Item("Dr").Value)
                     'result is 175
            End If
            rsReceipt.Close
            Set rsReceipt = Nothing
            getAvailablefunds = getAvailablefunds - dblAmt
'    End If
 
   'c   (-): Sum of Supplier amounts Paid/Refunded ( allocated /purchase payment
    If boolConsolidatedStatement = 1 Then
            whereProperty = "(PR.PropertyID in (" & ListOfProperties & ") OR isnull(PR.PropertyID)) AND "
    Else
            whereProperty = "PR.PropertyID in (" & ListOfProperties & ") AND "
    End If
    'we need to filter this sql by selected fund
'    szSQL = "Select  SUM(SWITCH(P.TYPE=24,S.Amount,P.TYPE=8,-S.Amount,P.TYPE=9,-S.Amount)) as AMT from tlbPayment P,tlbPaymentSplit S,Paytransactions PS,Fund F,Supplier SP,Property PR where  " & _
'            "PS.ToTran=P.TransactionID AND SP.SupplierID=P.SageAccountNumber AND PR.propertyID=p.UNITID AND P.TransactionID=S.PayHeader AND P.TYPE IN(6,7) AND (P.RentSumStatement='' or isnull(P.RentSumStatement)) AND " & _
'            "S.FundID=F.FundID and " & whereProperty & " P.BankCODE='" & szSelectedBankAccount & "' and  F.FundCode in (" & ListOfFunds & ") AND P.ClientID ='" & szSelectedClient & "' " & _
'            "AND P.PDate >=#" & Format(txtLastStatementDate1.text, "dd/mmm/yyyy") & "# AND P.PDate <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "#"
 szSQL = "Select  SUM(SWITCH(P.TYPE=24,S.Amount,P.TYPE=8,-S.Amount,P.TYPE=9,-S.Amount)) as AMT from tlbPayment P,tlbPaymentSplit S,Paytransactions PS,Fund F,Supplier SP,tlbPayment P1,Property PR  where  " & _
            "PS.FromTran=P.TransactionID AND SP.SupplierID=P.SageAccountNumber   AND PR.propertyID=p1.UNITID AND P.TransactionID=S.PayHeader AND P1.TYPE IN(6,7) AND  PS.TOTran=P1.TransactionID " & _
            "AND (P.RentSumStatement='' or isnull(P.RentSumStatement)) AND " & _
            "S.FundID=F.FundID and " & whereProperty & " P.BankCODE='" & szSelectedBankAccount & "' and  F.FundCode in (" & ListOfFunds & ") AND P.ClientID ='" & szSelectedClient & "' " & _
            "AND P.PDate >=#" & Format(txtLastStatementDate1.text, "dd/mmm/yyyy") & "# AND P.PDate <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "#"
    rsPayment.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
    If Not rsPayment.EOF Then
            dblAmt = IIf(IsNull(rsPayment.Fields.Item("AMT").Value), 0, rsPayment.Fields.Item("AMT").Value)
        'result is -837
    End If
    rsPayment.Close
    Set rsPayment = Nothing
    getAvailablefunds = getAvailablefunds + dblAmt
'Take the vat amount from the allocation table
'    If bolisAgentToSubmit = True Then
        szSQL = "Select  SUM(AL.VatAmount)  as AMT from tlbPayment P,PayTransactions AL,tlbPaymentSplit S,Fund F,Supplier SP,Property PR,GLobalData G  where  G.PropertyID=PR.PropertyID AND isAgentToSubmit=true  AND " & _
                "SP.SupplierID=P.SageAccountNumber AND AL.Deleteflag=false and PR.propertyID=p.UNITID AND P.TransactionID=S.PayHeader AND P.TYPE IN(8,9,24) AND (P.RentSumStatement='' or isnull(P.RentSumStatement)) AND " & _
                "S.FundID=F.FundID AND AL.FromTran=P.transactionID and " & whereProperty & " P.BankCODE='" & szSelectedBankAccount & "' and  F.FundCode in (" & ListOfFunds & ") AND P.ClientID ='" & szSelectedClient & "' " & _
                "AND P.PDate >=#" & Format(txtLastStatementDate1.text, "dd/mmm/yyyy") & "# AND P.PDate <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "#"
        rsPayment.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
        If Not rsPayment.EOF Then
                dblAmt = IIf(IsNull(rsPayment.Fields.Item("AMT").Value), 0, rsPayment.Fields.Item("AMT").Value)
                'result is -15
        End If
        rsPayment.Close
        Set rsPayment = Nothing
         getAvailablefunds = getAvailablefunds - dblAmt
'    End If
    'd)  Add (+): Sum of Bank payments and receipts
    If boolConsolidatedStatement = 1 Then
            whereProperty = "(B.PropertyID in (" & ListOfProperties & ") OR isnull(B.PropertyID) OR B.PropertyID ='' ) AND "
    Else
            whereProperty = "B.PropertyID in (" & ListOfProperties & ") AND "
    End If
    
     szSQL = "Select  SUM(SWITCH(TransactionType=11 ,-B.NET_AMOUNT,TransactionType=12 ,B.NET_AMOUNT)) as AMT from tlbBankPayment B, Fund F  where B.DEPT_ID=cstr(F.FundID) " & _
            "and " & whereProperty & " BANK_AC='" & szSelectedBankAccount & "' AND F.FundCode in (" & ListOfFunds & ") AND (B.RentSumStatement='' OR isnull(B.RentSumStatement)) and clientID='" & szSelectedClient & "' " & _
            "AND B.TRAN_DATE >=#" & Format(txtLastStatementDate1.text, "dd/mmm/yyyy") & "# AND B.TRAN_DATE <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "#"
    rsBankPaymentAndRcpt.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
    If Not rsBankPaymentAndRcpt.EOF Then
        dblAmt = IIf(IsNull(rsBankPaymentAndRcpt.Fields.Item("AMT").Value), 0, rsBankPaymentAndRcpt.Fields.Item("AMT").Value)
           'result is 0
    End If
    rsBankPaymentAndRcpt.Close
    Set rsBankPaymentAndRcpt = Nothing
    getAvailablefunds = getAvailablefunds + dblAmt
    'f)  Less (-): Supplier OS Account balances for the client selected
    dblAmt = GetSupplierOSAmount
    'If negative then ignore this
    getAvailablefunds = getAvailablefunds - IIf(dblAmt < 0, 0, dblAmt)
    'it should be -40

    Dim rsNLposting As New adodb.Recordset

'g)  Less (-): Client /Landlord OS balances for the client selected  and property selected amounts due to Client/Landlord not paid
         dblAmt = GetClientACBalance
         'COMING -35  dblAmt is negative then ignore
    'getAvailablefunds = getAvailablefunds + GetClientACBalance + GetLandLordACBalance
        getAvailablefunds = getAvailablefunds + IIf(dblAmt < 0, 0, dblAmt)
        dblAmt = GetLandLordACBalance
          'if dblAmt is negative then ignore
        getAvailablefunds = getAvailablefunds + IIf(dblAmt < 0, 0, dblAmt)
     
    'h)  Less (-): Managing Agent OS Balances for the client selected Management Fees due but not paid
    dblAmt = GetAgentBalance
    getAvailablefunds = getAvailablefunds - IIf(dblAmt < 0, 0, dblAmt)
    
    Debug.Print getAvailablefunds
    
    
    Dim ManagementFeeControl As String
    AccrualsControl = GetNominalCodeForControlAccount(adoconn, "Accruals Control Account (B/S)", szSelectedClient)
    'Dim rsNLPosting As New ADODB.Recordset
     If boolConsolidatedStatement = 1 Then
            whereProperty = "(PROPERTY_ID IN (" & ListOfProperties & ") OR isnull(PROPERTY_ID)) AND "
    Else
            whereProperty = "PROPERTY_ID in (" & ListOfProperties & ") AND "
    End If
    
    rsNLposting.Open "Select sum(AMOUNT) as AMT from NLPosting where " & whereProperty & _
                    "  NOMINAL_CODE='" & AccrualsControl & "' AND ClientID='" & _
                    szSelectedClient & "' AND DeleteFlag=false", adoconn, adOpenStatic, adLockReadOnly
    
    If Not rsNLposting.EOF Then
        dblAmt = IIf(IsNull(rsNLposting.Fields.Item("AMT").Value), 0, rsNLposting.Fields.Item("AMT").Value)
    End If
    rsNLposting.Close
    Set rsNLposting = Nothing
    getAvailablefunds = getAvailablefunds + dblAmt
    'j)  Less (-): Tenant Deposits received for the client selected
    'REMOVE THIS AS PER SPEC
    'getAvailablefunds = getAvailablefunds - GetRentDeposit
    
    getAvailablefunds = getAvailablefunds - Val(txtRetention.text)
    MsgBox "Available fund is: " & getAvailablefunds
    txtAvailableFunds.text = getAvailablefunds
    txtRentPayable.text = txtAvailableFunds.text
     
End Function
Private Sub GenerateSummaryStatement(szStatmentID As String)
'    Frame1(6).Visible = True
'    Frame1(6).Top = 135
'    Frame1(6).Left = 2025
    Dim adoconn As New adodb.Connection
    Dim rsRentSummaryStatement As New adodb.Recordset
    adoconn.Open getConnectionString
    Dim dblLasClosingBalance As Double
    Dim szSQL As String
    szSQL = "Select StatementClosingBal from RentSummaryStatement where StatementNo=" & GetLastStatementNoByClient & " AND ClientIDLandlordID='" & _
    szSelectedClient & "'"
    rsRentSummaryStatement.Open szSQL, adoconn, adOpenKeyset, adLockReadOnly
    If Not rsRentSummaryStatement.EOF Then
        dblLasClosingBalance = rsRentSummaryStatement!StatementClosingBal
    End If
   
    rsRentSummaryStatement.Close
    Set rsRentSummaryStatement = Nothing
    
    'szStatementID
    rsRentSummaryStatement.Open "Select * from RentSummaryStatement where 1=2", adoconn, adOpenDynamic, adLockOptimistic
    With rsRentSummaryStatement
            .AddNew
            !StatementID = szStatmentID 'we are setting this column atutomatically
             szCurrentStatementID = szStatmentID
            !StatementNo = GetLastStatementNoByClient + 1
            !ClientIDLandlordID = szSelectedClient
            !BankCode = szSelectedBankAccount
            !PreviousStatementDate = IIf(GetLastStatementDateByClient = "", Null, GetLastStatementDateByClient) 'This is Fromdate
            !StatementDate = Format(txtStatementDate1.text, "dd/mmmm/yyyy") 'This is todate
            !StatementOpBal = dblLasClosingBalance
            !Retentions = Val(txtRetention.text) 'we need to further analyse detail/add/deduct retension
            !Clearretentions = False 'Will need to come again
            
            !AccrualsAcBalance = GetAccrualsControlBalance
            !SupplierAcBalance = GetSupplierOSAmount 'GetBalance("Supplier") 'GetBalanceSupplier'wrong
            !ManagingAgentAcBalance = GetAgentBalance 'GetBalance("Agent") 'GetBalanceAgent'wrong
            !ClientACBalance = GetClientACBalance
            !LandLordACBalance = GetLandLordACBalance
            !ListOffundID = ListOfFundsForDBSave ' szSelectedFund
'            !ListOfPayableTypeID = ListOfPayableTypesForDBSave ' ListOfPayableTypes
            !TenantDepositsReceived = GetRentDeposit()
            !Availablefunds = getAvailablefunds(dblLasClosingBalance)
            !ListOfinputProperties = ListOfProperties
            !PaymentsonAccount = -GetPaymentsonAccount
            'New fields added 2021-01-24
            !TenantReceipts = GetTenantReceipts
            !SupplierPayments = GetSupplierPayment
            !BankPaymentReceipts = GetBankPaymentReceipts
            !ClientLandlordBalance = GetClientACBalance + GetLandLordACBalance
            !ClientPayments = GetClientPayments
            !LandlordPayments = GetLandLordPayments
            !ManagingAgentPayments = GetAGENTPayments
            !PayableAmount = txtRentPayable.text
            !StatementClosingBal = getClosingBalance(dblLasClosingBalance)
            !PINumber = ""
            !Generated_Date = Format(Now, "dd/mmmm/yyyy")
            !Printed = False
            !Emailed = False
            !Invoiced = False
            !PostTohistory = False
            .Update
    End With
    rsRentSummaryStatement.Close
    Set rsRentSummaryStatement = Nothing
    Call SaveRetentionDetails(adoconn)
    Call frmRentPayable.loadflxPayFees
    adoconn.Close
    Set adoconn = Nothing
End Sub
Private Sub SaveRetentionDetails(adoconn As adodb.Connection)
        Dim rsRetensionDetails As New adodb.Recordset
        If szCurrentStatementID = "" Then Exit Sub
        adoconn.Execute "Delete from RetentionDetails where statementID =" & szCurrentStatementID & ""
        'Enter data into grid only memory version
        Dim iRow As Integer
        rsRetensionDetails.Open "Select * from RetentionDetails where 1=2", adoconn, adOpenDynamic, adLockOptimistic
        For iRow = 1 To flxRetensionDetails.Rows - 1
            If flxRetensionDetails.TextMatrix(iRow, 3) <> "" Then
                If Val(flxRetensionDetails.TextMatrix(iRow, 4)) > 0 Then
                    With rsRetensionDetails
                            .AddNew
                            !StatementID = szCurrentStatementID
                            !SlNumber = iRow
                            !description = flxRetensionDetails.TextMatrix(iRow, 3)
                            !amount = Val(flxRetensionDetails.TextMatrix(iRow, 4))
                            .Update
                    End With
                End If
            End If
       Next iRow
        'then load everything to the grid
'        Call LoadflxRetensionDetails
End Sub
'Public Sub loadflxPayFees()
'        Call configflxPayFees
'        Dim adoconn As New ADODB.Connection
'        Dim rsRentSummaryStatement As New ADODB.Recordset
'        Dim i As Long
'        adoconn.Open getConnectionString
'            rsRentSummaryStatement.Open "Select * from RentSummaryStatement R,Supplier S where " & _
'                    "R.ClientIDLandlordID=S.SupplierID Order by StatementID", adoconn, adOpenDynamic, adLockOptimistic
'            i = 1
'            With rsRentSummaryStatement
'            While Not rsRentSummaryStatement.EOF
'
'                frmRentPayable.flxPayFees.AddItem ""
'                frmRentPayable.flxPayFees.TextMatrix(i, 1) = "SS" & !StatementID
'                frmRentPayable.flxPayFees.TextMatrix(i, 2) = !StatementNo
'                frmRentPayable.flxPayFees.TextMatrix(i, 3) = !ClientIDLandlordID
'                frmRentPayable.flxPayFees.TextMatrix(i, 4) = !BankCode
'                frmRentPayable.flxPayFees.TextMatrix(i, 5) = IIf(IsNull(!PreviousStatementDate), "", Format(!PreviousStatementDate, "dd/MM/yyyy")) 'check null here
'                frmRentPayable.flxPayFees.TextMatrix(i, 6) = !StatementDate
'                frmRentPayable.flxPayFees.TextMatrix(i, 7) = !StatementOpBal
'                frmRentPayable.flxPayFees.TextMatrix(i, 8) = !Retentions
'                frmRentPayable.flxPayFees.TextMatrix(i, 9) = !AccrualsAcBalance
'                frmRentPayable.flxPayFees.TextMatrix(i, 10) = IIf(IsNull(!SupplierAcBalance), 0, !SupplierAcBalance)
'                frmRentPayable.flxPayFees.TextMatrix(i, 11) = IIf(IsNull(!ManagingAgentAcBalance), 0, !ManagingAgentAcBalance)
'                frmRentPayable.flxPayFees.TextMatrix(i, 12) = IIf(IsNull(!ClientACBalance), 0, !ClientACBalance)
'                frmRentPayable.flxPayFees.TextMatrix(i, 13) = IIf(IsNull(!LandLordACBalance), 0, !LandLordACBalance)
'                frmRentPayable.flxPayFees.TextMatrix(i, 14) = IIf(IsNull(!TenantDepositsReceived), 0, !TenantDepositsReceived)
'                frmRentPayable.flxPayFees.TextMatrix(i, 15) = !Availablefunds
'                frmRentPayable.flxPayFees.TextMatrix(i, 16) = !PaymentsonAccount
'                frmRentPayable.flxPayFees.TextMatrix(i, 17) = !PayableAmount
'                frmRentPayable.flxPayFees.TextMatrix(i, 18) = !StatementClosingBal
'                frmRentPayable.flxPayFees.TextMatrix(i, 19) = IIf(IsNull(!Generated_Date), "", Format(!Generated_Date, "dd/MM/yyyy")) '!Generated_Date
'                frmRentPayable.flxPayFees.TextMatrix(i, 20) = !Printed
'                frmRentPayable.flxPayFees.TextMatrix(i, 21) = !Emailed
'                frmRentPayable.flxPayFees.TextMatrix(i, 22) = !Invoiced
'                rsRentSummaryStatement.MoveNext
'                i = i + 1
'            Wend
'            End With
'            adoconn.Close
'            Set adoconn = Nothing
'End Sub
Private Sub configflxPayFees()
        Dim szHeader As String
        frmRentPayable.flxPayFees.Clear
        szHeader$ = "|<StatementID|<StatementNo|<ClientIDLandlordID|<BankCode|<PreviousStatementDate|<StatementDate|<StatementOpBal|<Retentions" & _
            "|>AccrualsACBalance|<SupplierACBalance|>ManagingAgentACBalance|>ClientACBalance|<LandLordACBalance|<TenantDepositsreceived|<Availablefunds|<PaymentsonAccount" & _
            " |<PayableAmount|<StatementClosingBal|<Generated_Date|<Printed|<Emailed|<Invoiced"
        frmRentPayable.flxPayFees.FormatString = szHeader$
        'frmRentPayable.flxPayFees.Clear
        frmRentPayable.flxPayFees.Cols = 23
        frmRentPayable.flxPayFees.Rows = 2
        'frmRentPayable.flxPayFees.RowHeight(0) = 0
        frmRentPayable.flxPayFees.ColWidth(0) = 220
        frmRentPayable.flxPayFees.ColWidth(1) = 1500 'Label5(101).Left - Label5(78).Left'StatementID
        frmRentPayable.flxPayFees.ColWidth(2) = 1500 'Label5(101).Left - Label5(79).Left StatementNo
        frmRentPayable.flxPayFees.ColWidth(3) = 1500 'Label5(101).Left - Label5(80).Left ClientIDLandlordID
        frmRentPayable.flxPayFees.ColWidth(4) = 1500 'Label5(101).Left - Label5(81).Left BankCode
        frmRentPayable.flxPayFees.ColWidth(5) = 1500 'Label5(101).Left - Label5(82).Left PreviousStatementDate
        frmRentPayable.flxPayFees.ColWidth(6) = 1500 'Label5(101).Left - Label5(83).Left StatementDate
        frmRentPayable.flxPayFees.ColWidth(7) = 1500 'Label5(101).Left - Label5(84).Left StatementOpBal
        frmRentPayable.flxPayFees.ColWidth(8) = 1500 'Label5(101).Left - Label5(85).Left Retentions
        frmRentPayable.flxPayFees.ColWidth(9) = 1500 ' Label5(101).Left - Label5(86).Left AccrualsACBalance
        frmRentPayable.flxPayFees.ColWidth(10) = 1500 'Label5(101).Left - Label5(87).Left SupplierACBalance
        frmRentPayable.flxPayFees.ColWidth(11) = 1500 'Label5(101).Left - Label5(88).Left ManagingAgentACBalance
        frmRentPayable.flxPayFees.ColWidth(12) = 1500  'Label5(101).Left - Label5(89).Left ClientACBalance
        frmRentPayable.flxPayFees.ColWidth(13) = 1500  'Label5(101).Left - Label5(90).Left LandLordACBalance
        frmRentPayable.flxPayFees.ColWidth(14) = 1500  'Label5(101).Left - Label5(91).Left TenantDepositsreceived
        frmRentPayable.flxPayFees.ColWidth(15) = 1500  'Label5(101).Left - Label5(92).Left Availablefunds
        frmRentPayable.flxPayFees.ColWidth(16) = 1500  'Label5(101).Left - Label5(93).Left PaymentsonAccount
        frmRentPayable.flxPayFees.ColWidth(17) = 1500  'Label5(101).Left - Label5(94).Left PayableAmount
        frmRentPayable.flxPayFees.ColWidth(18) = 1800  'Label5(101).Left - Label5(95).Left StatementClosingBal
        frmRentPayable.flxPayFees.ColWidth(19) = 1700  'Label5(101).Left - Label5(96).Left Generated_Date
        frmRentPayable.flxPayFees.ColWidth(20) = 1200  'Label5(101).Left - Label5(97).Left !Printed
        frmRentPayable.flxPayFees.ColWidth(21) = 1200  'Label5(101).Left - Label5(98).Left  !Emailed 22 !Invoiced
        frmRentPayable.flxPayFees.ColWidth(21) = 1200  'Label5(101).Left - Label5(99).Left
        frmRentPayable.flxPayFees.ColWidth(22) = 1200  'Label5(101).Left - Label5(100).Left
        
    
End Sub
Private Sub GeneratePreview(szStatmentID As String)
'    Exit Sub
'    Frame1(6).Visible = True
'    Frame1(6).Top = 135
'    Frame1(6).Left = 2025
    
    Dim adoconn As New adodb.Connection
    Dim rsRentSummaryStatement As New adodb.Recordset
'    Dim rsConsolidatedStatement As New ADODB.Recordset
    adoconn.Open getConnectionString
'    rsConsolidatedStatement.Open "Select * from client where clientID ='" & szSelectedClient & "'", adocon, adOpenStatic, adLockReadOnly
'    If Not rsConsolidatedStatement.EOF Then
'        boolConsolidatedStatement = rsConsolidatedStatement("ConsolidatedStatement").Value
'    End If
'    rsConsolidatedStatement.Close
    'You need to make all null values into a empty value for the propertyID for the ease of putting where clause by empty string . You need to do that for that for the three table tlbpayment,receipt and bank receipt
'    adoconn.Execute "Update tlbPayment set UnitID ='' where  isnull(UnitID)" 'this unitID is actually a property ID
'    adoconn.Execute "Update tlbReceipt set UnitID ='' where  isnull(UnitID)" 'after join property ID will  be null so this method will not work
'    adoconn.Execute "Update tlbBankPayment set PropertyID ='' where  isnull(PropertyID)"
    
    Dim dblLasClosingBalance As Double
    Dim szSQL As String
    'Before writing this table you need to delete this table
    If szSelectedFund = "" Then
        MsgBox "Please select a fund", vbInformation, "Warning!"
        Exit Sub
    End If
    adoconn.Execute "Delete from  RentSummaryStatementPreview"
    szSQL = "Select StatementClosingBal from RentSummaryStatement where StatementNo=" & GetLastStatementNoByClient & " AND ClientIDLandlordID='" & szSelectedClient & "'"
    rsRentSummaryStatement.Open szSQL, adoconn, adOpenKeyset, adLockReadOnly
    If Not rsRentSummaryStatement.EOF Then
        dblLasClosingBalance = rsRentSummaryStatement!StatementClosingBal
    End If
    Dim X As String
    rsRentSummaryStatement.Close
    Set rsRentSummaryStatement = Nothing
    rsRentSummaryStatement.Open "Select * from RentSummaryStatementPreview where 1=2", adoconn, adOpenDynamic, adLockOptimistic
    With rsRentSummaryStatement
            .AddNew
            !StatementID = szStatmentID 'we are setting this column atutomatically
            !StatementNo = GetLastStatementNoByClient + 1
            !ClientIDLandlordID = szSelectedClient
            !BankCode = szSelectedBankAccount
            !PreviousStatementDate = IIf(GetLastStatementDateByClient = "", Null, GetLastStatementDateByClient)
            !StatementDate = Format(txtStatementDate1.text, "dd/mmmm/yyyy")
            !StatementOpBal = dblLasClosingBalance
            !Retentions = txtRetention.text 'we need to further analyse detail/add/deduct retension
            !Clearretentions = False 'Will need to come again
            !AccrualsAcBalance = GetAccrualsControlBalance
            !SupplierAcBalance = GetSupplierOSAmount 'GetBalance("Supplier") 'GetBalanceSupplier'wrong'' for consolidated I need not filter it by property but for other I need to filter
            !ManagingAgentAcBalance = GetAgentBalance 'GetBalance("Agent") 'GetBalanceAgent'wrong' for consolidated I need not filter it by property but for other I need to filter
            !ClientACBalance = GetClientACBalance
            !LandLordACBalance = GetLandLordACBalance
            !ListOffundID = ListOfFundsForDBSave
            !ListOfPayableTypeID = ListOfFundsForDBSave ' ListOfPayableTypes
            !ListOfinputProperties = ListOfProperties
            !TenantDepositsReceived = GetRentDeposit()
            !Availablefunds = getAvailablefunds(dblLasClosingBalance)
            !PaymentsonAccount = -GetPaymentsonAccount 'date  filter added
            'New fields added 2021-02-22
   
                !ClientPayments = GetClientPayments
                !LandlordPayments = GetLandLordPayments
                !ManagingAgentPayments = GetAGENTPayments
            
             'New fields added 2021-01-24
            !TenantReceipts = GetTenantReceipts
            !SupplierPayments = GetSupplierPayment 'Purchase payment
            !BankPaymentReceipts = GetBankPaymentReceipts
            'addded newly by anol 2021-08-19
            !BankPayment = GetBankPaymentPreview
            !BankReceipts = GetBankreceiptsPreview
            !BankACBalancePreview = BankAccBalance(adoconn, szSelectedBankAccount, szSelectedClient)
            !ClientLandlordBalance = GetClientACBalance + GetLandLordACBalance
            
            
            !PayableAmount = Val(txtRentPayable.text)
            !StatementClosingBal = BankAccBalance(adoconn, szSelectedBankAccount, szSelectedClient) - !Availablefunds 'getClosingBalance(dblLasClosingBalance)
            !PINumber = ""
            !Generated_Date = Format(Now, "dd/mmmm/yyyy")
            !Printed = False
            !Emailed = False
            !Invoiced = False
            !PostTohistory = False
            .Update
    End With
    rsRentSummaryStatement.Close
    Set rsRentSummaryStatement = Nothing
    adoconn.Close
    Set adoconn = Nothing
End Sub
'Private Function GetClientLandLordBalance() As Double
'    Dim rsPayment As New ADODB.Recordset
'    Dim szSQL As String
'    Dim adoConn As New ADODB.Connection
'    'F.CategoryCode = 1 Fund category 1 Means rent
'    'Implement switch here in SQL
'    adoConn.Open getConnectionString
'
'    szSQL = "SELECT  P.SageAccountNumber,SUM(SWITCH(P.Type = 6,P.Amount, P.Type = 24,P.Amount,P.Type = 7, " & _
'            "-P.Amount,P.Type = 8,-P.Amount,P.Type = 9,-P.Amount)) AS Dr " & _
'            "FROM tlbPayment AS P , Client WHERE  P.SageAccountNumber = Client.ClientID " & _
'            "AND P.ClientID ='" & szSelectedClient & "' " & _
'            "GROUP BY P.SageAccountNumber"
'
'    rsPayment.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'    If Not rsPayment.EOF Then
'        GetClientLandLordBalance = IIf(IsNull(rsPayment.Fields.Item("DR").Value), 0, rsPayment.Fields.Item("DR").Value)
'    End If
'    rsPayment.Close
'    adoConn.Close
'    Set adoConn = Nothing
'End Function
'Private Function GetClientLandLordPA() As Double
'    Dim rsPayment As New ADODB.Recordset
'    Dim szSQL As String
'    Dim adoconn As New ADODB.Connection
'    'F.CategoryCode = 1 Fund category 1 Means rent
'    'Implement switch here in SQL
'    adoconn.Open getConnectionString
'
'    szSQL = "SELECT  P.SageAccountNumber,SUM(SP.Amount) AS Dr " & _
'            "FROM tlbPayment AS P ,tlbPaymentSplit SP, Client,Supplier SS WHERE  SS.SupplierID=P.SageAccountNumber AND SP.Payheader=P.TransactionID AND " & _
'            "P.SageAccountNumber = Client.ClientID AND SS.Type in ('LLORD','CLIENT') " & _
'            "AND P.ClientID ='" & szSelectedClient & "' AND P.Type=9 " & _
'            "GROUP BY P.SageAccountNumber"
'
'    rsPayment.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
'    If Not rsPayment.EOF Then
'        GetClientLandLordPA = IIf(IsNull(rsPayment.Fields.Item("DR").Value), 0, rsPayment.Fields.Item("DR").Value)
'    End If
'    rsPayment.Close
'    adoconn.Close
'    Set adoconn = Nothing
'End Function
Private Function GetBankPaymentPreview() As Double
    Dim rsPayment As New adodb.Recordset
    Dim szSQL As String
    Dim adoconn As New adodb.Connection
    'F.CategoryCode = 1 Fund category 1 Means rent
    'Implement switch here in SQL
    adoconn.Open getConnectionString
    Dim whereProperty As String
    If boolConsolidatedStatement = 1 Then
            whereProperty = "(B.PropertyID in (" & ListOfProperties & ") OR isnull(B.PropertyID) OR B.PropertyID='' ) "
    Else
            whereProperty = "B.PropertyID in (" & ListOfProperties & ") "
    End If
    szSQL = "Select  SUM(SWITCH(TransactionType=12,B.NET_AMOUNT,TransactionType=11,-B.NET_AMOUNT)) as DR from tlbBankPayment B,Fund F where TransactionType IN (11) AND " & _
            "B.TRAN_DATE >=#" & Format(txtLastStatementDate1.text, "dd/mmm/yyyy") & "# AND B.TRAN_DATE <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "# AND " & _
            "B.DEPT_ID=cstr(F.FundID)  AND (B.RentSumStatement=''OR isnull(B.RentSumStatement))  AND B.ClientID ='" & szSelectedClient & "' AND " & whereProperty
    
    rsPayment.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
    If Not rsPayment.EOF Then
        GetBankPaymentPreview = IIf(IsNull(rsPayment.Fields.Item("DR").Value), 0, rsPayment.Fields.Item("DR").Value)
    End If
    rsPayment.Close
    adoconn.Close
    Set adoconn = Nothing
End Function
Private Function GetBankreceiptsPreview() As Double
    Dim rsPayment As New adodb.Recordset
    Dim szSQL As String
    Dim adoconn As New adodb.Connection
    'F.CategoryCode = 1 Fund category 1 Means rent
    'Implement switch here in SQL
    adoconn.Open getConnectionString
    Dim whereProperty As String
    If boolConsolidatedStatement = 1 Then
            whereProperty = "(B.PropertyID in (" & ListOfProperties & ") OR isnull(B.PropertyID) OR B.PropertyID='' ) "
    Else
            whereProperty = "B.PropertyID in (" & ListOfProperties & ") "
    End If
    szSQL = "Select  SUM(SWITCH(TransactionType=12,B.NET_AMOUNT,TransactionType=11,-B.NET_AMOUNT)) as DR from tlbBankPayment B,Fund F where TransactionType IN(12) AND " & _
            "B.TRAN_DATE >=#" & Format(txtLastStatementDate1.text, "dd/mmm/yyyy") & "# AND B.TRAN_DATE <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "# AND " & _
            "B.DEPT_ID=cstr(F.FundID)  AND (B.RentSumStatement=''OR isnull(B.RentSumStatement))  AND B.ClientID ='" & szSelectedClient & "' AND " & whereProperty
    
    rsPayment.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
    If Not rsPayment.EOF Then
        GetBankreceiptsPreview = IIf(IsNull(rsPayment.Fields.Item("DR").Value), 0, rsPayment.Fields.Item("DR").Value)
    End If
    rsPayment.Close
    adoconn.Close
    Set adoconn = Nothing
End Function
'Private Function GetBankACBalance() As Double
'    Dim rsPayment As New ADODB.Recordset
'    Dim szSQL As String
'    Dim adoconn As New ADODB.Connection
'    'F.CategoryCode = 1 Fund category 1 Means rent
'    'Implement switch here in SQL
'    adoconn.Open getConnectionString
'
'    szSQL = "Select  SUM(SWITCH(TransactionType=12,B.NET_AMOUNT,TransactionType=11,-B.NET_AMOUNT)) as DR from tlbBankPayment B,Fund F where TransactionType IN(12) AND " & _
'            "B.TRAN_DATE >=#" & Format(txtLastStatementDate1.text, "dd/mmm/yyyy") & "# AND B.TRAN_DATE <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "# AND " & _
'            "B.DEPT_ID=cstr(F.FundID)  AND (B.RentSumStatement=''OR isnull(B.RentSumStatement))  AND B.ClientID ='" & szSelectedClient & "'"
'
'    rsPayment.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
'    If Not rsPayment.EOF Then
'        GetBankACBalance = IIf(IsNull(rsPayment.Fields.Item("DR").Value), 0, rsPayment.Fields.Item("DR").Value)
'    End If
'    rsPayment.Close
'    adoconn.Close
'    Set adoconn = Nothing
'End Function
Private Function GetBankPaymentReceipts() As Double
    Dim rsPayment As New adodb.Recordset
    Dim szSQL As String
    Dim adoconn As New adodb.Connection
    'F.CategoryCode = 1 Fund category 1 Means rent
    'Implement switch here in SQL
    adoconn.Open getConnectionString

    szSQL = "Select  SUM(SWITCH(TransactionType=12,B.NET_AMOUNT,TransactionType=11,-B.NET_AMOUNT)) as DR from tlbBankPayment B,Fund F where TransactionType IN(11,12) AND " & _
            "B.TRAN_DATE >=#" & Format(txtLastStatementDate1.text, "dd/mmm/yyyy") & "# AND B.TRAN_DATE <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "# AND " & _
            "B.DEPT_ID=cstr(F.FundID)  AND (B.RentSumStatement=''OR isnull(B.RentSumStatement))  AND B.ClientID ='" & szSelectedClient & "'"
    
    rsPayment.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
    If Not rsPayment.EOF Then
        GetBankPaymentReceipts = IIf(IsNull(rsPayment.Fields.Item("DR").Value), 0, rsPayment.Fields.Item("DR").Value)
    End If
    rsPayment.Close
    adoconn.Close
    Set adoconn = Nothing
End Function
Private Function GetLandLordACBalance() As Double   'This function return result as minus This is getting LLORD balance
    Dim rsPayment As New adodb.Recordset
    Dim szSQL As String
    Dim adoconn As New adodb.Connection
    'F.CategoryCode = 1 Fund category 1 Means rent
    'Implement switch here in SQL
    'Bank code does not exits in PI,so do not put it in where clause
    adoconn.Open getConnectionString
    szSQL = "Select  SUM(SWITCH(P.TYPE =24,S.Amount,P.TYPE =6,S.Amount,P.TYPE =7,-S.Amount,P.TYPE =8,-S.Amount,P.TYPE =9,-S.Amount)) as AMT from tlbPayment P,tlbPaymentSplit S,Fund F,Supplier SP where " & _
            " SP.SupplierID=P.SageaccountNumber AND  P.TransactionID=S.PayHeader AND P.TYPE " & _
            "IN(6,7,8,9,24) AND S.FundID=F.FundID AND   F.FundCode in (" & _
             ListOfFunds & ") AND P.ClientID ='" & szSelectedClient & "'  and SP.TYPE in ('LLORD')"
    
    rsPayment.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
    If Not rsPayment.EOF Then
        GetLandLordACBalance = IIf(IsNull(rsPayment.Fields.Item("AMT").Value), 0, rsPayment.Fields.Item("AMT").Value)
    End If
    rsPayment.Close
    adoconn.Close
    Set adoconn = Nothing
End Function
Private Function GetClientACBalance() As Double   'This function return result as minus This is getting CLIENT balance
    Dim rsPayment As New adodb.Recordset
    Dim szSQL As String
    Dim adoconn As New adodb.Connection
    'F.CategoryCode = 1 Fund category 1 Means rent
    'Implement switch here in SQL
    'Bank code does not exits in PI,so do not put it in where clause
    adoconn.Open getConnectionString
    szSQL = "Select  SUM(SWITCH(P.TYPE =24,S.Amount,P.TYPE =6,S.Amount,P.TYPE =7,-S.Amount,P.TYPE =8,-S.Amount,P.TYPE =9,-S.Amount)) as AMT from tlbPayment P,tlbPaymentSplit S,Fund F,Supplier SP where " & _
            " SP.SupplierID=P.SageaccountNumber AND  P.TransactionID=S.PayHeader AND P.TYPE " & _
            "IN(6,7,8,9,24) AND S.FundID=F.FundID AND   F.FundCode in (" & _
             ListOfFunds & ") AND P.ClientID ='" & szSelectedClient & "'  and SP.TYPE in ('CLIENT')"
    
    rsPayment.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
    If Not rsPayment.EOF Then
        GetClientACBalance = IIf(IsNull(rsPayment.Fields.Item("AMT").Value), 0, rsPayment.Fields.Item("AMT").Value)
    End If
    rsPayment.Close
    adoconn.Close
    Set adoconn = Nothing
End Function
Private Function GetAgentBalance() As Double   'This function return result as minus This is getting Agent balance
    Dim rsPayment As New adodb.Recordset
    Dim szSQL As String
    Dim adoconn As New adodb.Connection
    'F.CategoryCode = 1 Fund category 1 Means rent
    'Implement switch here in SQL
    'Bank code does not exits in PI,so do not put it in where clause
    adoconn.Open getConnectionString
    Dim whereProperty As String
    If boolConsolidatedStatement = 1 Then
    'for consolidated I need not filter it by property but for other I need to filter
            'whereProperty = "(P.UNITID IN (" & ListOfProperties & ") OR isnull(P.UNITID)  OR P.UNITID='' ) AND "
    Else
            whereProperty = "P.UNITID in (" & ListOfProperties & ") AND "
    End If
    szSQL = "Select  SUM(SWITCH(P.TYPE =24,S.Amount,P.TYPE =6,S.Amount,P.TYPE =7,-S.Amount,P.TYPE =8,-S.Amount,P.TYPE =9,-S.Amount)) as AMT from tlbPayment P,tlbPaymentSplit S,Fund F,Supplier SP where " & _
            " SP.SupplierID=P.SageaccountNumber AND  " & whereProperty & "  P.TransactionID=S.PayHeader AND P.TYPE " & _
            "IN(6,7,8,9,24) AND S.FundID=F.FundID AND   F.FundCode in (" & _
             ListOfFunds & ") AND P.ClientID ='" & szSelectedClient & "'  and SP.TYPE in ('AGENT')"
    
    rsPayment.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
    If Not rsPayment.EOF Then
        GetAgentBalance = IIf(IsNull(rsPayment.Fields.Item("AMT").Value), 0, rsPayment.Fields.Item("AMT").Value)
    End If
    rsPayment.Close
    adoconn.Close
    Set adoconn = Nothing
End Function
Private Function GetSupplierOSAmount() As Double   'This function return result as minus'This is getting supplier balance
    Dim rsPayment As New adodb.Recordset
    Dim szSQL As String
    Dim adoconn As New adodb.Connection
    'F.CategoryCode = 1 Fund category 1 Means rent
    'Implement switch here in SQL
    'Bank code does not exits in PI,so do not put it in where clause
    adoconn.Open getConnectionString
    Dim whereProperty As String
    ' for consolidated I need not filter it by property but for other I need to filter
    If boolConsolidatedStatement = 1 Then
            'whereProperty = "(P.UNITID IN (" & ListOfProperties & ") OR isnull(P.UNITID)  OR P.UNITID='' ) AND "
    Else
            whereProperty = "P.UNITID in (" & ListOfProperties & ") AND "
    End If
    szSQL = "Select  SUM(SWITCH(P.TYPE =24,S.Amount,P.TYPE =6,S.Amount,P.TYPE =7,-S.Amount,P.TYPE =8,-S.Amount,P.TYPE =9,-S.Amount)) as AMT from tlbPayment P,tlbPaymentSplit S,Fund F,Supplier SP where " & _
            " SP.SupplierID=P.SageaccountNumber AND   " & whereProperty & " P.TransactionID=S.PayHeader AND P.TYPE " & _
            "IN(6,7,8,9,24) AND S.FundID=F.FundID AND   F.FundCode in (" & _
             ListOfFunds & ") AND P.ClientID ='" & szSelectedClient & "'  and SP.TYPE in ('Supplier')"
    
    rsPayment.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
    If Not rsPayment.EOF Then
        GetSupplierOSAmount = IIf(IsNull(rsPayment.Fields.Item("AMT").Value), 0, rsPayment.Fields.Item("AMT").Value)
    End If
    rsPayment.Close
    adoconn.Close
    Set adoconn = Nothing
End Function
Private Function GetClientPayments() As Double   'This function return result as minus'This is getting supplier balance
    Dim rsPayment As New adodb.Recordset
    Dim szSQL As String
    Dim adoconn As New adodb.Connection
    'F.CategoryCode = 1 Fund category 1 Means rent
    'Implement switch here in SQL
    'Bank code does not exits in PI,so do not put it in where clause
    Dim whereProperty As String
    adoconn.Open getConnectionString
    If boolConsolidatedStatement = 1 Then
            whereProperty = "(P.UNITID IN (" & ListOfProperties & ") OR isnull(P.UNITID)  OR P.UNITID='' ) AND "
    Else
            whereProperty = "P.UNITID in (" & ListOfProperties & ") AND "
    End If
    szSQL = "Select  SUM(SWITCH(P.TYPE =24,S.Amount,P.TYPE =6,S.Amount,P.TYPE =7,-S.Amount,P.TYPE =8,-S.Amount,P.TYPE =9,-S.Amount)) as AMT from tlbPayment P,tlbPaymentSplit S,Fund F,Supplier SP where " & _
            " SP.SupplierID=P.SageaccountNumber AND " & whereProperty & "  P.TransactionID=S.PayHeader AND P.TYPE " & _
            "IN(7,8,9) AND S.FundID=F.FundID AND   F.FundCode in (" & _
             ListOfFunds & ") AND P.ClientID ='" & szSelectedClient & "'  and SP.TYPE in ('Client')"
    
    rsPayment.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
    If Not rsPayment.EOF Then
        GetClientPayments = IIf(IsNull(rsPayment.Fields.Item("AMT").Value), 0, rsPayment.Fields.Item("AMT").Value)
    End If
    rsPayment.Close
    adoconn.Close
    Set adoconn = Nothing
End Function
Private Function GetLandLordPayments() As Double   'This function return result as minus'This is getting supplier balance
    Dim rsPayment As New adodb.Recordset
    Dim szSQL As String
    Dim adoconn As New adodb.Connection
    'F.CategoryCode = 1 Fund category 1 Means rent
    'Implement switch here in SQL
    'Bank code does not exits in PI,so do not put it in where clause
    Dim whereProperty As String
    adoconn.Open getConnectionString
      If boolConsolidatedStatement = 1 Then
            whereProperty = "(P.UNITID IN (" & ListOfProperties & ") OR isnull(P.UNITID)  OR P.UNITID='' ) AND "
    Else
            whereProperty = "P.UNITID in (" & ListOfProperties & ") AND "
    End If
    szSQL = "Select  SUM(SWITCH(P.TYPE =24,S.Amount,P.TYPE =6,S.Amount,P.TYPE =7,-S.Amount,P.TYPE =8,-S.Amount,P.TYPE =9,-S.Amount)) as AMT from tlbPayment P,tlbPaymentSplit S,Fund F,Supplier SP where " & _
            " SP.SupplierID=P.SageaccountNumber AND " & whereProperty & "   P.TransactionID=S.PayHeader AND P.TYPE " & _
            "IN(7,8,9) AND S.FundID=F.FundID AND   F.FundCode in (" & _
             ListOfFunds & ") AND P.ClientID ='" & szSelectedClient & "'  and SP.TYPE in ('LLORD')"
    
    rsPayment.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
    If Not rsPayment.EOF Then
        GetLandLordPayments = IIf(IsNull(rsPayment.Fields.Item("AMT").Value), 0, rsPayment.Fields.Item("AMT").Value)
    End If
    rsPayment.Close
    adoconn.Close
    Set adoconn = Nothing
End Function
Private Function GetAGENTPayments() As Double   'This function return result as minus'This is getting supplier balance
    Dim rsPayment As New adodb.Recordset
    Dim szSQL As String
    Dim adoconn As New adodb.Connection
    'F.CategoryCode = 1 Fund category 1 Means rent
    'Implement switch here in SQL
    'Bank code does not exits in PI,so do not put it in where clause
    Dim whereProperty As String
    adoconn.Open getConnectionString
      If boolConsolidatedStatement = 1 Then
            whereProperty = "(P.UNITID IN (" & ListOfProperties & ") OR isnull(P.UNITID)  OR P.UNITID='' ) AND "
    Else
            whereProperty = "P.UNITID in (" & ListOfProperties & ") AND "
    End If
    szSQL = "Select  SUM(SWITCH(P.TYPE =24,S.Amount,P.TYPE =6,S.Amount,P.TYPE =7,-S.Amount,P.TYPE =8,-S.Amount,P.TYPE =9,-S.Amount)) as AMT from tlbPayment P,tlbPaymentSplit S,Fund F,Supplier SP where " & _
            " SP.SupplierID=P.SageaccountNumber AND " & whereProperty & "  P.TransactionID=S.PayHeader AND P.TYPE " & _
            "IN(7,8,9) AND S.FundID=F.FundID AND   F.FundCode in (" & _
             ListOfFunds & ") AND P.ClientID ='" & szSelectedClient & "'  and SP.TYPE in ('AGENT')"
    
    rsPayment.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
    If Not rsPayment.EOF Then
        GetAGENTPayments = IIf(IsNull(rsPayment.Fields.Item("AMT").Value), 0, rsPayment.Fields.Item("AMT").Value)
    End If
    rsPayment.Close
    adoconn.Close
    Set adoconn = Nothing
End Function
Private Function GetSupplierPayment() As Double
    Dim rsPayment As New adodb.Recordset
    Dim rsReceipt As New adodb.Recordset
    Dim szSQL As String
    Dim adoconn As New adodb.Connection
    Dim whereProperty As String
    Dim dblAmt As Double
    'F.CategoryCode = 1 Fund category 1 Means rent
    'Implement switch here in SQL
     If boolConsolidatedStatement = 1 Then
     '(P.UNITID IN (" & ListOfProperties & " ))
            whereProperty = "AND (P.UNITID IN (" & ListOfProperties & ") OR isnull(P.UNITID) OR P.UNITID='') "
    Else
            whereProperty = "AND P.UNITID in (" & ListOfProperties & ") "
    End If
    adoconn.Open getConnectionString
    szSQL = "Select  SUM(SWITCH(P.TYPE =24,S.Amount,P.TYPE =8,-S.Amount,P.TYPE =9,-S.Amount)) as AMT from tlbPayment P,tlbPaymentSplit S,Fund F,Supplier SP where " & _
            "SP.SupplierID=P.SageaccountNumber AND  P.TransactionID=S.PayHeader AND P.TYPE IN(8,9,24) AND " & _
            "P.PDate >=#" & Format(txtLastStatementDate1.text, "dd/mmm/yyyy") & "# AND P.PDate <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "# AND " & _
            "S.FundID=F.FundID AND  P.BankCODE='" & szSelectedBankAccount & "'  AND (P.RentSumStatement='' OR isnull(P.RentSumStatement)) and  F.FundCode in (" & _
             ListOfFunds & ") AND P.ClientID ='" & szSelectedClient & "' AND (SP.Type='Supplier' OR SP.Type='Agent') " & whereProperty
    
    rsPayment.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
    If Not rsPayment.EOF Then
        GetSupplierPayment = IIf(IsNull(rsPayment.Fields.Item("AMT").Value), 0, rsPayment.Fields.Item("AMT").Value)
    End If
    rsPayment.Close
'       szSQL = "Select  SUM(AL.VatAmount) as DR from tlbReceipt R,rptTransactions AL,tlbReceiptSplit S,Fund F, Units U,GLobalData G where G.PropertyID=U.PropertyID AND G.isAgentToSubmit=true AND R.TransactionID= S.RptHeader AND TYPE IN(3,4,23) " & _
'            "AND S.FundID=F.FundID AND R.BankCODE='" & szSelectedBankAccount & "' and F.FundCode<>'TENANTDEPOSIT'  AND (R.RentSumStatement='' or isnull(R.RentSumStatement)) AND ClientID ='" & szSelectedClient & "' " & _
'            "AND R.UnitID=U.UnitNumber and AL.Deleteflag=false and AL.FromTran=R.TransactionID AND  F.FundCode in (" & ListOfFunds & ") AND " & whereProperty & " R.RDate >=#" & Format(txtLastStatementDate1.text, "dd/mmm/yyyy") & "# AND R.RDate <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "#"
'            rsReceipt.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'            If Not rsReceipt.EOF Then
'                    GetSupplierPayment = GetSupplierPayment - IIf(IsNull(rsReceipt.Fields.Item("Dr").Value), 0, rsReceipt.Fields.Item("Dr").Value)
'                     'result is 175
'            End If
'            rsReceipt.Close
'            Set rsReceipt = Nothing
If boolConsolidatedStatement = 1 Then
     '(P.UNITID IN (" & ListOfProperties & " ))
            whereProperty = "AND (G.PropertyID IN (" & ListOfProperties & ") OR isnull(G.PropertyID) OR G.PropertyID='') "
    Else
            whereProperty = "AND G.PropertyID in (" & ListOfProperties & ") "
    End If
             szSQL = "Select  SUM(AL.VatAmount)  as AMT from tlbPayment P,PayTransactions AL,tlbPaymentSplit S,Fund F,Supplier SP,Property PR,GLobalData G  where  G.PropertyID=PR.PropertyID AND isAgentToSubmit=true  AND " & _
                "SP.SupplierID=P.SageAccountNumber AND AL.Deleteflag=false and PR.propertyID=p.UNITID AND P.TransactionID=S.PayHeader AND P.TYPE IN(8,9,24) AND (P.RentSumStatement='' or isnull(P.RentSumStatement)) AND " & _
                "S.FundID=F.FundID AND AL.FromTran=P.transactionID " & whereProperty & " AND P.BankCODE='" & szSelectedBankAccount & "' and  F.FundCode in (" & ListOfFunds & ") AND P.ClientID ='" & szSelectedClient & "' " & _
                "AND P.PDate >=#" & Format(txtLastStatementDate1.text, "dd/mmm/yyyy") & "# AND P.PDate <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "#"
        rsPayment.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
        If Not rsPayment.EOF Then
                GetSupplierPayment = GetSupplierPayment + IIf(IsNull(rsPayment.Fields.Item("AMT").Value), 0, rsPayment.Fields.Item("AMT").Value)
                'result is -15
        End If
        rsPayment.Close
        Set rsPayment = Nothing
         GetSupplierPayment = GetSupplierPayment - dblAmt
    adoconn.Close
    Set adoconn = Nothing
End Function
Private Function GetTenantReceipts() As Double
    Dim rsPayment As New adodb.Recordset
    Dim szSQL As String
    Dim adoconn As New adodb.Connection
    Dim rsReceipt As New adodb.Recordset
    Dim whereProperty As String
    If boolConsolidatedStatement = 1 Then
            whereProperty = "(U.PropertyID in (" & ListOfProperties & ") OR isnull(U.PropertyID)) AND "
    Else
            whereProperty = "U.PropertyID  in (" & ListOfProperties & ") AND "
    End If
    
    adoconn.Open getConnectionString
    szSQL = "Select  SUM(SWITCH(R.TYPE=23,-RS.Amount,R.TYPE=3,RS.Amount,R.TYPE=4,RS.Amount)) as AMT from tlbReceipt R,tlbReceiptSplit RS,Fund F,Units U where " & _
            "R.TransactionID=RS.RptHeader AND R.TYPE IN(3,4,23) AND " & _
            "R.RDate >=#" & Format(txtLastStatementDate1.text, "dd/mmm/yyyy") & "# AND R.RDate <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "# " & _
            "AND RS.FundID=F.FundID AND  R.BankCODE='" & szSelectedBankAccount & "'  AND (R.RentSumStatement='' OR isnull(R.RentSumStatement)) and  F.FundCode in (" & _
             ListOfFunds & ") AND R.UnitID=U.UnitNumber AND " & whereProperty & " R.ClientID ='" & szSelectedClient & "' "
    
    rsPayment.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
    If Not rsPayment.EOF Then
        GetTenantReceipts = IIf(IsNull(rsPayment.Fields.Item("AMT").Value), 0, rsPayment.Fields.Item("AMT").Value)
    End If
    rsPayment.Close

                szSQL = "Select  SUM(AL.VatAmount) as DR from tlbReceipt R,rptTransactions AL,tlbReceiptSplit S,Fund F, Units U,GLobalData G where G.PropertyID=U.PropertyID AND G.isAgentToSubmit=true AND R.TransactionID= S.RptHeader AND TYPE IN(3,4,23) " & _
            "AND S.FundID=F.FundID AND R.BankCODE='" & szSelectedBankAccount & "' and F.FundCode<>'TENANTDEPOSIT'  AND (R.RentSumStatement='' or isnull(R.RentSumStatement)) AND ClientID ='" & szSelectedClient & "' " & _
            "AND R.UnitID=U.UnitNumber and AL.Deleteflag=false and AL.FromTran=R.TransactionID AND  F.FundCode in (" & ListOfFunds & ") AND " & whereProperty & " R.RDate >=#" & Format(txtLastStatementDate1.text, "dd/mmm/yyyy") & "# AND R.RDate <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "#"
            rsReceipt.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
            If Not rsReceipt.EOF Then
                    GetTenantReceipts = GetTenantReceipts - IIf(IsNull(rsReceipt.Fields.Item("Dr").Value), 0, rsReceipt.Fields.Item("Dr").Value)
                     'result is 175
            End If
            rsReceipt.Close
            Set rsReceipt = Nothing
    adoconn.Close
    Set adoconn = Nothing
End Function

Private Function GetPaymentsonAccount() As Double
    Dim rsPayment As New adodb.Recordset
    Dim adoconn As New adodb.Connection
    adoconn.Open getConnectionString
    Dim szSQL As String
    szSQL = "Select  SUM(S.Amount) as AMT from tlbPayment P,tlbPaymentSplit S,Fund F,Supplier SP where  SP.SupplierID=P.SAGEACCOUNTNUMBER AND " & _
            "P.TransactionID=S.PayHeader AND P.PDate >=#" & Format(txtLastStatementDate1.text, "dd/mmm/yyyy") & "# AND P.PDate <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "# AND P.TYPE  " & _
            "IN(9) AND S.FundID=F.FundID AND  P.BankCODE='" & szSelectedBankAccount & "' and  F.FundCode in (" & _
             ListOfFunds & ") AND SP.SupplierID ='" & szSelectedClient & "' AND (P.RentSumStatement='' OR isnull(P.RentSumStatement)) AND SP.Type in ('CLIENT','LLORD') "
    rsPayment.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
    If Not rsPayment.EOF Then
        GetPaymentsonAccount = IIf(IsNull(rsPayment.Fields.Item("AMT").Value), 0, rsPayment.Fields.Item("AMT").Value)
    End If
    rsPayment.Close
    Set rsPayment = Nothing
    adoconn.Close
    Set adoconn = Nothing
End Function




Private Sub chkAllProperties_Click()
    Dim i As Integer
    If chkAllProperties.Value = 0 Then
        For i = 1 To flxProperties.Rows - 1
             flxProperties.TextMatrix(i, 0) = ""
        Next i
        Call ConfigFlxInFunds
    Else
         For i = 1 To flxProperties.Rows - 1
             If flxProperties.TextMatrix(i, 1) <> "" Then
                   flxProperties.TextMatrix(i, 0) = "X"
             End If
        Next i
        Call LoadflxInFunds
    End If
End Sub

Private Sub chkInFunds_Click()
    Dim i As Integer
    If chkInFunds.Value = 0 Then
        For i = 1 To flxInFunds.Rows - 1
             flxInFunds.TextMatrix(i, 0) = ""
        Next i
    Else
         For i = 1 To flxInFunds.Rows - 1
             flxInFunds.TextMatrix(i, 0) = "X"
        Next i
    End If
End Sub

Private Sub cmdAddToGrid_Click()
'            If Val(txtRetensionAmount1.text) = 0 Then
'                    MsgBox "Please enter amount greater than zero", vbInformation, "Warning"
'                    Exit Sub
'            End If
'            flxRetensionDetails.Enabled = True
'            'Enter data into grid only memory version
'            'statementId you shall generate it when you finally save the statement
'            flxRetensionDetails.TextMatrix(flxRetensionDetails.Rows - 1, 0) = IIf(Option1.Value = True, "+", "-")
'            flxRetensionDetails.TextMatrix(flxRetensionDetails.Rows - 1, 2) = flxRetensionDetails.Rows - 1 'This is slNumber
'            flxRetensionDetails.TextMatrix(flxRetensionDetails.Rows - 1, 3) = txtRetentionDescriptions.text 'This is Description
'            flxRetensionDetails.TextMatrix(flxRetensionDetails.Rows - 1, 4) = Format(Val(txtRetensionAmount1.text), "0.00") 'This is amount
'            flxRetensionDetails.AddItem ""
'           ' txtRetensionAmount1.Visible = False
'
'            txtRetensionAmount1.text = "0.00"
'            FocusControl txtRetensionAmount1
'            txtRetensionAmount1.SelStart = 0
'            txtRetensionAmount1.SelLength = Len(txtRetensionAmount1.text)
'            Call MakeSummaryRetention
End Sub

Private Sub cmdAddRetention_Click()
    'frmRetentionDetails.Top = Me.Top + 600
    Dim rCount As Integer
    Dim selRow As Integer
    Dim iIncDec As Integer
    For rCount = 1 To flxClients.Rows - 1
         If flxClients.TextMatrix(rCount, 0) = "X" Then
             iIncDec = iIncDec + 1
             selRow = rCount
         End If
    Next
    If iIncDec <> 1 Then
       MsgBox "Please select a client.", vbInformation + vbOKOnly, "Client Selection"
       Exit Sub
    End If
    frmRetentionDetails.szClienyIDForRetention = flxClients.TextMatrix(selRow, 1)
    'frmRetentionDetails.ConfigflxRetensionDetailsMainMain
    frmRetentionDetails.Top = Me.Top + 500
    frmRetentionDetails.Show
    If flxRetensionDetails.TextMatrix(1, 2) <> "" Then
            Dim iCol As Integer
            Dim iRow As Integer
            
            'We are setting clientID in this form because form is listing the retention details for the client with cumulative balance
            frmRetentionDetails.ConfigflxRetensionDetails
            frmRetentionDetails.flxRetensionDetails.Cols = flxRetensionDetails.Cols
            frmRetentionDetails.flxRetensionDetails.Rows = flxRetensionDetails.Rows
            For iRow = 1 To flxRetensionDetails.Rows - 1
                For iCol = 1 To flxRetensionDetails.Cols - 1
                     frmRetentionDetails.flxRetensionDetails.TextMatrix(iRow, iCol) = flxRetensionDetails.TextMatrix(iRow, iCol)
                Next
            Next
    End If
    frmRetentionDetails.ZOrder 0
End Sub

Private Sub cmdCalculateAvailableFund_Click()
    If ListOfProperties = "" Then
         MsgBox "Please select a Property", vbInformation, "Property!!!"
         flxProperties.SetFocus
         Exit Sub
    End If
    If ListOfFunds = "" Then
         MsgBox "Please select a fund", vbInformation, "Fund!!!"
         flxProperties.SetFocus
         Exit Sub
    End If
    If szSelectedBankAccount = "" Then
        MsgBox "Please select a Bank account", vbInformation, "Warning "
        Exit Sub
    End If
    Dim adoconn As New adodb.Connection
    Dim rsRentSummaryStatement As New adodb.Recordset
    adoconn.Open getConnectionString
    Dim dblLasClosingBalance As Double
    Dim szSQL As String
    'Before writing this table you need to delete this table
    adoconn.Execute "Delete from  RentSummaryStatementPreview"
    szSQL = "Select StatementClosingBal from RentSummaryStatement where StatementNo=" & GetLastStatementNoByClient & " AND ClientIDLandlordID='" & szSelectedClient & "'"
    rsRentSummaryStatement.Open szSQL, adoconn, adOpenKeyset, adLockReadOnly
    If Not rsRentSummaryStatement.EOF Then
        dblLasClosingBalance = rsRentSummaryStatement!StatementClosingBal
    End If
    rsRentSummaryStatement.Close
    Set rsRentSummaryStatement = Nothing
    adoconn.Close
    Set adoconn = Nothing
    txtAvailableFunds.text = Format(getAvailablefunds(dblLasClosingBalance), "0.00")
    txtRentPayable.text = txtAvailableFunds.text
End Sub


Private Sub cmdClose_Click(Index As Integer)
         Unload Me
End Sub

Private Sub cmdGenAll_Click()
    
End Sub

Private Sub cmdClose1_Click()
    Frame1(6).Visible = False
End Sub

Private Sub cmdGenReport_Click() 'produce clientsummarystatement
    'This table shall write into the table clientsummarystatement
    'In this table StatementID is the primary key , Detail shall be loaded from the tlbPayment,tlbReceipt and tlbBankPayment and receipt
    
    '******************** Inputs for Populating this table**************************
    '1.Take input form the input frame
    '2. Then read tlbPayable
    
    
    
End Sub
'Private Function GetControlAccountForPayable(adoconn As ADODB.Connection, szSelectedPayableTypeID As String) As Boolean
'    Dim rsPayableTypes As New ADODB.Recordset
'    rsPayableTypes.Open "Select * from  PayableTypes where ID=" & szSelectedPayableTypeID & "", adoconn, adOpenStatic, adLockReadOnly
'    If Not rsPayableTypes.EOF Then
'            GetControlAccountForPayable = rsPayableTypes!isUseControlAccount 'PayNCAmt
'    End If
'    rsPayableTypes.Close
'    Set rsPayableTypes = Nothing
'
'End Function

Private Function GetControlAccountForPayableString(adoconn As adodb.Connection, szSelectedPayableTypeID As String) As Boolean
    Dim rsPayableTypes As New adodb.Recordset
    rsPayableTypes.Open "Select * from  PayableTypes where ID=" & szSelectedPayableTypeID & "", adoconn, adOpenStatic, adLockReadOnly
    If Not rsPayableTypes.EOF Then
            GetControlAccountForPayableString = rsPayableTypes!PayNCAmt 'PayNCAmt
    End If
    rsPayableTypes.Close
    Set rsPayableTypes = Nothing

End Function

Private Sub cmdOKInouts_Click()
    'validaton
    'at least one bank , one fund, one property ,one payable type is selected.
    'this procedure shall make visible only that
    If Val(txtRentPayable.text) > Val(txtAvailableFunds.text) Then
        MsgBox "Rent Payable amount cannot be greater than the Available funds", vbInformation, "Warning!"
        Exit Sub
    End If
    If DateDiff("d", txtStatementDate1.text, txtLastStatementDate1.text) > 0 Then
        MsgBox "Current statment date cannot be greater than last statement date", vbInformation, "Statement Date!!!"
        Exit Sub
    End If
    'Validation for Client
    Dim rCount As Integer
    Dim selRow As Integer
    Dim iIncDec As Long
    For rCount = 1 To flxClients.Rows - 1
         If flxClients.TextMatrix(rCount, 0) = "X" Then
             iIncDec = iIncDec + 1
             selRow = rCount
         End If
    Next
    If iIncDec <> 1 Then
       MsgBox "Please select a client.", vbInformation + vbOKOnly, "Client Selection"
       Exit Sub
    End If
    
    
    iIncDec = 0
    If ListOfProperties = "" Then
         MsgBox "Please select a Property", vbInformation, "Property!!!"
         flxProperties.SetFocus
         Exit Sub
    End If
    
    'Validation for Property
    For rCount = 1 To flxProperties.Rows - 1
         If flxProperties.TextMatrix(rCount, 0) = "X" Then
             iIncDec = iIncDec + 1
             selRow = rCount
         End If
    Next
    If iIncDec <> 1 And boolConsolidatedStatement = 0 Then
       MsgBox "Please select one property.", vbInformation + vbOKOnly, "Property Selection"
       Exit Sub
    End If
    
    
    
    If ListOfFunds = "" Then
         MsgBox "Please select a fund", vbInformation, "Fund!!!"
         flxProperties.SetFocus
         Exit Sub
    End If
    If szSelectedBankAccount = "" Then
        MsgBox "Please select a Bank account", vbInformation, "Warning!"
        Exit Sub
    End If
'    If ListOfPayableTypes = "" Then
'        MsgBox "Please select a Payable Type", vbInformation, "Warning!"
'        Exit Sub
'    End If
            Dim szStatmentID As String
            szStatmentID = GetLastStatementID + 1
            whichFieldToCheck = "RentSumStatementPreview"
            Call GeneratePreview(szStatmentID)
            Call MarkAllTransactionsWithSS(szStatmentID)
            'run TestReportForRentSummary.rpt
            Dim reportApp As New CRAXDRT.Application
            Dim Report As CRAXDRT.Report
            
            Set Report = reportApp.OpenReport(App.Path & szReportPath & "\RentPayablePreview.rpt")
            
            Report.EnableParameterPrompting = False
            Report.DiscardSavedData
            Report.ParameterFields(1).AddCurrentValue CInt(szStatmentID)
            
            '               Report.ParameterFields(1).AddCurrentValue CStr(txtLLID.text)
            '               Report.ParameterFields(2).AddCurrentValue CDate(txtFromDate.text)
            '               Report.ParameterFields(3).AddCurrentValue CDate(txtToDate.text)
            '                Report.ParameterFields(4).AddCurrentValue cboCategory.text
            Load frmReport
            frmReport.LoadReportViewer Report
           ' Unload Me
'            Frame1(6).Visible = False
End Sub
'Private Function ListOfPayableTypesForDBSave() As String
'   Dim i As Integer
'
'   For i = 1 To flxPayableTypes.Rows - 1
'      If flxPayableTypes.TextMatrix(i, 0) = "X" Then
'         ListOfPayableTypesForDBSave = ListOfPayableTypesForDBSave & "" & flxPayableTypes.TextMatrix(i, 1) & ","
'      End If
'   Next i
'   If Len(ListOfPayableTypesForDBSave) > 0 Then ListOfPayableTypesForDBSave = Left(ListOfPayableTypesForDBSave, Len(ListOfPayableTypesForDBSave) - 1)
'End Function
'Private Function ListOfPayableTypes() As String
'   Dim i As Integer
'
'   For i = 1 To flxPayableTypes.Rows - 1
'      If flxPayableTypes.TextMatrix(i, 0) = "X" Then
'         ListOfPayableTypes = ListOfPayableTypes & " " & flxPayableTypes.TextMatrix(i, 1) & ", "
'      End If
'   Next i
'   If Len(ListOfPayableTypes) > 0 Then ListOfPayableTypes = Left(ListOfPayableTypes, Len(ListOfPayableTypes) - 2)
'End Function
Private Function ListOfProperties() As String
   Dim i As Integer
  ' ListOfProperties = "''," ' This shall always include No Property
   For i = 1 To flxProperties.Rows - 1
      If flxProperties.TextMatrix(i, 0) = "X" Then
         ListOfProperties = ListOfProperties & " '" & flxProperties.TextMatrix(i, 1) & "', "
      End If
   Next i
   If Len(ListOfProperties) > 0 Then ListOfProperties = Left(ListOfProperties, Len(ListOfProperties) - 2)
End Function
'Private Function ListOfProperties2(rsCheck As ADODB.Recordset) As String
'    'This function not seems to be completed
'    Dim i As Integer
'    While Not rsCheck.EOF
'        ListOfProperties2 = ListOfProperties & " '" & flxProperties.TextMatrix(i, 1) & "', "
'        rsCheck.MoveNext
'    Wend
'    If Len(ListOfProperties2) > 0 Then ListOfProperties2 = Left(ListOfProperties2, Len(ListOfProperties2) - 2)
'End Function

Private Function ListOfFunds() As String
    Dim i As Integer
    For i = 1 To flxInFunds.Rows - 1
         If flxInFunds.TextMatrix(i, 0) = "X" Then
            'structure of grid
'            flxInFunds.TextMatrix(iRow, 1) = adoRst!fundID
'            flxInFunds.TextMatrix(iRow, 2) = adoRst!FundName
'            flxInFunds.TextMatrix(iRow, 3) = adoRst!FundCode
                ListOfFunds = ListOfFunds & "'" & flxInFunds.TextMatrix(i, 3) & "', "
         End If
    Next
    If Len(ListOfFunds) > 0 Then ListOfFunds = Left(ListOfFunds, Len(ListOfFunds) - 2)
End Function
Private Function ListOfFundsForDBSave() As String
    Dim i As Integer
    For i = 1 To flxInFunds.Rows - 1
         If flxInFunds.TextMatrix(i, 0) = "X" Then
            'structure of grid
'            flxInFunds.TextMatrix(iRow, 1) = adoRst!fundID
'            flxInFunds.TextMatrix(iRow, 2) = adoRst!FundName
'            flxInFunds.TextMatrix(iRow, 3) = adoRst!FundCode
                ListOfFundsForDBSave = ListOfFundsForDBSave & "" & flxInFunds.TextMatrix(i, 1) & ","
         End If
    Next
    If Len(ListOfFundsForDBSave) > 0 Then ListOfFundsForDBSave = Left(ListOfFundsForDBSave, Len(ListOfFundsForDBSave) - 1)
End Function

Private Function GetBalanceSupplier() As Double
    Dim PurchaseLedgerControl As String
    Dim dblAmt As Double
    Dim adoconn As New adodb.Connection
    adoconn.Open getConnectionString
    PurchaseLedgerControl = GetNominalCodeForControlAccount(adoconn, "Purchase Ledger Control", szSelectedClient)
    Dim rsNLposting As New adodb.Recordset
    rsNLposting.Open "Select sum(AMOUNT) as dr from NLPosting where ACCOUNT_NUMBER ='" & _
                    szSelectedBankAccount & "'  AND NOMINAL_CODE='" & PurchaseLedgerControl & "' AND ClientID='" & _
                    szSelectedClient & "' ", adoconn, adOpenStatic, adLockReadOnly
    
    If Not rsNLposting.EOF Then
        dblAmt = IIf(IsNull(rsNLposting.Fields.Item("Dr").Value), 0, rsNLposting.Fields.Item("Dr").Value)
    End If
    rsNLposting.Close
    Set rsNLposting = Nothing
    adoconn.Close
    Set adoconn = Nothing
    GetBalanceSupplier = dblAmt
End Function
Private Function GetBalanceAgent() As Double
    Dim ManagementFeesControl As String
    Dim dblAmt As Double
    Dim adoconn As New adodb.Connection
    adoconn.Open getConnectionString
    ManagementFeesControl = GetNominalCodeForControlAccount(adoconn, "Management Fees Control Account (B/S)", szSelectedClient)
    Dim rsNLposting As New adodb.Recordset
    rsNLposting.Open "Select sum(AMOUNT) as dr from NLPosting where ACCOUNT_NUMBER ='" & _
                    szSelectedBankAccount & "'  AND NOMINAL_CODE='" & ManagementFeesControl & "' AND ClientID='" & _
                    szSelectedClient & "' ", adoconn, adOpenStatic, adLockReadOnly
    
    If Not rsNLposting.EOF Then
        dblAmt = IIf(IsNull(rsNLposting.Fields.Item("Dr").Value), 0, rsNLposting.Fields.Item("Dr").Value)
    End If
    rsNLposting.Close
    Set rsNLposting = Nothing
    adoconn.Close
    Set adoconn = Nothing
    GetBalanceAgent = dblAmt
End Function

Private Function GetRentDeposit() As Double
    Dim szSQL As String
'    Dim szSQL1 As String
    Dim szSQL2 As String
'    Dim szSQL3 As String
    Dim rsPayment As New adodb.Recordset
    Dim rsReceipt1 As New adodb.Recordset
    Dim rsReceipt2 As New adodb.Recordset
    Dim rsReceipt3 As New adodb.Recordset
    Dim rsReceipt As New adodb.Recordset
    Dim adoconn As New adodb.Connection
    Dim dblAmt, dblamt1, dblamt2, dblamt3 As Double
    adoconn.Open getConnectionString
    'tlbBankPayment
    'BANK_AC
    'TRAN_TYPE
    'DEPT_ID
    'propertyID
    'clientID
    'NET_AMOUNT
    Dim whereProperty  As String
    If boolConsolidatedStatement = 1 Then
            whereProperty = "(U.PROPERTYID IN (" & ListOfProperties & ") OR isnull(U.PROPERTYID) or U.PROPERTYID='' ) AND "
    Else
            whereProperty = "U.PROPERTYID  in (" & ListOfProperties & ") AND "
    End If
    
    'From unit Id i Ned to build a relation with the selected properties
    szSQL = "Select  SUM(SWITCH(TYPE=1,R.Amount,TYPE=2,R.Amount,TYPE=3,-R.Amount,TYPE=4,-R.Amount,TYPE=23,-R.Amount)) as DR from tlbReceipt R,Fund F, Units U " & _
            "where R.UnitID=U.UnitNumber AND " & whereProperty & "  TYPE IN(1,2,3,4,23) AND R.FundID=F.FundID and F.FundCode='RENTDEPOSIT' AND ClientID ='" & _
             szSelectedClient & "'"
'    szSQL1 = "Select  SUM(R.Amount)  as DR from tlbReceipt R,Fund F where TYPE IN(3,4,23) AND R.FundID=F.FundID and F.FundCode='RENTDEPOSIT' AND ClientID ='" & _
'              szSelectedClient & "'"
    'PropertyID
    If boolConsolidatedStatement = 1 Then
            whereProperty = "(B.PROPERTYID IN (" & ListOfProperties & ") OR isnull(B.PROPERTYID) or B.PROPERTYID='' ) AND "
    Else
            whereProperty = "B.PROPERTYID  in (" & ListOfProperties & ") AND "
    End If
    
    
    szSQL2 = "Select  SUM(SWITCH(TransactionType=11,B.NET_AMOUNT,TransactionType=12,-B.NET_AMOUNT)) as DR from tlbBankPayment B,Fund F where TransactionType " & _
            " IN(11,12) AND " & whereProperty & " B.DEPT_ID=cstr(F.FundID) and " & _
            "F.FundCode='RENTDEPOSIT' AND B.ClientID ='" & szSelectedClient & "'"
'    szSQL3 = "Select  SUM(B.NET_AMOUNT)  as DR from tlbBankPayment B,Fund F where TransactionType IN(12) AND B.DEPT_ID= cstr(F.FundID) and " & _
'            "F.FundCode='RENTDEPOSIT' AND B.ClientID ='" & szSelectedClient & "'"
   
   ' vat Deducting  on reciept
'            szSQL = "Select  SUM(AL.VatAmount) as DR from tlbReceipt R,rptTransactions AL,tlbReceiptSplit S,Fund F, Units B,GLobalData G where G.PropertyID=B.PropertyID AND G.isAgentToSubmit=true AND R.TransactionID= S.RptHeader AND TYPE IN(3,4,23) " & _
'            "AND S.FundID=F.FundID AND R.BankCODE='" & szSelectedBankAccount & "' and F.FundCode<>'TENANTDEPOSIT'  AND (R.RentSumStatement='' or isnull(R.RentSumStatement)) AND ClientID ='" & szSelectedClient & "' " & _
'            "AND R.UnitID=B.UnitNumber and AL.Deleteflag=false and AL.FromTran=R.TransactionID AND  F.FundCode in (" & ListOfFunds & ") AND " & whereProperty & " R.RDate >=#" & Format(txtLastStatementDate1.text, "dd/mmm/yyyy") & "# AND R.RDate <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "#"
'            rsReceipt.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'            If Not rsReceipt.EOF Then
'                    dblAmt = dblAmt - IIf(IsNull(rsReceipt.Fields.Item("Dr").Value), 0, rsReceipt.Fields.Item("Dr").Value)
'                     'result is 175
'            End If
'            rsReceipt.Close
'            Set rsReceipt = Nothing
            
            
    
    rsReceipt.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
    If Not rsReceipt.EOF Then
        dblAmt = IIf(IsNull(rsReceipt.Fields.Item("Dr").Value), 0, rsReceipt.Fields.Item("Dr").Value)
    End If
    rsReceipt.Close
'    rsReceipt1.Open szSQL1, adoconn, adOpenStatic, adLockReadOnly
'    If Not rsReceipt1.EOF Then
'         dblamt1 = IIf(IsNull(rsReceipt1.Fields.Item("Dr").Value), 0, rsReceipt1.Fields.Item("Dr").Value)
'    End If
'    rsReceipt1.Close
    
    rsReceipt2.Open szSQL2, adoconn, adOpenStatic, adLockReadOnly
    If Not rsReceipt2.EOF Then
        dblamt2 = IIf(IsNull(rsReceipt2.Fields.Item("Dr").Value), 0, rsReceipt2.Fields.Item("Dr").Value)
    End If
    rsReceipt2.Close
    
'    rsReceipt3.Open szSQL3, adoconn, adOpenStatic, adLockReadOnly
'    If Not rsReceipt3.EOF Then
'        dblamt3 = IIf(IsNull(rsReceipt3.Fields.Item("Dr").Value), 0, rsReceipt3.Fields.Item("Dr").Value)
'    End If
'    rsReceipt3.Close
'    dblamt2 = 0
'    dblamt1 = 0
'    dblAmt = 0
    GetRentDeposit = dblAmt + dblamt2
    adoconn.Close
    Set adoconn = Nothing
End Function
Private Function GetBalance(szType As String) As Double
    'szType is the suppplier type from supplier table
   Dim szSQL   As String
   Dim szSqlPI As String
   Dim szSQLSI As String
   Dim i       As Integer
   Dim iSI     As Integer
   Dim iPI     As Integer
   Dim iIndex  As Integer
   Dim adoconn As New adodb.Connection
   Dim adoPayDr As New adodb.Recordset, adoPayCr As New adodb.Recordset
   Dim adoRptDr As New adodb.Recordset, adoRptCr As New adodb.Recordset
   adoconn.Open getConnectionString
   Dim szaClientBal(1, 1) As String

   szSQL = "SELECT  SUM(P.Amount) AS Dr " & _
           "FROM tlbPayment AS P, Client C, Supplier S " & _
           "WHERE (P.Type = 6 OR P.Type = 24) AND C.ClientID=S.SupplierID AND P.SageAccountNumber = C.ClientID " & _
           "and  C.ClientID='" & szSelectedClient & "' AND S.Type='" & szType & "'  "

   adoPayDr.Open szSQL, adoconn, adOpenStatic, adLockReadOnly

   iIndex = 0
   While Not adoPayDr.EOF
      szaClientBal(1, iIndex) = IIf(IsNull(adoPayDr.Fields.Item("Dr").Value), 0, adoPayDr.Fields.Item("Dr").Value)
      adoPayDr.MoveNext
   Wend
   adoPayDr.Close

   szSQL = "SELECT  SUM(P.Amount) AS Cr " & _
           "FROM tlbPayment AS P, Client C,  Supplier S " & _
           "WHERE P.Type <> 6 AND P.Type <> 24 AND P.SageAccountNumber = C.ClientID and  C.ClientID='" & szSelectedClient & "' " & _
           "AND C.ClientID=S.SupplierID  AND S.Type='" & szType & "'"

   adoPayCr.Open szSQL, adoconn, adOpenStatic, adLockReadOnly

   While Not adoPayCr.EOF
         szaClientBal(1, iIndex) = IIf(IsNull(adoPayCr.Fields.Item("Cr").Value), 0, adoPayCr.Fields.Item("Cr").Value) 'adoPayCr.Fields.Item("Cr").Value
         adoPayCr.MoveNext
   Wend

   adoPayCr.Close
   GetBalance = szaClientBal(1, 0)
   Set adoRptDr = Nothing
   Set adoRptCr = Nothing
   adoconn.Close
   Set adoconn = Nothing
End Function
Private Function GetManagingAgentACBalance() As Double
    Dim adoconn As New adodb.Connection
    Dim rsAccrualsControlBalance As New adodb.Recordset
    adoconn.Open getConnectionString
    Dim szSQL As String
'    szSQL = "Select Sum(AMOUNT) as SumAmount from NLPOSTING where NOMINAL_CODE='" & _
'            NominalCode & "' AND ClientID='" & szSelectedClient & "' AND PROPERTY_ID in (" & ListOfProperties & ")"
    rsAccrualsControlBalance.Open szSQL, adoconn, adOpenKeyset, adLockReadOnly
    If Not rsAccrualsControlBalance.EOF Then
        GetManagingAgentACBalance = rsAccrualsControlBalance("SumAmount").Value
    End If
    rsAccrualsControlBalance.Close
    Set rsAccrualsControlBalance = Nothing
End Function
Private Function GetAccrualsControlBalance() As Double
    'include no property when  calculating accruals
    Dim adoconn As New adodb.Connection
    Dim rsAccrualsControlBalance As New adodb.Recordset
    adoconn.Open getConnectionString
    Dim szSQL As String
    Dim AccrualsCode As String
    If ListOfProperties = "'" Then Exit Function
    AccrualsCode = GetNominalCodeForControlAccount(adoconn, "Accruals Control Account (B/S)", szSelectedClient)
    
    szSQL = "Select Sum(AMOUNT) as SumAmount from NLPOSTING where NOMINAL_CODE='" & AccrualsCode & "' AND ClientID='" & szSelectedClient & "' AND PROPERTY_ID in (" & ListOfProperties & ")"
    rsAccrualsControlBalance.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
    If Not rsAccrualsControlBalance.EOF Then
        GetAccrualsControlBalance = IIf(IsNull(rsAccrualsControlBalance("SumAmount").Value), 0, rsAccrualsControlBalance("SumAmount").Value)
    End If
    rsAccrualsControlBalance.Close
    Set rsAccrualsControlBalance = Nothing
End Function

'Private Function getAvailablefunds() As Double
'    Dim intmaxStatementNo As Integer
'    Dim adoconn As New ADODB.Connection
'    Dim rsRentSummaryStatement As New ADODB.Recordset
'    adoconn.Open getConnectionString
'    Dim szSQL As String
'    'Rent Summary Statement Opening Balance (=Closing balance of previous statement)
'    'This is by client
'    'Get ID by Client max ID from RentSummaryStatement
'    szSQL = "Select max(StatementNo) as IDbyCL from RentSummaryStatement where ClientID='" & szSelectedClient & "'"
'    rsRentSummaryStatement.Open szSQL, adoconn, adOpenKeyset, adLockReadOnly
'    If Not rsRentSummaryStatement.EOF Then
'        getAvailablefunds = rsRentSummaryStatement!IDbyCL
'    End If
'    rsRentSummaryStatement.Close
'    Set rsRentSummaryStatement = Nothing
'    adoconn.Close
'    Set adoconn = Nothing
'End Function
Private Function GetLastStatementNoByClient() As Integer
    Dim intmaxStatementNo As Integer
    Dim adoconn As New adodb.Connection
    Dim rsRentSummaryStatement As New adodb.Recordset
    adoconn.Open getConnectionString
    Dim szSQL As String
    'Rent Summary Statement Opening Balance (=Closing balance of previous statement)
    'This is by client
    'Get ID by Client max ID from RentSummaryStatement
    szSQL = "Select max(StatementNo) as IDbyCL from RentSummaryStatement where ClientIDLandlordID='" & szSelectedClient & "'"
    rsRentSummaryStatement.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
    If Not rsRentSummaryStatement.EOF Then
        GetLastStatementNoByClient = IIf(IsNull(rsRentSummaryStatement!IDbyCL), 0, rsRentSummaryStatement!IDbyCL)
    End If
    rsRentSummaryStatement.Close
    Set rsRentSummaryStatement = Nothing
    adoconn.Close
    Set adoconn = Nothing
End Function
Private Function GetLastStatementID() As Long 'this is not by client
    Dim intmaxStatementNo As Integer
    Dim adoconn As New adodb.Connection
    Dim rsRentSummaryStatement As New adodb.Recordset
    adoconn.Open getConnectionString
    Dim szSQL As String
    'Rent Summary Statement Opening Balance (=Closing balance of previous statement)
    'This is by client
    'Get ID by Client max ID from RentSummaryStatement
    szSQL = "Select max(StatementID) as IDbyCL from RentSummaryStatement"
    rsRentSummaryStatement.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
    If Not rsRentSummaryStatement.EOF Then
        GetLastStatementID = IIf(IsNull(rsRentSummaryStatement!IDbyCL), 0, rsRentSummaryStatement!IDbyCL)
    End If
    rsRentSummaryStatement.Close
    Set rsRentSummaryStatement = Nothing
    adoconn.Close
    Set adoconn = Nothing
End Function
Private Function GetLastStatementDateByClient() As String
    Dim intmaxStatementNo As Integer
    Dim adoconn As New adodb.Connection
    Dim rsRentSummaryStatement As New adodb.Recordset
    adoconn.Open getConnectionString
    Dim szSQL As String
    'Rent Summary Statement Opening Balance (=Closing balance of previous statement)
    'This is by client
    'Get ID by Client max ID from RentSummaryStatement
    szSQL = "Select StatementDate from RentSummaryStatement where ClientIDLandlordID='" & szSelectedClient & "' order by StatementNo Desc"
    rsRentSummaryStatement.Open szSQL, adoconn, adOpenKeyset, adLockReadOnly
    If Not rsRentSummaryStatement.EOF Then
        GetLastStatementDateByClient = rsRentSummaryStatement!StatementDate
    End If
    rsRentSummaryStatement.Close
    Set rsRentSummaryStatement = Nothing
    adoconn.Close
    Set adoconn = Nothing
End Function
Private Sub cmdPostRP_Click()
   If MsgBox("Do you want to post to the history?", vbQuestion + vbYesNo, "Post") = vbNo Then Exit Sub
   Dim adoconn As New adodb.Connection
   Dim adoRst As adodb.Recordset
   Dim sSQLQuery As String

   adoconn.Open getConnectionString
   Set adoRst = New adodb.Recordset

   sSQLQuery = "UPDATE tblPurInv " & _
               "SET tblPurInv.HISTORY = TRUE " & _
               "WHERE tblPurInv.TTP = " & CByte(TransactionTakePlace("TTP", "RENT PAYABLE", adoconn)) & " AND " & _
                  "tblPurInv.HISTORY = FALSE AND tblPurInv.UPDATE_SAGE = TRUE"
   adoRst.Open sSQLQuery, adoconn, adOpenStatic, adLockReadOnly
   
   Set adoRst = Nothing
   Set adoconn = Nothing
End Sub





Private Sub cmdPrintThis_Click()
   Dim iIncDec As Long
   iIncDec = 0
   Dim rCount As Integer
   Dim selRow As Integer
   For rCount = 1 To frmRentPayable.flxPayFees.Rows - 1
        If frmRentPayable.flxPayFees.TextMatrix(rCount, 0) = "X" Then
            iIncDec = iIncDec + 1
            selRow = rCount
        End If
   Next
   If iIncDec < 1 Then
      MsgBox "Please select one statement only.", vbInformation + vbOKOnly, "statement Selection"
      Exit Sub
   End If
'            whichFieldToCheck = "RentSumStatement"
'            Call GeneratePreview(szCurrentStatementID)
            'Call MarkAllTransactionsWithSS(szCurrentStatementID)
            szCurrentStatementID = frmRentPayable.flxPayFees.TextMatrix(selRow, 1)
            'run TestReportForRentSummary.rpt
            Dim reportApp As New CRAXDRT.Application
            Dim Report As CRAXDRT.Report
            
            Set Report = reportApp.OpenReport(App.Path & szReportPath & "\TestReportForRentSummary.rpt")
            
            Report.EnableParameterPrompting = False
            Report.DiscardSavedData
            Report.ParameterFields(1).AddCurrentValue CInt(Right(szCurrentStatementID, Len(szCurrentStatementID) - 2))
            
            '               Report.ParameterFields(1).AddCurrentValue CStr(txtLLID.text)
            '               Report.ParameterFields(2).AddCurrentValue CDate(txtFromDate.text)
            '               Report.ParameterFields(3).AddCurrentValue CDate(txtToDate.text)
            '                Report.ParameterFields(4).AddCurrentValue cboCategory.text
            Load frmReport
            frmReport.LoadReportViewer Report
End Sub


'Private Sub cmdSavePI_Click()
'      If txtStatementDate1.text = "" Then
'            MsgBox "Please enter statement date", vbInformation, "Warning"
'            Exit Sub
'      End If
'      Dim lSlNumber As Long
'      Dim adoConn As New ADODB.Connection
'      Dim adoPIHeader As New ADODB.Recordset
'      Dim szSQL As String
'      adoConn.Open getConnectionString
'      szSQL = "SELECT * FROM tblPurInv"
'      lSlNumber = SlNumber("PI", "tblPurInv", adoConn)
'      With adoPIHeader
'                .Open szSQL, adoConn, adOpenDynamic, adLockPessimistic
'                .AddNew
'                .Fields.Item("SlNumber").Value = lSlNumber
'                .Fields.Item("SUPP_AC").Value = szSelectedClient
'                .Fields.Item("TRAN_DATE").Value = Format(txtStatementDate1.text, "DD/MMMM/YYYY")
'                .Fields.Item("TransactionType").Value = 6
'                .Fields.Item("INV_NO").Value = szSelectedStatement 'txtInv(0).text
'                .Fields.Item("TOTAL_AMOUNT").Value = CCur(txtRentPayable.text)
'                .Fields.Item("TTP").Value = "PURCHASE INVOICE" 'CByte(TransactionTakePlace("TTP", "PURCHASE INVOICE", adoconn))
'                .Fields.Item("History").Value = False
'                .Fields.Item("TrfPayment").Value = False
'                .Fields.Item("PropertyID").Value = ""
'                .Fields.Item("CL_ID").Value = szSelectedClient
'                .Fields.Item("NLPost").Value = False
'                .Fields.Item("DueDate").Value = Format(txtStatementDate1.text, "DD/MMMM/YYYY")
'                .Fields.Item("PostingDate").Value = Format(txtStatementDate1.text, "DD/MMMM/YYYY")
'                .Update
'      End With
'      adoConn.Close
'      Set adoConn = Nothing
'End Sub

Private Sub cmdSave_Click()
        'validation
    'at least one bank , one fund, one property ,one payable type is selected.
    'this procedure shall make visible only that
     'Validation for Client
    Dim rCount As Integer
    Dim selRow As Integer
    Dim iIncDec As Long
    If Val(txtRentPayable.text) > Val(txtAvailableFunds.text) Then
        MsgBox "Rent Payable amount cannot be greater than the Available funds", vbInformation, "Warning!"
        Exit Sub
    End If
    
    If DateDiff("d", txtStatementDate1.text, txtLastStatementDate1.text) > 0 Then
        MsgBox "Current statment date cannot be greater than last statement date", vbInformation, "Statement Date!!!"
        Exit Sub
    End If
    
    For rCount = 1 To flxClients.Rows - 1
         If flxClients.TextMatrix(rCount, 0) = "X" Then
             iIncDec = iIncDec + 1
             selRow = rCount
         End If
    Next
    If iIncDec <> 1 Then
       MsgBox "Please select a client.", vbInformation + vbOKOnly, "Client Selection"
       Exit Sub
    End If
    
    
    iIncDec = 0
    If ListOfProperties = "" Then
         MsgBox "Please select a Property", vbInformation, "Property!!!"
         flxProperties.SetFocus
         Exit Sub
    End If
    
    'Validation for Property
    For rCount = 1 To flxProperties.Rows - 1
         If flxProperties.TextMatrix(rCount, 0) = "X" Then
             iIncDec = iIncDec + 1
             selRow = rCount
         End If
    Next
    If iIncDec < 1 Then
       MsgBox "Please select a property.", vbInformation + vbOKOnly, "Property Selection"
       Exit Sub
    End If
    If ListOfProperties = "" Then
         MsgBox "Please select a Property", vbInformation, "Property!!!"
         flxProperties.SetFocus
         Exit Sub
    End If
    If ListOfFunds = "" Then
         MsgBox "Please select a fund", vbInformation, "Fund!!!"
         flxProperties.SetFocus
         Exit Sub
    End If
    For rCount = 1 To flxBankAccounts.Rows - 1
         If flxBankAccounts.TextMatrix(rCount, 0) = "X" Then
            szSelectedBankAccount = flxBankAccounts.TextMatrix(rCount, 2)
            Exit For
         End If
    Next
    If szSelectedBankAccount = "" Then
        MsgBox "Please select a Bank account", vbInformation, "Warning "
        Exit Sub
    End If
    'If bPreviewMode = False Then
    Dim szStatmentID As String
    If MsgBox("Are you sure, you want to save this statement?", vbYesNo, "Please confirm") = vbYes Then
        If bEditMode = False Then
            szStatmentID = GetLastStatementID + 1
        Else
                szStatmentID = szCurrentStatementID
              'When you are modifying a statement then it is using the selected statment ID from the parant form(rentPayable  form) )
                Dim adoconn As New adodb.Connection
                adoconn.Open getConnectionString
                adoconn.Execute "Delete from RentSummaryStatement where  StatementID=" & szStatmentID & ""
                adoconn.Execute "Delete from RentSummaryStatementPreview where  StatementID=" & szStatmentID & ""
                adoconn.Execute "Update tlbBankPayment Set RentSumStatement='' where  RentSumStatement='" & szStatmentID & "'"
                adoconn.Execute "Update tlbPayment Set RentSumStatement='' where  RentSumStatement='" & szStatmentID & "'"
                adoconn.Execute "Update tlbReceipt Set RentSumStatement='' where  RentSumStatement='" & szStatmentID & "'"
                adoconn.Close
        End If
        whichFieldToCheck = "RentSumStatement"
        Call GenerateSummaryStatement(szStatmentID)   'Write into SummaryStatement table in this function
        Call MarkAllTransactionsWithSS(szStatmentID)
    
        'run TestReportForRentSummary.rpt
        Call frmRentPayable.loadflxPayFees
        Dim reportApp As New CRAXDRT.Application
        Dim Report As CRAXDRT.Report
      
        Set Report = reportApp.OpenReport(App.Path & szReportPath & "\TestReportForRentSummary.rpt")
        Report.EnableParameterPrompting = False
        Report.DiscardSavedData
        Report.ParameterFields(1).AddCurrentValue CInt(szStatmentID)
        Load frmReport
        frmReport.LoadReportViewer Report
    End If
'    Frame1(6).Visible = False
    Unload Me
End Sub

Private Sub cmdTestReport_Click()
            'Call GetLandLordBalance
'        Dim reportApp As New CRAXDRT.Application
'        Dim Report As CRAXDRT.Report
'
'        Set Report = reportApp.OpenReport(App.Path & szReportPath & "\TestReportForRentSummary.rpt")
'
'        Report.EnableParameterPrompting = False
'        Report.DiscardSavedData
'        Load frmReport
'        frmReport.LoadReportViewer Report



            'run TestReportForRentSummary.rpt
'            Dim reportApp As New CRAXDRT.Application
'            Dim Report As CRAXDRT.Report
'            Set Report = reportApp.OpenReport(App.Path & szReportPath & "\TestReportForRentSummary.rpt")
'            Report.EnableParameterPrompting = False
'            Report.DiscardSavedData
''               Report.ParameterFields(1).AddCurrentValue CStr(txtLLID.text)
''               Report.ParameterFields(2).AddCurrentValue CDate(txtFromDate.text)
''               Report.ParameterFields(3).AddCurrentValue CDate(txtToDate.text)
''                Report.ParameterFields(4).AddCurrentValue cboCategory.text
'            Load frmReport
'            frmReport.LoadReportViewer Report
End Sub

Private Sub Command1_Click()
        'This is recalculate rent summary sub procedure where we are clearing all the flags when we press recalculate button
        Dim adoconn As New adodb.Connection
        adoconn.Open getConnectionString
        adoconn.Execute "Delete from RentSummaryStatement"
        adoconn.Execute "Delete from RentSummaryStatementPreview"
        adoconn.Execute "Update tlbBankPayment Set RentSumStatement=''"
        adoconn.Execute "Update tlbPayment Set RentSumStatement=''"
        adoconn.Execute "Update tlbReceipt Set RentSumStatement=''"
        Call frmRentPayable.loadflxPayFees
'    adoconn.Execute "Update tlbBankPayment Set RentSumStatement='' where RentSumStatement='" & szCurrentStatementID & "'"
'    adoconn.Execute "Update tlbPayment Set RentSumStatement='' where RentSumStatement='" & szCurrentStatementID & "'"
'    adoconn.Execute "Update tlbReceipt Set RentSumStatement='' where RentSumStatement='" & szCurrentStatementID & "'"
    adoconn.Close
    Set adoconn = Nothing
    MsgBox "Flag has been cleared"
End Sub

Private Sub Command2_Click()

End Sub

'Private Sub Command2_Click()
'    Frame2.Visible = True
'    Frame2.Left = 315
'    Frame2.Top = 4320
'    txtRetensionAmount1.Visible = True
'    txtRetensionAmount1.SelStart = 0
'    txtRetensionAmount1.SelLength = Len(txtRetensionAmount1.text)
'    flxRetensionDetails.Enabled = False
'    FocusControl txtRetensionAmount1
'End Sub

'Private Sub Command3_Click()
'        Frame2.Visible = True
'        flxRetensionDetails.Enabled = True
'        txtRetensionAmount1.Visible = False
'        FocusControl flxRetensionDetails
'        Dim iRow As Integer
'        For iRow = 1 To flxRetensionDetails.Rows - 1
'            If flxRetensionDetails.TextMatrix(iRow, 2) <> "" Then
'                    flxRetensionDetails.TextMatrix(iRow, 0) = "-"
'            End If
'        Next
'End Sub




Private Sub flxBankAccounts_Click()
    Dim iRow As Integer
    chkAllProperties.Value = 0
    'addeb by anol 20210903
     For iRow = 1 To flxClients.Rows - 1
                If flxClients.TextMatrix(iRow, 0) = "X" Then
                        szSelectedClient = flxClients.TextMatrix(iRow, 1)
                End If
      Next
    
    
    If flxBankAccounts.TextMatrix(flxBankAccounts.row, 1) = "" Then Exit Sub
    SelectOnly1RowFlxGrid flxBankAccounts, flxBankAccounts.row, 0
    If flxBankAccounts.TextMatrix(flxBankAccounts.row, 0) = "X" Then
            szSelectedBankAccount = flxBankAccounts.TextMatrix(flxBankAccounts.row, 2)
            addPropertiesTowizard szSelectedClient
'            Call LoadflxInFunds
    End If
    hasSelBankAccounts = False
    For iRow = 1 To flxBankAccounts.Rows - 1
            If flxBankAccounts.TextMatrix(iRow, 0) = "X" Then
                hasSelBankAccounts = True
                Exit For
            End If
    Next
    If hasSelBankAccounts = False Then
        Call ConfigFlxProperties
    End If
End Sub



Private Sub flxClients_Click()
'     SelectOnly1RowFlxGrid flxClients, flxClients.row, 0
     'SelectOnly1RowFlxGrid flxBankAccounts, flxBankAccounts.row, 0
'     SelectOnly1RowFlxGrid flxClients, flxClients.row, 0
     Select1RowFlxGrid flxClients, flxClients.row, 0
     szSelectedClient = flxClients.TextMatrix(flxClients.row, 1)
     'Auto select Properties bases on whatever you select at client
     Dim iRow As Integer
     Dim rCount As Integer
   Dim adoConn1 As New adodb.Connection
   'Dim szSelectedClient As String
   Dim szSelectedClientName As String
   Dim PurchaseLedgerControl As String
   
     For iRow = 1 To flxClients.Rows - 1
                If flxClients.TextMatrix(iRow, 0) = "X" Then
                        szSelectedClient = flxClients.TextMatrix(iRow, 1)
                        szSelectedClientName = flxClients.TextMatrix(rCount, 2)
                        '                    addPropertiesTowizard flxClients.TextMatrix(iRow, 1)
                        Call LoadFlxBankAccounts(szSelectedClient)
                        '                    Call LoadflxPayableTypes 'szSelectedClient is a glbal variable and this function is loadiing the values according to client
                        '                    Call LoadflxInFunds
                        '                    Call LoadFlxBankAccounts(szSelectedClient)
                        Call LoadLaststatementdate
                        Exit For
                Else
                    Call ConfigFlxBankAccounts
                    Call ConfigFlxProperties
                    Call ConfigFlxInFunds
'                    txtLastStatementDate1.text = Format(Date, "dd/mm/yyyy")
'                    txtLastStatementDate1.text = ""
                End If
      Next
    

'   For rCount = 1 To flxClients.Rows - 1
'         If flxClients.TextMatrix(rCount, 0) = "X" Then
'            szSelectedClient = flxClients.TextMatrix(rCount, 1)
'            szSelectedClientName = flxClients.TextMatrix(rCount, 2)
'            Exit For
'         End If
'   Next
'
    
   

   If adoConn1.State = 0 Then
        adoConn1.Open getConnectionString
   End If
    If szSelectedClient <> "" Then
        PurchaseLedgerControl = GetNominalCodeForControlAccount(adoConn1, "Client/Landlord Control Account (B/S)", szSelectedClient)
        If (PurchaseLedgerControl = "") Then
            MsgBox "Please set up Rent and Other Amounts Payable control accounts for '" & szSelectedClient & "' "
            Exit Sub
        End If
    
    
        PurchaseLedgerControl = GetNominalCodeForControlAccount(adoConn1, "Rent & Other Amounts Payable (P&L)", szSelectedClient)
        If (PurchaseLedgerControl = "") Then
            MsgBox "Please set up Rent and Other Amounts Payable  control accounts for '" & szSelectedClientName & "' "
            Exit Sub
        End If
    End If
    adoConn1.Close
    
    Set adoConn1 = Nothing
End Sub
Private Sub addPropertiesTowizard(strClientID As String)
        Call ConfigFlxProperties
        Dim iRow As Integer
        Dim szSQL As String
        Dim adoconn As New adodb.Connection
        Dim adoRst As New adodb.Recordset
        adoconn.Open getConnectionString
        Dim rsConsolidatedStatement As New adodb.Recordset
        'adoConn.Open getConnectionString
        rsConsolidatedStatement.Open "Select * from client where clientID ='" & szSelectedClient & "'", adoconn, adOpenStatic, adLockReadOnly
        If Not rsConsolidatedStatement.EOF Then
            If IsNull(rsConsolidatedStatement("ConsolidatedStatement").Value) Then
                MsgBox "The Consolidated Client Statement option has not been set for this client. Please set this option on the client record", vbInformation, "Warning"
            End If
            
            boolConsolidatedStatement = IIf(IsNull(rsConsolidatedStatement("ConsolidatedStatement").Value), 0, rsConsolidatedStatement("ConsolidatedStatement"))
        End If
        rsConsolidatedStatement.Close
    
        szSQL = "SELECT   PROPERTYID, PROPERTYNAME, ClientID " & _
               "FROM  PROPERTY where clientID = '" & strClientID & "'" & _
               "ORDER BY ClientID,PROPERTYID;"
        adoRst.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
        iRow = 1

        While Not adoRst.EOF 'While Not adoRst.EOF
            flxProperties.AddItem ""
            flxProperties.TextMatrix(iRow, 0) = ""
            flxProperties.TextMatrix(iRow, 1) = adoRst("PROPERTYID").Value
            flxProperties.TextMatrix(iRow, 2) = adoRst("PROPERTYNAME").Value
            flxProperties.TextMatrix(iRow, 3) = adoRst("ClientID").Value
           ' flxProperties.RowHeight(iRow) = 280
'            If iRow > 1 Then
'                flxProperties.AddItem ""
'            End If
            iRow = iRow + 1
    
            adoRst.MoveNext
        'End If
        Wend
       adoconn.Close
       Set adoconn = Nothing
       If boolConsolidatedStatement = 1 Then
            chkAllProperties.Value = 1
       End If
End Sub





Private Sub flxInFunds_Click()
    SelectFlxGridRow 0, flxInFunds, flxInFunds.row
    
    '  SelectOnly1RowFlxGrid flxInFunds, flxInFunds.row, 0
      szSelectedFund = flxInFunds.TextMatrix(flxInFunds.row, 1)
End Sub



Public Function SelectFlxGridRow(iColID As Integer, conFlxGrid As MSHFlexGrid, iSelRow As Integer) As Integer
   Dim iRow As Integer

   If conFlxGrid.TextMatrix(iSelRow, iColID) = "X" Then
      conFlxGrid.TextMatrix(iSelRow, iColID) = ""
      conFlxGrid.row = iSelRow
      For iRow = conFlxGrid.Cols - 1 To 1
         conFlxGrid.col = iRow
         conFlxGrid.CellBackColor = RGB(255, 255, 255)
      Next iRow
      SelectFlxGridRow = -1
   Else
        'Here I have Implemented if no value in the grid row then do not select anol 2020-11-04
      If conFlxGrid.TextMatrix(iSelRow, iColID + 1) <> "" Then
            conFlxGrid.TextMatrix(iSelRow, iColID) = "X"
            conFlxGrid.row = iSelRow
            For iRow = conFlxGrid.Cols - 1 To 1
               conFlxGrid.col = iRow
               conFlxGrid.CellBackColor = RGB(174, 179, 233)
            Next iRow
            SelectFlxGridRow = 1
      Else
            SelectFlxGridRow = -1
      End If
   End If
End Function
Public Sub SelectOnly1RowFlxGrid(conFlxGrid As Control, iNewRow As Integer, Optional iColID As Integer = 0)
   Dim iRow       As Integer
   Dim iCol       As Integer
   Dim iColPaint  As Integer

   iColPaint = IIf(iColID = 0, 1, 0)
   
   For iRow = conFlxGrid.Rows - 1 To 1 Step -1
      If conFlxGrid.TextMatrix(iRow, iColID) = "X" Then
         If iRow = iNewRow And conFlxGrid.TextMatrix(iRow, iColID) = "X" Then
                conFlxGrid.TextMatrix(iRow, iColID) = ""
                conFlxGrid.TextMatrix(iRow, iColID) = ""
                conFlxGrid.row = iRow
                For iCol = iColPaint To conFlxGrid.Cols - 1
                   conFlxGrid.col = iCol
                   conFlxGrid.CellBackColor = vbWhite
                Next iCol
                Exit Sub
         End If
'         conFlxGrid.TextMatrix(iRow, iColID) = ""
'         conFlxGrid.row = iRow
'         For iCol = iColPaint To conFlxGrid.Cols - 1
'            conFlxGrid.col = iColvj
'            conFlxGrid.CellBackColor = vbWhite
'         Next iCol
      End If
   Next iRow

   conFlxGrid.TextMatrix(iNewRow, iColID) = "X"
   conFlxGrid.row = iNewRow

   For iCol = conFlxGrid.Cols - 1 To iColPaint Step -1
      conFlxGrid.col = iCol
      conFlxGrid.CellBackColor = RGB(174, 179, 233)
   Next iCol
End Sub
Private Sub flxProperties_Click()
'    SelectFlxGridRow 0, flxProperties, flxProperties.row
    Dim iIncDec As Integer
    Dim iRow As Integer
    iIncDec = iIncDec + SelectFlxGridRow(0, flxProperties, flxProperties.row) 'Returns 1 or -1 depends on selection
'    Call LoadflxInFunds
    hasSelProperty = False
    For iRow = 1 To flxProperties.Rows - 1
            If flxProperties.TextMatrix(iRow, 0) = "X" Then
                hasSelProperty = True
                Exit For
            End If
    Next
    
    If hasSelProperty Then
          Call LoadflxInFunds
          'Call LoadFlxBankAccounts(szSelectedClient)
    Else
          Call ConfigFlxInFunds
    End If
End Sub

Private Sub Form_Activate()
    Dim adoconn As New adodb.Connection
    Dim iRow As Integer
    Dim szSelectedFunds1  As String
    Dim szSelectedProperties1 As String
    Dim szSelectedBankAC1 As String
    If bEditMode = True Then
        If szCurrentStatementID = "" Then Exit Sub
        adoconn.Open getConnectionString
        Dim szSQL As String
        Dim rsRentSummaryStatement As New adodb.Recordset
'        Dim adoconn As New ADODB.Connection
'        adoconn.Open getConnectionString
        szSQL = "Select ClientIDLandlordID,ListOfFundId,ListOfinputProperties,BankCode,Retentions,PayableAmount,AvailableFunds from RentSummaryStatement where StatementID=" & szCurrentStatementID & ""
        rsRentSummaryStatement.Open szSQL, adoconn, adOpenKeyset, adLockReadOnly
        If Not rsRentSummaryStatement.EOF Then
            For iRow = 1 To flxClients.Rows - 1
                If flxClients.TextMatrix(iRow, 1) = rsRentSummaryStatement("ClientIDLandlordID").Value Then
                    flxClients.TextMatrix(iRow, 0) = "X"
                    szSelectedClient = flxClients.TextMatrix(iRow, 1)
                    szSelectedFunds1 = rsRentSummaryStatement("ListOfFundId").Value
                    szSelectedProperties1 = rsRentSummaryStatement("ListOfinputProperties").Value
                    szSelectedBankAC1 = rsRentSummaryStatement("BankCode").Value
                    txtRetention.text = Format(rsRentSummaryStatement("Retentions").Value, "0.00")
                    txtAvailableFunds.text = Format(rsRentSummaryStatement("AvailableFunds").Value, "0.00")
                    txtRentPayable.text = Format(rsRentSummaryStatement("PayableAmount").Value, "0.00")
                    Call LoadFlxBankAccounts(szSelectedClient)
                    Call addPropertiesTowizard(szSelectedClient)
                    Exit For
                End If
            Next
        End If
        rsRentSummaryStatement.Close
        Set rsRentSummaryStatement = Nothing
        Call LoadflxInFundsONEditInput(szSelectedClient)
        For iRow = 1 To flxBankAccounts.Rows - 1
             If InStr(1, szSelectedBankAC1, flxBankAccounts.TextMatrix(iRow, 2)) > 0 Then
                    flxBankAccounts.TextMatrix(iRow, 0) = "X"
                Exit For
             End If
        Next
        For iRow = 1 To flxInFunds.Rows - 1
             If InStr(1, szSelectedFunds1, flxInFunds.TextMatrix(iRow, 1)) > 0 Then
                    flxInFunds.TextMatrix(iRow, 0) = "X"
             End If
        Next
        For iRow = 1 To flxProperties.Rows - 1
             If InStr(1, szSelectedProperties1, flxProperties.TextMatrix(iRow, 1)) > 0 Then
                        flxProperties.TextMatrix(iRow, 0) = "X"
             End If
        Next
        'Here you need to auto select those items
        adoconn.Close
        Set adoconn = Nothing
    End If
End Sub

'Private Sub flxRetensionDetails_Click()
'    If flxRetensionDetails.TextMatrix(flxRetensionDetails.row, 0) = "-" Then
'        flxRetensionDetails.RemoveItem flxRetensionDetails.row
'        Call MakeSummaryRetention
'    End If
'End Sub

Private Sub Form_Load()
    Dim iRow As Integer
    Dim adoconn As New adodb.Connection
    Me.BackColor = MODULEBACKCOLOR
    Frame1(6).BackColor = MODULEBACKCOLOR
    Frame1(12).BackColor = MODULEBACKCOLOR
    Frame1(8).BackColor = MODULEBACKCOLOR
    Frame1(9).BackColor = MODULEBACKCOLOR
    Frame1(10).BackColor = MODULEBACKCOLOR
    chkAllProperties.BackColor = MODULEBACKCOLOR
    Frame1(13).BackColor = MODULEBACKCOLOR
    Frame1(7).BackColor = MODULEBACKCOLOR
    Frame1(14).BackColor = MODULEBACKCOLOR
    If Len(txtLastStatementDate1.text) < 10 Then txtLastStatementDate1.text = Format(Date, "dd/mm/yyyy")
    SelTxtInCtrl txtLastStatementDate1
    chkInFunds.BackColor = MODULEBACKCOLOR
    Call ConfigflxRetensionDetails
    Dim szSelectedFunds1 As String
    Dim szSelectedProperties1 As String
    Dim szSelectedBankAC1 As String
    txtStatementDate1.text = Format(Now, "dd/mm/yyyy")
    Me.Width = 14220
    Call LoadFlxClients
    If bEditMode = True Then
'        If szCurrentStatementID = "" Then Exit Sub
'        adoConn.Open getConnectionString
'        Dim szSQL As String
'        Dim rsRentSummaryStatement As New ADODB.Recordset
''        Dim adoconn As New ADODB.Connection
''        adoconn.Open getConnectionString
'        szSQL = "Select ClientIDLandlordID,ListOfFundId,ListOfinputProperties,BankCode from RentSummaryStatement where StatementID=" & szCurrentStatementID & ""
'        rsRentSummaryStatement.Open szSQL, adoConn, adOpenKeyset, adLockReadOnly
'        If Not rsRentSummaryStatement.EOF Then
'            For iRow = 1 To flxClients.Rows - 1
'                If flxClients.TextMatrix(iRow, 1) = rsRentSummaryStatement("ClientIDLandlordID").Value Then
'                    flxClients.TextMatrix(iRow, 0) = "X"
'                    szSelectedClient = flxClients.TextMatrix(iRow, 1)
'                    szSelectedFunds1 = rsRentSummaryStatement("ListOfFundId").Value
'                    szSelectedProperties1 = rsRentSummaryStatement("ListOfinputProperties").Value
'                    szSelectedBankAC1 = rsRentSummaryStatement("BankCode").Value
'                    Call LoadFlxBankAccounts(szSelectedClient)
'                    Call addPropertiesTowizard(szSelectedClient)
'                    Exit For
'                End If
'            Next
'        End If
'        rsRentSummaryStatement.Close
'        Set rsRentSummaryStatement = Nothing
'        Call LoadflxInFundsONEditInput(szSelectedClient)
'        For iRow = 1 To flxBankAccounts.Rows - 1
'             If InStr(1, szSelectedBankAC1, flxBankAccounts.TextMatrix(iRow, 2)) > 0 Then
'                    flxBankAccounts.TextMatrix(iRow, 0) = "X"
'                Exit For
'             End If
'        Next
'        For iRow = 1 To flxInFunds.Rows - 1
'             If InStr(1, szSelectedFunds1, flxInFunds.TextMatrix(iRow, 1)) > 0 Then
'                    flxInFunds.TextMatrix(iRow, 0) = "X"
'             End If
'        Next
'        For iRow = 1 To flxProperties.Rows - 1
'             If InStr(1, szSelectedProperties1, flxProperties.TextMatrix(iRow, 1)) > 0 Then
'                        flxProperties.TextMatrix(iRow, 0) = "X"
'             End If
'        Next
'        'Here you need to auto select those items
'        adoConn.Close
'        Set adoConn = Nothing
    Else
        Call LoadflxInFunds
'        Call loadflxPayFees
    End If
    Call WheelHook(Me.hWnd)
End Sub
Private Sub LoadLaststatementdate()
    Dim szSQL As String
    Dim rsRentSummaryStatement As New adodb.Recordset
    Dim adoconn As New adodb.Connection
    adoconn.Open getConnectionString
    szSQL = "Select StatementDate from RentSummaryStatement where StatementNo=" & GetLastStatementNoByClient & " AND ClientIDLandlordID='" & szSelectedClient & "'"
    rsRentSummaryStatement.Open szSQL, adoconn, adOpenKeyset, adLockReadOnly
    If Not rsRentSummaryStatement.EOF Then
        txtLastStatementDate1.text = rsRentSummaryStatement!StatementDate
    Else
'        txtLastStatementDate1.text = Format(Date, "dd/mm/yyyy")
    End If
    
    rsRentSummaryStatement.Close
    Set rsRentSummaryStatement = Nothing
    adoconn.Close
    Set adoconn = Nothing
End Sub

'Private Sub txtAvailableFund1_KeyPress(KeyAscii As Integer)
'    DigitTextKeyPress txtAvailableFund1, KeyAscii
'End Sub

Private Sub txtAvailableFunds_KeyPress(KeyAscii As Integer)
     DigitTextKeyPress txtAvailableFunds, KeyAscii
End Sub

Private Sub txtLastStatementDate1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        FocusControl txtClientSearch
    End If
    TextBoxKeyPrsDate txtLastStatementDate1, KeyAscii
End Sub
Private Sub txtLastStatementDate1_LostFocus()
    If txtLastStatementDate1.text <> "" Then TextBoxFormatDate txtLastStatementDate1
End Sub
Private Sub txtLastStatementDate1_GotFocus()
   If Len(txtLastStatementDate1.text) < 10 Then txtLastStatementDate1.text = Format(Date, "dd/mm/yyyy")
   SelTxtInCtrl txtLastStatementDate1
End Sub
Private Sub txtLastStatementDate1_Change()
    TextBoxChangeDate txtLastStatementDate1
End Sub

Private Function NextID(adoconn As adodb.Connection) As Long
   Dim szSQL As String
   Dim adoRst As New adodb.Recordset
   szSQL = "SELECT MAX(Cint(StatementID))+1 AS Ref FROM RentSummaryStatement;"
   adoRst.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
        NextID = IIf(adoRst.EOF, 1, IIf(IsNull(adoRst!ref), 1, adoRst!ref))
   adoRst.Close
   Set adoRst = Nothing
End Function

'Private Sub LoadflxPayableTypes()
'   Dim rstClient   As New ADODB.Recordset
'   Dim szSQL       As String
'   Dim iRow As Integer
'   Dim conClient As New ADODB.Connection
'   On Error GoTo ErrorHandler
'   Call ConfigflxPayableTypes
'   conClient.Open getConnectionString
'   szSQL = "SELECT D.*, " & _
'                  "IIF(ISNULL(P.PropertyName), 'All Properties', P.PropertyName) AS PropertyName, " & _
'                  "IIF(ISNULL(P.ClientID), 'All Clients', P.ClientID) AS ClientID, " & _
'                  "IIF(ISNULL(P.ClientID), '', C.ClientName) AS ClientName " & _
'                "FROM (PayableTypes AS D INNER JOIN Property AS P ON " & _
'                      "D.PropertyID = P.PropertyID) INNER JOIN Client AS C ON P.ClientID = C.ClientID " & _
'                      "where C.ClientID='" & szSelectedClient & "' " & _
'                " ORDER BY ClientName, PropertyName, D.PayType, D.ID;"
'
'
'   rstClient.Open szSQL, conClient, adOpenStatic, adLockReadOnly
'
'   iRow = 1
'
'   While Not rstClient.EOF
'      flxPayableTypes.TextMatrix(iRow, 1) = rstClient!id
'      flxPayableTypes.TextMatrix(iRow, 2) = rstClient!PayType
'      'flxPayableTypes.TextMatrix(iRow, 3) = rstClient!Id
'      rstClient.MoveNext
'      If Not rstClient.EOF Then flxPayableTypes.AddItem ""
'      iRow = iRow + 1
'   Wend
'
'NoRes:
'   rstClient.Close
'   Set rstClient = Nothing
'   conClient.Close
'   Set conClient = Nothing
'   Exit Sub
'
'ErrorHandler:
'   MsgBox ERR.description & "::" & ERR.Number
'
'   Set rstClient = Nothing
'End Sub
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
Private Sub LoadflxInFundsONEditInput(szSelectedClient As String)
   Dim adoRst   As New adodb.Recordset
   Dim szSQL       As String
   Dim iRow As Integer
   Dim conClient As New adodb.Connection
   
   ConfigFlxInFunds
   conClient.Open getConnectionString
    Dim rsFundMatrix As New adodb.Recordset
    Dim iSel As Integer
    szSQL = "SELECT Distinct FundID, FundName, FundCode,CategoryCode FROM Fund LEFT JOIN tlbPayable PB ON PB.Pay_fund=cstr(Fund.FUNDID) where PB.clientID='" & _
            szSelectedClient & "' order by fundID;"

    adoRst.Open szSQL, conClient, adOpenStatic, adLockReadOnly
   

   iRow = 1

   While Not adoRst.EOF
      If iRow = 1 Then
'            flxInFunds.TextMatrix(iRow, 0) = "X"
'            szSelectedFund = adoRst!fundID
      End If
      flxInFunds.TextMatrix(iRow, 1) = adoRst!fundID
      flxInFunds.TextMatrix(iRow, 2) = adoRst!FundName
      flxInFunds.TextMatrix(iRow, 3) = adoRst!FundCode
      If flxInFunds.TextMatrix(iRow, 3) = "TENANTDEPOSIT" Then
                 flxInFunds.TextMatrix(iRow, 0) = "X"
                 flxInFunds.RowHeight(iRow) = 0
      End If
      flxInFunds.ColWidth(4) = 0
      If iSel = 1 Then
            flxInFunds.ColWidth(4) = 1500
            flxInFunds.TextMatrix(iRow, 4) = adoRst!propertyID
      End If
      adoRst.MoveNext
      If Not adoRst.EOF Then flxInFunds.AddItem ""
      iRow = iRow + 1
   Wend

NoRes:
   adoRst.Close
   Set adoRst = Nothing
   Exit Sub

ErrorHandler:
   MsgBox ERR.description & "::" & ERR.Number

   Set adoRst = Nothing
End Sub
Private Sub LoadflxInFunds()
   Dim adoRst   As New adodb.Recordset
   Dim szSQL       As String
   Dim iRow As Integer
   Dim conClient As New adodb.Connection
'   On Error GoTo ErrorHandler
   
        If ListOfProperties = "" Then
            Exit Sub
        End If
        If szSelectedClient = "" Then
                Exit Sub
        End If
   
   ConfigFlxInFunds
   conClient.Open getConnectionString
'   szSQL = "SELECT F.FundID, F.FundName, S.Value " & _
'           "FROM Fund AS F, SecondaryCode AS S " & _
'           "WHERE F.CategoryCode = CBYTE(S.Code) AND S.PrimaryCode = 'DCTG' " & _
'           "ORDER BY FundID;"
'
'   adoRst.Open szSQL, conClient, adOpenStatic, adLockReadOnly
    Dim rsFundMatrix As New adodb.Recordset
    Dim iSel As Integer
    szSQL = "SELECT Distinct FundID, FundName, FundCode,CategoryCode FROM Fund LEFT JOIN tlbPayable PB ON PB.Pay_fund=cstr(Fund.FUNDID) where PB.clientID='" & _
            szSelectedClient & "' order by fundID;"
     
     
'    rsFundMatrix.Open "Select isfundAssign from shoppingcentre", conClient, adOpenStatic, adLockReadOnly
'    If rsFundMatrix("isfundAssign").Value = False Then
'        iSel = 0
'        szSQL = "SELECT FundID, FundName, FundCode,CategoryCode FROM Fund order by fundID;"
'    Else
'        iSel = 1
'        szSQL = "Select F.*,M.PropertyID from Fund F,fundMatrix M where F.FundID=M.FundID AND PropertyID  in (" & _
'                ListOfProperties & ") and ClientID='" & szSelectedClient & "' and isDeleted=false order by F.fundID"
'    End If
'    rsFundMatrix.Close
    adoRst.Open szSQL, conClient, adOpenStatic, adLockReadOnly
   

   iRow = 1

   While Not adoRst.EOF
      If iRow = 1 Then
'            flxInFunds.TextMatrix(iRow, 0) = "X"
'            szSelectedFund = adoRst!fundID
      End If
      flxInFunds.TextMatrix(iRow, 1) = adoRst!fundID
      flxInFunds.TextMatrix(iRow, 2) = adoRst!FundName
      flxInFunds.TextMatrix(iRow, 3) = adoRst!FundCode
      If flxInFunds.TextMatrix(iRow, 3) = "TENANTDEPOSIT" Then
                 flxInFunds.TextMatrix(iRow, 0) = "X"
                 flxInFunds.RowHeight(iRow) = 0
      End If
      flxInFunds.ColWidth(4) = 0
      If iSel = 1 Then
            flxInFunds.ColWidth(4) = 1500
            flxInFunds.TextMatrix(iRow, 4) = adoRst!propertyID
      End If
      adoRst.MoveNext
      If Not adoRst.EOF Then flxInFunds.AddItem ""
      iRow = iRow + 1
   Wend

NoRes:
   adoRst.Close
   Set adoRst = Nothing
   Exit Sub

ErrorHandler:
   MsgBox ERR.description & "::" & ERR.Number

   Set adoRst = Nothing
End Sub
Private Sub ConfigFlxGrids()
   Dim szHeader As String

'   flxClients.RowHeight(0) = 0
'   flxClients.ColWidth(0) = 300
'   flxClients.ColWidth(1) = 1000
'   flxClients.ColWidth(2) = 2350

   szHeader$ = "<|Property ID<|Property Name<|<"
   flxProperties.FormatString = szHeader
   flxProperties.Cols = 4
   flxProperties.RowHeight(0) = 0
   flxProperties.ColWidth(0) = 300                   '"X"
   flxProperties.ColWidth(1) = 1000                'Property ID
   flxProperties.ColWidth(2) = 2350                'Property Name
   flxProperties.ColWidth(3) = 0                   'Client ID

'   flxDemandTypes.Cols = 5
'   flxDemandTypes.RowHeight(0) = 0
'   flxDemandTypes.ColWidth(0) = 300                  '"X"
'   flxDemandTypes.ColWidth(1) = 900                  'Property ID
'   flxDemandTypes.ColWidth(2) = 0               'Demand Type ID
'   flxDemandTypes.ColAlignment(2) = vbRightJustify
'   flxDemandTypes.ColWidth(3) = 4000               'Demand Type Name
'   flxDemandTypes.ColWidth(4) = 0                  'Demand Category
'
'   flxCategory.RowHeight(0) = 0
'   flxCategory.ColWidth(0) = 0
'   flxCategory.ColWidth(1) = 0
'   flxCategory.ColWidth(2) = flxCategory.Width - 250
End Sub
Private Sub ConfigFlxProperties()
   Dim szHeader As String
   flxProperties.Clear
   flxProperties.Rows = 2
   szHeader$ = "|<Property ID|<Property Name|<|<"
   With flxProperties
      .FormatString = szHeader
      .Cols = 4
     ' .RowHeight(0) = 0
      .ColWidth(0) = 200 'Label2(0).Left - .Left '200                 '"X"
      .ColWidth(1) = 2000 'Label2(1).Left - Label2(0).Left 'Property ID
      .ColWidth(2) = 2500 'Label2(2).Left - Label2(1).Left 'Property Name
      .ColWidth(3) = 0 '.Width + .Left - Label2(2).Left - 300 'Client ID
   End With
End Sub
Private Sub ConfigFlxInFunds()
   Dim szHeader As String

   flxInFunds.Cols = 5
   flxInFunds.Clear
   szHeader$ = "|<ID|<Name|<Fund Code|<Property ID"
   flxInFunds.FormatString = szHeader$
   flxInFunds.ColWidth(0) = 280        'Selection column
   flxInFunds.ColWidth(1) = 400        'fundID
   flxInFunds.ColWidth(2) = 2800       'FundName
   flxInFunds.ColWidth(3) = 1500       'FundCode
   flxInFunds.ColWidth(4) = 0       '
   flxInFunds.Rows = 2
End Sub

'Private Sub ConfigflxPayableTypes()
'   Dim szHeader As String
'
'   flxPayableTypes.Cols = 4
'   flxPayableTypes.Clear
'   szHeader$ = "|<ID|<Payable Types|<Category"
'   flxPayableTypes.FormatString = szHeader$
'   flxPayableTypes.ColWidth(0) = 280        'Solid column
'   flxPayableTypes.ColWidth(1) = 400        'ID
'   flxPayableTypes.ColWidth(2) = 2800       'Name
'   flxPayableTypes.ColWidth(3) = 1500       'empty text
'   flxPayableTypes.Rows = 2
'
'   flxPayableTypes.RowHeightMin = 255
'End Sub
'Private Sub loadflxProperties()
'    Dim szSQL   As String
'    Dim r       As Integer
'    Dim adoConn As New ADODB.Connection
'    Dim adoRst As New ADODB.Recordset
'    adoConn.Open getConnectionString
'     szSQL = "SELECT   PROPERTYID, PROPERTYNAME, ClientID " & _
'           "FROM     PROPERTY where ClientID='" & szSelectedClient & "' " & _
'           "ORDER BY PROPERTYID;"
'           'where ClientID='" & txtClientID.text & "'
'   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'   ConfigFlxProperties
'   r = 1
'
'
'   While Not adoRst.EOF
'      flxProperties.TextMatrix(r, 1) = adoRst.Fields.Item("PROPERTYID").Value
'      flxProperties.TextMatrix(r, 2) = adoRst.Fields.Item("PROPERTYNAME").Value
'      flxProperties.TextMatrix(r, 3) = adoRst.Fields.Item("ClientID").Value
'      flxProperties.RowHeight(r) = 240
'      r = r + 1
'
'      adoRst.MoveNext
'      If Not adoRst.EOF Then flxProperties.AddItem ""
'   Wend
'    Debug.Print r
'   adoRst.Close
'   Set adoRst = Nothing
'   adoConn.Close
'   Set adoConn = Nothing
'   flxProperties.row = 0
'End Sub
Private Sub LoadFlxBankAccounts(szClientID As String)
   Dim conClient As New adodb.Connection
   Dim rstClient   As New adodb.Recordset
   Dim szSQL       As String
   Dim iRow As Integer

   On Error GoTo ErrorHandler
   conClient.Open getConnectionString
   ConfigFlxBankAccounts
   szSQL = "SELECT MY_ID, NominalCode, Bank_AC_Name, CLIENT_ID " & _
           "FROM tlbClientBanks where CLIENT_ID='" & szClientID & "'" & _
           "ORDER BY NominalCode;"

   rstClient.Open szSQL, conClient, adOpenStatic, adLockReadOnly

   iRow = 1

   While Not rstClient.EOF
      If iRow = 1 Then
'         flxBankAccounts.TextMatrix(iRow, 0) = "X"
'         szSelectedBankAccount = rstClient!nominalCode
      End If
      flxBankAccounts.TextMatrix(iRow, 1) = rstClient!My_ID
      flxBankAccounts.TextMatrix(iRow, 2) = rstClient!nominalCode
      flxBankAccounts.TextMatrix(iRow, 3) = rstClient!Bank_AC_Name
      flxBankAccounts.TextMatrix(iRow, 4) = rstClient!CLIENT_ID
      rstClient.MoveNext
      If Not rstClient.EOF Then flxBankAccounts.AddItem ""
      iRow = iRow + 1
   Wend

NoRes:
   rstClient.Close
   Set rstClient = Nothing
   conClient.Close
   Set conClient = Nothing
   Exit Sub

ErrorHandler:
   MsgBox ERR.description & "::" & ERR.Number

   Set rstClient = Nothing
End Sub
Private Sub ConfigFlxClients()
    Dim szHeader As String
    flxClients.Clear
    szHeader$ = "|<ClientID|<ClientName|<.."
    flxClients.FormatString = szHeader$
    flxClients.Cols = 4
    flxClients.Rows = 2
    flxClients.RowHeight(0) = 0
    flxClients.ColWidth(0) = 200
    flxClients.ColWidth(1) = 1200
    flxClients.ColWidth(2) = 4000
    flxClients.ColWidth(3) = 0
    
End Sub
Private Sub LoadFlxClients()
   Call ConfigFlxClients
   Dim szSQL As String, r As Integer
   Dim adoconn As New adodb.Connection
   Dim adoRst As New adodb.Recordset

'   connect to database
   adoconn.Open getConnectionString

   szSQL = "SELECT SupplierID, SupplierName " & _
           "FROM Supplier where type in ('Client');"
   adoRst.Open szSQL, adoconn, adOpenDynamic, adLockPessimistic

   r = 1
   flxClients.Rows = 1

   While Not adoRst.EOF
      flxClients.AddItem ""
      flxClients.TextMatrix(r, 1) = adoRst.Fields.Item("SupplierID").Value
      flxClients.TextMatrix(r, 2) = adoRst.Fields.Item("SupplierNAME").Value
      r = r + 1
      adoRst.MoveNext
   Wend

'        If r > 1 Then
'                SelectOnly1RowFlxGrid flxClients, 1, 0
'                szSelectedClient = flxClients.TextMatrix(1, 1) 'saving the first propertyID in the list in a variable
'                addPropertiesTowizard szSelectedClient
'                LoadFlxBankAccounts szSelectedClient
'        End If
   adoRst.Close
End Sub
Private Sub LoadFreq()
   Dim adoRstFreq As adodb.Recordset
   Dim adoconn As adodb.Connection
   Dim strSQLTitles As String

   Set adoconn = New adodb.Connection
   Set adoRstFreq = New adodb.Recordset
   adoconn.Open getConnectionString
   strSQLTitles = "SELECT * FROM FREQUENCIES;"
   adoRstFreq.Open strSQLTitles, adoconn, adOpenStatic, adLockReadOnly

   ReDim szaFreq(adoRstFreq.RecordCount) As String

   While Not adoRstFreq.EOF
      szaFreq(adoRstFreq.Fields("ID").Value) = adoRstFreq.Fields("CALDAYS").Value
      adoRstFreq.MoveNext
   Wend
   adoRstFreq.Close
   adoconn.Close
   Set adoRstFreq = Nothing
   Set adoconn = Nothing
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   frmMMain.MousePointer = vbArrow
End Sub

Private Sub Form_Unload(Cancel As Integer)
'   Call WheelUnHook(Me.hWnd)
'   frmMMain.fraCmdButton.Enabled = True
   UnLoadForm Me
End Sub

'Private Sub PrepareList()
'   FlxDemandsConfigure flxClientList
'   LoadAllClientFlxGrd
'End Sub






'Private Sub tabFees_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   frmMMain.MousePointer = vbArrow
'End Sub

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

Private Sub txtClientSearch_Change()
   Dim i As Integer
    For i = flxClients.Rows - 1 To 1 Step -1
            flxClients.TextMatrix(i, 0) = ""
   Next i
   
   
   For i = flxClients.Rows - 1 To 1 Step -1
            flxClients.RowHeight(i) = 240
            If InStr(1, UCase(flxClients.TextMatrix(i, 1)), UCase(txtClientSearch.text), vbTextCompare) = 0 And txtClientSearch.text <> "" Then
                flxClients.RowHeight(i) = 0
            End If
         flxClients.TextMatrix(i, 0) = ""
      If flxClients.RowHeight(i) = 240 Then
            flxClients.row = i
      End If
   Next i
End Sub

'Private Sub LoadflxRetensionDetails()
'     Dim adoConn As New ADODB.Connection
'     adoConn.Open getConnectionString
'     Dim rsRetensionDetails As New ADODB.Recordset
'     Dim iRow As Integer
'     iRow = 1
'     rsRetensionDetails.Open "Select * from RetentionDetails where statementID=" & szCurrentStatementID & "", adoConn, adOpenStatic, adLockReadOnly
'     While Not rsRetensionDetails.EOF
'            flxRetensionDetails.AddItem ""
'            flxRetensionDetails.TextMatrix(iRow, 1) = rsRetensionDetails("statementID").Value
'            flxRetensionDetails.TextMatrix(iRow, 2) = rsRetensionDetails("SLNumber").Value
'            flxRetensionDetails.TextMatrix(iRow, 3) = rsRetensionDetails("Description").Value
'            flxRetensionDetails.TextMatrix(iRow, 4) = rsRetensionDetails("Amount").Value
'            iRow = iRow + 1
'            rsRetensionDetails.MoveNext
'     Wend
'     rsRetensionDetails.Close
'     Set rsRetensionDetails = Nothing
'     adoConn.Close
'     Set adoConn = Nothing
'End Sub
Public Sub ConfigflxRetensionDetails()
        flxRetensionDetails.Clear
        Dim szHeader As String
        szHeader$ = "|<StatementID|<SlNumber|<Amount"
        flxRetensionDetails.FormatString = szHeader$

        flxRetensionDetails.Cols = 5
        flxRetensionDetails.Rows = 2
        flxRetensionDetails.RowHeight(0) = 0
        flxRetensionDetails.ColWidth(0) = 250   'Selection Row put plus or minus sign
        flxRetensionDetails.ColWidth(1) = 0 'This is statementId
        flxRetensionDetails.ColWidth(2) = 1200 'This is slNumber
        flxRetensionDetails.ColWidth(3) = 1200  'This is Description
        flxRetensionDetails.ColWidth(4) = 1200  'This is amount
        flxRetensionDetails.ColAlignment(3) = vbLeftJustify

End Sub
'Private Sub LoadFlxFundList()
'        Call ConfigFlxFundList
'        Dim adoConn As New ADODB.Connection
'        Dim rstRec As New ADODB.Recordset
'        Dim szSQL As String
'
'        Dim rRow As Integer
'        adoConn.Open getConnectionString
'
'        szSQL = "SELECT FundID, FundName, FundCode,CategoryCode FROM Fund;"
'        rstRec.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'
'        rRow = 1
'        While Not rstRec.EOF
'                flxFundList.TextMatrix(rRow, 0) = ""
'                flxFundList.TextMatrix(rRow, 1) = rstRec.Fields.Item("FundID").Value
'                flxFundList.TextMatrix(rRow, 2) = rstRec.Fields.Item("FundCode").Value
'                flxFundList.TextMatrix(rRow, 3) = rstRec.Fields.Item("FundName").Value
'                flxFundList.RowHeight(rRow) = 280
'                rstRec.MoveNext
'                If Not rstRec.EOF Then flxFundList.AddItem ""
'                rRow = rRow + 1
'        Wend
'
'        rstRec.Close
'        adoConn.Close
'        Set rstRec = Nothing
'        Set adoConn = Nothing
'End Sub
'Private Sub ConfigFlxFundList()
'        flxFundList.Clear
'        Dim szHeader As String
'        szHeader$ = "|<FundID|<FundCode|<FundName"
'        flxFundList.FormatString = szHeader$
'
'        flxFundList.Cols = 4
'        flxFundList.Rows = 2
'        flxFundList.RowHeight(0) = 0
'        flxFundList.ColWidth(0) = 250   'Selection Row put plus or minus sign
'        flxFundList.ColWidth(1) = 0 'FundID
'        flxFundList.ColWidth(2) = 2000 ' FundCode
'        flxFundList.ColWidth(3) = 2000  ' FundName
'        flxFundList.ColAlignment(0) = vbLeftJustify
'        flxFundList.ColAlignment(1) = vbLeftJustify
'        flxFundList.ColAlignment(2) = vbLeftJustify
'        flxFundList.ColAlignment(3) = vbLeftJustify
'End Sub

'Private Sub txtPayableDate2_GotFocus()
'    txtPayableDate2.SelStart = 0
'    txtPayableDate2.SelLength = Len(txtPayableDate2.text)
'End Sub



Private Sub txtRentPayable_KeyPress(KeyAscii As Integer)
    DigitTextKeyPress txtRentPayable, KeyAscii
End Sub

'Private Sub txtRentPayable1_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        FocusControl txtPayableDate2
'    End If
'    DigitTextKeyPress txtRentPayable1, KeyAscii
'End Sub

'Private Sub txtRetensionAmount1_GotFocus()
'     txtRetensionAmount1.SelStart = 0
'     txtRetensionAmount1.SelLength = Len(txtRetensionAmount1.text)
'End Sub

'Private Sub txtRetensionAmount1_KeyPress(KeyAscii As Integer)
'     If KeyAscii = 13 Then
''            flxRetensionDetails.Enabled = True
''            'Enter data into grid only memory version
''            'statementId you shall generate it when you finally save the statement
''            flxRetensionDetails.TextMatrix(flxRetensionDetails.Rows - 1, 0) = IIf(Option1.Value = True, "+", "-")
''            flxRetensionDetails.TextMatrix(flxRetensionDetails.Rows - 1, 2) = flxRetensionDetails.Rows - 1 'This is slNumber
''            flxRetensionDetails.TextMatrix(flxRetensionDetails.Rows - 1, 3) = txtRetentionDescriptions.text 'This is Description
''            flxRetensionDetails.TextMatrix(flxRetensionDetails.Rows - 1, 4) = Format(Val(txtRetensionAmount1.text), "0.00") 'This is amount
''            flxRetensionDetails.AddItem ""
''           ' txtRetensionAmount1.Visible = False
''
''            txtRetensionAmount1.text = "0.00"
''            txtRetensionAmount1.SelStart = 0
''            txtRetensionAmount1.SelLength = Len(txtRetensionAmount1.text)
''            Call MakeSummaryRetention
'        FocusControl txtRetentionDescriptions
'     End If
''     If KeyAscii = 27 Then 'escape ascii key
''        txtRetensionAmount1.Visible = False
''     End If
'     DigitTextKeyPress txtRetensionAmount1, KeyAscii
'End Sub
'
'Private Sub MakeSummaryRetention()
'    Dim iRow As Long
'    Dim dblAmt As Double
'    For iRow = 1 To flxRetensionDetails.Rows - 1
'            If flxRetensionDetails.TextMatrix(iRow, 2) <> "" Then
'                    If flxRetensionDetails.TextMatrix(iRow, 0) = "+" Then
'                            dblAmt = dblAmt + flxRetensionDetails.TextMatrix(iRow, 4)
'                    Else
'                            dblAmt = dblAmt - flxRetensionDetails.TextMatrix(iRow, 4)
'                    End If
'            End If
'        Next
'    txtRetention.text = dblAmt
'End Sub
'
'Private Sub txtRetensionAmount1_LostFocus()
'    txtRetensionAmount1.text = Format(txtRetensionAmount1.text, "0.00")
'End Sub

'Private Sub txtRetentionDescriptions_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        FocusControl cmdAddToGrid
'    End If
'End Sub

Private Sub txtStatementDate1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        FocusControl txtClientSearch
    End If
    TextBoxKeyPrsDate txtStatementDate1, KeyAscii
End Sub
Private Sub txtStatementDate1_LostFocus()
    If txtStatementDate1.text <> "" Then TextBoxFormatDate txtStatementDate1
End Sub
Private Sub txtStatementDate1_GotFocus()
   If Len(txtStatementDate1.text) < 10 Then txtStatementDate1.text = Format(Date, "dd/mm/yyyy")
   SelTxtInCtrl txtStatementDate1
End Sub
Private Sub txtStatementDate1_Change()
    TextBoxChangeDate txtStatementDate1
End Sub
'Private Sub txtPayableDate2_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'         FocusControl cmdFundListForCreatePI
'    End If
'    TextBoxKeyPrsDate txtPayableDate2, KeyAscii
'End Sub
'Private Sub txtPayableDate2_LostFocus()
'    If txtPayableDate2.text <> "" Then TextBoxFormatDate txtPayableDate2
'End Sub

'Private Sub txtPayableDate2_Change()
'    TextBoxChangeDate txtPayableDate2
'End Sub


