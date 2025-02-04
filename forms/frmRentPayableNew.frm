VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmRentPayableNew 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "New Client Statement"
   ClientHeight    =   12630
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14895
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
   Icon            =   "frmRentPayableNew.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   12630
   ScaleWidth      =   14895
   Begin VB.Frame Frame1 
      Caption         =   "Produce Client Statement"
      Height          =   11220
      Index           =   6
      Left            =   45
      TabIndex        =   11
      Top             =   0
      Width           =   14775
      Begin VB.CheckBox chkExcludeSupOS 
         Caption         =   "Incl. Supplier OS"
         Height          =   210
         Left            =   8235
         TabIndex        =   36
         Top             =   9630
         UseMaskColor    =   -1  'True
         Width           =   1590
      End
      Begin VB.CommandButton cmdViewAllRetention 
         Caption         =   "View All Retentions"
         Height          =   375
         Left            =   1845
         TabIndex        =   35
         Top             =   10080
         Width           =   2175
      End
      Begin VB.CommandButton cmdAddRetention 
         Caption         =   "Add Retentions"
         Height          =   375
         Left            =   135
         TabIndex        =   34
         Top             =   10080
         Width           =   1545
      End
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
         Height          =   285
         Left            =   1170
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   33
         Text            =   "0.00"
         Top             =   9630
         Width           =   1125
      End
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
         Left            =   6075
         MaxLength       =   10
         TabIndex        =   31
         Text            =   "0.00"
         Top             =   9630
         Width           =   1395
      End
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
         Left            =   3600
         MaxLength       =   10
         TabIndex        =   30
         Text            =   "0.00"
         Top             =   9630
         Width           =   1305
      End
      Begin VB.CommandButton cmdfix 
         Caption         =   "fix"
         Height          =   330
         Left            =   11160
         TabIndex        =   27
         Top             =   11340
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CheckBox chkShowDue 
         Caption         =   "Incl. Mngt Fees Due"
         Height          =   210
         Left            =   12600
         TabIndex        =   26
         Top             =   9630
         UseMaskColor    =   -1  'True
         Width           =   1995
      End
      Begin VB.CheckBox chkExcludeReceipt 
         Caption         =   "Exclude Advanced Receipts"
         Height          =   210
         Left            =   9990
         TabIndex        =   25
         Top             =   9630
         UseMaskColor    =   -1  'True
         Width           =   2580
      End
      Begin VB.CheckBox chkConsolidatedCS 
         Caption         =   "Print Details"
         Height          =   210
         Left            =   7380
         TabIndex        =   24
         Top             =   11295
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   1680
      End
      Begin VB.TextBox txtComparenextDueDate1 
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
         Left            =   11925
         MaxLength       =   10
         TabIndex        =   23
         Top             =   135
         Visible         =   0   'False
         Width           =   1575
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxRetensionDetails 
         Height          =   2100
         Left            =   2655
         TabIndex        =   22
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
         Left            =   90
         TabIndex        =   17
         Top             =   4950
         Width           =   7095
         Begin VB.CheckBox chkAllProperties 
            Caption         =   "All Properties"
            Height          =   255
            Left            =   90
            TabIndex        =   18
            Top             =   240
            Width           =   2025
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxProperties 
            Height          =   3810
            Left            =   90
            TabIndex        =   5
            Top             =   495
            Width           =   6810
            _ExtentX        =   12012
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
         Caption         =   "Funds:"
         Height          =   4455
         Index           =   10
         Left            =   7245
         TabIndex        =   15
         Top             =   4950
         Width           =   7410
         Begin VB.CheckBox chkInFunds 
            Caption         =   "All Funds"
            Height          =   255
            Left            =   180
            TabIndex        =   16
            Top             =   270
            Width           =   1095
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxInFunds 
            Height          =   3765
            Left            =   120
            TabIndex        =   6
            Top             =   570
            Width           =   7170
            _ExtentX        =   12647
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
         Left            =   7245
         TabIndex        =   14
         Top             =   765
         Width           =   7455
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxBankAccounts 
            Height          =   3810
            Left            =   90
            TabIndex        =   4
            Top             =   300
            Width           =   7260
            _ExtentX        =   12806
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
         Caption         =   "&Preview Statement"
         Height          =   375
         Left            =   10620
         TabIndex        =   8
         Top             =   10035
         Width           =   1620
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
         TabIndex        =   2
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
         TabIndex        =   1
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
         TabIndex        =   0
         Top             =   450
         Width           =   1575
      End
      Begin VB.Frame Frame1 
         Caption         =   "Clients:"
         Height          =   3780
         Index           =   12
         Left            =   90
         TabIndex        =   13
         Top             =   1170
         Width           =   7095
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxClients 
            Height          =   3495
            Left            =   90
            TabIndex        =   3
            Top             =   225
            Width           =   6825
            _ExtentX        =   12039
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
         Left            =   3825
         TabIndex        =   12
         Top             =   11205
         Visible         =   0   'False
         Width           =   3705
      End
      Begin VB.CommandButton cmdCalculateAvailableFund 
         Caption         =   "Calculate Available Funds"
         Height          =   375
         Left            =   7875
         TabIndex        =   7
         Top             =   10035
         Width           =   2520
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Produce Statement"
         Height          =   375
         Left            =   12375
         TabIndex        =   9
         Top             =   10035
         Width           =   2115
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
         Left            =   13275
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   10575
         Width           =   1200
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Retentions"
         Height          =   210
         Left            =   135
         TabIndex        =   32
         Top             =   9630
         Width           =   840
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Rent Payable"
         Height          =   210
         Left            =   4995
         TabIndex        =   29
         Top             =   9630
         Width           =   1020
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Available Fund"
         Height          =   210
         Left            =   2340
         TabIndex        =   28
         Top             =   9630
         Width           =   1170
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
         TabIndex        =   21
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
         TabIndex        =   20
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
         TabIndex        =   19
         Top             =   450
         Width           =   1110
      End
   End
End
Attribute VB_Name = "frmRentPayableNew"
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
Public szSelectedClient As String
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
Dim szLastStatementDate As String
Dim szStatementNo As Long
Dim bEditDone As Boolean
'Dim szSelectedPayableTypeID As String
'Dim szCurrentRentsummarySTID As String
Private Sub MarkAllTransactionsWithCSID(strSSID As String)
   Dim adoConn As New ADODB.Connection
    adoConn.Open getConnectionString
    Dim szSQL As String
    Dim whereProperty As String
    
    szSQL = "Update tlbReceipt R,tlbReceiptSplit S,Fund F SET ClientStatementID=" & strSSID & "  where R.TransactionID= S.RptHeader AND TYPE IN(3,4,23) " & _
            "AND S.FundID=F.FundID AND R.BankCODE='" & szSelectedBankAccount & "' AND F.FundCode<>'TENANTDEPOSIT' and S.amount>S.OSamount  AND isnull(S.ClientStatementID) " & _
            "AND  S.PropertyID in (" & ListOfProperties & ") AND R.RDate <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "# " & _
            "AND F.FundCode in (" & ListOfFunds & ") AND ClientID ='" & szSelectedClient & "'"
    adoConn.Execute szSQL
    
     'ONE SSR do not have property ID in it. and stackholder wants that to be in CS. SO I need to consider empty Property ID for SSR , include that for marking 2023-08-20
    szSQL = "Update tlbReceipt R,tlbReceiptSplit S,Fund F SET ClientStatementID='" & strSSID & "'  where R.TransactionID= S.RptHeader AND TYPE IN(3,4,23) " & _
            "AND S.FundID=F.FundID AND R.BankCODE='" & szSelectedBankAccount & "' AND F.FundCode<>'TENANTDEPOSIT' and S.amount>S.OSamount  AND  isnull(S.ClientStatementID) " & _
            "AND R.RDate <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "# " & _
            "AND F.FundCode in (" & ListOfFunds & ") AND  (isnull(S.PropertyID) OR  S.PropertyID='')  AND ClientID ='" & szSelectedClient & "'"
    adoConn.Execute szSQL
    
    
    whereProperty = "AND (P.UnitID in (" & ListOfProperties & ") OR isnull(P.UnitID) OR P.UnitID='' ) "
    szSQL = "Update tlbPayment P,tlbPaymentSplit S,Fund F,Supplier SP SET ClientStatementID=" & strSSID & "  where " & _
            "P.PDate <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "# AND P.BankCODE='" & szSelectedBankAccount & "'  AND " & _
            "SP.SupplierID=P.SageaccountNumber and  S.Amount>S.OSAmount  AND  P.TransactionID=S.PayHeader AND P.TYPE IN(24,8,9) AND S.FundID=F.FundID AND F.FundCode in (" & _
             ListOfFunds & ") AND  isnull(S.ClientStatementID) AND P.ClientID ='" & szSelectedClient & "' " & whereProperty & ""
    adoConn.Execute szSQL
    
    szSQL = "Update tlbBankPayment B, Fund F  SET RentSumStatement='" & strSSID & "' where B.DEPT_ID=F.FundID " & _
            "AND B.TRAN_DATE <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "# " & _
            "AND B.PropertyID in (" & ListOfProperties & ") AND BANK_AC='" & szSelectedBankAccount & "' AND F.FundCode in (" & ListOfFunds & ") AND (B.RentSumStatement='' OR isnull(B.RentSumStatement)) and clientID='" & szSelectedClient & "' "
            
    adoConn.Execute szSQL
    
    adoConn.Close
    Set adoConn = Nothing
End Sub
Private Sub MarkAllTransactionsWithCSIDPreview(strSSID As String)
   Dim adoConn As New ADODB.Connection
    adoConn.Open getConnectionString
    Dim szSQL As String
    Dim whereProperty As String
    szSQL = "Update tlbReceiptSplit S SET ClientStatementPrevID=NULL "
    adoConn.Execute szSQL
    szSQL = "Update tlbPaymentSplit S SET ClientStatementPrevID=NULL "
    adoConn.Execute szSQL
    szSQL = "Update tlbBankPayment S SET RentSumStatementPreview=NULL "
    adoConn.Execute szSQL
    
    szSQL = "Update tlbReceipt R,tlbReceiptSplit S,Fund F SET ClientStatementPrevID='" & strSSID & "'  where R.TransactionID= S.RptHeader AND TYPE IN(3,4,23) " & _
            "AND S.FundID=F.FundID AND R.BankCODE='" & szSelectedBankAccount & "' AND F.FundCode<>'TENANTDEPOSIT' and S.amount>S.OSamount  AND  isnull(S.ClientStatementID) " & _
            "AND R.RDate <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "# " & _
            "AND F.FundCode in (" & ListOfFunds & ") AND  S.PropertyID in (" & ListOfProperties & ")  AND ClientID ='" & szSelectedClient & "'"
    adoConn.Execute szSQL
    
    'ONE SSR do not have property ID in it. and stackholder wants that to be in CS. SO I need to consider empty Property ID for SSR , include that for marking 2023-08-20
    szSQL = "Update tlbReceipt R,tlbReceiptSplit S,Fund F SET ClientStatementPrevID='" & strSSID & "'  where R.TransactionID= S.RptHeader AND TYPE IN(3,4,23) " & _
            "AND S.FundID=F.FundID AND R.BankCODE='" & szSelectedBankAccount & "' AND F.FundCode<>'TENANTDEPOSIT' and S.amount>S.OSamount  AND  isnull(S.ClientStatementID) " & _
            "AND R.RDate <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "# " & _
            "AND F.FundCode in (" & ListOfFunds & ") AND  (isnull(S.PropertyID) OR  S.PropertyID='')  AND ClientID ='" & szSelectedClient & "'"
    adoConn.Execute szSQL
    
    
    whereProperty = "AND (P.UnitID in (" & ListOfProperties & ") OR isnull(P.UnitID) OR P.UnitID='' ) "
    szSQL = "Update tlbPayment P,tlbPaymentSplit S,Fund F,Supplier SP SET ClientStatementPrevID='" & strSSID & "'  where " & _
            "P.PDate <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "# AND P.BankCODE='" & szSelectedBankAccount & "' AND SP.SupplierID=P.SageaccountNumber" & _
            " and  P.Amount>P.OSAmount  AND  P.TransactionID=S.PayHeader and S.amount>S.OSamount AND P.TYPE IN(24,8,9) AND S.FundID=F.FundID AND F.FundCode in (" & _
             ListOfFunds & ") AND  isnull(S.ClientStatementID) AND P.ClientID ='" & szSelectedClient & "' " & whereProperty & ""
             
    adoConn.Execute szSQL
    
    szSQL = "Update tlbBankPayment B, Fund F  SET RentSumStatementPreview='" & strSSID & "'  where B.DEPT_ID=F.FundID " & _
            "AND  B.PropertyID in (" & ListOfProperties & ") AND B.TRAN_DATE <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "# " & _
            "and BANK_AC='" & szSelectedBankAccount & "' AND F.FundCode in (" & ListOfFunds & ") AND (B.RentSumStatement='' OR isnull(B.RentSumStatement)) and clientID='" & szSelectedClient & "' "
            
    adoConn.Execute szSQL
    
    adoConn.Close
    Set adoConn = Nothing
    
    
End Sub
Private Function getAvailablefundsPreview(dblLasClosingBalance As Double, ByVal trtoinclude As Long) As Double 'this one I am using while finalize
    'Pass propery as parameter for selected property
    'No property spec:
    'Exit Function
    Dim adoConn As New ADODB.Connection
    Dim rsReceipt As New ADODB.Recordset
    Dim rsPayment As New ADODB.Recordset
    Dim rsBankPaymentAndRcpt As New ADODB.Recordset
    Dim dblAmt As Double
    adoConn.Open getConnectionString
    Dim szSQL As String
    Dim AccrualsControl As String
    Dim whereProperty As String
    'we are not using property filter here
    'B )***********************  Sum of Rent received Paid/Refunded ***********************************
    'AND tlbReceipt.RDate >=#" & Format(txtLastStatementDate1.text, "dd/mmm/yyyy") & "# AND tlbReceipt.RDate <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "#"
    Dim rsGlobalData1 As New ADODB.Recordset
    Dim bolVatOptionEnabled As Boolean
    Dim bolOptedTotax As String
    Dim bolisAgentToSubmit As Boolean
    Dim strManagingAgentID As String
    'OR P.RentSumStatement='" & trToinclude & "'
    Dim dblTotalReceipt As Double
    Dim dblSupplierPayment As Double
    Dim dblSupplierOsAmount As Double
    Dim dblagentBalance As Double
    Dim dblManagementFee As Double
    Dim dblAgentPayment As Double
    
    rsGlobalData1.Open "Select vatOptionEnabled,isAgentToSubmit from Globaldata G,Property P where P.PropertyID=G.PropertyID AND P.PropertyID='" & _
                       szSQL & "' ", adoConn, adOpenStatic, adLockReadOnly
                       
    If Not rsGlobalData1.EOF Then
           bolVatOptionEnabled = rsGlobalData1("vatOptionEnabled").Value
           bolisAgentToSubmit = rsGlobalData1("isAgentToSubmit").Value
    End If
    rsGlobalData1.Close

                
    If boolConsolidatedStatement = 1 Then
            whereProperty = "(U.PropertyID in (" & ListOfProperties & ")OR isnull(U.PropertyID) OR U.PropertyID='' ) AND "
    Else
            whereProperty = "U.PropertyID in (" & ListOfProperties & ") AND "
    End If
    
    'Receipt total
    szSQL = "Select  SUM(switch(type=3,S.Amount,type=4,S.Amount,type=23,-S.Amount)) as DR from tlbReceipt R,tlbReceiptSplit S where R.TransactionID= S.RptHeader " & _
    "AND ClientID ='" & szSelectedClient & "' and S.ClientStatementPrevID=" & trtoinclude & ""
    rsReceipt.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
    If Not rsReceipt.EOF Then
            dblAmt = IIf(IsNull(rsReceipt.Fields.Item("Dr").Value), 0, rsReceipt.Fields.Item("Dr").Value)
            dblTotalReceipt = dblAmt
             'result is 175
    End If
    rsReceipt.Close
    Set rsReceipt = Nothing
    
    
    getAvailablefundsPreview = dblLasClosingBalance + dblAmt
    'Vat calculation of B)take the  VAT amount from the allocation table and deduct it
            szSQL = "Select  SUM(AL.VatAmount) as DR from tlbReceipt R,rptTransactionsSplit AL,tlbReceiptSplit S,Fund F, Units U,GlobalData G where G.PropertyID=U.PropertyID " & _
            "AND G.isAgentToSubmit=true AND R.TransactionID= S.RptHeader AND AL.Deleteflag=False AND S.FundID=F.FundID AND R.BankCODE='" & szSelectedBankAccount & "' " & _
            "AND S.ClientStatementPrevID=" & trtoinclude & " AND ClientID ='" & szSelectedClient & "' AND R.UnitID=U.UnitNumber " & _
            "and AL.Deleteflag=false and AL.FromTran=R.TransactionID AND R.RDate <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "#"
            rsReceipt.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
            If Not rsReceipt.EOF Then
                    dblAmt = IIf(IsNull(rsReceipt.Fields.Item("Dr").Value), 0, rsReceipt.Fields.Item("Dr").Value)
                     'result is 175
            End If
            rsReceipt.Close
            Set rsReceipt = Nothing
            getAvailablefundsPreview = getAvailablefundsPreview - dblAmt
'    End If
 
   'c   (-): Sum of Supplier amounts Paid/Refunded ( allocated /purchase payment
'    If boolConsolidatedStatement = 1 Then
            whereProperty = "(P.UnitID in (" & ListOfProperties & ") OR isnull(P.UnitID) OR P.UnitID='' ) AND "
'    Else
'            whereProperty = "P.UnitID in (" & ListOfProperties & ") AND "
'    End If
'supplier payment
'Rem by anol 2023-08-27
'    szSQL = "Select   SUM(SWITCH(P.TYPE=24,PS.Amount,P.TYPE=8,-PS.Amount,P.TYPE=9,-PS.Amount)) as AMT from tlbPayment P,tlbPaymentSplit PS,Supplier SP where  " & _
'            "PS.Payheader=P.TransactionID AND SP.SupplierID=P.SageAccountNumber AND PS.Amount>PS.OSAmount  AND PS.ClientStatementPrevID=" & trtoinclude & " AND " & _
'            "Sp.Type='Supplier' AND P.ClientID ='" & szSelectedClient & "' " & _
'            "AND P.PDate <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "#"
        szSQL = "Select SUM(SWITCH(P.TYPE=24,PS.PaymentAmount,P.TYPE=8,-PS.PaymentAmount,P.TYPE=9,-PS.PaymentAmount)) as AMT from tlbPayment P,tlbPaymentSplit S," & _
            "PaytransactionsSplit PS,Supplier SP where (  PS.TransactionID=S.PayTransactionIDSplit OR PS.TOTran=P.TransactionID ) AND SP.SupplierID=P.SageAccountNumber AND P.TransactionID=S.PayHeader " & _
            "AND PS.Deleteflag=False AND S.ClientStatementPrevID=" & trtoinclude & " AND Sp.Type='Supplier' AND P.ClientID ='" & szSelectedClient & "' " & _
            "AND P.PDate <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "#"
     'I have added  OR PS.ToTran=P.TransactionID this part because now PPR is in action, and I am taking account of it.

    rsPayment.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
    If Not rsPayment.EOF Then
            dblAmt = IIf(IsNull(rsPayment.Fields.Item("AMT").Value), 0, rsPayment.Fields.Item("AMT").Value)
            dblSupplierPayment = dblAmt
        'result is -837
    End If
    rsPayment.Close
    Set rsPayment = Nothing
    getAvailablefundsPreview = getAvailablefundsPreview + dblAmt
'Take the vat amount from the allocation table
'    If bolisAgentToSubmit = True Then
        szSQL = "Select  SUM(AL.VatAmount)  as AMT from tlbPayment P,PayTransactionsSplit AL,tlbPaymentSplit S,Fund F,Supplier SP,Property PR,GLobalData G " & _
                "where  G.PropertyID=PR.PropertyID AND isAgentToSubmit=true  AND " & _
                "AL.TransactionID=S.PayTransactionIDSplit and SP.SupplierID=P.SageAccountNumber AND AL.Deleteflag=false and PR.propertyID=P.UNITID and " & _
                "Sp.Type='Supplier' AND P.TransactionID=S.PayHeader AND S.ClientStatementPrevID=" & trtoinclude & "" & _
                "AND S.FundID=F.FundID AND AL.FromTran=P.transactionID AND P.ClientID ='" & szSelectedClient & "' " & _
                "AND P.PDate <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "#"
        rsPayment.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
        If Not rsPayment.EOF Then
                dblAmt = IIf(IsNull(rsPayment.Fields.Item("AMT").Value), 0, rsPayment.Fields.Item("AMT").Value)
                'result is -15
        End If
        rsPayment.Close
        Set rsPayment = Nothing
         getAvailablefundsPreview = getAvailablefundsPreview - dblAmt
'    End If
    'd)  Add (+): Sum of Bank payments and receipts
'    If boolConsolidatedStatement = 1 Then
            whereProperty = "(B.PropertyID in (" & ListOfProperties & ") OR isnull(B.PropertyID) OR B.PropertyID ='' ) AND "
'    Else
'            whereProperty = "B.PropertyID in (" & ListOfProperties & ") AND "
'    End If
    
     szSQL = "Select  SUM(SWITCH(TransactionType=11 ,-B.NET_AMOUNT,TransactionType=12 ,B.NET_AMOUNT)) as AMT from tlbBankPayment B, Fund F  where B.DEPT_ID=F.FundID " & _
            "AND B.RentSumStatementPreview='" & trtoinclude & "' and clientID='" & szSelectedClient & "' " & _
            "AND  B.TRAN_DATE <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "#"
    rsBankPaymentAndRcpt.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
    If Not rsBankPaymentAndRcpt.EOF Then
        dblAmt = IIf(IsNull(rsBankPaymentAndRcpt.Fields.Item("AMT").Value), 0, rsBankPaymentAndRcpt.Fields.Item("AMT").Value)
           'result is 0
    End If
    rsBankPaymentAndRcpt.Close
    Set rsBankPaymentAndRcpt = Nothing
    getAvailablefundsPreview = getAvailablefundsPreview + dblAmt
    'f)  Less (-): Supplier OS Account balances for the client selected
    
        dblAmt = GetSupplierOSAmount
   
    dblSupplierOsAmount = -dblAmt
     
    'If negative then ignore this
    getAvailablefundsPreview = getAvailablefundsPreview - IIf(dblAmt < 0, 0, dblAmt)
    'it should be -40

    Dim rsNLposting As New ADODB.Recordset

'g)  Less (-): Client /Landlord OS balances for the client selected  and property selected amounts due to Client/Landlord not paid
         dblAmt = GetClientACBalance ' GetClientACBalanceModPreview(trToinclude)  ' -GetClientACBalance
          getAvailablefundsPreview = getAvailablefundsPreview + dblAmt
         'client payment
                whereProperty = "(P.UnitID in (" & ListOfProperties & ") OR isnull(P.UnitID) OR P.UnitID='' ) AND "
                szSQL = "Select   SUM(SWITCH(P.TYPE=24,PS.Amount,P.TYPE=8,-PS.Amount,P.TYPE=9,-PS.Amount)) as AMT from tlbPayment P,tlbPaymentSplit PS,Fund F,Supplier SP," & _
                "tblPurInv V,tlbPayment PI,PayTransactions PT where PT.fromtran=PI.transactionID and P.transactionID=PT.Totran and " & _
                "PS.Payheader=P.TransactionID AND SP.SupplierID=P.SageAccountNumber AND PS.Amount>PS.OSAmount AND PS.ClientStatementPrevID=" & trtoinclude & " AND " & _
                "PS.FundID=F.FundID and Sp.Type='Client' AND P.ClientID ='" & szSelectedClient & "' AND V.MY_ID=PI.PI AND isRentPayable=false " & _
                 "AND P.PDate <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "#"

                rsPayment.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
                If Not rsPayment.EOF Then
                        dblAmt = IIf(IsNull(rsPayment.Fields.Item("AMT").Value), 0, rsPayment.Fields.Item("AMT").Value)
                    'result is -837
                End If
                rsPayment.Close
                Set rsPayment = Nothing
    
    
         'COMING -35  dblAmt is negative then ignore
    'getAvailablefundsPreview = getAvailablefundsPreview + GetClientACBalance + GetLandLordACBalance
        getAvailablefundsPreview = getAvailablefundsPreview + dblAmt

          
        'landlord payment
          whereProperty = "(P.UnitID in (" & ListOfProperties & ") OR isnull(P.UnitID) OR P.UnitID='') AND "
'                szSQL = "Select   SUM(SWITCH(P.TYPE=24,PS.Amount,P.TYPE=8,-PS.Amount,P.TYPE=9,-PS.Amount)) as AMT from tlbPayment P,tlbPaymentSplit PS,Fund F,Supplier SP where  " & _
'                "PS.Payheader=P.TransactionID AND SP.SupplierID=P.SageAccountNumber AND PS.Amount>PS.OSAmount AND PS.ClientStatementPrevID=" & trtoinclude & " AND " & _
'                "PS.FundID=F.FundID and Sp.Type='LLORD' and  P.ClientID ='" & szSelectedClient & "' " & _
'                "AND P.PDate <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "#"
                szSQL = "Select   SUM(SWITCH(P.TYPE=24,PS.Amount,P.TYPE=8,-PS.Amount,P.TYPE=9,-PS.Amount)) as AMT from tlbPayment P,tlbPaymentSplit PS,Fund F,Supplier SP," & _
                "tblPurInv V,tlbPayment PI,PayTransactions PT where PT.fromtran=PI.transactionID and P.transactionID=PT.Totran and " & _
                "PS.Payheader=P.TransactionID AND SP.SupplierID=P.SageAccountNumber AND PS.Amount>PS.OSAmount AND PS.ClientStatementPrevID=" & trtoinclude & " AND " & _
                "PS.FundID=F.FundID and Sp.Type='LLORD' AND P.ClientID ='" & szSelectedClient & "' AND V.MY_ID=PI.PI AND isRentPayable=false " & _
                 "AND P.PDate <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "#"

                rsPayment.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
                If Not rsPayment.EOF Then
                        dblAmt = IIf(IsNull(rsPayment.Fields.Item("AMT").Value), 0, rsPayment.Fields.Item("AMT").Value)
                    'result is -837
                End If
                rsPayment.Close
                Set rsPayment = Nothing
    
    
         'COMING -35  dblAmt is negative then ignore
    'getAvailablefundsPreview = getAvailablefundsPreview + GetClientACBalance + GetLandLordACBalance
        getAvailablefundsPreview = getAvailablefundsPreview + dblAmt
        
        
        
        dblAmt = -GetLandLordACBalance '-GetLandLordACBalanceMODpreview(trToinclude) '-GetLandLordACBalance
        'GetLandLordACBalanceMODpreview
          'if dblAmt is negative then ignore
        getAvailablefundsPreview = getAvailablefundsPreview + dblAmt
     
     
    
    'getAvailablefundsPreview = getAvailablefundsPreview + GetClientACBalance + GetLandLordACBalance
       ' getAvailablefundsPreview = getAvailablefundsPreview + dblAmt
        
    'h)  Less (-): Managing Agent OS Balances for the client selected Management Fees due but not paid
    'GetAgentBalanceModPreview
    If chkShowDue.Value = 1 Then
        dblAmt = -GetAgentBalance ' -GetAgentBalanceModPreview(trToinclude) 'GetAgentBalance '
    End If
    
    dblagentBalance = -GetAgentBalance
    getAvailablefundsPreview = getAvailablefundsPreview + dblAmt
    '
    dblAmt = GetAGENTPaymentsPreview(trtoinclude)
    dblAgentPayment = dblAmt
    getAvailablefundsPreview = getAvailablefundsPreview + dblAmt
    getAvailablefundsPreview = Round(getAvailablefundsPreview, 2)
    
    Debug.Print getAvailablefundsPreview
    
    
    Dim ManagementFeeControl As String
    AccrualsControl = GetNominalCodeForControlAccount(adoConn, "Accruals Control Account (B/S)", szSelectedClient)
    'Dim rsNLPosting As New ADODB.Recordset
     If boolConsolidatedStatement = 1 Then
            whereProperty = "(PROPERTY_ID IN (" & ListOfProperties & ") OR isnull(PROPERTY_ID)) AND "
    Else
            whereProperty = "PROPERTY_ID in (" & ListOfProperties & ") AND "
    End If
    
    rsNLposting.Open "Select sum(AMOUNT) as AMT from NLPosting where " & whereProperty & _
                    "  NOMINAL_CODE='" & AccrualsControl & "' AND ClientID='" & _
                    szSelectedClient & "' AND DeleteFlag=false", adoConn, adOpenStatic, adLockReadOnly
    
    If Not rsNLposting.EOF Then
        dblAmt = IIf(IsNull(rsNLposting.Fields.Item("AMT").Value), 0, rsNLposting.Fields.Item("AMT").Value)
    End If
    rsNLposting.Close
    Set rsNLposting = Nothing
    getAvailablefundsPreview = getAvailablefundsPreview + dblAmt
    'j)  Less (-): Tenant Deposits received for the client selected
    'REMOVE THIS AS PER SPEC
    'getAvailablefundsPreview = getAvailablefundsPreview - GetRentDeposit
    rsNLposting.Open "Select sum(AMOUNT) as AMT from RetentionDetails where  isDeleted=false and BankCode='" & szSelectedBankAccount & "' AND " & _
                    "ClientID='" & szSelectedClient & "' AND  isDeleted=false AND RDate<=#" & _
                    Format(txtStatementDate1.text, " dd MMM yyyy") & "#  and isnull(statementID)", adoConn, adOpenStatic, adLockReadOnly
    If Not rsNLposting.EOF Then
        txtRetention.text = IIf(IsNull(rsNLposting.Fields.Item("AMT").Value), 0, rsNLposting.Fields.Item("AMT").Value)
    End If
    rsNLposting.Close
    
    getAvailablefundsPreview = getAvailablefundsPreview - Val(txtRetention.text)
    'newly added on 2023-02-10
    Dim rsMgtFee As New ADODB.Recordset
    rsMgtFee.Open "Select sum(MgtFeeAmtTotal) as Total from ManagementFeePreview", adoConn, adOpenStatic, adLockReadOnly
    If Not rsMgtFee.EOF Then
            getAvailablefundsPreview = getAvailablefundsPreview - IIf(IsNull(rsMgtFee("Total").Value), 0, rsMgtFee("Total").Value)
            dblManagementFee = -IIf(IsNull(rsMgtFee("Total").Value), 0, rsMgtFee("Total").Value)
    End If
    rsMgtFee.Close
    MsgBox "Available fund is: " & Round(getAvailablefundsPreview, 2)
    txtAvailableFunds.text = Round(getAvailablefundsPreview, 2)
    If bEditMode = False Then
        txtRentPayable.text = txtAvailableFunds.text
    End If

    Debug.Print dblTotalReceipt
    Debug.Print dblSupplierPayment
    Debug.Print dblSupplierOsAmount
   ' Debug.Print dblagentBalance
    Debug.Print dblManagementFee
    Debug.Print dblAgentPayment
    Debug.Print -txtRetention.text

End Function
Private Function GetRentDeposit(ByVal trxToinclude As Long) As Double
    Dim szSQL As String
'    Dim szSQL1 As String
    Dim szSQL2 As String
'    Dim szSQL3 As String
    Dim rsPayment As New ADODB.Recordset
    Dim rsReceipt1 As New ADODB.Recordset
    Dim rsReceipt2 As New ADODB.Recordset
    Dim rsReceipt3 As New ADODB.Recordset
    Dim rsReceipt As New ADODB.Recordset
    Dim adoConn As New ADODB.Connection
    Dim dblAmt, dblamt1, dblamt2, dblamt3 As Double
    adoConn.Open getConnectionString
    'tlbBankPayment
    'BANK_AC
    'TRAN_TYPE
    'DEPT_ID
    'propertyID
    'clientID
    'NET_AMOUNT
    Dim whereProperty  As String
'    If boolConsolidatedStatement = 1 Then
'            whereProperty = "(U.PROPERTYID IN (" & ListOfProperties & ") OR isnull(U.PROPERTYID) or U.PROPERTYID='' ) AND "
'    Else
'            whereProperty = "U.PROPERTYID  in (" & ListOfProperties & ") AND "
'    End If
   
    szSQL = "Select  SUM(SWITCH(TYPE=1,RS.Amount,TYPE=2,RS.Amount,TYPE=3,-RS.Amount,TYPE=4,-RS.Amount,TYPE=23,-RS.Amount)) as DR from tlbReceipt R,tlbReceiptSplit RS,Fund F, Units U " & _
            "where R.UnitID=U.UnitNumber AND R.TransactionID=RS.rptHeader AND RS.ClientStatementID=" & trxToinclude & " AND " & whereProperty & "  TYPE IN(3,4,23) AND R.FundID=F.FundID and F.FundCode='RENTDEPOSIT' AND ClientID ='" & _
             szSelectedClient & "' AND R.RDate <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "# "

'    If boolConsolidatedStatement = 1 Then
'            whereProperty = "(B.PROPERTYID IN (" & ListOfProperties & ") OR isnull(B.PROPERTYID) or B.PROPERTYID='' ) AND "
'    Else
'            whereProperty = "B.PROPERTYID  in (" & ListOfProperties & ") AND "
'    End If
    
    
    szSQL2 = "Select  SUM(SWITCH(TransactionType=11,B.NET_AMOUNT,TransactionType=12,-B.NET_AMOUNT)) as DR from tlbBankPayment B,Fund F where TransactionType " & _
            " IN(11,12) AND " & whereProperty & " B.DEPT_ID=F.FundID and F.FundCode='RENTDEPOSIT' AND B.RentSumStatement=" & trxToinclude & " AND B.ClientID ='" & szSelectedClient & "'" & _
             "AND B.TRAN_DATE <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "# "
    
    rsReceipt.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
    If Not rsReceipt.EOF Then
        dblAmt = IIf(IsNull(rsReceipt.Fields.Item("Dr").Value), 0, rsReceipt.Fields.Item("Dr").Value)
    End If
    rsReceipt.Close

    
    rsReceipt2.Open szSQL2, adoConn, adOpenStatic, adLockReadOnly
    If Not rsReceipt2.EOF Then
        dblamt2 = IIf(IsNull(rsReceipt2.Fields.Item("Dr").Value), 0, rsReceipt2.Fields.Item("Dr").Value)
    End If
    rsReceipt2.Close

    GetRentDeposit = dblAmt + dblamt2
    adoConn.Close
    Set adoConn = Nothing
End Function
'Dim szCurrentRentsummarySTID As String
Private Function GetRentDepositPreview(ByVal trxToinclude As Long) As Double
    Dim szSQL As String
    Dim szSQL2 As String
    Dim rsPayment As New ADODB.Recordset
    Dim rsReceipt1 As New ADODB.Recordset
    Dim rsReceipt2 As New ADODB.Recordset
    Dim rsReceipt3 As New ADODB.Recordset
    Dim rsReceipt As New ADODB.Recordset
    Dim adoConn As New ADODB.Connection
    Dim dblAmt, dblamt1, dblamt2, dblamt3 As Double
    adoConn.Open getConnectionString
    
    Dim whereProperty  As String
    szSQL = "Select  SUM(SWITCH(TYPE=1,RS.Amount,TYPE=2,RS.Amount,TYPE=3,-RS.Amount,TYPE=4,-RS.Amount,TYPE=23,-RS.Amount)) as DR from tlbReceipt R,tlbReceiptSplit RS,Fund F, Units U " & _
            "where R.UnitID=U.UnitNumber AND R.TransactionID=RS.rptHeader AND RS.ClientStatementPrevID=" & trxToinclude & " AND " & _
            whereProperty & "  TYPE IN(3,4,23) AND R.FundID=F.FundID and F.FundCode='RENTDEPOSIT' AND ClientID ='" & _
             szSelectedClient & "' AND R.RDate <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "# "
    
    szSQL2 = "Select  SUM(SWITCH(TransactionType=11,B.NET_AMOUNT,TransactionType=12,-B.NET_AMOUNT)) as DR from tlbBankPayment B,Fund F where TransactionType " & _
            " IN(11,12) AND " & whereProperty & " B.DEPT_ID=F.FundID and F.FundCode='RENTDEPOSIT' AND B.RentSumStatementPreview='" & trxToinclude & "'  AND B.ClientID ='" & szSelectedClient & "'" & _
             "AND B.TRAN_DATE <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "# "
    
    rsReceipt.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
    If Not rsReceipt.EOF Then
        dblAmt = IIf(IsNull(rsReceipt.Fields.Item("Dr").Value), 0, rsReceipt.Fields.Item("Dr").Value)
    End If
    rsReceipt.Close
    
    rsReceipt2.Open szSQL2, adoConn, adOpenStatic, adLockReadOnly
    If Not rsReceipt2.EOF Then
        dblamt2 = IIf(IsNull(rsReceipt2.Fields.Item("Dr").Value), 0, rsReceipt2.Fields.Item("Dr").Value)
    End If
    rsReceipt2.Close

    GetRentDepositPreview = dblAmt + dblamt2
    adoConn.Close
    Set adoConn = Nothing
End Function
Private Function getAvailablefundsProduce(dblLasClosingBalance As Double, ByVal trtoinclude As Long) As Double 'this one I am using while finalize
    'Pass propery as parameter for selected property
    'No property spec:
    'Exit Function
    Dim adoConn As New ADODB.Connection
    Dim rsReceipt As New ADODB.Recordset
    Dim rsPayment As New ADODB.Recordset
    Dim rsBankPaymentAndRcpt As New ADODB.Recordset
    Dim dblAmt As Double
    adoConn.Open getConnectionString
    Dim szSQL As String
    Dim AccrualsControl As String
    Dim whereProperty As String
    'we are not using property filter here
    'B )***********************  Sum of Rent received Paid/Refunded ***********************************
    'AND tlbReceipt.RDate >=#" & Format(txtLastStatementDate1.text, "dd/mmm/yyyy") & "# AND tlbReceipt.RDate <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "#"
    Dim rsGlobalData1 As New ADODB.Recordset
    Dim bolVatOptionEnabled As Boolean
    Dim bolOptedTotax As String
    Dim bolisAgentToSubmit As Boolean
    Dim strManagingAgentID As String
    'OR P.RentSumStatement='" & trToinclude & "'
    
    rsGlobalData1.Open "Select vatOptionEnabled,isAgentToSubmit from Globaldata G,Property P where P.PropertyID=G.PropertyID AND P.PropertyID='" & _
                       szSQL & "' ", adoConn, adOpenStatic, adLockReadOnly
                       
    If Not rsGlobalData1.EOF Then
           bolVatOptionEnabled = rsGlobalData1("vatOptionEnabled").Value
           bolisAgentToSubmit = rsGlobalData1("isAgentToSubmit").Value
    End If
    rsGlobalData1.Close

                
    If boolConsolidatedStatement = 1 Then
            whereProperty = "(U.PropertyID in (" & ListOfProperties & ")OR isnull(U.PropertyID) OR U.PropertyID='' ) AND "
    Else
            whereProperty = "U.PropertyID in (" & ListOfProperties & ") AND "
    End If
    
'total Reciept
    szSQL = "Select  SUM(switch(type=3,S.Amount,type=4,S.Amount,type=23,-S.Amount)) as DR from tlbReceipt R,tlbReceiptSplit S where R.TransactionID= S.RptHeader " & _
    "AND ClientID ='" & szSelectedClient & "' and S.ClientStatementID=" & trtoinclude & ""
    rsReceipt.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
    If Not rsReceipt.EOF Then
            dblAmt = IIf(IsNull(rsReceipt.Fields.Item("Dr").Value), 0, rsReceipt.Fields.Item("Dr").Value)
             'result is 175
    End If
    rsReceipt.Close
    Set rsReceipt = Nothing
    
    
    getAvailablefundsProduce = dblLasClosingBalance + dblAmt
    'Vat calculation of B)take the  VAT amount from the allocation table and deduct it
            szSQL = "Select  SUM(AL.VatAmount) as DR from tlbReceipt R,rptTransactionsSplit AL,tlbReceiptSplit S, Units U,GlobalData G where G.PropertyID=U.PropertyID " & _
            "AND G.isAgentToSubmit=true AND R.TransactionID= S.RptHeader AND AL.Deleteflag=False  AND R.BankCODE='" & szSelectedBankAccount & "' " & _
            "AND S.ClientStatementID=" & trtoinclude & " AND ClientID ='" & szSelectedClient & "' AND R.UnitID=U.UnitNumber " & _
            "AND AL.Deleteflag=false and AL.FromTran=R.TransactionID AND R.RDate <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "#"
            rsReceipt.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
            If Not rsReceipt.EOF Then
                    dblAmt = IIf(IsNull(rsReceipt.Fields.Item("Dr").Value), 0, rsReceipt.Fields.Item("Dr").Value)
                     'result is 175
            End If
            rsReceipt.Close
            Set rsReceipt = Nothing
            getAvailablefundsProduce = getAvailablefundsProduce - dblAmt
'    End If
 
   'c   (-): Sum of Supplier amounts Paid/Refunded ( allocated /purchase payment
'    If boolConsolidatedStatement = 1 Then
            whereProperty = "(P.UnitID in (" & ListOfProperties & ") OR isnull(P.UnitID) OR P.UnitID='' ) AND "
'    Else
'            whereProperty = "P.UnitID in (" & ListOfProperties & ") AND "
'    End If

'
'    szSQL = "Select SUM(SWITCH(P.TYPE=24,PS.PaymentAmount,P.TYPE=8,-PS.PaymentAmount,P.TYPE=9,-PS.PaymentAmount)) as AMT from tlbPayment P,tlbPaymentSplit S," & _
'            "PaytransactionsSplit PS,Fund F,Supplier SP where  PS.TransactionID=S.PayTransactionIDSplit AND SP.SupplierID=P.SageAccountNumber AND P.TransactionID=S.PayHeader " & _
'            "AND PS.Deleteflag=False AND S.ClientStatementID=" & trtoinclude & " AND Sp.Type='Supplier' AND P.ClientID ='" & szSelectedClient & "' " & _
'            "AND P.PDate <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "#"
'szSQL = "Select   SUM(SWITCH(P.TYPE=24,PS.Amount,P.TYPE=8,-PS.Amount,P.TYPE=9,-PS.Amount)) as AMT from tlbPayment P,tlbPaymentSplit PS,Supplier SP where  " & _
'            "PS.Payheader=P.TransactionID AND SP.SupplierID=P.SageAccountNumber AND PS.Amount>PS.OSAmount  AND PS.ClientStatementID=" & trtoinclude & " AND " & _
'            "Sp.Type='Supplier' AND P.ClientID ='" & szSelectedClient & "' " & _
'            "AND P.PDate <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "#"
'Rem by anol 2023-08-27
'szSQL = "Select SUM(SWITCH(P.TYPE=24,PS.PaymentAmount,P.TYPE=8,-PS.PaymentAmount,P.TYPE=9,-PS.PaymentAmount)) as AMT from tlbPayment P,tlbPaymentSplit S," & _
'            "PaytransactionsSplit PS,Supplier SP where  PS.TransactionID=S.PayTransactionIDSplit AND SP.SupplierID=P.SageAccountNumber AND P.TransactionID=S.PayHeader " & _
'            "AND PS.Deleteflag=False AND S.ClientStatementID=" & trtoinclude & " AND Sp.Type='Supplier' AND P.ClientID ='" & szSelectedClient & "' " & _
'            "AND P.PDate <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "#"
szSQL = "Select SUM(SWITCH(P.TYPE=24,PS.PaymentAmount,P.TYPE=8,-PS.PaymentAmount,P.TYPE=9,-PS.PaymentAmount)) as AMT from tlbPayment P,tlbPaymentSplit S," & _
            "PaytransactionsSplit PS,Supplier SP where (  PS.TransactionID=S.PayTransactionIDSplit OR PS.TOTran=P.TransactionID ) AND SP.SupplierID=P.SageAccountNumber AND P.TransactionID=S.PayHeader " & _
            "AND PS.Deleteflag=False AND S.ClientStatementID=" & trtoinclude & " AND Sp.Type='Supplier' AND P.ClientID ='" & szSelectedClient & "' " & _
            "AND P.PDate <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "#"
'I have added  OR PS.ToTran=P.TransactionID this part because now PPR is in action, and I am taking account of it.


            
    rsPayment.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
    If Not rsPayment.EOF Then
            dblAmt = IIf(IsNull(rsPayment.Fields.Item("AMT").Value), 0, rsPayment.Fields.Item("AMT").Value)
        'result is -837
    End If
    rsPayment.Close
    Set rsPayment = Nothing
    getAvailablefundsProduce = getAvailablefundsProduce + dblAmt
'Take the vat amount from the allocation table
'    If bolisAgentToSubmit = True Then
        szSQL = "Select  SUM(AL.VatAmount)  as AMT from tlbPayment P,PayTransactionsSplit AL,tlbPaymentSplit S,Supplier SP,Property PR,GLobalData G " & _
                "where  G.PropertyID=PR.PropertyID AND isAgentToSubmit=true  AND AL.TransactionID=S.PayTransactionIDSplit " & _
                "AND SP.SupplierID=P.SageAccountNumber AND AL.Deleteflag=false and PR.propertyID=P.UNITID and " & _
                "Sp.Type='Supplier' AND P.TransactionID=S.PayHeader AND S.ClientStatementID=" & trtoinclude & "" & _
                "AND AL.FromTran=P.transactionID AND P.ClientID ='" & szSelectedClient & "' " & _
                "AND P.PDate <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "#"
        rsPayment.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
        If Not rsPayment.EOF Then
                dblAmt = IIf(IsNull(rsPayment.Fields.Item("AMT").Value), 0, rsPayment.Fields.Item("AMT").Value)
                'result is -15
        End If
        rsPayment.Close
        Set rsPayment = Nothing
         getAvailablefundsProduce = getAvailablefundsProduce - dblAmt
'    End If
    'd)  Add (+): Sum of Bank payments and receipts
'    If boolConsolidatedStatement = 1 Then
            whereProperty = "(B.PropertyID in (" & ListOfProperties & ") OR isnull(B.PropertyID) OR B.PropertyID ='' ) AND "
'    Else
'            whereProperty = "B.PropertyID in (" & ListOfProperties & ") AND "
'    End If
    
     szSQL = "Select  SUM(SWITCH(TransactionType=11 ,-B.NET_AMOUNT,TransactionType=12 ,B.NET_AMOUNT)) as AMT from tlbBankPayment B where " & _
            "B.RentSumStatement='" & trtoinclude & "' and clientID='" & szSelectedClient & "' " & _
            "AND  B.TRAN_DATE <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "#"
    rsBankPaymentAndRcpt.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
    If Not rsBankPaymentAndRcpt.EOF Then
        dblAmt = IIf(IsNull(rsBankPaymentAndRcpt.Fields.Item("AMT").Value), 0, rsBankPaymentAndRcpt.Fields.Item("AMT").Value)
           'result is 0
    End If
    rsBankPaymentAndRcpt.Close
    Set rsBankPaymentAndRcpt = Nothing
    getAvailablefundsProduce = getAvailablefundsProduce + dblAmt
    'f)  Less (-): Supplier OS Account balances for the client selected
    dblAmt = GetSupplierOSAmount
    'If negative then ignore this
    getAvailablefundsProduce = getAvailablefundsProduce - IIf(dblAmt < 0, 0, dblAmt)
    'it should be -40

    Dim rsNLposting As New ADODB.Recordset

'g)  Less (-): Client /Landlord OS balances for the client selected  and property selected amounts due to Client/Landlord not paid
         dblAmt = GetClientACBalance ' GetClientACBalanceModPreview(trToinclude)  ' -GetClientACBalance
          getAvailablefundsProduce = getAvailablefundsProduce + dblAmt
         'client payment
                whereProperty = "(P.UnitID in (" & ListOfProperties & ") OR isnull(P.UnitID) OR P.UnitID='' ) AND "
'                szSQL = "Select   SUM(SWITCH(P.TYPE=24,PS.Amount,P.TYPE=8,-PS.Amount,P.TYPE=9,-PS.Amount)) as AMT from tlbPayment P,tlbPaymentSplit PS,Supplier SP where  " & _
'                "PS.PayHeader=P.TransactionID AND SP.SupplierID=P.SageAccountNumber  AND PS.Amount>PS.OSAmount AND PS.ClientStatementID=" & trtoinclude & " AND " & _
'                "Sp.Type='Client' AND P.ClientID ='" & szSelectedClient & "' " & _
'                "AND P.PDate <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "#"
                
                 szSQL = "Select   SUM(SWITCH(P.TYPE=24,PS.Amount,P.TYPE=8,-PS.Amount,P.TYPE=9,-PS.Amount)) as AMT from tlbPayment P,tlbPaymentSplit PS,Fund F,Supplier SP," & _
                "tblPurInv V,tlbPayment PI,PayTransactions PT where PT.fromtran=PI.transactionID and P.transactionID=PT.Totran and " & _
                "PS.Payheader=P.TransactionID AND SP.SupplierID=P.SageAccountNumber AND PS.Amount>PS.OSAmount AND PS.ClientStatementID=" & trtoinclude & " AND " & _
                "PS.FundID=F.FundID and Sp.Type='Client' AND P.ClientID ='" & szSelectedClient & "' AND V.MY_ID=PI.PI AND isRentPayable=false " & _
                 "AND P.PDate <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "#"

                rsPayment.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
                If Not rsPayment.EOF Then
                        dblAmt = IIf(IsNull(rsPayment.Fields.Item("AMT").Value), 0, rsPayment.Fields.Item("AMT").Value)
                    'result is -837
                End If
                rsPayment.Close
                Set rsPayment = Nothing
    
    
         'COMING -35  dblAmt is negative then ignore
    'getAvailablefundsProduce = getAvailablefundsProduce + GetClientACBalance + GetLandLordACBalance
        getAvailablefundsProduce = getAvailablefundsProduce + dblAmt

          
        'landlord payment
          whereProperty = "(P.UnitID in (" & ListOfProperties & ") OR isnull(P.UnitID) OR P.UnitID='') AND "
'                szSQL = "Select   SUM(SWITCH(P.TYPE=24,PS.Amount,P.TYPE=8,-PS.Amount,P.TYPE=9,-PS.Amount)) as AMT from tlbPayment P,tlbPaymentSplit PS,Supplier SP where  " & _
'                "PS.PayHeader=P.TransactionID AND SP.SupplierID=P.SageAccountNumber  AND PS.Amount>PS.OSAmount AND  PS.ClientStatementID=" & trtoinclude & " AND " & _
'                "Sp.Type='LLORD' and  P.ClientID ='" & szSelectedClient & "' " & _
'                "AND P.PDate <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "#"
                szSQL = "Select   SUM(SWITCH(P.TYPE=24,PS.Amount,P.TYPE=8,-PS.Amount,P.TYPE=9,-PS.Amount)) as AMT from tlbPayment P,tlbPaymentSplit PS,Fund F,Supplier SP," & _
                "tblPurInv V,tlbPayment PI,PayTransactions PT where PT.fromtran=PI.transactionID and P.transactionID=PT.Totran and " & _
                "PS.Payheader=P.TransactionID AND SP.SupplierID=P.SageAccountNumber AND PS.Amount>PS.OSAmount AND PS.ClientStatementID=" & trtoinclude & " AND " & _
                "PS.FundID=F.FundID and Sp.Type='LLORD' AND P.ClientID ='" & szSelectedClient & "' AND V.MY_ID=PI.PI AND isRentPayable=false " & _
                 "AND P.PDate <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "#"

                rsPayment.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
                If Not rsPayment.EOF Then
                        dblAmt = IIf(IsNull(rsPayment.Fields.Item("AMT").Value), 0, rsPayment.Fields.Item("AMT").Value)
                    'result is -837
                End If
                rsPayment.Close
                Set rsPayment = Nothing
    
    
         'COMING -35  dblAmt is negative then ignore
    'getAvailablefundsProduce = getAvailablefundsProduce + GetClientACBalance + GetLandLordACBalance
        getAvailablefundsProduce = getAvailablefundsProduce + dblAmt
        
        
        
        dblAmt = -GetLandLordACBalance '-GetLandLordACBalanceMODpreview(trToinclude) '-GetLandLordACBalance
        'GetLandLordACBalanceMODpreview
          'if dblAmt is negative then ignore
        getAvailablefundsProduce = getAvailablefundsProduce + dblAmt
     
     
    
    'getAvailablefundsPreview = getAvailablefundsPreview + GetClientACBalance + GetLandLordACBalance
        getAvailablefundsProduce = getAvailablefundsProduce + dblAmt
        
    'h)  Less (-): Managing Agent OS Balances for the client selected Management Fees due but not paid
    'GetAgentBalanceModPreview
      If chkShowDue.Value = 1 Then
        dblAmt = -GetAgentBalance ' -GetAgentBalanceModPreview(trToinclude) 'GetAgentBalance '
    End If
    
    'dblAmt = -GetAgentBalance ' -GetAgentBalanceModPreview(trToinclude) 'GetAgentBalance ' rem by anol 2023-08-23
    getAvailablefundsProduce = getAvailablefundsProduce + dblAmt
    '
    dblAmt = GetAGENTPayments(trtoinclude)
    getAvailablefundsProduce = getAvailablefundsProduce + dblAmt
    getAvailablefundsProduce = Round(getAvailablefundsProduce, 2)
    
    Debug.Print getAvailablefundsProduce
    
    
    Dim ManagementFeeControl As String
    AccrualsControl = GetNominalCodeForControlAccount(adoConn, "Accruals Control Account (B/S)", szSelectedClient)
    'Dim rsNLPosting As New ADODB.Recordset
     If boolConsolidatedStatement = 1 Then
            whereProperty = "(PROPERTY_ID IN (" & ListOfProperties & ") OR isnull(PROPERTY_ID)) AND "
    Else
            whereProperty = "PROPERTY_ID in (" & ListOfProperties & ") AND "
    End If
    
    rsNLposting.Open "Select sum(AMOUNT) as AMT from NLPosting where " & whereProperty & _
                    "  NOMINAL_CODE='" & AccrualsControl & "' AND ClientID='" & _
                    szSelectedClient & "' AND DeleteFlag=false", adoConn, adOpenStatic, adLockReadOnly
    
    If Not rsNLposting.EOF Then
        dblAmt = IIf(IsNull(rsNLposting.Fields.Item("AMT").Value), 0, rsNLposting.Fields.Item("AMT").Value)
    End If
    rsNLposting.Close
    Set rsNLposting = Nothing
    getAvailablefundsProduce = getAvailablefundsProduce + dblAmt
    'j)  Less (-): Tenant Deposits received for the client selected
    'REMOVE THIS AS PER SPEC
    'getAvailablefundsProduce = getAvailablefundsProduce - GetRentDeposit
'    rsNLposting.Open "Select sum(AMOUNT) as AMT from RetentionDetails where BankCode='" & szSelectedBankAccount & "' and " & _
'                    "ClientID='" & szSelectedClient & "' and isDeleted=false AND RDate<=#" & Format(txtStatementDate1.text, " dd MMM yyyy") & "# ", adoconn, adOpenStatic, adLockReadOnly

'     rsNLposting.Open "Select sum(AMOUNT) as AMT from RetentionDetails where  isDeleted=false  and StatementID=" & trtoinclude & "", adoconn, adOpenStatic, adLockReadOnly
'
'    If Not rsNLposting.EOF Then
'        txtRetention.text = IIf(IsNull(rsNLposting.Fields.Item("AMT").Value), 0, rsNLposting.Fields.Item("AMT").Value)
'    End If
'    rsNLposting.Close
    rsNLposting.Open "Select sum(AMOUNT) as AMT from RetentionDetails where  isDeleted=false and BankCode='" & szSelectedBankAccount & "' AND " & _
                    "ClientID='" & szSelectedClient & "' AND statementID=" & _
                    trtoinclude & " ", adoConn, adOpenStatic, adLockReadOnly
    If Not rsNLposting.EOF Then
        txtRetention.text = IIf(IsNull(rsNLposting.Fields.Item("AMT").Value), 0, rsNLposting.Fields.Item("AMT").Value)
    End If
    rsNLposting.Close

    getAvailablefundsProduce = getAvailablefundsProduce - Val(txtRetention.text)
    MsgBox "Available fund is: " & Round(getAvailablefundsProduce, 2)
    txtAvailableFunds.text = Round(getAvailablefundsProduce, 2)
    If bEditMode = False Then
        txtRentPayable.text = txtAvailableFunds.text
    End If
     
End Function
Private Function GetPaymentsonAccount(ByVal trxID As Long) As Double
    Dim rsPayment As New ADODB.Recordset
    Dim adoConn As New ADODB.Connection
    adoConn.Open getConnectionString
    Dim szSQL As String
    szSQL = "Select  SUM(S.Amount) as AMT from tlbPayment P,tlbPaymentSplit S,Fund F,Supplier SP where  SP.SupplierID=P.SAGEACCOUNTNUMBER AND " & _
            "P.TransactionID=S.PayHeader AND P.PDate <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "# AND P.TYPE  " & _
            "IN(9) AND S.FundID=F.FundID AND  P.BankCODE='" & szSelectedBankAccount & "' and  SP.SupplierID ='" & _
            szSelectedClient & "' AND S.ClientStatementID=" & trxID & " AND SP.Type in ('CLIENT','LLORD') "
    rsPayment.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
    If Not rsPayment.EOF Then
        GetPaymentsonAccount = IIf(IsNull(rsPayment.Fields.Item("AMT").Value), 0, rsPayment.Fields.Item("AMT").Value)
    End If
    rsPayment.Close
    Set rsPayment = Nothing
    adoConn.Close
    Set adoConn = Nothing
    
End Function
Private Function GetPaymentsonAccountPreview(ByVal trxID As Long) As Double
    Dim rsPayment As New ADODB.Recordset
    Dim adoConn As New ADODB.Connection
    adoConn.Open getConnectionString
    Dim szSQL As String
    szSQL = "Select  SUM(S.Amount) as AMT from tlbPayment P,tlbPaymentSplit S,Fund F,Supplier SP where  SP.SupplierID=P.SAGEACCOUNTNUMBER AND " & _
            "P.TransactionID=S.PayHeader AND P.PDate <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "# AND P.TYPE  " & _
            "IN(9) AND S.FundID=F.FundID AND  P.BankCODE='" & szSelectedBankAccount & "' and  SP.SupplierID ='" & _
            szSelectedClient & "' AND S.ClientStatementPrevID=" & trxID & " AND SP.Type in ('CLIENT','LLORD') "
    rsPayment.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
    If Not rsPayment.EOF Then
        GetPaymentsonAccountPreview = IIf(IsNull(rsPayment.Fields.Item("AMT").Value), 0, rsPayment.Fields.Item("AMT").Value)
    End If
    rsPayment.Close
    Set rsPayment = Nothing
    adoConn.Close
    Set adoConn = Nothing
    
End Function
Private Function GetTenantReceiptsFinalized(ByVal trtoinclude As Long) As Double
    Dim rsPayment As New ADODB.Recordset
    Dim szSQL As String
    Dim adoConn As New ADODB.Connection
    Dim rsReceipt As New ADODB.Recordset
    Dim whereProperty As String
'    If boolConsolidatedStatement = 1 Then
'            whereProperty = "(U.PropertyID in (" & ListOfProperties & ") OR isnull(U.PropertyID)) AND "
'    Else
'            whereProperty = "U.PropertyID  in (" & ListOfProperties & ") AND "
'    End If
    
        adoConn.Open getConnectionString
        
        szSQL = "Select  SUM(SWITCH(R.TYPE=23,-RS.Amount,R.TYPE=3,RS.Amount,R.TYPE=4,RS.Amount)) as AMT from tlbReceipt R,tlbReceiptSplit RS,Fund F,Units U where " & _
            "R.TransactionID=RS.RptHeader AND R.TYPE IN(3,4,23) AND R.RDate <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "# " & _
            "AND RS.FundID=F.FundID AND  R.BankCODE='" & szSelectedBankAccount & "'  AND  RS.ClientStatementID=" & trtoinclude & " and  R.UnitID=U.UnitNumber AND " & whereProperty & " R.ClientID ='" & szSelectedClient & "' "
        
        rsPayment.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
        If Not rsPayment.EOF Then
            GetTenantReceiptsFinalized = IIf(IsNull(rsPayment.Fields.Item("AMT").Value), 0, rsPayment.Fields.Item("AMT").Value)
        End If
        rsPayment.Close
        
        szSQL = "Select  SUM(AL.VatAmount) as DR from tlbReceipt R,rptTransactions AL,tlbReceiptSplit S,Fund F, Units U,GLobalData G where G.PropertyID=U.PropertyID AND " & _
        "G.isAgentToSubmit=true AND R.TransactionID= S.RptHeader AND TYPE IN(3,4,23) AND S.FundID=F.FundID AND R.BankCODE='" & szSelectedBankAccount & "' " & _
        " and F.FundCode<>'TENANTDEPOSIT'  AND S.ClientStatementID=" & trtoinclude & " AND ClientID ='" & szSelectedClient & "' " & _
        "AND R.UnitID=U.UnitNumber and AL.Deleteflag=false and AL.FromTran=R.TransactionID  AND " & _
        whereProperty & "  R.RDate <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "#"
        rsReceipt.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
        If Not rsReceipt.EOF Then
            GetTenantReceiptsFinalized = GetTenantReceiptsFinalized - IIf(IsNull(rsReceipt.Fields.Item("Dr").Value), 0, rsReceipt.Fields.Item("Dr").Value)
         'result is 175
        End If
        rsReceipt.Close
        Set rsReceipt = Nothing
        adoConn.Close
        Set adoConn = Nothing
End Function
Private Function GetTenantReceiptsPreview(ByVal trtoinclude As Long) As Double
    Dim rsPayment As New ADODB.Recordset
    Dim szSQL As String
    Dim adoConn As New ADODB.Connection
    Dim rsReceipt As New ADODB.Recordset
    Dim whereProperty As String
'    If boolConsolidatedStatement = 1 Then
'            whereProperty = "(U.PropertyID in (" & ListOfProperties & ") OR isnull(U.PropertyID)) AND "
'    Else
'            whereProperty = "U.PropertyID  in (" & ListOfProperties & ") AND "
'    End If
    
        adoConn.Open getConnectionString
        
        szSQL = "Select  SUM(SWITCH(R.TYPE=23,-RS.Amount,R.TYPE=3,RS.Amount,R.TYPE=4,RS.Amount)) as AMT from tlbReceipt R,tlbReceiptSplit RS,Fund F,Units U where " & _
            "R.TransactionID=RS.RptHeader AND R.TYPE IN(3,4,23) AND R.RDate <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "# " & _
            "AND RS.FundID=F.FundID AND  R.BankCODE='" & szSelectedBankAccount & "'  AND  RS.ClientStatementPrevID=" & _
            trtoinclude & " and  R.UnitID=U.UnitNumber AND " & whereProperty & " R.ClientID ='" & szSelectedClient & "' "
        
        rsPayment.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
        If Not rsPayment.EOF Then
            GetTenantReceiptsPreview = IIf(IsNull(rsPayment.Fields.Item("AMT").Value), 0, rsPayment.Fields.Item("AMT").Value)
        End If
        rsPayment.Close
        
        szSQL = "Select  SUM(AL.VatAmount) as DR from tlbReceipt R,rptTransactions AL,tlbReceiptSplit S,Fund F, Units U,GLobalData G where G.PropertyID=U.PropertyID AND " & _
        "G.isAgentToSubmit=true AND R.TransactionID= S.RptHeader AND TYPE IN(3,4,23) AND S.FundID=F.FundID AND R.BankCODE='" & szSelectedBankAccount & "' " & _
        "AND F.FundCode<>'TENANTDEPOSIT'  AND  S.ClientStatementPrevID=" & trtoinclude & " AND ClientID ='" & szSelectedClient & "' " & _
        "AND R.UnitID=U.UnitNumber and AL.Deleteflag=false and AL.FromTran=R.TransactionID  AND " & _
        whereProperty & "  R.RDate <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "#"
        rsReceipt.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
        If Not rsReceipt.EOF Then
            GetTenantReceiptsPreview = GetTenantReceiptsPreview - IIf(IsNull(rsReceipt.Fields.Item("Dr").Value), 0, rsReceipt.Fields.Item("Dr").Value)
         'result is 175
        End If
        rsReceipt.Close
        Set rsReceipt = Nothing
        adoConn.Close
        Set adoConn = Nothing
End Function
Private Function GetSupplierPayment(ByVal trxToinclude As Long) As Double
    Dim rsPayment As New ADODB.Recordset
    Dim rsReceipt As New ADODB.Recordset
    Dim szSQL As String
    Dim adoConn As New ADODB.Connection
    Dim whereProperty As String
    Dim dblAmt As Double
    'F.CategoryCode = 1 Fund category 1 Means rent
    'Implement switch here in SQL
'     If boolConsolidatedStatement = 1 Then
'            whereProperty = "AND (P.UNITID IN (" & ListOfProperties & ") OR isnull(P.UNITID) OR P.UNITID='') "
'    Else
'            whereProperty = "AND P.UNITID in (" & ListOfProperties & ") "
'    End If
    adoConn.Open getConnectionString
    szSQL = "Select  SUM(SWITCH(P.TYPE =24,S.Amount,P.TYPE =8,-S.Amount,P.TYPE =9,-S.Amount)) as AMT from tlbPayment P,tlbPaymentSplit S,Fund F,Supplier SP where " & _
            "SP.SupplierID=P.SageaccountNumber AND  P.TransactionID=S.PayHeader AND P.TYPE IN(8,9,24) AND " & _
            "P.PDate <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "# AND " & _
            "S.FundID=F.FundID AND S.ClientStatementID=" & trxToinclude & " AND P.ClientID ='" & szSelectedClient & "' AND (SP.Type='Supplier' ) " & whereProperty
    
    rsPayment.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
    If Not rsPayment.EOF Then
        GetSupplierPayment = IIf(IsNull(rsPayment.Fields.Item("AMT").Value), 0, rsPayment.Fields.Item("AMT").Value)
    End If
    rsPayment.Close
'    If boolConsolidatedStatement = 1 Then
'            whereProperty = "AND (G.PropertyID IN (" & ListOfProperties & ") OR isnull(G.PropertyID) OR G.PropertyID='') "
'    Else
'            whereProperty = "AND G.PropertyID in (" & ListOfProperties & ") "
'    End If
             szSQL = "Select  SUM(AL.VatAmount)  as AMT from tlbPayment P,PayTransactions AL,tlbPaymentSplit S,Fund F,Supplier SP,Property PR,GLobalData G  where  G.PropertyID=PR.PropertyID AND isAgentToSubmit=true  AND " & _
                "SP.SupplierID=P.SageAccountNumber AND AL.Deleteflag=false and PR.propertyID=p.UNITID AND P.TransactionID=S.PayHeader AND P.TYPE IN(8,9,24) AND S.ClientStatementID=" & trxToinclude & " AND " & _
                "S.FundID=F.FundID AND AL.FromTran=P.transactionID " & whereProperty & " AND P.ClientID ='" & szSelectedClient & "' " & _
                "AND P.PDate <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "#  AND (SP.Type='Supplier' ) "
        rsPayment.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
        If Not rsPayment.EOF Then
                GetSupplierPayment = GetSupplierPayment + IIf(IsNull(rsPayment.Fields.Item("AMT").Value), 0, rsPayment.Fields.Item("AMT").Value)
                'result is -15
        End If
        rsPayment.Close
        Set rsPayment = Nothing
         GetSupplierPayment = GetSupplierPayment - dblAmt
    adoConn.Close
    Set adoConn = Nothing
End Function
Private Function GetBankPaymentReceipts(ByVal trxToinc As Long) As Double
    Dim rsPayment As New ADODB.Recordset
    Dim szSQL As String
    Dim adoConn As New ADODB.Connection
    'F.CategoryCode = 1 Fund category 1 Means rent
    'Implement switch here in SQL
    adoConn.Open getConnectionString

    szSQL = "Select  SUM(SWITCH(TransactionType=12,B.NET_AMOUNT,TransactionType=11,-B.NET_AMOUNT)) as DR from tlbBankPayment B,Fund F where TransactionType IN(11,12) AND " & _
            "B.TRAN_DATE <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "# AND " & _
            "B.DEPT_ID=F.FundID  AND B.RentSumStatement='" & trxToinc & "' AND B.ClientID ='" & szSelectedClient & "'"

    rsPayment.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
    If Not rsPayment.EOF Then
        GetBankPaymentReceipts = IIf(IsNull(rsPayment.Fields.Item("DR").Value), 0, rsPayment.Fields.Item("DR").Value)
    End If
    rsPayment.Close
    adoConn.Close
    Set adoConn = Nothing
End Function
Private Function GetBankPaymentReceiptsPreview(ByVal trxToinc As Long) As Double
    Dim rsPayment As New ADODB.Recordset
    Dim szSQL As String
    Dim adoConn As New ADODB.Connection
    'F.CategoryCode = 1 Fund category 1 Means rent
    'Implement switch here in SQL
    adoConn.Open getConnectionString

    szSQL = "Select  SUM(SWITCH(TransactionType=12,B.NET_AMOUNT,TransactionType=11,-B.NET_AMOUNT)) as DR from tlbBankPayment B,Fund F where TransactionType IN(11,12) AND " & _
            "B.TRAN_DATE <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "# AND " & _
            "B.DEPT_ID=F.FundID  AND B.RentSumStatementPreview='" & trxToinc & "' AND B.ClientID ='" & szSelectedClient & "'"

    rsPayment.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
    If Not rsPayment.EOF Then
        GetBankPaymentReceiptsPreview = IIf(IsNull(rsPayment.Fields.Item("DR").Value), 0, rsPayment.Fields.Item("DR").Value)
    End If
    rsPayment.Close
    adoConn.Close
    Set adoConn = Nothing
End Function
Private Function GetClientPayments(ByVal trxToiclude As Long) As Double    'This function return result as minus'This is getting supplier balance
    Dim rsPayment As New ADODB.Recordset
    Dim szSQL As String
    Dim adoConn As New ADODB.Connection
    'F.CategoryCode = 1 Fund category 1 Means rent
    'Implement switch here in SQL
    'Bank code does not exits in PI,so do not put it in where clause
    Dim whereProperty As String
    adoConn.Open getConnectionString
'    If boolConsolidatedStatement = 1 Then
'            whereProperty = "(P.UNITID IN (" & ListOfProperties & ") OR isnull(P.UNITID)  OR P.UNITID='' ) AND "
'    Else
'            whereProperty = "P.UNITID in (" & ListOfProperties & ") AND "
'    End If
'    szSQL = "Select  SUM(SWITCH(P.TYPE =24,S.Amount,P.TYPE =6,S.Amount,P.TYPE =7,-S.Amount,P.TYPE =8,-S.Amount,P.TYPE =9,-S.Amount)) as AMT " & _
'            "from tlbPayment P,tlbPaymentSplit S,Supplier SP where " & _
'            " SP.SupplierID=P.SageaccountNumber AND " & whereProperty & "  P.TransactionID=S.PayHeader  AND P.ClientID ='" & szSelectedClient & "'  and SP.TYPE in ('Client') " & _
'               "AND S.ClientStatementID=" & trxToiclude & " AND P.PDate <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "#"
    
     szSQL = "Select   SUM(SWITCH(P.TYPE=24,PS.Amount,P.TYPE=8,-PS.Amount,P.TYPE=9,-PS.Amount)) as AMT from tlbPayment P,tlbPaymentSplit PS,Fund F,Supplier SP," & _
                "tblPurInv V,tlbPayment PI,PayTransactions PT where PT.fromtran=PI.transactionID and P.transactionID=PT.Totran and " & _
                "PS.Payheader=P.TransactionID AND SP.SupplierID=P.SageAccountNumber AND PS.Amount>PS.OSAmount AND PS.ClientStatementID=" & trxToiclude & " AND " & _
                "PS.FundID=F.FundID and Sp.Type='Client' AND P.ClientID ='" & szSelectedClient & "' AND V.MY_ID=PI.PI AND isRentPayable=false " & _
                 "AND P.PDate <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "#"
                 
    rsPayment.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
    If Not rsPayment.EOF Then
        GetClientPayments = IIf(IsNull(rsPayment.Fields.Item("AMT").Value), 0, rsPayment.Fields.Item("AMT").Value)
    End If
    rsPayment.Close
    adoConn.Close
    Set adoConn = Nothing
End Function
Private Function GetClientPaymentsPreview(ByVal trxToiclude As Long) As Double    'This function return result as minus'This is getting supplier balance
    Dim rsPayment As New ADODB.Recordset
    Dim szSQL As String
    Dim adoConn As New ADODB.Connection
    'F.CategoryCode = 1 Fund category 1 Means rent
    'Implement switch here in SQL
    'Bank code does not exits in PI,so do not put it in where clause
    Dim whereProperty As String
    adoConn.Open getConnectionString
'    If boolConsolidatedStatement = 1 Then
'            whereProperty = "(P.UNITID IN (" & ListOfProperties & ") OR isnull(P.UNITID)  OR P.UNITID='' ) AND "
'    Else
'            whereProperty = "P.UNITID in (" & ListOfProperties & ") AND "
'    End If
'    szSQL = "Select  SUM(SWITCH(P.TYPE =24,S.Amount,P.TYPE =6,S.Amount,P.TYPE =7,-S.Amount,P.TYPE =8,-S.Amount,P.TYPE =9,-S.Amount)) as AMT " & _
'            "from tlbPayment P,tlbPaymentSplit S,Supplier SP where " & _
'            " SP.SupplierID=P.SageaccountNumber AND " & whereProperty & "  P.TransactionID=S.PayHeader AND P.ClientID ='" & szSelectedClient & "'  and SP.TYPE in ('Client') " & _
'               "AND S.ClientStatementPrevID=" & trxToiclude & " AND P.PDate <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "#"
     szSQL = "Select   SUM(SWITCH(P.TYPE=24,PS.Amount,P.TYPE=8,-PS.Amount,P.TYPE=9,-PS.Amount)) as AMT from tlbPayment P,tlbPaymentSplit PS,Fund F,Supplier SP," & _
                "tblPurInv V,tlbPayment PI,PayTransactions PT where PT.fromtran=PI.transactionID and P.transactionID=PT.Totran and " & _
                "PS.Payheader=P.TransactionID AND SP.SupplierID=P.SageAccountNumber AND PS.Amount>PS.OSAmount AND PS.ClientStatementPrevID=" & trxToiclude & " AND " & _
                "PS.FundID=F.FundID and Sp.Type='Client' AND P.ClientID ='" & szSelectedClient & "' AND V.MY_ID=PI.PI AND isRentPayable=false " & _
                 "AND P.PDate <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "#"
    
    rsPayment.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
    If Not rsPayment.EOF Then
        GetClientPaymentsPreview = IIf(IsNull(rsPayment.Fields.Item("AMT").Value), 0, rsPayment.Fields.Item("AMT").Value)
    End If
    rsPayment.Close
    adoConn.Close
    Set adoConn = Nothing
End Function
Private Function GetLandLordPayments(ByVal trxToinclude As Long) As Double    'This function return result as minus'This is getting supplier balance
    Dim rsPayment As New ADODB.Recordset
    Dim szSQL As String
    Dim adoConn As New ADODB.Connection
    'F.CategoryCode = 1 Fund category 1 Means rent
    'Implement switch here in SQL
    'Bank code does not exits in PI,so do not put it in where clause
    Dim whereProperty As String
    adoConn.Open getConnectionString
'      If boolConsolidatedStatement = 1 Then
'            whereProperty = "(P.UNITID IN (" & ListOfProperties & ") OR isnull(P.UNITID)  OR P.UNITID='' ) AND "
'    Else
'            whereProperty = "P.UNITID in (" & ListOfProperties & ") AND "
'    End If
    szSQL = "Select  SUM(SWITCH(P.TYPE =24,S.Amount,P.TYPE =6,S.Amount,P.TYPE =7,-S.Amount,P.TYPE =8,-S.Amount,P.TYPE =9,-S.Amount)) as AMT from " & _
            "tlbPayment P,tlbPaymentSplit S,Supplier SP where " & _
            " SP.SupplierID=P.SageaccountNumber AND " & whereProperty & "   P.TransactionID=S.PayHeader AND P.TYPE AND S.ClientStatementID=" & trxToinclude & " " & _
            "IN(7,8,9) AND P.ClientID ='" & szSelectedClient & "'  and SP.TYPE in ('LLORD')" & _
            "AND P.PDate <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "#"

    rsPayment.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
    If Not rsPayment.EOF Then
        GetLandLordPayments = IIf(IsNull(rsPayment.Fields.Item("AMT").Value), 0, rsPayment.Fields.Item("AMT").Value)
    End If
    rsPayment.Close
    adoConn.Close
    Set adoConn = Nothing
End Function
Private Function GetLandLordPaymentsPreview(ByVal trxToinclude As Long) As Double    'This function return result as minus'This is getting supplier balance
    Dim rsPayment As New ADODB.Recordset
    Dim szSQL As String
    Dim adoConn As New ADODB.Connection
    'F.CategoryCode = 1 Fund category 1 Means rent
    'Implement switch here in SQL
    'Bank code does not exits in PI,so do not put it in where clause
    Dim whereProperty As String
    adoConn.Open getConnectionString
'      If boolConsolidatedStatement = 1 Then
'            whereProperty = "(P.UNITID IN (" & ListOfProperties & ") OR isnull(P.UNITID)  OR P.UNITID='' ) AND "
'    Else
'            whereProperty = "P.UNITID in (" & ListOfProperties & ") AND "
'    End If
    szSQL = "Select  SUM(SWITCH(P.TYPE =24,S.Amount,P.TYPE =6,S.Amount,P.TYPE =7,-S.Amount,P.TYPE =8,-S.Amount,P.TYPE =9,-S.Amount)) as AMT from " & _
            "tlbPayment P,tlbPaymentSplit S,Supplier SP where " & _
            " SP.SupplierID=P.SageaccountNumber AND " & whereProperty & "   P.TransactionID=S.PayHeader AND P.TYPE AND S.ClientStatementPrevID=" & trxToinclude & " " & _
            "IN(7,8,9)  AND P.ClientID ='" & szSelectedClient & "'  and SP.TYPE in ('LLORD')" & _
            "AND P.PDate <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "#"

    rsPayment.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
    If Not rsPayment.EOF Then
        GetLandLordPaymentsPreview = IIf(IsNull(rsPayment.Fields.Item("AMT").Value), 0, rsPayment.Fields.Item("AMT").Value)
    End If
    rsPayment.Close
    adoConn.Close
    Set adoConn = Nothing
End Function
Private Function GetAGENTFinalized(ByVal trtoinclude As Long) As Double    'This function return result as minus'This is getting supplier balance
    Dim rsPayment As New ADODB.Recordset
    Dim szSQL As String
    Dim adoConn As New ADODB.Connection
    'F.CategoryCode = 1 Fund category 1 Means rent
    'Implement switch here in SQL
    'Bank code does not exits in PI,so do not put it in where clause
    Dim whereProperty As String
    adoConn.Open getConnectionString
'      If boolConsolidatedStatement = 1 Then
            whereProperty = "(P.UNITID IN (" & ListOfProperties & ") OR isnull(P.UNITID)  OR P.UNITID='' ) AND "
'    Else
'            whereProperty = "P.UNITID in (" & ListOfProperties & ") AND "
'    End If

             
    szSQL = "Select   SUM(PS.PaymentAmount) as AMT from tlbPayment P,Paytransactions PS,Fund F,Supplier SP where  " & _
            "PS.FROMTran=P.TransactionID AND SP.SupplierID=P.SageAccountNumber  AND PS.Deleteflag=False AND  P.RentSumStatement='" & trtoinclude & "' AND " & _
            "PS.FundID=F.FundID and P.type=24 and " & whereProperty & " P.BankCODE='" & szSelectedBankAccount & "' and Sp.Type='AGENT' and F.FundCode in (" & _
            ListOfFunds & ") AND P.ClientID ='" & szSelectedClient & "' " & _
            "AND P P.PDate <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "#"

rsPayment.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
    If Not rsPayment.EOF Then
        GetAGENTFinalized = IIf(IsNull(rsPayment.Fields.Item("AMT").Value), 0, rsPayment.Fields.Item("AMT").Value)
    End If
    rsPayment.Close
    whereProperty = "(S.PropertyID IN (" & ListOfProperties & ") OR isnull(S.PropertyID)  OR S.PropertyID='' ) AND "

'Otherwise you shall get duplicated value
         szSQL = "Select   SUM(PS.PaymentAmount) as AMT from tlbPayment P,tlbPaymentSplit S,PaytransactionsSplit PS,Fund F,Supplier SP where  " & _
            "PS.TransactionID=S.PayTransactionIDSplit and PS.FromTran=P.TransactionID AND SP.SupplierID=P.SageAccountNumber  AND PS.Deleteflag=False " & _
            "AND  P.RentSumStatement='" & trtoinclude & "' AND " & _
            " P.TransactionID=S.PayHeader AND PS.FundID=F.FundID and P.type IN(8,9) and " & whereProperty & " P.BankCODE='" & szSelectedBankAccount & "' " & _
            "and Sp.Type='AGENT'  AND P.ClientID ='" & szSelectedClient & "'  AND AND P.PDate <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "#"

    
    rsPayment.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
    If Not rsPayment.EOF Then
        GetAGENTFinalized = GetAGENTFinalized - IIf(IsNull(rsPayment.Fields.Item("AMT").Value), 0, rsPayment.Fields.Item("AMT").Value)
    End If
    rsPayment.Close
    adoConn.Close
    Set adoConn = Nothing
End Function
Private Function GetBankPaymentPreview(ByVal idToinclude As Long) As Double
    Dim rsPayment As New ADODB.Recordset
    Dim szSQL As String
    Dim adoConn As New ADODB.Connection
    'F.CategoryCode = 1 Fund category 1 Means rent
    'Implement switch here in SQL
    adoConn.Open getConnectionString
    Dim whereProperty As String
    If boolConsolidatedStatement = 1 Then
            whereProperty = "(B.PropertyID in (" & ListOfProperties & ") OR isnull(B.PropertyID) OR B.PropertyID='' ) "
    Else
            whereProperty = "B.PropertyID in (" & ListOfProperties & ") "
    End If
    szSQL = "Select  SUM(SWITCH(TransactionType=12,B.NET_AMOUNT,TransactionType=11,-B.NET_AMOUNT)) as DR from tlbBankPayment B,Fund F where TransactionType IN (11) AND " & _
            "B.TRAN_DATE <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "# AND " & _
            "B.DEPT_ID=F.FundID  AND  B.RentSumStatementPreview='" & idToinclude & "' AND B.ClientID ='" & _
            szSelectedClient & "' AND " & whereProperty
    
    rsPayment.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
    If Not rsPayment.EOF Then
        GetBankPaymentPreview = IIf(IsNull(rsPayment.Fields.Item("DR").Value), 0, rsPayment.Fields.Item("DR").Value)
    End If
    rsPayment.Close
    adoConn.Close
    Set adoConn = Nothing
End Function
Private Function GetBankPayment(ByVal idToinclude As Long) As Double
    Dim rsPayment As New ADODB.Recordset
    Dim szSQL As String
    Dim adoConn As New ADODB.Connection
    'F.CategoryCode = 1 Fund category 1 Means rent
    'Implement switch here in SQL
    adoConn.Open getConnectionString
    Dim whereProperty As String
    If boolConsolidatedStatement = 1 Then
            whereProperty = "(B.PropertyID in (" & ListOfProperties & ") OR isnull(B.PropertyID) OR B.PropertyID='' ) "
    Else
            whereProperty = "B.PropertyID in (" & ListOfProperties & ") "
    End If
    szSQL = "Select  SUM(SWITCH(TransactionType=12,B.NET_AMOUNT,TransactionType=11,-B.NET_AMOUNT)) as DR from tlbBankPayment B,Fund F where TransactionType IN (11) AND " & _
            "B.TRAN_DATE <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "# AND " & _
            "B.DEPT_ID=F.FundID  AND  B.RentSumStatement='" & idToinclude & "' AND B.ClientID ='" & _
            szSelectedClient & "' AND " & whereProperty
    
    rsPayment.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
    If Not rsPayment.EOF Then
        GetBankPayment = IIf(IsNull(rsPayment.Fields.Item("DR").Value), 0, rsPayment.Fields.Item("DR").Value)
    End If
    rsPayment.Close
    adoConn.Close
    Set adoConn = Nothing
End Function
Private Function GetBankreceiptsPreview(ByVal idToinclude As Long) As Double
    Dim rsPayment As New ADODB.Recordset
    Dim szSQL As String
    Dim adoConn As New ADODB.Connection
    'F.CategoryCode = 1 Fund category 1 Means rent
    'Implement switch here in SQL
    adoConn.Open getConnectionString
    Dim whereProperty As String
'    If boolConsolidatedStatement = 1 Then
            whereProperty = "(B.PropertyID in (" & ListOfProperties & ") OR isnull(B.PropertyID) OR B.PropertyID='' ) "
'    Else
'            whereProperty = "B.PropertyID in (" & ListOfProperties & ") "
'    End If
    szSQL = "Select  SUM(SWITCH(TransactionType=12,B.NET_AMOUNT,TransactionType=11,-B.NET_AMOUNT)) as DR from tlbBankPayment B,Fund F where TransactionType IN(12) AND " & _
            "B.TRAN_DATE <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "# AND " & _
            "B.DEPT_ID=F.FundID  AND B.RentSumStatementPreview='" & idToinclude & "'  AND B.ClientID ='" & szSelectedClient & "' AND " & whereProperty
    
    rsPayment.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
    If Not rsPayment.EOF Then
        GetBankreceiptsPreview = IIf(IsNull(rsPayment.Fields.Item("DR").Value), 0, rsPayment.Fields.Item("DR").Value)
    End If
    rsPayment.Close
    adoConn.Close
    Set adoConn = Nothing
End Function
Private Function GetBankreceipts(ByVal idToinclude As Long) As Double
    Dim rsPayment As New ADODB.Recordset
    Dim szSQL As String
    Dim adoConn As New ADODB.Connection
    'F.CategoryCode = 1 Fund category 1 Means rent
    'Implement switch here in SQL
    adoConn.Open getConnectionString
    Dim whereProperty As String
'    If boolConsolidatedStatement = 1 Then
            whereProperty = "(B.PropertyID in (" & ListOfProperties & ") OR isnull(B.PropertyID) OR B.PropertyID='' ) "
'    Else
'            whereProperty = "B.PropertyID in (" & ListOfProperties & ") "
'    End If
    szSQL = "Select  SUM(SWITCH(TransactionType=12,B.NET_AMOUNT,TransactionType=11,-B.NET_AMOUNT)) as DR from tlbBankPayment B,Fund F where TransactionType IN(12) AND " & _
            "B.TRAN_DATE <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "# AND " & _
            "B.DEPT_ID=F.FundID  AND B.RentSumStatement='" & idToinclude & "'  AND B.ClientID ='" & szSelectedClient & "' AND " & whereProperty
    
    rsPayment.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
    If Not rsPayment.EOF Then
        GetBankreceipts = IIf(IsNull(rsPayment.Fields.Item("DR").Value), 0, rsPayment.Fields.Item("DR").Value)
    End If
    rsPayment.Close
    adoConn.Close
    Set adoConn = Nothing
End Function
Private Function getClosingBalanceFinalized(dblLasClosingBalance As Double, ByVal trxToinclude As Long) As Double
    'Pass propery as parameter for selected property
    'No property spec:
    Dim adoConn As New ADODB.Connection
    Dim rsReceipt As New ADODB.Recordset
    Dim rsPayment As New ADODB.Recordset
    Dim rsBankPaymentAndRcpt As New ADODB.Recordset
    Dim dblAmt As Double
    adoConn.Open getConnectionString
    Dim szSQL As String
    Dim AccrualsControl As String
    'we are not using property filter here
    'B )***********************  Sum of Rent received Paid/Refunded ***********************************


    szSQL = "Select  SUM(switch(type=3,S.Amount,type=4,S.Amount,type=23,-S.Amount)) as DR from tlbReceipt R,tlbReceiptSplit S,Fund F where R.TransactionID= S.RptHeader AND TYPE IN(3,4,23) " & _
    "AND S.FundID=F.FundID AND R.BankCODE='" & szSelectedBankAccount & "' and F.FundCode<>'TENANTDEPOSIT'  AND  R.RentSumStatement='" & trxToinclude & " ' AND ClientID ='" & szSelectedClient & "' " & _
    "AND S.Amount>S.OSAmount AND R.RDate <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "#"
    rsReceipt.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
    If Not rsReceipt.EOF Then
            dblAmt = IIf(IsNull(rsReceipt.Fields.Item("Dr").Value), 0, rsReceipt.Fields.Item("Dr").Value)
             'result is 255
    End If
    rsReceipt.Close
    Set rsReceipt = Nothing
    getClosingBalanceFinalized = dblLasClosingBalance + dblAmt
 
   'c   (-): Sum of Supplier amounts Paid/Refunded (Both allocated and unallocated)
 
    szSQL = "Select  SUM(SWITCH(P.TYPE=24,S.Amount,P.TYPE=8,-S.Amount,P.TYPE=9,-S.Amount)) as AMT from tlbPayment P,tlbPaymentSplit S,Fund F,Supplier SP where  " & _
            "SP.SupplierID=P.SageAccountNumber AND P.TransactionID=S.PayHeader AND P.TYPE IN(8,9,24) AND  P.RentSumStatement='" & trxToinclude & "' AND " & _
            "P.PDate <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "# AND " & _
            "S.FundID=F.FundID and  P.BankCODE='" & szSelectedBankAccount & "' AND ClientID ='" & szSelectedClient & "' "
    rsPayment.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
    If Not rsPayment.EOF Then
            dblAmt = IIf(IsNull(rsPayment.Fields.Item("AMT").Value), 0, rsPayment.Fields.Item("AMT").Value)
        'result is -50
    End If
    rsPayment.Close
    Set rsPayment = Nothing
    getClosingBalanceFinalized = getClosingBalanceFinalized + dblAmt

    Dim whereProperty As String

        
        'd)  Add (+): Sum of Bank payments and receipts
        
        
     szSQL = "Select  SUM(SWITCH(TransactionType=11 ,-B.NET_AMOUNT,TransactionType=12 ,B.NET_AMOUNT)) as AMT from tlbBankPayment B, Fund F  where B.DEPT_ID=F.FundID " & _
            "and BANK_AC='" & szSelectedBankAccount & "' AND  B.RentSumStatement='" & trxToinclude & "' and clientID='" & szSelectedClient & "' " & _
            "AND B.TRAN_DATE <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "#"
    rsBankPaymentAndRcpt.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
    If Not rsBankPaymentAndRcpt.EOF Then
        dblAmt = IIf(IsNull(rsBankPaymentAndRcpt.Fields.Item("AMT").Value), 0, rsBankPaymentAndRcpt.Fields.Item("AMT").Value)
           'result is 0
    End If
     getClosingBalanceFinalized = getClosingBalanceFinalized + dblAmt
    rsBankPaymentAndRcpt.Close
    Set rsBankPaymentAndRcpt = Nothing
 End Function
Private Function getAvailablefunds(dblLasClosingBalance As Double) As Double
    'Pass propery as parameter for selected property
    'No property spec:
    'Exit Function
    Dim adoConn As New ADODB.Connection
    Dim rsReceipt As New ADODB.Recordset
    Dim rsPayment As New ADODB.Recordset
    Dim rsBankPaymentAndRcpt As New ADODB.Recordset
    Dim dblAmt As Double
    adoConn.Open getConnectionString
    Dim szSQL As String
    Dim AccrualsControl As String
    Dim whereProperty As String
    'we are not using property filter here
    'B )***********************  Sum of Rent received Paid/Refunded ***********************************
    'AND tlbReceipt.RDate >=#" & Format(txtLastStatementDate1.text, "dd/mmm/yyyy") & "# AND tlbReceipt.RDate <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "#"
    Dim rsGlobalData1 As New ADODB.Recordset
    Dim bolVatOptionEnabled As Boolean
    Dim bolOptedTotax As String
    Dim bolisAgentToSubmit As Boolean
    Dim strManagingAgentID As String
    
    rsGlobalData1.Open "Select vatOptionEnabled,isAgentToSubmit from Globaldata G,Property P where P.PropertyID=G.PropertyID AND P.PropertyID='" & _
                       szSQL & "' ", adoConn, adOpenStatic, adLockReadOnly
                       
    If Not rsGlobalData1.EOF Then
           bolVatOptionEnabled = rsGlobalData1("vatOptionEnabled").Value
           bolisAgentToSubmit = rsGlobalData1("isAgentToSubmit").Value
    End If
    rsGlobalData1.Close

                
    If boolConsolidatedStatement = 1 Then
            whereProperty = "(U.PropertyID in (" & ListOfProperties & ")OR isnull(U.PropertyID) OR U.PropertyID='' ) AND "
    Else
            whereProperty = "U.PropertyID in (" & ListOfProperties & ") AND "
    End If
    'Exit Function
    'sum of the receipt
'    szSQL = "Select  SUM(switch(type=3,S.Amount,type=4,S.Amount,type=23,-S.Amount)) as DR from tlbReceipt R,tlbReceiptSplit S,Fund F, Units U where R.TransactionID= S.RptHeader " & _
'    "AND TYPE IN(3,4,23) AND S.FundID=F.FundID AND R.BankCODE='" & szSelectedBankAccount & "' and F.FundCode<>'TENANTDEPOSIT'  AND (R.RentSumStatement='' or isnull(R.RentSumStatement)) " & _
'    "AND ClientID ='" & szSelectedClient & "' AND R.UnitID=U.UnitNumber and  F.FundCode in (" & ListOfFunds & ") AND " & whereProperty & _
'    "(R.RDate >=#" & Format(txtLastStatementDate1.text, "dd/mmm/yyyy") & "# AND R.RDate <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "# ) OR " & _
'    " (R.RDate < #" & Format(txtLastStatementDate1.text, "dd/mmm/yyyy") & "# AND (RentSumStatement='' OR isnull(RentSumStatement))"
    
      szSQL = "Select  SUM(switch(type=3,S.Amount,type=4,S.Amount,type=23,-S.Amount)) as DR from tlbReceipt R,tlbReceiptSplit S,Fund F, Units U where R.TransactionID= S.RptHeader " & _
    "AND TYPE IN(3,4,23) AND S.FundID=F.FundID AND R.BankCODE='" & szSelectedBankAccount & "' and F.FundCode<>'TENANTDEPOSIT'  AND (R.RentSumStatement='' or isnull(R.RentSumStatement)) " & _
    "AND ClientID ='" & szSelectedClient & "' and S.Amount>S.OSAmount AND R.UnitID=U.UnitNumber and  F.FundCode in (" & ListOfFunds & ") AND " & whereProperty & _
    "R.RDate >#" & Format(txtLastStatementDate1.text, "dd/mmm/yyyy") & "# AND R.RDate <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "#  "
    rsReceipt.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
    If Not rsReceipt.EOF Then
            dblAmt = IIf(IsNull(rsReceipt.Fields.Item("Dr").Value), 0, rsReceipt.Fields.Item("Dr").Value)
             'result is 175
    End If
    rsReceipt.Close
    Set rsReceipt = Nothing
    getAvailablefunds = dblLasClosingBalance + dblAmt
    'Vat calculation of B)
'    If bolisAgentToSubmit = True Then' we have alreay done this if in SQL where G.isAgentToSubmit=true
            'take the  VAT amount from the allocation table and deduct it
            szSQL = "Select  SUM(AL.VatAmount) as DR from tlbReceipt R,rptTransactionsSplit AL,tlbReceiptSplit S,Fund F, Units U,GLobalData G where G.PropertyID=U.PropertyID " & _
            "AND G.isAgentToSubmit=true AND R.TransactionID= S.RptHeader AND TYPE IN(3,4,23) AND AL.Deleteflag=False AND S.FundID=F.FundID AND R.BankCODE='" & szSelectedBankAccount & "' " & _
            "and F.FundCode<>'TENANTDEPOSIT'  AND (R.RentSumStatement='' or isnull(R.RentSumStatement)) AND ClientID ='" & szSelectedClient & "' AND R.UnitID=U.UnitNumber " & _
            "and AL.Deleteflag=false and AL.FromTran=R.TransactionID AND  F.FundCode in (" & ListOfFunds & ") AND " & whereProperty & " R.RDate >#" & _
            Format(txtLastStatementDate1.text, "dd/mmm/yyyy") & "# AND R.RDate <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "#"
            rsReceipt.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
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
            whereProperty = "(P.UnitID in (" & ListOfProperties & ") OR isnull(P.UnitID) OR P.UnitID='') AND "
    Else
            whereProperty = "P.UnitID in (" & ListOfProperties & ") AND "
    End If

'supplier payment
  szSQL = "Select   SUM(SWITCH(P.TYPE=24,PS.PaymentAmount,P.TYPE=8,-PS.PaymentAmount,P.TYPE=9,-PS.PaymentAmount)) as AMT from tlbPayment P,Paytransactions PS,Fund F,Supplier SP where  " & _
            "PS.FROMTran=P.TransactionID AND SP.SupplierID=P.SageAccountNumber  AND PS.Deleteflag=False AND (P.RentSumStatement='' or isnull(P.RentSumStatement)) AND " & _
            "PS.FundID=F.FundID and " & whereProperty & " P.BankCODE='" & szSelectedBankAccount & "' and Sp.Type='Supplier' and F.FundCode in (" & ListOfFunds & ") AND P.ClientID ='" & szSelectedClient & "' " & _
            "AND P.PDate >#" & Format(txtLastStatementDate1.text, "dd/mmm/yyyy") & "# AND P.PDate <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "#"

    rsPayment.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
    If Not rsPayment.EOF Then
            dblAmt = IIf(IsNull(rsPayment.Fields.Item("AMT").Value), 0, rsPayment.Fields.Item("AMT").Value)
        'result is -837
    End If
    rsPayment.Close
    Set rsPayment = Nothing
    getAvailablefunds = getAvailablefunds + dblAmt
'Take the vat amount from the allocation table
'    If bolisAgentToSubmit = True Then
        szSQL = "Select  SUM(AL.VatAmount)  as AMT from tlbPayment P,PayTransactionsSplit AL,tlbPaymentSplit S,Fund F,Supplier SP,Property PR,GLobalData G  where  G.PropertyID=PR.PropertyID AND isAgentToSubmit=true  AND " & _
                "AL.TransactionID=S.PayTransactionIDSplit and SP.SupplierID=P.SageAccountNumber AND AL.Deleteflag=false and PR.propertyID=p.UNITID and Sp.Type='Supplier' AND P.TransactionID=S.PayHeader AND P.TYPE IN(8,9,24) AND (P.RentSumStatement='' or isnull(P.RentSumStatement)) AND " & _
                "S.FundID=F.FundID AND AL.FromTran=P.transactionID and " & whereProperty & " P.BankCODE='" & szSelectedBankAccount & "' and  F.FundCode in (" & ListOfFunds & ") AND P.ClientID ='" & szSelectedClient & "' " & _
                "AND P.PDate >#" & Format(txtLastStatementDate1.text, "dd/mmm/yyyy") & "# AND P.PDate <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "#"
        rsPayment.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
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
    
     szSQL = "Select  SUM(SWITCH(TransactionType=11 ,-B.NET_AMOUNT,TransactionType=12 ,B.NET_AMOUNT)) as AMT from tlbBankPayment B, Fund F  where B.DEPT_ID=F.FundID " & _
            "and " & whereProperty & " BANK_AC='" & szSelectedBankAccount & "' AND F.FundCode in (" & ListOfFunds & ") AND (B.RentSumStatement='' OR isnull(B.RentSumStatement)) and clientID='" & szSelectedClient & "' " & _
            "AND B.TRAN_DATE >#" & Format(txtLastStatementDate1.text, "dd/mmm/yyyy") & "# AND B.TRAN_DATE <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "#"
    rsBankPaymentAndRcpt.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
    If Not rsBankPaymentAndRcpt.EOF Then
        dblAmt = IIf(IsNull(rsBankPaymentAndRcpt.Fields.Item("AMT").Value), 0, rsBankPaymentAndRcpt.Fields.Item("AMT").Value)
           'result is 0
    End If
    rsBankPaymentAndRcpt.Close
    Set rsBankPaymentAndRcpt = Nothing
    getAvailablefunds = getAvailablefunds + dblAmt
    'f)  Less (-): Supplier OS Account balances for the client selected
    '*********************23/12/2021*************we removed this temporarily .
    dblAmt = GetSupplierOSAmount
    'If negative then ignore this
    getAvailablefunds = getAvailablefunds - IIf(dblAmt < 0, 0, dblAmt)
    'it should be -40

    Dim rsNLposting As New ADODB.Recordset

'g)  Less (-): Client /Landlord OS balances for the client selected  and property selected amounts due to Client/Landlord not paid
         dblAmt = GetClientACBalance
         'COMING -35  dblAmt is negative then ignore
          getAvailablefunds = getAvailablefunds + IIf(dblAmt < 0, 0, dblAmt)
          
          'client payment
          whereProperty = "(P.UnitID in (" & ListOfProperties & ") OR isnull(P.UnitID) OR P.UnitID='' ) AND "
         szSQL = "Select   SUM(SWITCH(P.TYPE=24,PS.PaymentAmount,P.TYPE=8,-PS.PaymentAmount,P.TYPE=9,-PS.PaymentAmount)) as AMT from tlbPayment P,Paytransactions PS,Fund F,Supplier SP where  " & _
            "PS.FROMTran=P.TransactionID AND SP.SupplierID=P.SageAccountNumber  AND PS.Deleteflag=False AND (P.RentSumStatement=''  or isnull(P.RentSumStatement)) AND " & _
            "PS.FundID=F.FundID and " & whereProperty & " P.BankCODE='" & szSelectedBankAccount & "' and Sp.Type='Client'  and F.FundCode in (" & ListOfFunds & ") AND P.ClientID ='" & szSelectedClient & "' " & _
            "AND P.PDate >#" & Format(txtLastStatementDate1.text, "dd/mmm/yyyy") & "# AND P.PDate <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "#"

    rsPayment.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
    If Not rsPayment.EOF Then
            dblAmt = IIf(IsNull(rsPayment.Fields.Item("AMT").Value), 0, rsPayment.Fields.Item("AMT").Value)
        'result is -837
    End If
    rsPayment.Close
    Set rsPayment = Nothing
     getAvailablefunds = getAvailablefunds + dblAmt

    'agent payment
          whereProperty = "(P.UnitID in (" & ListOfProperties & ") OR isnull(P.UnitID) OR P.UnitID='' ) AND "
         szSQL = "Select   SUM(SWITCH(P.TYPE=24,PS.PaymentAmount,P.TYPE=8,-PS.PaymentAmount,P.TYPE=9,-PS.PaymentAmount)) as AMT from tlbPayment P,Paytransactions PS,Fund F,Supplier SP where  " & _
            "PS.FROMTran=P.TransactionID AND SP.SupplierID=P.SageAccountNumber  AND PS.Deleteflag=False AND (P.RentSumStatement='' Or isnull(P.RentSumStatement)) AND " & _
            "PS.FundID=F.FundID and " & whereProperty & " P.BankCODE='" & szSelectedBankAccount & "' and Sp.Type='AGENT'  and F.FundCode in (" & ListOfFunds & ") AND P.ClientID ='" & szSelectedClient & "' " & _
            "AND P.PDate >#" & Format(txtLastStatementDate1.text, "dd/mmm/yyyy") & "# AND P.PDate <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "#"

    rsPayment.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
    If Not rsPayment.EOF Then
            dblAmt = IIf(IsNull(rsPayment.Fields.Item("AMT").Value), 0, rsPayment.Fields.Item("AMT").Value)
        'result is -837
    End If
    rsPayment.Close
    Set rsPayment = Nothing
        getAvailablefunds = getAvailablefunds + dblAmt
        
        'landlord payment
          whereProperty = "(P.UnitID in (" & ListOfProperties & ") OR isnull(P.UnitID) OR P.UnitID='') AND "
                szSQL = "Select   SUM(SWITCH(P.TYPE=24,PS.PaymentAmount,P.TYPE=8,-PS.PaymentAmount,P.TYPE=9,-PS.PaymentAmount)) as AMT from tlbPayment P,Paytransactions PS,Fund F,Supplier SP where  " & _
                "PS.FROMTran=P.TransactionID AND SP.SupplierID=P.SageAccountNumber  AND PS.Deleteflag=False AND (P.RentSumStatement='' or isnull(P.RentSumStatement)) AND " & _
                "PS.FundID=F.FundID and " & whereProperty & " P.BankCODE='" & szSelectedBankAccount & "' and Sp.Type='LLORD' and F.FundCode in (" & ListOfFunds & ") AND P.ClientID ='" & szSelectedClient & "' " & _
                 "AND P.PDate >#" & Format(txtLastStatementDate1.text, "dd/mmm/yyyy") & "# AND P.PDate <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "#"

                rsPayment.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
                If Not rsPayment.EOF Then
                        dblAmt = IIf(IsNull(rsPayment.Fields.Item("AMT").Value), 0, rsPayment.Fields.Item("AMT").Value)
                    'result is -837
                End If
                rsPayment.Close
                Set rsPayment = Nothing
      getAvailablefunds = getAvailablefunds + dblAmt
      
      

        dblAmt = -GetLandLordACBalance '-GetLandLordACBalanceMODpreview(trToinclude) '-GetLandLordACBalance
        'GetLandLordACBalanceMODpreview
          'if dblAmt is negative then ignore
        getAvailablefunds = getAvailablefunds + dblAmt
     
    'h)  Less (-): Managing Agent OS Balances for the client selected Management Fees due but not paid
    'GetAgentBalanceModPreview
    dblAmt = -GetAgentBalance ' -GetAgentBalanceModPreview(trToinclude) 'GetAgentBalance '
    getAvailablefunds = getAvailablefunds + dblAmt
    '
'    dblAmt = GetAGENTPaymentsPreview(trtoinclude)
'    getAvailablefunds = getAvailablefunds + dblAmt
    getAvailablefunds = Round(getAvailablefunds, 2)
    
    Debug.Print getAvailablefunds
    
    
    Dim ManagementFeeControl As String
    AccrualsControl = GetNominalCodeForControlAccount(adoConn, "Accruals Control Account (B/S)", szSelectedClient)
    'Dim rsNLPosting As New ADODB.Recordset
     If boolConsolidatedStatement = 1 Then
            whereProperty = "(PROPERTY_ID IN (" & ListOfProperties & ") OR isnull(PROPERTY_ID)) AND "
    Else
            whereProperty = "PROPERTY_ID in (" & ListOfProperties & ") AND "
    End If
    
    rsNLposting.Open "Select sum(AMOUNT) as AMT from NLPosting where " & whereProperty & _
                    "  NOMINAL_CODE='" & AccrualsControl & "' AND ClientID='" & _
                    szSelectedClient & "' AND DeleteFlag=false", adoConn, adOpenStatic, adLockReadOnly
    
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
    MsgBox "Available fund is: " & Round(getAvailablefunds, 2)
    txtAvailableFunds.text = Round(getAvailablefunds, 2)
    If bEditMode = False Then
        txtRentPayable.text = txtAvailableFunds.text
    End If
    
    
''    'getAvailablefunds = getAvailablefunds + GetClientACBalance + GetLandLordACBalance
''        getAvailablefunds = getAvailablefunds + IIf(dblAmt < 0, 0, dblAmt)
''        dblAmt = GetLandLordACBalance
''          'if dblAmt is negative then ignore
''        getAvailablefunds = getAvailablefunds + IIf(dblAmt < 0, 0, dblAmt)
''
''    'h)  Less (-): Managing Agent OS Balances for the client selected Management Fees due but not paid
''    dblAmt = GetAgentBalance
''    getAvailablefunds = getAvailablefunds - IIf(dblAmt < 0, 0, dblAmt)
''
''    dblAmt = GetAGENTPayments
''    getAvailablefunds = getAvailablefunds + dblAmt
''
''    Debug.Print getAvailablefunds
''
''
''    Dim ManagementFeeControl As String
''    AccrualsControl = GetNominalCodeForControlAccount(adoConn, "Accruals Control Account (B/S)", szSelectedClient)
''    'Dim rsNLPosting As New ADODB.Recordset
''     If boolConsolidatedStatement = 1 Then
''            whereProperty = "(PROPERTY_ID IN (" & ListOfProperties & ") OR isnull(PROPERTY_ID)) AND "
''    Else
''            whereProperty = "PROPERTY_ID in (" & ListOfProperties & ") AND "
''    End If
''
''    rsNLposting.Open "Select sum(AMOUNT) as AMT from NLPosting where " & whereProperty & _
''                    "  NOMINAL_CODE='" & AccrualsControl & "' AND ClientID='" & _
''                    szSelectedClient & "' AND DeleteFlag=false", adoConn, adOpenStatic, adLockReadOnly
''
''    If Not rsNLposting.EOF Then
''        dblAmt = IIf(IsNull(rsNLposting.Fields.Item("AMT").Value), 0, rsNLposting.Fields.Item("AMT").Value)
''    End If
''    rsNLposting.Close
''    Set rsNLposting = Nothing
''    getAvailablefunds = getAvailablefunds + dblAmt
''    'j)  Less (-): Tenant Deposits received for the client selected
''    'REMOVE THIS AS PER SPEC
''    'getAvailablefunds = getAvailablefunds - GetRentDeposit
''
''    getAvailablefunds = getAvailablefunds - Val(txtRetention.text)
''    MsgBox "Available fund is: " & Round(getAvailablefunds, 2)
''    txtAvailableFunds.text = Round(getAvailablefunds, 2)
''    txtRentPayable.text = txtAvailableFunds.text
     
End Function


Private Function getClosingBalance(dblLasClosingBalance As Double) As Double
    'Pass propery as parameter for selected property
    'No property spec:
    Dim adoConn As New ADODB.Connection
    Dim rsReceipt As New ADODB.Recordset
    Dim rsPayment As New ADODB.Recordset
    Dim rsBankPaymentAndRcpt As New ADODB.Recordset
    Dim dblAmt As Double
    adoConn.Open getConnectionString
    Dim szSQL As String
    Dim AccrualsControl As String
    'we are not using property filter here
    'B )***********************  Sum of Rent received Paid/Refunded ***********************************


    szSQL = "Select  SUM(switch(type=3,S.Amount,type=4,S.Amount,type=23,-S.Amount)) as DR from tlbReceipt R,tlbReceiptSplit S,Fund F where R.TransactionID= S.RptHeader AND TYPE IN(3,4,23) " & _
    "AND S.FundID=F.FundID AND R.BankCODE='" & szSelectedBankAccount & "' and F.FundCode<>'TENANTDEPOSIT'  AND isnull(S.ClientStatementID) AND ClientID ='" & szSelectedClient & "' " & _
    "AND R.RDate <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "#"
    rsReceipt.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
    If Not rsReceipt.EOF Then
            dblAmt = IIf(IsNull(rsReceipt.Fields.Item("Dr").Value), 0, rsReceipt.Fields.Item("Dr").Value)
             'result is 255
    End If
    rsReceipt.Close
    Set rsReceipt = Nothing
    getClosingBalance = dblLasClosingBalance + dblAmt
 
   'c   (-): Sum of Supplier amounts Paid/Refunded (Both allocated and unallocated)
 
    szSQL = "Select  SUM(SWITCH(P.TYPE=24,S.Amount,P.TYPE=8,-S.Amount,P.TYPE=9,-S.Amount)) as AMT from tlbPayment P,tlbPaymentSplit S,Fund F,Supplier SP where  " & _
            "SP.SupplierID=P.SageAccountNumber AND P.TransactionID=S.PayHeader AND P.TYPE IN(8,9,24) AND  isnull(S.ClientStatementID) AND " & _
            "P.PDate <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "# AND " & _
            "S.FundID=F.FundID and  P.BankCODE='" & szSelectedBankAccount & "' and  ClientID ='" & szSelectedClient & "' "
    rsPayment.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
    If Not rsPayment.EOF Then
            dblAmt = IIf(IsNull(rsPayment.Fields.Item("AMT").Value), 0, rsPayment.Fields.Item("AMT").Value)
        'result is -50
    End If
    rsPayment.Close
    Set rsPayment = Nothing
    getClosingBalance = getClosingBalance + dblAmt

    'd)  Add (+): Sum of Bank payments and receipts
   
     szSQL = "Select  SUM(SWITCH(TransactionType=11 ,-B.NET_AMOUNT,TransactionType=12 ,B.NET_AMOUNT)) as AMT from tlbBankPayment B, Fund F  where B.DEPT_ID=F.FundID " & _
            "and BANK_AC='" & szSelectedBankAccount & "' AND F.FundCode in (" & ListOfFunds & ") AND (B.RentSumStatement='' OR isnull(B.RentSumStatement)) and clientID='" & szSelectedClient & "' " & _
            "AND B.TRAN_DATE >#" & Format(txtLastStatementDate1.text, "dd/mmm/yyyy") & "# AND B.TRAN_DATE <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "#"
    rsBankPaymentAndRcpt.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
    If Not rsBankPaymentAndRcpt.EOF Then
        dblAmt = IIf(IsNull(rsBankPaymentAndRcpt.Fields.Item("AMT").Value), 0, rsBankPaymentAndRcpt.Fields.Item("AMT").Value)
           'result is 0
    End If
     getClosingBalance = getClosingBalance + dblAmt
     getClosingBalance = Round(getClosingBalance, 2)
     rsBankPaymentAndRcpt.Close
     Set rsBankPaymentAndRcpt = Nothing
 End Function
' Private Function getClosingBalanceFinalized(dblLasClosingBalance As Double, ByVal trxToinclude As Long) As Double
'    'Pass propery as parameter for selected property
'    'No property spec:
'    Dim adoConn As New ADODB.Connection
'    Dim rsReceipt As New ADODB.Recordset
'    Dim rsPayment As New ADODB.Recordset
'    Dim rsBankPaymentAndRcpt As New ADODB.Recordset
'    Dim dblAmt As Double
'    adoConn.Open getConnectionString
'    Dim szSQL As String
'    Dim AccrualsControl As String
'    'we are not using property filter here
'    'B )***********************  Sum of Rent received Paid/Refunded ***********************************
'
'
'    szSQL = "Select  SUM(switch(type=3,S.Amount,type=4,S.Amount,type=23,-S.Amount)) as DR from tlbReceipt R,tlbReceiptSplit S,Fund F where R.TransactionID= S.RptHeader AND TYPE IN(3,4,23) " & _
'    "AND S.FundID=F.FundID AND R.BankCODE='" & szSelectedBankAccount & "' and F.FundCode<>'TENANTDEPOSIT'  AND (R.RentSumStatement='' or R.RentSumStatement=" & trxToinclude & "  or isnull(R.RentSumStatement)) AND ClientID ='" & szSelectedClient & "' " & _
'    "AND R.RDate >#" & Format(txtLastStatementDate1.text, "dd/mmm/yyyy") & "# AND R.RDate <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "#"
'    rsReceipt.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'    If Not rsReceipt.EOF Then
'            dblAmt = IIf(IsNull(rsReceipt.Fields.Item("Dr").Value), 0, rsReceipt.Fields.Item("Dr").Value)
'             'result is 255
'    End If
'    rsReceipt.Close
'    Set rsReceipt = Nothing
'    getClosingBalanceFinalized = dblLasClosingBalance + dblAmt
'
'   'c   (-): Sum of Supplier amounts Paid/Refunded (Both allocated and unallocated)
'
'    szSQL = "Select  SUM(SWITCH(P.TYPE=24,S.Amount,P.TYPE=8,-S.Amount,P.TYPE=9,-S.Amount)) as AMT from tlbPayment P,tlbPaymentSplit S,Fund F,Supplier SP where  " & _
'            "SP.SupplierID=P.SageAccountNumber AND P.TransactionID=S.PayHeader AND P.TYPE IN(8,9,24) AND (P.RentSumStatement='' or P.RentSumStatement=" & trxToinclude & "  or isnull(P.RentSumStatement)) AND " & _
'            "P.PDate >#" & Format(txtLastStatementDate1.text, "dd/mmm/yyyy") & "# AND P.PDate <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "# AND " & _
'            "S.FundID=F.FundID and  P.BankCODE='" & szSelectedBankAccount & "' and  F.FundCode in (" & ListOfFunds & ") AND ClientID ='" & szSelectedClient & "' "
'    rsPayment.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'    If Not rsPayment.EOF Then
'            dblAmt = IIf(IsNull(rsPayment.Fields.Item("AMT").Value), 0, rsPayment.Fields.Item("AMT").Value)
'        'result is -50
'    End If
'    rsPayment.Close
'    Set rsPayment = Nothing
'    getClosingBalanceFinalized = getClosingBalanceFinalized + dblAmt
'
'    'd)  Add (+): Sum of Bank payments and receipts
'
'     szSQL = "Select  SUM(SWITCH(TransactionType=11 ,-B.NET_AMOUNT,TransactionType=12 ,B.NET_AMOUNT)) as AMT from tlbBankPayment B, Fund F  where B.DEPT_ID=cstr(F.FundID) " & _
'            "and BANK_AC='" & szSelectedBankAccount & "' AND F.FundCode in (" & ListOfFunds & ") AND (B.RentSumStatement='' or B.RentSumStatement=" & trxToinclude & "  OR isnull(B.RentSumStatement)) and clientID='" & szSelectedClient & "' " & _
'            "AND B.TRAN_DATE >#" & Format(txtLastStatementDate1.text, "dd/mmm/yyyy") & "# AND B.TRAN_DATE <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "#"
'    rsBankPaymentAndRcpt.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'    If Not rsBankPaymentAndRcpt.EOF Then
'        dblAmt = IIf(IsNull(rsBankPaymentAndRcpt.Fields.Item("AMT").Value), 0, rsBankPaymentAndRcpt.Fields.Item("AMT").Value)
'           'result is 0
'    End If
'     getClosingBalanceFinalized = getClosingBalanceFinalized + dblAmt
'    rsBankPaymentAndRcpt.Close
'    Set rsBankPaymentAndRcpt = Nothing
' End Function



Private Sub ProduceClientStatement(szStatmentID As String, ByVal szReportGenID As Long)

    Dim adoConn As New ADODB.Connection
    Dim rsRentSummaryStatement As New ADODB.Recordset
    adoConn.Open getConnectionString
    Dim dblLasClosingBalance As Double
    Dim szSQL As String
    szSQL = "Select StatementClosingBal from RentSummaryStatement where StatementNo=" & GetLastStatementNoByClient & " AND ClientIDLandlordID='" & _
    szSelectedClient & "'"
    rsRentSummaryStatement.Open szSQL, adoConn, adOpenKeyset, adLockReadOnly
    If Not rsRentSummaryStatement.EOF Then
        dblLasClosingBalance = rsRentSummaryStatement!StatementClosingBal
    End If
   
    rsRentSummaryStatement.Close
    Set rsRentSummaryStatement = Nothing
    
    'szStatementID
    rsRentSummaryStatement.Open "Select * from RentSummaryStatement where 1=2", adoConn, adOpenDynamic, adLockOptimistic
    With rsRentSummaryStatement
            .AddNew
            !statementID = szStatmentID 'we are setting this column atutomatically
            !CreatedBy = User
            !CreatedDate = Now
             szCurrentStatementID = szStatmentID
            !statementNo = GetLastStatementNoByClient + 1
            !ClientIDLandlordID = szSelectedClient
            !BankCode = szSelectedBankAccount
            !PreviousStatementDate = IIf(GetLastStatementDateByClient(szStatmentID) = "", "01/01/2000", GetLastStatementDateByClient(szStatmentID)) 'This is Fromdate
            !StatementDate = Format(txtStatementDate1.text, "dd/mmmm/yyyy") 'This is todate
            !StatementOpBal = dblLasClosingBalance
           
            !Clearretentions = False 'Will need to come again
            ''******************************Produce client statement*********************************************************************
            !AccrualsAcBalance = GetAccrualsControlBalance
            !SupplierAcBalance = GetSupplierOSAmount 'GetBalance("Supplier") 'GetBalanceSupplier'wrong
            !ManagingAgentAcBalance = GetAgentBalance 'GetBalance("Agent") 'GetBalanceAgent'wrong
            !ClientACBalance = GetClientACBalance
            !LandlordACBalance = GetLandLordACBalance
            !ListOffundID = ListOfFundsForDBSave ' szSelectedFund
'            !ListOfPayableTypeID = ListOfPayableTypesForDBSave ' ListOfPayableTypes
            !TenantDepositsReceived = GetRentDeposit(szStatmentID)
            Call MarkRetentionDetails(adoConn)
            !Availablefunds = getAvailablefundsProduce(dblLasClosingBalance, szStatmentID)
            !Retentions = Val(txtRetention.text) 'we need to further analyse detail/add/deduct retension
            !ListOfinputProperties = ListOfProperties
            !PaymentsonAccount = -GetPaymentsonAccount(szStatmentID)
            'New fields added 2021-01-24
            !TenantReceipts = GetTenantReceiptsFinalized(szStatmentID)
            !SupplierPayments = GetSupplierPayment(szStatmentID)
            !BankPaymentReceipts = GetBankPaymentReceipts(szStatmentID)
            !ClientLandlordBalance = GetClientACBalance + GetLandLordACBalance
            !ClientPayments = GetClientPayments(szStatmentID)
            !LandlordPayments = GetLandLordPayments(szStatmentID)
            !ManagingAgentPayments = GetAGENTPayments(szStatmentID)
            !PayableAmount = 0 'txtRentPayable.text
            !StatementClosingBal = !Availablefunds  'getClosingBalanceFinalized(dblLasClosingBalance, szStatmentID) ' getClosingBalance(dblLasClosingBalance)
            !BankPayment = GetBankPayment(szStatmentID)
            !BankReceipts = GetBankreceipts(szStatmentID)
            !PINumber = ""
            !Generated_Date = Format(Now, "dd/mmmm/yyyy")
            !BankACBalance = BankAccBalance(adoConn, szSelectedBankAccount, szSelectedClient)
            !Printed = False
            !Emailed = False
            !Invoiced = False
            !PostToHistory = False
            !ReportGenID = szReportGenID
            !InclSupplierOS = IIf(chkExcludeSupOS.Value = 1, True, False)
            !InclMngtFeesDue = IIf(chkShowDue.Value = 1, True, False)
            .Update
    End With
    rsRentSummaryStatement.Close
    Set rsRentSummaryStatement = Nothing
'    Call MarkRetentionDetails(adoconn)
    Call WorkOnMgtfeedueSupplierOS(adoConn, szStatmentID) ' added by anol 15/08/2023
    Call frmRentPayable.loadflxPayFees("")
    adoConn.Close
    Set adoConn = Nothing
End Sub
''''Private Function WorkOnMgtfeedueSupplierOS(adoconn As ADODB.Connection, szCurrentStatementID As String)
''''    adoconn.Execute "DELETE FROM ClientStatementPurchasesSnapshot where StatementID=" & szCurrentStatementID & ""
''''    Dim rsRentSummaryStatement As New ADODB.Recordset
''''    Dim SQLforInsert As String
''''    rsRentSummaryStatement.Open "Select InclSupplierOS ,InclMngtFeesDue,ListOfFundId,ClientIDLandlordID  From RentSummaryStatement where StatementID=" & szCurrentStatementID & "", adoconn, adOpenStatic, adLockReadOnly
''''    Dim szWhere As String
''''    If Not rsRentSummaryStatement.EOF Then
''''            If IsNull(rsRentSummaryStatement("InclMngtFeesDue").Value) Then   'null means false
''''                   ' if  Or rsRentSummaryStatement("InclSupplierOS").Value = False
''''                'Enter data into snapshot table where there is I am not consedering osamount<amount because I want save all mgt fee (means fully paid or partially paid) and type management fee
''''
''''            ElseIf rsRentSummaryStatement("InclMngtFeesDue").Value = True Then
''''                    szWhere = " AND C.osamount>0" 'if there is any allocation against PI this PI shall come in the snapshot table
''''            End If
''''            'Inserting mananagement fee into snapshot table (conditional) this is the only place for report I am inserting management fee?? ans: No
''''            ' take only PI where there is no allocation
''''
''''           SQLforInsert = "select " & szCurrentStatementID & " as StatementID,P.TransactionType as Type, P.MY_ID,S.TRAN_ID,C.TransactionID,C.ClientID,C.UnitID,C.PDate,C.SageAccountNumber," & _
''''                                 "NOMINAL_CODE,'PI'& P.slnumber,S.NET_AMOUNT,VAT, 0,0,C.OSAmount,'1' " & _
''''                                 "from tblPurInv P,tblPurInvSRec S,tlbPayment C, Fund F Where C.PI=P.MY_ID and P.MY_ID= " & _
''''                                  "S.parentID and F.FundID=S.dept_ID and F.FundID in (" & rsRentSummaryStatement("ListOfFundId").Value & ") and C.ClientID='" & rsRentSummaryStatement("ClientIDLandlordID").Value & _
''''                                  "' and P.TransactionType=6 AND P.isManagementFee=true " & szWhere
''''
''''            'this one inserting allocated and unallocated both management fee
''''            adoconn.Execute "Insert into ClientStatementPurchasesSnapshot (StatementID,Type,MY_ID,SplitID,TransactionID,ClientID,PropertyID,TranDate,SupplierID,NOMINAL_CODE,PaymentRef," & _
''''                            "NetAmount,VATAmount,PaymentAmount,CreditAmount,osAmount,isManagementFee)" & _
''''                            SQLforInsert
''''
''''
''''            If IsNull(rsRentSummaryStatement("InclSupplierOS").Value) Then   'null means false
''''
''''                'Enter data into snapshot table where there is osamount<amount (means fully paid or partially paid) and type management fee
''''            ElseIf rsRentSummaryStatement("InclSupplierOS").Value = True Then
''''                   szWhere = " AND amount>osamount and osamount<>0" 'Fully unpaid
''''            End If
''''            '********* 'Inserting mananagement fee into snapshot table (conditional)*****************
''''                    'Type:6
''''                'This records do not have allocation record and they are fully unallocated
''''                'Select PI SPlit lines those have the selected fund and osAmount=amount in tlbpursplit table
''''                            SQLforInsert = "select " & szCurrentStatementID & " as StatementID,P.TransactionType as Type, P.MY_ID,S.TRAN_ID,C.TransactionID,C.ClientID,C.UnitID,C.PDate,C.SageAccountNumber," & _
''''                                "NOMINAL_CODE,'PI'& P.slnumber,S.NET_AMOUNT,VAT, 0,0,C.OSAmount,0 " & _
''''                                "from tblPurInv P,tblPurInvSRec S,tlbPayment C, Fund F Where C.PI=P.MY_ID and P.MY_ID= " & _
''''                       "S.parentID and F.FundID=S.dept_ID and F.FundID in (" & rsRentSummaryStatement("ListOfFundId").Value & ") and C.ClientID='" & _
''''                       rsRentSummaryStatement("ClientIDLandlordID").Value & "' and P.TransactionType=6 AND P.isManagementFee=false AND P.isRentPayable=false " & szWhere
''''
''''
''''                 adoconn.Execute "Insert into ClientStatementPurchasesSnapshot (StatementID,Type,MY_ID,SplitID,TransactionID,ClientID,PropertyID,TranDate,SupplierID,NOMINAL_CODE,PaymentRef," & _
''''                           "NetAmount,VATAmount,PaymentAmount,CreditAmount,osAmount,isManagementFee)" & _
''''                           SQLforInsert
''''
''''                 'Type:7
''''                   SQLforInsert = "select " & szCurrentStatementID & " as StatementID,P.TransactionType as Type, P.MY_ID,S.TRAN_ID,C.TransactionID,C.ClientID,C.UnitID,C.PDate,C.SageAccountNumber," & _
''''                                "NOMINAL_CODE,-S.NET_AMOUNT,-VAT, 0,0,-(C.OSAmount),0 " & _
''''                                "from tblPurInv P,tblPurInvSRec S,tlbPayment C, Fund F Where C.PI=P.MY_ID and P.MY_ID= " & _
''''                       "S.parentID and osamount<amount and F.FundID=S.dept_ID and F.FundID in (" & rsRentSummaryStatement("ListOfFundId").Value & ")  and C.ClientID='" & rsRentSummaryStatement("ClientIDLandlordID").Value & "' and P.TransactionType=7" & szWhere
''''
''''
''''                 adoconn.Execute "Insert into ClientStatementPurchasesSnapshot(StatementID,Type,MY_ID,SplitID,TransactionID,ClientID,PropertyID,TranDate,SupplierID,NOMINAL_CODE," & _
''''                           "NetAmount,VATAmount,PaymentAmount,CreditAmount,osAmount,isManagementFee)" & _
''''                           SQLforInsert
''''                 '****************end********************************************************
''''
''''        End If
''''        rsRentSummaryStatement.Close
''''End Function
Private Function GetStatementNobyProperty(strPropertyID As String) As Long  'this is not by client
    Dim intmaxStatementNo As Integer
    Dim adoConn As New ADODB.Connection
    Dim rsRentSummaryStatement As New ADODB.Recordset
    adoConn.Open getConnectionString
    Dim szSQL As String
    'Rent Summary Statement Opening Balance (=Closing balance of previous statement)
    'This is by client
    'Get ID by Client max ID from RentSummaryStatement
    szSQL = "Select max(StatementNoByProperty) as IDbyPR from RentSummaryStatement"
    rsRentSummaryStatement.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
    If Not rsRentSummaryStatement.EOF Then
        GetStatementNobyProperty = IIf(IsNull(rsRentSummaryStatement!IDbyPR), 0, rsRentSummaryStatement!IDbyPR)
    End If
    rsRentSummaryStatement.Close
    Set rsRentSummaryStatement = Nothing
    adoConn.Close
    Set adoConn = Nothing
End Function
Private Sub MarkRetentionDetails(adoConn As ADODB.Connection)
        Dim rsRetensionDetails As New ADODB.Recordset
        If szCurrentStatementID = "" Then Exit Sub
       ' adoConn.Execute "Delete from RetentionDetails where statementID =" & szCurrentStatementID & ""
        'Enter data into grid only memory version
        Dim iRow As Integer
        iRow = 1
        rsRetensionDetails.Open "Select * from RetentionDetails where isnull(statementID) and BankCode='" & _
                        szSelectedBankAccount & "' and ClientID='" & szSelectedClient & "'", adoConn, adOpenDynamic, adLockOptimistic
        While Not rsRetensionDetails.EOF
            With rsRetensionDetails
                    !statementID = szCurrentStatementID
                    !SlNumber = iRow
                    .Update
                    iRow = iRow + 1
            End With
            rsRetensionDetails.MoveNext
        Wend
        
        rsRetensionDetails.Close

End Sub

Private Sub GenerateCSPreview(szStatmentID As String)
    Dim X As String
    Dim adoConn As New ADODB.Connection
    Dim rsRentSummaryStatement As New ADODB.Recordset
'    Dim rsConsolidatedStatement As New ADODB.Recordset
    adoConn.Open getConnectionString
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
    Dim CSID As String
    CSID = Replace(szStatmentID, "CS", "")
    'Before writing this table you need to delete this table
    If ListOfFunds = "" Then
        MsgBox "Please select a fund", vbInformation, "Warning!"
        Exit Sub
    End If
    adoConn.Execute "Delete from  RentSummaryStatementPreview"
    If bEditMode Then
            szSQL = "Select StatementClosingBal from RentSummaryStatement where StatementNo=" & GetLastStatementNoByClient - 1 & " AND ClientIDLandlordID='" & szSelectedClient & "'"
    Else
        szSQL = "Select StatementClosingBal from RentSummaryStatement where StatementNo=" & GetLastStatementNoByClient & " AND ClientIDLandlordID='" & szSelectedClient & "'"
    End If
    rsRentSummaryStatement.Open szSQL, adoConn, adOpenKeyset, adLockReadOnly
    If Not rsRentSummaryStatement.EOF Then
        dblLasClosingBalance = rsRentSummaryStatement!StatementClosingBal
    End If

    rsRentSummaryStatement.Close
    Set rsRentSummaryStatement = Nothing
    rsRentSummaryStatement.Open "Select * from RentSummaryStatementPreview where 1=2", adoConn, adOpenDynamic, adLockOptimistic
    With rsRentSummaryStatement
            .AddNew
            !statementID = szStatmentID 'we are setting this column atutomatically
            !statementNo = GetLastStatementNoByClient + 1
            !ClientIDLandlordID = szSelectedClient
            !BankCode = szSelectedBankAccount
            !PreviousStatementDate = IIf(GetLastStatementDateByClient(szStatmentID) = "", "01/01/2000", GetLastStatementDateByClient(szStatmentID))
            !StatementDate = Format(txtStatementDate1.text, "dd/mmmm/yyyy")
            !StatementOpBal = dblLasClosingBalance
            
            !Clearretentions = False 'Will need to come again
            !AccrualsAcBalance = GetAccrualsControlBalance 'you need to check
            !SupplierAcBalance = GetSupplierOSAmount 'GetBalance("Supplier") 'GetBalanceSupplier'wrong'' for consolidated I need not filter it by property but for other I need to filter
            !ManagingAgentAcBalance = -GetAgentBalance 'GetBalance("Agent") 'GetBalanceAgent'wrong' for consolidated I need not filter it by property but for other I need to filter
            !ClientACBalance = GetClientACBalance
            !LandlordACBalance = GetLandLordACBalance
            !ListOffundID = ListOfFundsForDBSave
            !ListOfPayableTypeID = ListOfFundsForDBSave ' ListOfPayableTypes
            !ListOfinputProperties = ListOfProperties
            !TenantDepositsReceived = GetRentDepositPreview(CSID)
            !Availablefunds = getAvailablefundsPreview(dblLasClosingBalance, CSID)
            !Retentions = txtRetention.text 'we need to further analyse detail/add/deduct retension
            !PaymentsonAccount = -GetPaymentsonAccountPreview(CSID) 'date  filter added
            'New fields added 2021-02-22
            !ClientPayments = GetClientPaymentsPreview(CSID)
            !LandlordPayments = GetLandLordPaymentsPreview(CSID)
            !ManagingAgentPayments = GetAGENTPaymentsPreview(CSID) 'GetAGENTPaymentsPreview(CLng(szStatmentID))
             'New fields added 2021-01-24
            !TenantReceipts = GetTenantReceiptsPreview(CSID) 'GetTenantReceiptsPreview(CLng(szStatmentID))
            !SupplierPayments = GetSupplierPaymentPreview(CSID) ' GetSupplierPaymentPreview(CLng(szStatmentID)) 'Purchase payment
            !BankPaymentReceipts = GetBankPaymentReceiptsPreview(CSID)
            'addded newly by anol 2021-08-19
            !BankPayment = GetBankPaymentPreview(CSID)
            !BankReceipts = GetBankreceiptsPreview(CSID)
            !BankACBalancePreview = BankAccBalance(adoConn, szSelectedBankAccount, szSelectedClient)
            !ClientLandlordBalance = GetClientACBalance + GetLandLordACBalance


            Dim iIncDec As Long
            iIncDec = 0
            Dim rCount As Integer
            Dim selRow As Integer
            Dim isitPlus As Boolean
            For rCount = 1 To frmRentPayable.flxPayFees.Rows - 1
                 If frmRentPayable.flxPayFees.TextMatrix(rCount, 0) = "X" Then
                     If frmRentPayable.flxPayFees.TextMatrix(rCount, 1) = "+" Then
                        isitPlus = True
                     Else
                        isitPlus = False
                     End If
                     iIncDec = iIncDec + 1
                     selRow = rCount
                 End If
            Next

            If bEditMode = True And frmRentPayable.flxPayFees.TextMatrix(selRow, 29) <> "" Then
                 !PayableAmount = Val(txtRentPayable.text)
            Else
                 !PayableAmount = 0 'Val(txtRentPayable.text)'only first before we g n R P  !PayableAmount = 0
            End If
            '!StatementClosingBal = BankAccBalance(adoconn, szSelectedBankAccount, szSelectedClient) - !Availablefunds 'getClosingBalance(dblLasClosingBalance)
            !StatementClosingBal = !Availablefunds  'getClosingBalance(dblLasClosingBalance)
            !PINumber = ""
            !Generated_Date = Format(Now, "dd/mmmm/yyyy")
            !Printed = False
            !Emailed = False
            !Invoiced = False
            !PostToHistory = False
            .Update
    End With
    rsRentSummaryStatement.Close
    Set rsRentSummaryStatement = Nothing
    adoConn.Close
    Set adoConn = Nothing
End Sub

Private Function GetLandLordACBalanceMODpreview(trIdtoInclude As Long) As Double    'This function return result as minus This is getting LLORD balance
    Dim rsPayment As New ADODB.Recordset
    Dim szSQL As String
    Dim adoConn As New ADODB.Connection
    'F.CategoryCode = 1 Fund category 1 Means rent
    'Implement switch here in SQL
    'Bank code does not exits in PI,so do not put it in where clause
    adoConn.Open getConnectionString
    szSQL = "Select  SUM(SWITCH(P.TYPE =24,S.Amount,P.TYPE =6,S.Amount,P.TYPE =7,-S.Amount,P.TYPE =8,-S.Amount,P.TYPE =9,-S.Amount)) as AMT from tlbPayment P,tlbPaymentSplit S,Fund F,Supplier SP where " & _
            " SP.SupplierID=P.SageaccountNumber AND (P.RentSumStatement='' OR P.RentSumStatement='" & trIdtoInclude & "' or isnull(P.RentSumStatement))  AND  P.TransactionID=S.PayHeader AND P.TYPE " & _
            "IN(6,7,8,9,24) AND S.FundID=F.FundID AND   F.FundCode in (" & _
             ListOfFunds & ") AND P.ClientID ='" & szSelectedClient & "'  and SP.TYPE in ('LLORD')"
    
    rsPayment.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
    If Not rsPayment.EOF Then
        GetLandLordACBalanceMODpreview = IIf(IsNull(rsPayment.Fields.Item("AMT").Value), 0, rsPayment.Fields.Item("AMT").Value)
    End If
    rsPayment.Close
    adoConn.Close
    Set adoConn = Nothing
End Function


Private Function GetClientACBalanceModPreview(trIdtoInclude As Long) As Double    'This function return result as minus This is getting CLIENT balance
    Dim rsPayment As New ADODB.Recordset
    Dim szSQL As String
    Dim adoConn As New ADODB.Connection
    'F.CategoryCode = 1 Fund category 1 Means rent
    'Implement switch here in SQL
    'Bank code does not exits in PI,so do not put it in where clause
    adoConn.Open getConnectionString
    szSQL = "Select  SUM(SWITCH(P.TYPE =24,S.Amount,P.TYPE =6,S.Amount,P.TYPE =7,-S.Amount,P.TYPE =8,-S.Amount,P.TYPE =9,-S.Amount)) as AMT" & _
            " from tlbPayment P,tlbPaymentSplit S,Fund F,Supplier SP where " & _
            " SP.SupplierID=P.SageaccountNumber AND  P.TransactionID=S.PayHeader AND P.TYPE " & _
            "IN(6,7,8,9,24) AND S.FundID=F.FundID AND (P.RentSumStatement='' OR P.RentSumStatement='" & trIdtoInclude & "' or isnull(P.RentSumStatement))  AND   F.FundCode in (" & _
             ListOfFunds & ") AND P.ClientID ='" & szSelectedClient & "'  and SP.TYPE in ('CLIENT')"
    
    rsPayment.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
    If Not rsPayment.EOF Then
        GetClientACBalanceModPreview = IIf(IsNull(rsPayment.Fields.Item("AMT").Value), 0, rsPayment.Fields.Item("AMT").Value)
    End If
    rsPayment.Close
    adoConn.Close
    Set adoConn = Nothing
End Function

Private Function GetAgentBalance() As Double   'This function return result as minus This is getting Agent balance
    Dim rsPayment As New ADODB.Recordset
    Dim szSQL As String
    Dim adoConn As New ADODB.Connection
    'F.CategoryCode = 1 Fund category 1 Means rent
    'Implement switch here in SQL
    'Bank code does not exits in PI,so do not put it in where clause
    adoConn.Open getConnectionString
    Dim whereProperty As String

    whereProperty = "(P.UNITID IN (" & ListOfProperties & ") OR isnull(P.UNITID)  OR P.UNITID='' ) AND "
    szSQL = "Select  SUM(S.Amount) as AMT from tlbPayment P,tlbPaymentSplit S,Fund F,Supplier SP where " & _
            " SP.SupplierID=P.SageaccountNumber AND  " & whereProperty & "  P.TransactionID=S.PayHeader AND P.TYPE " & _
            "IN(6,24) AND S.FundID=F.FundID AND   F.FundCode in (" & _
             ListOfFunds & ") AND P.ClientID ='" & szSelectedClient & "'  and SP.TYPE in ('AGENT') AND  P.PDate <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "#"
    
    rsPayment.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
    If Not rsPayment.EOF Then
        GetAgentBalance = IIf(IsNull(rsPayment.Fields.Item("AMT").Value), 0, rsPayment.Fields.Item("AMT").Value)
    End If
    rsPayment.Close
    
    whereProperty = "(S.PropertyID IN (" & ListOfProperties & ") OR isnull(S.PropertyID)  OR S.PropertyID='' ) AND "
     
    szSQL = "Select  SUM(SWITCH(P.TYPE =7,-S.Amount,P.TYPE =8,-S.Amount,P.TYPE =9,-S.Amount)) as AMT from tlbPayment P, " & _
            "tlbPaymentSplit S,Fund F,Supplier SP where SP.SupplierID=P.SageaccountNumber AND  " & whereProperty & "  P.TransactionID=S.PayHeader AND P.TYPE " & _
            "IN(7,8,9) AND S.FundID=F.FundID AND   F.FundCode in (" & _
             ListOfFunds & ") AND P.ClientID ='" & szSelectedClient & "'  and SP.TYPE in ('AGENT')  AND  P.PDate <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "#"
    
    rsPayment.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
    If Not rsPayment.EOF Then
        GetAgentBalance = Round(GetAgentBalance, 2) + Round(IIf(IsNull(rsPayment.Fields.Item("AMT").Value), 0, rsPayment.Fields.Item("AMT").Value), 2)
    End If
    rsPayment.Close
    
    adoConn.Close
    Set adoConn = Nothing
End Function
Private Function GetSupplierOSAmount() As Double   'This function return result as minus'This is getting supplier balance
    If chkExcludeSupOS.Value = 1 Then 'actually include
            'Temporarily remming it 2023-08-11 by anol
            Dim rsPayment As New ADODB.Recordset
            Dim szSQL As String
            Dim adoConn As New ADODB.Connection
            'F.CategoryCode = 1 Fund category 1 Means rent
            'Implement switch here in SQL
            'Bank code does not exits in PI,so do not put it in where clause
            adoConn.Open getConnectionString
            Dim whereProperty As String
            ' for consolidated I need not filter it by property but for other I need to filter
            'If boolConsolidatedStatement = 1 Then
                    whereProperty = "(P.UNITID IN (" & ListOfProperties & ") OR isnull(P.UNITID)  OR P.UNITID='' ) AND "
        '    Else
        '            whereProperty = "P.UNITID in (" & ListOfProperties & ") AND "
        '    End If
            szSQL = "Select  SUM(S.Amount) as AMT from tlbPayment P,tlbPaymentSplit S,Fund F,Supplier SP where " & _
                    " SP.SupplierID=P.SageaccountNumber AND   " & whereProperty & " P.TransactionID=S.PayHeader AND P.TYPE " & _
                    "IN(6,24) AND S.FundID=F.FundID AND   F.FundCode in (" & _
                     ListOfFunds & ") AND P.ClientID ='" & szSelectedClient & "'  and SP.TYPE in ('Supplier')   AND  P.PDate <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "#"
        
            rsPayment.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
            If Not rsPayment.EOF Then
                GetSupplierOSAmount = IIf(IsNull(rsPayment.Fields.Item("AMT").Value), 0, rsPayment.Fields.Item("AMT").Value)
            End If
            rsPayment.Close
            'I have modified the code according to added field in paymentsplit and splite into 2 SQL modified 20211024
        '    If boolConsolidatedStatement = 1 Then
                    whereProperty = "(S.PropertyID IN (" & ListOfProperties & ") OR isnull(S.PropertyID)  OR S.PropertyID='' ) AND "
        '    Else
        '            whereProperty = "S.PropertyID in (" & ListOfProperties & ") AND "
        '    End If
            szSQL = "Select   SUM(S.Amount) as AMT from tlbPayment P,tlbPaymentSplit S,Fund F,Supplier SP where " & _
                    " SP.SupplierID=P.SageaccountNumber AND   " & whereProperty & " P.TransactionID=S.PayHeader AND P.TYPE " & _
                    "IN(7,8,9) AND S.FundID=F.FundID AND   F.FundCode in (" & _
                     ListOfFunds & ") AND P.ClientID ='" & szSelectedClient & "'  and SP.TYPE in ('Supplier')   AND  P.PDate <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "#"
        
            rsPayment.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
            If Not rsPayment.EOF Then
                GetSupplierOSAmount = GetSupplierOSAmount - IIf(IsNull(rsPayment.Fields.Item("AMT").Value), 0, rsPayment.Fields.Item("AMT").Value)
            End If
            rsPayment.Close
            adoConn.Close
            Set adoConn = Nothing
    Else
        GetSupplierOSAmount = 0
    End If
End Function
Private Function GetSupplierOSAmountAllTime() As Double   'This function return result as minus'This is getting supplier balance
    
            'Temporarily remming it 2023-08-11 by anol
            Dim rsPayment As New ADODB.Recordset
            Dim szSQL As String
            Dim adoConn As New ADODB.Connection
            'F.CategoryCode = 1 Fund category 1 Means rent
            'Implement switch here in SQL
            'Bank code does not exits in PI,so do not put it in where clause
            adoConn.Open getConnectionString
            Dim whereProperty As String
            ' for consolidated I need not filter it by property but for other I need to filter
            'If boolConsolidatedStatement = 1 Then
                    whereProperty = "(P.UNITID IN (" & ListOfProperties & ") OR isnull(P.UNITID)  OR P.UNITID='' ) AND "
        '    Else
        '            whereProperty = "P.UNITID in (" & ListOfProperties & ") AND "
        '    End If
            szSQL = "Select  SUM(S.OSAmount) as AMT from tlbPayment P,tlbPaymentSplit S,Fund F,Supplier SP where " & _
                    " SP.SupplierID=P.SageaccountNumber AND   " & whereProperty & " P.TransactionID=S.PayHeader AND P.TYPE " & _
                    "IN(6,24) AND S.FundID=F.FundID AND   F.FundCode in (" & _
                     ListOfFunds & ") AND P.ClientID ='" & szSelectedClient & "'  and SP.TYPE in ('Supplier')   AND  P.PDate <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "#"
        
            rsPayment.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
            If Not rsPayment.EOF Then
                GetSupplierOSAmountAllTime = IIf(IsNull(rsPayment.Fields.Item("AMT").Value), 0, rsPayment.Fields.Item("AMT").Value)
            End If
            rsPayment.Close
            'I have modified the code according to added field in paymentsplit and splite into 2 SQL modified 20211024
        '    If boolConsolidatedStatement = 1 Then
                    whereProperty = "(S.PropertyID IN (" & ListOfProperties & ") OR isnull(S.PropertyID)  OR S.PropertyID='' ) AND "
        '    Else
        '            whereProperty = "S.PropertyID in (" & ListOfProperties & ") AND "
        '    End If
'            szSQL = "Select   SUM(S.OSAmount) as AMT from tlbPayment P,tlbPaymentSplit S,Fund F,Supplier SP where " & _
'                    " SP.SupplierID=P.SageaccountNumber AND   " & whereProperty & " P.TransactionID=S.PayHeader AND P.TYPE " & _
'                    "IN(7,8,9) AND S.FundID=F.FundID AND   F.FundCode in (" & _
'                     ListOfFunds & ") AND P.ClientID ='" & szSelectedClient & "'  and SP.TYPE in ('Supplier')   AND  P.PDate <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "#"
'
'            rsPayment.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
'            If Not rsPayment.EOF Then
'                GetSupplierOSAmountAllTime = GetSupplierOSAmountAllTime - IIf(IsNull(rsPayment.Fields.Item("AMT").Value), 0, rsPayment.Fields.Item("AMT").Value)
'            End If
'            rsPayment.Close
            adoConn.Close
            Set adoConn = Nothing
    
End Function
Private Function GetClientOSAmount() As Double   'This function return result as minus'This is getting supplier balance
    Dim rsPayment As New ADODB.Recordset
    Dim szSQL As String
    Dim adoConn As New ADODB.Connection
    'F.CategoryCode = 1 Fund category 1 Means rent
    'Implement switch here in SQL
    'Bank code does not exits in PI,so do not put it in where clause
    adoConn.Open getConnectionString
    Dim whereProperty As String
    ' for consolidated I need not filter it by property but for other I need to filter
    'If boolConsolidatedStatement = 1 Then
            whereProperty = "(P.UNITID IN (" & ListOfProperties & ") OR isnull(P.UNITID)  OR P.UNITID='' ) AND "
'    Else
'            whereProperty = "P.UNITID in (" & ListOfProperties & ") AND "
'    End If
    szSQL = "Select  SUM(S.Amount) as AMT from tlbPayment P,tlbPaymentSplit S,Fund F,Supplier SP where " & _
            " SP.SupplierID=P.SageaccountNumber AND   " & whereProperty & " P.TransactionID=S.PayHeader AND P.TYPE " & _
            "IN(6,24) AND S.FundID=F.FundID AND   F.FundCode in (" & _
             ListOfFunds & ") AND P.ClientID ='" & szSelectedClient & "'  and SP.TYPE in ('Client')   AND  P.PDate <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "#"
    
    rsPayment.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
    If Not rsPayment.EOF Then
        GetClientOSAmount = IIf(IsNull(rsPayment.Fields.Item("AMT").Value), 0, rsPayment.Fields.Item("AMT").Value)
    End If
    rsPayment.Close
    'I have modified the code according to added field in paymentsplit and splite into 2 SQL modified 20211024
'    If boolConsolidatedStatement = 1 Then
            whereProperty = "(S.PropertyID IN (" & ListOfProperties & ") OR isnull(S.PropertyID)  OR S.PropertyID='' ) AND "
'    Else
'            whereProperty = "S.PropertyID in (" & ListOfProperties & ") AND "
'    End If
    szSQL = "Select   SUM(S.Amount) as AMT from tlbPayment P,tlbPaymentSplit S,Fund F,Supplier SP where " & _
            " SP.SupplierID=P.SageaccountNumber AND   " & whereProperty & " P.TransactionID=S.PayHeader AND P.TYPE " & _
            "IN(7,8,9) AND S.FundID=F.FundID AND   F.FundCode in (" & _
             ListOfFunds & ") AND P.ClientID ='" & szSelectedClient & "'  and SP.TYPE in ('Client')   AND  P.PDate <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "#"
    
    rsPayment.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
    If Not rsPayment.EOF Then
        GetClientOSAmount = GetClientOSAmount - IIf(IsNull(rsPayment.Fields.Item("AMT").Value), 0, rsPayment.Fields.Item("AMT").Value)
    End If
    rsPayment.Close
    adoConn.Close
    Set adoConn = Nothing
End Function
Private Function GetAgentOSAmount() As Double   'This function return result as minus'This is getting supplier balance
    Dim rsPayment As New ADODB.Recordset
    Dim szSQL As String
    Dim adoConn As New ADODB.Connection
    'F.CategoryCode = 1 Fund category 1 Means rent
    'Implement switch here in SQL
    'Bank code does not exits in PI,so do not put it in where clause
    adoConn.Open getConnectionString
    Dim whereProperty As String
    ' for consolidated I need not filter it by property but for other I need to filter
    'If boolConsolidatedStatement = 1 Then
            whereProperty = "(P.UNITID IN (" & ListOfProperties & ") OR isnull(P.UNITID)  OR P.UNITID='' ) AND "
'    Else
'            whereProperty = "P.UNITID in (" & ListOfProperties & ") AND "
'    End If
    szSQL = "Select  SUM(S.Amount) as AMT from tlbPayment P,tlbPaymentSplit S,Fund F,Supplier SP where " & _
            " SP.SupplierID=P.SageaccountNumber AND   " & whereProperty & " P.TransactionID=S.PayHeader AND P.TYPE " & _
            "IN(6,24) AND S.FundID=F.FundID AND   F.FundCode in (" & _
             ListOfFunds & ") AND P.ClientID ='" & szSelectedClient & "'  and SP.TYPE in ('Agent')   AND  P.PDate <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "#"
    
    rsPayment.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
    If Not rsPayment.EOF Then
        GetAgentOSAmount = IIf(IsNull(rsPayment.Fields.Item("AMT").Value), 0, rsPayment.Fields.Item("AMT").Value)
    End If
    rsPayment.Close
    'I have modified the code according to added field in paymentsplit and splite into 2 SQL modified 20211024
'    If boolConsolidatedStatement = 1 Then
            whereProperty = "(S.PropertyID IN (" & ListOfProperties & ") OR isnull(S.PropertyID)  OR S.PropertyID='' ) AND "
'    Else
'            whereProperty = "S.PropertyID in (" & ListOfProperties & ") AND "
'    End If
    szSQL = "Select   SUM(S.Amount) as AMT from tlbPayment P,tlbPaymentSplit S,Fund F,Supplier SP where " & _
            " SP.SupplierID=P.SageaccountNumber AND   " & whereProperty & " P.TransactionID=S.PayHeader AND P.TYPE " & _
            "IN(7,8,9) AND S.FundID=F.FundID AND   F.FundCode in (" & _
             ListOfFunds & ") AND P.ClientID ='" & szSelectedClient & "'  and SP.TYPE in ('Agent')   AND  P.PDate <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "#"
    
    rsPayment.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
    If Not rsPayment.EOF Then
        GetAgentOSAmount = GetAgentOSAmount - IIf(IsNull(rsPayment.Fields.Item("AMT").Value), 0, rsPayment.Fields.Item("AMT").Value)
    End If
    rsPayment.Close
    adoConn.Close
    Set adoConn = Nothing
End Function
Private Function GetClientACBalance() As Double   'This function return result as minus This is getting CLIENT balance
    Dim rsPayment As New ADODB.Recordset
    Dim szSQL As String
    Dim adoConn As New ADODB.Connection
    'F.CategoryCode = 1 Fund category 1 Means rent
    'Implement switch here in SQL
    'Bank code does not exits in PI,so do not put it in where clause
    adoConn.Open getConnectionString
    szSQL = "Select  SUM(SWITCH(P.TYPE =24,S.Amount,P.TYPE =6,S.Amount,P.TYPE =7,-S.Amount,P.TYPE =8,-S.Amount,P.TYPE =9,-S.Amount)) as AMT" & _
            " from tlbPayment P,tlbPaymentSplit S,Fund F,Supplier SP where " & _
            " SP.SupplierID=P.SageaccountNumber AND  P.TransactionID=S.PayHeader AND P.TYPE " & _
            "IN(6,7,8,9,24) AND S.FundID=F.FundID AND   F.FundCode in (" & _
             ListOfFunds & ") AND P.ClientID ='" & szSelectedClient & "'  and SP.TYPE in ('CLIENT')   AND  P.PDate <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "#"
    
    rsPayment.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
    If Not rsPayment.EOF Then
        GetClientACBalance = -IIf(IsNull(rsPayment.Fields.Item("AMT").Value), 0, rsPayment.Fields.Item("AMT").Value)
    End If
    rsPayment.Close
    adoConn.Close
    Set adoConn = Nothing
End Function
Private Function GetLandLordACBalance() As Double   'This function return result as minus This is getting LLORD balance
    Dim rsPayment As New ADODB.Recordset
    Dim szSQL As String
    Dim adoConn As New ADODB.Connection
    'F.CategoryCode = 1 Fund category 1 Means rent
    'Implement switch here in SQL
    'Bank code does not exits in PI,so do not put it in where clause
    adoConn.Open getConnectionString
    szSQL = "Select  SUM(SWITCH(P.TYPE =24,S.Amount,P.TYPE =6,S.Amount,P.TYPE =7,-S.Amount,P.TYPE =8,-S.Amount,P.TYPE =9,-S.Amount)) as AMT from tlbPayment P,tlbPaymentSplit S,Fund F,Supplier SP where " & _
            " SP.SupplierID=P.SageaccountNumber AND  P.TransactionID=S.PayHeader AND P.TYPE " & _
            "IN(6,7,8,9,24) AND S.FundID=F.FundID AND   F.FundCode in (" & _
             ListOfFunds & ") AND P.ClientID ='" & szSelectedClient & "'  and SP.TYPE in ('LLORD') AND  P.PDate <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "#"
    
    rsPayment.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
    If Not rsPayment.EOF Then
        GetLandLordACBalance = IIf(IsNull(rsPayment.Fields.Item("AMT").Value), 0, rsPayment.Fields.Item("AMT").Value)
    End If
    rsPayment.Close
    adoConn.Close
    Set adoConn = Nothing
End Function
Private Function GetAgentBalanceModPreview(trIdtoInclude As Long) As Double       'This function return result as minus This is getting Agent balance
    Dim rsPayment As New ADODB.Recordset
    Dim szSQL As String
    Dim adoConn As New ADODB.Connection
    'F.CategoryCode = 1 Fund category 1 Means rent
    'Implement switch here in SQL
    'Bank code does not exits in PI,so do not put it in where clause
    adoConn.Open getConnectionString
    Dim whereProperty As String

    whereProperty = "(P.UNITID IN (" & ListOfProperties & ") OR isnull(P.UNITID)  OR P.UNITID='' ) AND "
    szSQL = "Select  SUM(S.Amount) as AMT from tlbPayment P,tlbPaymentSplit S,Fund F,Supplier SP where " & _
            " SP.SupplierID=P.SageaccountNumber  AND (P.RentSumStatement='' OR P.RentSumStatement='" & trIdtoInclude & "' or isnull(P.RentSumStatement))  AND  " & whereProperty & "  P.TransactionID=S.PayHeader AND P.TYPE " & _
            "IN(6,24) AND S.FundID=F.FundID AND   F.FundCode in (" & _
             ListOfFunds & ") AND P.ClientID ='" & szSelectedClient & "'  and SP.TYPE in ('AGENT')"
    
    rsPayment.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
    If Not rsPayment.EOF Then
        GetAgentBalanceModPreview = IIf(IsNull(rsPayment.Fields.Item("AMT").Value), 0, rsPayment.Fields.Item("AMT").Value)
    End If
    rsPayment.Close
    
    whereProperty = "(S.PropertyID IN (" & ListOfProperties & ") OR isnull(S.PropertyID)  OR S.PropertyID='' ) AND "
     
    szSQL = "Select  SUM(SWITCH(P.TYPE =7,-S.Amount,P.TYPE =8,-S.Amount,P.TYPE =9,-S.Amount)) as AMT from tlbPayment P, " & _
            "tlbPaymentSplit S,Fund F,Supplier SP where SP.SupplierID=P.SageaccountNumber   AND (P.RentSumStatement='' OR P.RentSumStatement='" & trIdtoInclude & _
            "' or isnull(P.RentSumStatement)) AND  " & whereProperty & "  P.TransactionID=S.PayHeader AND P.TYPE " & _
            "IN(7,8,9) AND S.FundID=F.FundID AND   F.FundCode in (" & _
             ListOfFunds & ") AND P.ClientID ='" & szSelectedClient & "'  and SP.TYPE in ('AGENT')"
    
    rsPayment.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
    If Not rsPayment.EOF Then
        GetAgentBalanceModPreview = Round(GetAgentBalance, 2) + IIf(IsNull(rsPayment.Fields.Item("AMT").Value), 0, rsPayment.Fields.Item("AMT").Value)
    End If
    rsPayment.Close
    
    adoConn.Close
    Set adoConn = Nothing
End Function
Private Function GetAgentBalanceNonConsolidated(ByVal propertyID As String) As Double    'This function return result as minus This is getting Agent balance
    Dim rsPayment As New ADODB.Recordset
    Dim szSQL As String
    Dim adoConn As New ADODB.Connection
    'Big problem propertyID can be UnitID when type 6 and 24 means PI and PPR
    'now after allocation it is writing to tlbPaymentSplit
    ' TO solve this I need to split the SQL query
    'F.CategoryCode = 1 Fund category 1 Means rent
    'Implement switch here in SQL
    'Bank code does not exits in PI,so do not put it in where clause
    adoConn.Open getConnectionString
    Dim whereProperty As String

    whereProperty = "P.UNITID ='" & propertyID & "' AND "

    szSQL = "Select  SUM(SWITCH(P.TYPE =24,S.Amount,P.TYPE =6,S.Amount,P.TYPE =7,-S.Amount,P.TYPE =8,-S.Amount,P.TYPE =9,-S.Amount)) as AMT from tlbPayment P,tlbPaymentSplit S,Fund F,Supplier SP where " & _
            " SP.SupplierID=P.SageaccountNumber AND  " & whereProperty & "  P.TransactionID=S.PayHeader AND P.TYPE " & _
            "IN(6,24) AND S.FundID=F.FundID AND   F.FundCode in (" & _
             ListOfFunds & ") AND P.ClientID ='" & szSelectedClient & "'  and SP.TYPE in ('AGENT')"
    
    rsPayment.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
    If Not rsPayment.EOF Then
        GetAgentBalanceNonConsolidated = IIf(IsNull(rsPayment.Fields.Item("AMT").Value), 0, rsPayment.Fields.Item("AMT").Value)
    End If
    rsPayment.Close
    
    whereProperty = "S.PropertyID ='" & propertyID & "' AND "
       
    szSQL = "Select  SUM(SWITCH(P.TYPE =24,S.Amount,P.TYPE =6,S.Amount,P.TYPE =7,-S.Amount,P.TYPE =8,-S.Amount,P.TYPE =9,-S.Amount)) as AMT from tlbPayment P,tlbPaymentSplit S,Fund F,Supplier SP where " & _
            " SP.SupplierID=P.SageaccountNumber AND  " & whereProperty & "  P.TransactionID=S.PayHeader AND P.TYPE " & _
            "IN(7,8,9) AND S.FundID=F.FundID AND   F.FundCode in (" & _
             ListOfFunds & ") AND P.ClientID ='" & szSelectedClient & "'  and SP.TYPE in ('AGENT')"
    
    rsPayment.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
    If Not rsPayment.EOF Then
        GetAgentBalanceNonConsolidated = GetAgentBalanceNonConsolidated + IIf(IsNull(rsPayment.Fields.Item("AMT").Value), 0, rsPayment.Fields.Item("AMT").Value)
    End If
    rsPayment.Close
    
    
    adoConn.Close
    Set adoConn = Nothing
End Function

Private Function GetSupplierOSAmountNonConsolidated(ByVal szPropertyID) As Double   'This function return result as minus'This is getting supplier balance
    Dim rsPayment As New ADODB.Recordset
    Dim szSQL As String
    Dim adoConn As New ADODB.Connection
    'F.CategoryCode = 1 Fund category 1 Means rent
    'Implement switch here in SQL
    'Bank code does not exits in PI,so do not put it in where clause
    adoConn.Open getConnectionString
    Dim whereProperty As String
    ' for consolidated I need not filter it by property but for other I need to filter
    whereProperty = "S.PropertyID ='" & szPropertyID & "' AND "

    szSQL = "Select  SUM(S.Amount) as AMT from tlbPayment P,tlbPaymentSplit S,Fund F,Supplier SP where " & _
            " SP.SupplierID=P.SageaccountNumber AND   " & whereProperty & " P.TransactionID=S.PayHeader AND P.TYPE " & _
            "IN(6,24) AND S.FundID=F.FundID AND   F.FundCode in (" & _
             ListOfFunds & ") AND P.ClientID ='" & szSelectedClient & "'  and SP.TYPE in ('Supplier')"
    
    rsPayment.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
    If Not rsPayment.EOF Then
        GetSupplierOSAmountNonConsolidated = IIf(IsNull(rsPayment.Fields.Item("AMT").Value), 0, rsPayment.Fields.Item("AMT").Value)
    End If
    rsPayment.Close
    'I have modified the code according to added field in paymentsplit and splite into 2 SQL modified 20211024
    szSQL = "Select   SUM(S.Amount) as AMT from tlbPayment P,tlbPaymentSplit S,Fund F,Supplier SP where " & _
            " SP.SupplierID=P.SageaccountNumber AND   " & whereProperty & " P.TransactionID=S.PayHeader AND P.TYPE " & _
            "IN(7,8,9) AND S.FundID=F.FundID AND   F.FundCode in (" & _
             ListOfFunds & ") AND P.ClientID ='" & szSelectedClient & "'  and SP.TYPE in ('Supplier')"
    
    rsPayment.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
    If Not rsPayment.EOF Then
        GetSupplierOSAmountNonConsolidated = GetSupplierOSAmount - IIf(IsNull(rsPayment.Fields.Item("AMT").Value), 0, rsPayment.Fields.Item("AMT").Value)
    End If
    rsPayment.Close
    adoConn.Close
    Set adoConn = Nothing
End Function
'Private Function GetClientPayments() As Double   'This function return result as minus'This is getting supplier balance
'    Dim rsPayment As New ADODB.Recordset
'    Dim szSQL As String
'    Dim adoConn As New ADODB.Connection
'    'F.CategoryCode = 1 Fund category 1 Means rent
'    'Implement switch here in SQL
'    'Bank code does not exits in PI,so do not put it in where clause
'    Dim whereProperty As String
'    adoConn.Open getConnectionString
'    If boolConsolidatedStatement = 1 Then
'            whereProperty = "(P.UNITID IN (" & ListOfProperties & ") OR isnull(P.UNITID)  OR P.UNITID='' ) AND "
'    Else
'            whereProperty = "P.UNITID in (" & ListOfProperties & ") AND "
'    End If
'    szSQL = "Select  SUM(SWITCH(P.TYPE =24,S.Amount,P.TYPE =6,S.Amount,P.TYPE =7,-S.Amount,P.TYPE =8,-S.Amount,P.TYPE =9,-S.Amount)) as AMT " & _
'            "from tlbPayment P,tlbPaymentSplit S,Fund F,Supplier SP where " & _
'            " SP.SupplierID=P.SageaccountNumber AND " & whereProperty & "  P.TransactionID=S.PayHeader AND P.TYPE " & _
'            "IN(7,8,9) AND S.FundID=F.FundID AND   F.FundCode in (" & _
'             ListOfFunds & ") AND P.ClientID ='" & szSelectedClient & "'  and SP.TYPE in ('Client')" & _
'                 "AND P.PDate >#" & Format(txtLastStatementDate1.text, "dd/mmm/yyyy") & "# AND P.PDate <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "#"
'
'
'    rsPayment.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'    If Not rsPayment.EOF Then
'        GetClientPayments = IIf(IsNull(rsPayment.Fields.Item("AMT").Value), 0, rsPayment.Fields.Item("AMT").Value)
'    End If
'    rsPayment.Close
'    adoConn.Close
'    Set adoConn = Nothing
'End Function
'Private Function GetLandLordPayments() As Double   'This function return result as minus'This is getting supplier balance
'    Dim rsPayment As New ADODB.Recordset
'    Dim szSQL As String
'    Dim adoConn As New ADODB.Connection
'    'F.CategoryCode = 1 Fund category 1 Means rent
'    'Implement switch here in SQL
'    'Bank code does not exits in PI,so do not put it in where clause
'    Dim whereProperty As String
'    adoConn.Open getConnectionString
'      If boolConsolidatedStatement = 1 Then
'            whereProperty = "(P.UNITID IN (" & ListOfProperties & ") OR isnull(P.UNITID)  OR P.UNITID='' ) AND "
'    Else
'            whereProperty = "P.UNITID in (" & ListOfProperties & ") AND "
'    End If
'    szSQL = "Select  SUM(SWITCH(P.TYPE =24,S.Amount,P.TYPE =6,S.Amount,P.TYPE =7,-S.Amount,P.TYPE =8,-S.Amount,P.TYPE =9,-S.Amount)) as AMT" & _
'            " from tlbPayment P,tlbPaymentSplit S,Fund F,Supplier SP where " & _
'            " SP.SupplierID=P.SageaccountNumber AND " & whereProperty & "   P.TransactionID=S.PayHeader AND P.TYPE " & _
'            "IN(7,8,9) AND S.FundID=F.FundID AND   F.FundCode in (" & _
'             ListOfFunds & ") AND P.ClientID ='" & szSelectedClient & "'  and SP.TYPE in ('LLORD')" & _
'             "AND P.PDate >#" & Format(txtLastStatementDate1.text, "dd/mmm/yyyy") & "# AND P.PDate <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "#"
'
'    rsPayment.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'    If Not rsPayment.EOF Then
'        GetLandLordPayments = IIf(IsNull(rsPayment.Fields.Item("AMT").Value), 0, rsPayment.Fields.Item("AMT").Value)
'    End If
'    rsPayment.Close
'    adoConn.Close
'    Set adoConn = Nothing
'End Function
'Private Function GetAGENTPayments(ByVal trxToinclude As Long) As Double    'This function return result as minus'This is getting supplier balance
'    Dim rsPayment As New ADODB.Recordset
'    Dim szSQL As String
'    Dim adoConn As New ADODB.Connection
'    'F.CategoryCode = 1 Fund category 1 Means rent
'    'Implement switch here in SQL
'    'Bank code does not exits in PI,so do not put it in where clause
'    Dim whereProperty As String
'    adoConn.Open getConnectionString
''      If boolConsolidatedStatement = 1 Then
'            whereProperty = "(P.UNITID IN (" & ListOfProperties & ") OR isnull(P.UNITID)  OR P.UNITID='' ) AND "
''    Else
''            whereProperty = "P.UNITID in (" & ListOfProperties & ") AND "
''    End If
'
'
'    szSQL = "Select   SUM(PS.PaymentAmount) as AMT from tlbPayment P,tlbPaymentSplit S,Paytransactions PS,Fund F,Supplier SP where  " & _
'            "PS.FROMTran=P.TransactionID AND SP.SupplierID=P.SageAccountNumber  AND PS.Deleteflag=False AND P.TransactionID=S.payHeader AND S.ClientStatementID=" & trxToinclude & " AND " & _
'            "PS.FundID=F.FundID and P.type=24 and " & whereProperty & " P.BankCODE='" & szSelectedBankAccount & "' and Sp.Type='AGENT' and P.ClientID ='" & szSelectedClient & "' " & _
'            "AND P.PDate <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "#"
'
'rsPayment.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'    If Not rsPayment.EOF Then
'        GetAGENTPayments = IIf(IsNull(rsPayment.Fields.Item("AMT").Value), 0, rsPayment.Fields.Item("AMT").Value)
'    End If
'    rsPayment.Close
'    whereProperty = "(S.PropertyID IN (" & ListOfProperties & ") OR isnull(S.PropertyID)  OR S.PropertyID='' ) AND "
'
''Otherwise you shall get duplicated value
'         szSQL = "Select   SUM(PS.PaymentAmount) as AMT from tlbPayment P,tlbPaymentSplit S,PaytransactionsSplit PS,Fund F,Supplier SP where  " & _
'            "PS.TransactionID=S.PayTransactionIDSplit and PS.FromTran=P.TransactionID AND SP.SupplierID=P.SageAccountNumber  AND PS.Deleteflag=False AND S.ClientStatementID=" & trxToinclude & " AND " & _
'            " P.TransactionID=S.PayHeader AND PS.FundID=F.FundID and P.type IN(8,9) and " & whereProperty & " P.BankCODE='" & szSelectedBankAccount & "' " & _
'            "and Sp.Type='AGENT' AND P.ClientID ='" & szSelectedClient & "'  AND  P.PDate <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "#"
'
'
'    rsPayment.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'    If Not rsPayment.EOF Then
'        GetAGENTPayments = GetAGENTPayments - IIf(IsNull(rsPayment.Fields.Item("AMT").Value), 0, rsPayment.Fields.Item("AMT").Value)
'    End If
'    rsPayment.Close
'    adoConn.Close
'    Set adoConn = Nothing
'End Function
Private Function GetAGENTPaymentsPreview(ByVal idToinclude As Long) As Double   'This function return result as minus'This is getting supplier balance
    Dim rsPayment As New ADODB.Recordset
    Dim szSQL As String
    Dim adoConn As New ADODB.Connection
    'F.CategoryCode = 1 Fund category 1 Means rent
    'Implement switch here in SQL
    'Bank code does not exits in PI,so do not put it in where clause
    Dim whereProperty As String
    adoConn.Open getConnectionString
'      If boolConsolidatedStatement = 1 Then
            whereProperty = "(P.UNITID IN (" & ListOfProperties & ") OR isnull(P.UNITID)  OR P.UNITID='' ) AND "
'    Else
'            whereProperty = "P.UNITID in (" & ListOfProperties & ") AND "
'    End If

             
    szSQL = "Select   SUM(S.Amount) as AMT from tlbPayment P,tlbPaymentSplit S,Fund F,Supplier SP where  " & _
            "S.Payheader=P.TransactionID AND SP.SupplierID=P.SageAccountNumber  AND S.Amount>S.OSAmount  AND S.ClientStatementPrevID=" & idToinclude & " AND " & _
            "S.FundID=F.FundID and P.type=24 and " & whereProperty & " P.BankCODE='" & szSelectedBankAccount & "' and Sp.Type='AGENT' and P.ClientID ='" & szSelectedClient & "' " & _
            "AND P.PDate <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "#"

rsPayment.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
    If Not rsPayment.EOF Then
        GetAGENTPaymentsPreview = IIf(IsNull(rsPayment.Fields.Item("AMT").Value), 0, rsPayment.Fields.Item("AMT").Value)
    End If
    rsPayment.Close
    whereProperty = "(S.PropertyID IN (" & ListOfProperties & ") OR isnull(S.PropertyID)  OR S.PropertyID='' ) AND "

'Otherwise you shall get duplicated value
         szSQL = "Select   SUM(S.Amount) as AMT from tlbPayment P,tlbPaymentSplit S,Fund F,Supplier SP where  " & _
            "S.Payheader=P.TransactionID  AND SP.SupplierID=P.SageAccountNumber  AND  S.Amount>S.OSAmount AND S.ClientStatementPrevID=" & idToinclude & " AND " & _
            " P.TransactionID=S.PayHeader AND S.FundID=F.FundID and P.type IN(8,9) and " & whereProperty & " P.BankCODE='" & szSelectedBankAccount & "' " & _
            "and Sp.Type='AGENT' AND P.ClientID ='" & szSelectedClient & "'  AND  P.PDate <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "#"

    
    rsPayment.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
    If Not rsPayment.EOF Then
        GetAGENTPaymentsPreview = GetAGENTPaymentsPreview - IIf(IsNull(rsPayment.Fields.Item("AMT").Value), 0, rsPayment.Fields.Item("AMT").Value)
    End If
    rsPayment.Close
    adoConn.Close
    Set adoConn = Nothing
    
End Function
Private Function GetAGENTPayments(ByVal idToinclude As Long) As Double   'This function return result as minus'This is getting supplier balance
    Dim rsPayment As New ADODB.Recordset
    Dim szSQL As String
    Dim adoConn As New ADODB.Connection
    'F.CategoryCode = 1 Fund category 1 Means rent
    'Implement switch here in SQL
    'Bank code does not exits in PI,so do not put it in where clause
    Dim whereProperty As String
    adoConn.Open getConnectionString
'      If boolConsolidatedStatement = 1 Then
            whereProperty = "(P.UNITID IN (" & ListOfProperties & ") OR isnull(P.UNITID)  OR P.UNITID='' ) AND "
'    Else
'            whereProperty = "P.UNITID in (" & ListOfProperties & ") AND "
'    End If

             
    szSQL = "Select   SUM(S.Amount) as AMT from tlbPayment P,tlbPaymentSplit S,Fund F,Supplier SP where  " & _
            "S.Payheader=P.TransactionID AND SP.SupplierID=P.SageAccountNumber  AND S.Amount>S.OSAmount  AND S.ClientStatementID=" & idToinclude & " AND " & _
            "S.FundID=F.FundID and P.type=24 and " & whereProperty & " P.BankCODE='" & szSelectedBankAccount & "' and Sp.Type='AGENT' and P.ClientID ='" & szSelectedClient & "' " & _
            "AND P.PDate <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "#"

rsPayment.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
    If Not rsPayment.EOF Then
        GetAGENTPayments = IIf(IsNull(rsPayment.Fields.Item("AMT").Value), 0, rsPayment.Fields.Item("AMT").Value)
    End If
    rsPayment.Close
    whereProperty = "(S.PropertyID IN (" & ListOfProperties & ") OR isnull(S.PropertyID)  OR S.PropertyID='' ) AND "

'Otherwise you shall get duplicated value
         szSQL = "Select   SUM(S.Amount) as AMT from tlbPayment P,tlbPaymentSplit S,Fund F,Supplier SP where  " & _
            "S.Payheader=P.TransactionID  AND SP.SupplierID=P.SageAccountNumber  AND  S.Amount>S.OSAmount AND S.ClientStatementID=" & idToinclude & " AND " & _
            "S.FundID=F.FundID and P.type IN(8,9) and " & whereProperty & " P.BankCODE='" & szSelectedBankAccount & "' " & _
            "and Sp.Type='AGENT' AND P.ClientID ='" & szSelectedClient & "'  AND  P.PDate <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "#"

    
    rsPayment.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
    If Not rsPayment.EOF Then
        GetAGENTPayments = GetAGENTPayments - IIf(IsNull(rsPayment.Fields.Item("AMT").Value), 0, rsPayment.Fields.Item("AMT").Value)
    End If
    rsPayment.Close
    adoConn.Close
    Set adoConn = Nothing
    
End Function
'Private Function GetSupplierPayment() As Double
'    Dim rsPayment As New ADODB.Recordset
'    Dim rsReceipt As New ADODB.Recordset
'    Dim szSQL As String
'    Dim adoConn As New ADODB.Connection
'    Dim whereProperty As String
'    Dim dblAmt As Double
'    'F.CategoryCode = 1 Fund category 1 Means rent
'    'Implement switch here in SQL
'     If boolConsolidatedStatement = 1 Then
'            whereProperty = "AND (P.UNITID IN (" & ListOfProperties & ") OR isnull(P.UNITID) OR P.UNITID='') "
'    Else
'            whereProperty = "AND P.UNITID in (" & ListOfProperties & ") "
'    End If
'    adoConn.Open getConnectionString
'    szSQL = "Select  SUM(SWITCH(P.TYPE =24,S.Amount,P.TYPE =8,-S.Amount,P.TYPE =9,-S.Amount)) as AMT from tlbPayment P,tlbPaymentSplit S,Fund F,Supplier SP where " & _
'            "SP.SupplierID=P.SageaccountNumber AND  P.TransactionID=S.PayHeader AND P.TYPE IN(8,9,24) AND " & _
'            "P.PDate >#" & Format(txtLastStatementDate1.text, "dd/mmm/yyyy") & "# AND P.PDate <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "# AND " & _
'            "S.FundID=F.FundID AND  P.BankCODE='" & szSelectedBankAccount & "'  AND (P.RentSumStatement='' OR isnull(P.RentSumStatement)) and  F.FundCode in (" & _
'             ListOfFunds & ") AND P.ClientID ='" & szSelectedClient & "' AND (SP.Type='Supplier' ) " & whereProperty
'
'    rsPayment.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'    If Not rsPayment.EOF Then
'        GetSupplierPayment = IIf(IsNull(rsPayment.Fields.Item("AMT").Value), 0, rsPayment.Fields.Item("AMT").Value)
'    End If
'    rsPayment.Close
'    If boolConsolidatedStatement = 1 Then
'            whereProperty = "AND (G.PropertyID IN (" & ListOfProperties & ") OR isnull(G.PropertyID) OR G.PropertyID='') "
'    Else
'            whereProperty = "AND G.PropertyID in (" & ListOfProperties & ") "
'    End If
'             szSQL = "Select  SUM(AL.VatAmount)  as AMT from tlbPayment P,PayTransactions AL,tlbPaymentSplit S,Fund F,Supplier SP,Property PR,GLobalData G  where  G.PropertyID=PR.PropertyID AND isAgentToSubmit=true  AND " & _
'                "SP.SupplierID=P.SageAccountNumber AND AL.Deleteflag=false and PR.propertyID=p.UNITID AND P.TransactionID=S.PayHeader AND P.TYPE IN(8,9,24) AND (P.RentSumStatement='' or isnull(P.RentSumStatement)) AND " & _
'                "S.FundID=F.FundID AND AL.FromTran=P.transactionID " & whereProperty & " AND P.BankCODE='" & szSelectedBankAccount & "' and  F.FundCode in (" & ListOfFunds & ") AND P.ClientID ='" & szSelectedClient & "' " & _
'                "AND P.PDate >#" & Format(txtLastStatementDate1.text, "dd/mmm/yyyy") & "# AND P.PDate <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "#  AND (SP.Type='Supplier' ) "
'        rsPayment.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'        If Not rsPayment.EOF Then
'                GetSupplierPayment = GetSupplierPayment + IIf(IsNull(rsPayment.Fields.Item("AMT").Value), 0, rsPayment.Fields.Item("AMT").Value)
'                'result is -15
'        End If
'        rsPayment.Close
'        Set rsPayment = Nothing
'         GetSupplierPayment = GetSupplierPayment - dblAmt
'    adoConn.Close
'    Set adoConn = Nothing
'End Function
Private Function GetSupplierPaymentNonConsolidated() As Double
    Dim rsPayment As New ADODB.Recordset
    Dim rsReceipt As New ADODB.Recordset
    Dim szSQL As String
    Dim adoConn As New ADODB.Connection
    Dim whereProperty As String
    Dim dblAmt As Double
    'F.CategoryCode = 1 Fund category 1 Means rent
    'Implement switch here in SQL

        whereProperty = "AND P.UNITID in (" & ListOfProperties & ") "

        adoConn.Open getConnectionString
        szSQL = "Select  SUM(SWITCH(P.TYPE =24,S.Amount,P.TYPE =8,-S.Amount,P.TYPE =9,-S.Amount)) as AMT from tlbPayment P,tlbPaymentSplit S,Fund F,Supplier SP where " & _
        "SP.SupplierID=P.SageaccountNumber AND  P.TransactionID=S.PayHeader AND P.TYPE IN(8,9,24) AND " & _
        "P.PDate >#" & Format(txtLastStatementDate1.text, "dd/mmm/yyyy") & "# AND P.PDate <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "# AND " & _
        "S.FundID=F.FundID AND  P.BankCODE='" & szSelectedBankAccount & "'  AND (P.RentSumStatement='' OR isnull(P.RentSumStatement)) and  F.FundCode in (" & _
         ListOfFunds & ") AND P.ClientID ='" & szSelectedClient & "' AND (SP.Type='Supplier' OR SP.Type='Agent') " & whereProperty
        
        rsPayment.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
        If Not rsPayment.EOF Then
            GetSupplierPaymentNonConsolidated = IIf(IsNull(rsPayment.Fields.Item("AMT").Value), 0, rsPayment.Fields.Item("AMT").Value)
        End If
        rsPayment.Close


        whereProperty = "AND G.PropertyID in (" & ListOfProperties & ") "
    
        szSQL = "Select  SUM(AL.VatAmount)  as AMT from tlbPayment P,PayTransactions AL,tlbPaymentSplit S,Fund F,Supplier SP,Property PR,GLobalData G  where  G.PropertyID=PR.PropertyID AND isAgentToSubmit=true  AND " & _
            "SP.SupplierID=P.SageAccountNumber AND AL.Deleteflag=false and PR.propertyID=p.UNITID AND P.TransactionID=S.PayHeader AND P.TYPE IN(8,9,24) AND (P.RentSumStatement='' or isnull(P.RentSumStatement)) AND " & _
            "S.FundID=F.FundID AND AL.FromTran=P.transactionID " & whereProperty & " AND P.BankCODE='" & szSelectedBankAccount & "' and  F.FundCode in (" & ListOfFunds & ") AND P.ClientID ='" & szSelectedClient & "' " & _
            "AND P.PDate >#" & Format(txtLastStatementDate1.text, "dd/mmm/yyyy") & "# AND P.PDate <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "#"
        rsPayment.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
        If Not rsPayment.EOF Then
                GetSupplierPaymentNonConsolidated = GetSupplierPaymentNonConsolidated + IIf(IsNull(rsPayment.Fields.Item("AMT").Value), 0, rsPayment.Fields.Item("AMT").Value)
                'result is -15
        End If
        rsPayment.Close
        Set rsPayment = Nothing
        
    adoConn.Close
    Set adoConn = Nothing
End Function
Private Function GetSupplierPaymentPreview(ByVal idToinclude As Long) As Double
    Dim rsPayment As New ADODB.Recordset
    Dim rsReceipt As New ADODB.Recordset
    Dim szSQL As String
    Dim adoConn As New ADODB.Connection
    Dim whereProperty As String
    Dim dblAmt As Double
    'F.CategoryCode = 1 Fund category 1 Means rent
            whereProperty = "AND (S.PropertyID IN (" & ListOfProperties & ") OR isnull(P.UNITID) OR P.UNITID='') "
'    Else
'            whereProperty = "AND P.PropertyID in (" & ListOfProperties & ") "
'    End If
    adoConn.Open getConnectionString
    szSQL = "Select  SUM(SWITCH(P.TYPE =24,S.Amount,P.TYPE =8,-S.Amount,P.TYPE =9,-S.Amount)) as AMT from tlbPayment P,tlbPaymentSplit S,Fund F,Supplier SP where " & _
            "SP.SupplierID=P.SageaccountNumber AND  P.TransactionID=S.PayHeader AND P.TYPE IN(8,9,24) AND " & _
            "P.PDate >#" & Format(txtLastStatementDate1.text, "dd/mmm/yyyy") & "# AND P.PDate <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "# AND " & _
            "S.FundID=F.FundID AND  P.BankCODE='" & szSelectedBankAccount & "'  AND  S.ClientStatementPrevID=" & idToinclude & "" & _
            " AND P.ClientID ='" & szSelectedClient & "' AND (SP.Type='Supplier' ) " & whereProperty
    
    rsPayment.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
    If Not rsPayment.EOF Then
        GetSupplierPaymentPreview = IIf(IsNull(rsPayment.Fields.Item("AMT").Value), 0, rsPayment.Fields.Item("AMT").Value)
    End If
    rsPayment.Close
            whereProperty = "AND (G.PropertyID IN (" & ListOfProperties & ") OR isnull(G.PropertyID) OR G.PropertyID='') "
'    Else
'            whereProperty = "AND G.PropertyID in (" & ListOfProperties & ") "
'    End If
    szSQL = "Select  SUM(AL.VatAmount)  as AMT from tlbPayment P,PayTransactionsSplit AL,tlbPaymentSplit S,Fund F,Supplier SP,Property PR,GLobalData G  where  G.PropertyID=PR.PropertyID AND isAgentToSubmit=true  AND " & _
       "AL.TransactionID=S.PayTransactionIDSplit and SP.SupplierID=P.SageAccountNumber AND AL.Deleteflag=false and PR.propertyID=p.UNITID AND P.TransactionID=S.PayHeader AND P.TYPE IN(8,9,24) AND S.ClientStatementPrevID=" & idToinclude & " AND " & _
       "S.FundID=F.FundID AND AL.FromTran=P.transactionID " & whereProperty & " AND P.BankCODE='" & szSelectedBankAccount & "'  AND P.ClientID ='" & szSelectedClient & "' " & _
       "AND P.PDate <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "#"
        rsPayment.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
        If Not rsPayment.EOF Then
                GetSupplierPaymentPreview = GetSupplierPaymentPreview + IIf(IsNull(rsPayment.Fields.Item("AMT").Value), 0, rsPayment.Fields.Item("AMT").Value)
                'result is -15
        End If
        rsPayment.Close
        Set rsPayment = Nothing
         GetSupplierPaymentPreview = GetSupplierPaymentPreview - dblAmt
    adoConn.Close
    Set adoConn = Nothing
End Function
Private Function GetTenantReceipts() As Double
    Dim rsPayment As New ADODB.Recordset
    Dim szSQL As String
    Dim adoConn As New ADODB.Connection
    Dim rsReceipt As New ADODB.Recordset
    Dim whereProperty As String
    If boolConsolidatedStatement = 1 Then
            whereProperty = "(U.PropertyID in (" & ListOfProperties & ") OR isnull(U.PropertyID)) AND "
    Else
            whereProperty = "U.PropertyID  in (" & ListOfProperties & ") AND "
    End If
    
    adoConn.Open getConnectionString
    szSQL = "Select  SUM(SWITCH(R.TYPE=23,-RS.Amount,R.TYPE=3,RS.Amount,R.TYPE=4,RS.Amount)) as AMT from tlbReceipt R,tlbReceiptSplit RS,Fund F,Units U where " & _
            "R.TransactionID=RS.RptHeader AND R.TYPE IN(3,4,23) AND " & _
            "R.RDate >#" & Format(txtLastStatementDate1.text, "dd/mmm/yyyy") & "# AND R.RDate <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "# " & _
            "AND RS.FundID=F.FundID AND  R.BankCODE='" & szSelectedBankAccount & "'  AND (R.RentSumStatement='' OR isnull(R.RentSumStatement)) and  F.FundCode in (" & _
             ListOfFunds & ") AND R.UnitID=U.UnitNumber AND " & whereProperty & " R.ClientID ='" & szSelectedClient & "' "
    
    rsPayment.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
    If Not rsPayment.EOF Then
        GetTenantReceipts = IIf(IsNull(rsPayment.Fields.Item("AMT").Value), 0, rsPayment.Fields.Item("AMT").Value)
    End If
    rsPayment.Close

                szSQL = "Select  SUM(AL.VatAmount) as DR from tlbReceipt R,rptTransactions AL,tlbReceiptSplit S,Fund F, Units U,GLobalData G where G.PropertyID=U.PropertyID AND G.isAgentToSubmit=true AND R.TransactionID= S.RptHeader AND TYPE IN(3,4,23) " & _
            "AND S.FundID=F.FundID AND R.BankCODE='" & szSelectedBankAccount & "' and F.FundCode<>'TENANTDEPOSIT'  AND (R.RentSumStatement='' or isnull(R.RentSumStatement)) AND ClientID ='" & szSelectedClient & "' " & _
            "AND R.UnitID=U.UnitNumber and AL.Deleteflag=false and AL.FromTran=R.TransactionID AND  F.FundCode in (" & ListOfFunds & ") AND " & whereProperty & " R.RDate >#" & Format(txtLastStatementDate1.text, "dd/mmm/yyyy") & "# AND R.RDate <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "#"
            rsReceipt.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
            If Not rsReceipt.EOF Then
                    GetTenantReceipts = GetTenantReceipts - IIf(IsNull(rsReceipt.Fields.Item("Dr").Value), 0, rsReceipt.Fields.Item("Dr").Value)
                     'result is 175
            End If
            rsReceipt.Close
            Set rsReceipt = Nothing
    adoConn.Close
    Set adoConn = Nothing
End Function
'Private Function GetTenantReceiptsFinalized(ByVal trtoinclude As Long) As Double
'    Dim rsPayment As New ADODB.Recordset
'    Dim szSQL As String
'    Dim adoConn As New ADODB.Connection
'    Dim rsReceipt As New ADODB.Recordset
'    Dim whereProperty As String
'    If boolConsolidatedStatement = 1 Then
'            whereProperty = "(U.PropertyID in (" & ListOfProperties & ") OR isnull(U.PropertyID)) AND "
'    Else
'            whereProperty = "U.PropertyID  in (" & ListOfProperties & ") AND "
'    End If
'
'    adoConn.Open getConnectionString
'    szSQL = "Select  SUM(SWITCH(R.TYPE=23,-RS.Amount,R.TYPE=3,RS.Amount,R.TYPE=4,RS.Amount)) as AMT from tlbReceipt R,tlbReceiptSplit RS,Fund F,Units U where " & _
'            "R.TransactionID=RS.RptHeader AND R.TYPE IN(3,4,23) AND " & _
'            "R.RDate >#" & Format(txtLastStatementDate1.text, "dd/mmm/yyyy") & "# AND R.RDate <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "# " & _
'            "AND RS.FundID=F.FundID AND  R.BankCODE='" & szSelectedBankAccount & "'  AND (R.RentSumStatement=''  OR R.RentSumStatement=" & trtoinclude & " OR isnull(R.RentSumStatement)) and  F.FundCode in (" & _
'             ListOfFunds & ") AND R.UnitID=U.UnitNumber AND " & whereProperty & " R.ClientID ='" & szSelectedClient & "' "
'
'    rsPayment.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'    If Not rsPayment.EOF Then
'        GetTenantReceiptsFinalized = IIf(IsNull(rsPayment.Fields.Item("AMT").Value), 0, rsPayment.Fields.Item("AMT").Value)
'    End If
'    rsPayment.Close
'
'                szSQL = "Select  SUM(AL.VatAmount) as DR from tlbReceipt R,rptTransactions AL,tlbReceiptSplit S,Fund F, Units U,GLobalData G where G.PropertyID=U.PropertyID AND G.isAgentToSubmit=true AND R.TransactionID= S.RptHeader AND TYPE IN(3,4,23) " & _
'            "AND S.FundID=F.FundID AND R.BankCODE='" & szSelectedBankAccount & "' and F.FundCode<>'TENANTDEPOSIT'  AND (R.RentSumStatement=''  OR R.RentSumStatement=" & trtoinclude & " or isnull(R.RentSumStatement)) AND ClientID ='" & szSelectedClient & "' " & _
'            "AND R.UnitID=U.UnitNumber and AL.Deleteflag=false and AL.FromTran=R.TransactionID AND  F.FundCode in (" & ListOfFunds & ") AND " & whereProperty & " R.RDate >#" & Format(txtLastStatementDate1.text, "dd/mmm/yyyy") & "# AND R.RDate <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "#"
'            rsReceipt.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'            If Not rsReceipt.EOF Then
'                    GetTenantReceiptsFinalized = GetTenantReceipts - IIf(IsNull(rsReceipt.Fields.Item("Dr").Value), 0, rsReceipt.Fields.Item("Dr").Value)
'                     'result is 175
'            End If
'            rsReceipt.Close
'            Set rsReceipt = Nothing
'    adoConn.Close
'    Set adoConn = Nothing
'End Function
Private Function GetTenantReceiptsNonConsolidated(ByVal szPropertyID As String) As Double
    Dim rsPayment As New ADODB.Recordset
    Dim szSQL As String
    Dim adoConn As New ADODB.Connection
    Dim rsReceipt As New ADODB.Recordset
    Dim whereProperty As String

    whereProperty = "U.PropertyID ='" & szPropertyID & "' AND "

    
    adoConn.Open getConnectionString
    szSQL = "Select  SUM(SWITCH(R.TYPE=23,-RS.Amount,R.TYPE=3,RS.Amount,R.TYPE=4,RS.Amount)) as AMT from tlbReceipt R,tlbReceiptSplit RS,Fund F,Units U where " & _
            "R.TransactionID=RS.RptHeader AND R.TYPE IN(3,4,23) AND " & _
            "R.RDate >#" & Format(txtLastStatementDate1.text, "dd/mmm/yyyy") & "# AND R.RDate <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "# " & _
            "AND RS.FundID=F.FundID AND  R.BankCODE='" & szSelectedBankAccount & "'  AND (R.RentSumStatement='' OR isnull(R.RentSumStatement)) and  F.FundCode in (" & _
             ListOfFunds & ") AND R.UnitID=U.UnitNumber AND " & whereProperty & " R.ClientID ='" & szSelectedClient & "' "
    
            rsPayment.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
            If Not rsPayment.EOF Then
                GetTenantReceiptsNonConsolidated = IIf(IsNull(rsPayment.Fields.Item("AMT").Value), 0, rsPayment.Fields.Item("AMT").Value)
            End If
            rsPayment.Close

     szSQL = "Select  SUM(AL.VatAmount) as DR from tlbReceipt R,rptTransactions AL,tlbReceiptSplit S,Fund F, Units U,GLobalData G where G.PropertyID=U.PropertyID AND G.isAgentToSubmit=true AND R.TransactionID= S.RptHeader AND TYPE IN(3,4,23) " & _
            "AND S.FundID=F.FundID AND R.BankCODE='" & szSelectedBankAccount & "' and F.FundCode<>'TENANTDEPOSIT'  AND (R.RentSumStatement='' or isnull(R.RentSumStatement)) AND ClientID ='" & szSelectedClient & "' " & _
            "AND R.UnitID=U.UnitNumber and AL.Deleteflag=false and AL.FromTran=R.TransactionID AND  F.FundCode in (" & ListOfFunds & ") AND " & whereProperty & " R.RDate >#" & Format(txtLastStatementDate1.text, "dd/mmm/yyyy") & "# AND R.RDate <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "#"
            rsReceipt.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
            If Not rsReceipt.EOF Then
                    GetTenantReceiptsNonConsolidated = GetTenantReceiptsNonConsolidated - IIf(IsNull(rsReceipt.Fields.Item("Dr").Value), 0, rsReceipt.Fields.Item("Dr").Value)
                     'result is 175
            End If
            rsReceipt.Close
            Set rsReceipt = Nothing
    adoConn.Close
    Set adoConn = Nothing
End Function
'Private Function GetTenantReceiptsPreview(ByRef idToinclude As Long) As Double
'    Dim rsPayment As New ADODB.Recordset
'    Dim szSQL As String
'    Dim adoConn As New ADODB.Connection
'    Dim rsReceipt As New ADODB.Recordset
'    Dim whereProperty As String
'    If boolConsolidatedStatement = 1 Then
'            whereProperty = "(U.PropertyID in (" & ListOfProperties & ") OR isnull(U.PropertyID)) AND "
'    Else
'            whereProperty = "U.PropertyID  in (" & ListOfProperties & ") AND "
'    End If
'
'    adoConn.Open getConnectionString
'    szSQL = "Select  SUM(SWITCH(R.TYPE=23,-RS.Amount,R.TYPE=3,RS.Amount,R.TYPE=4,RS.Amount)) as AMT from tlbReceipt R,tlbReceiptSplit RS,Fund F,Units U where " & _
'            "R.TransactionID=RS.RptHeader AND R.TYPE IN(3,4,23) AND " & _
'            "R.RDate >#" & Format(txtLastStatementDate1.text, "dd/mmm/yyyy") & "# AND R.RDate <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "# " & _
'            "AND R.amount>R.OSamount AND RS.FundID=F.FundID AND  R.BankCODE='" & szSelectedBankAccount & "'  AND (R.RentSumStatement='' OR R.RentSumStatement='" & idToinclude & "' OR isnull(R.RentSumStatement)) and  F.FundCode in (" & _
'             ListOfFunds & ") AND R.UnitID=U.UnitNumber AND " & whereProperty & " R.ClientID ='" & szSelectedClient & "' "
'
'    rsPayment.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'    If Not rsPayment.EOF Then
'        GetTenantReceiptsPreview = IIf(IsNull(rsPayment.Fields.Item("AMT").Value), 0, rsPayment.Fields.Item("AMT").Value)
'    End If
'    rsPayment.Close
'
'            szSQL = "Select  SUM(AL.VatAmount) as DR from tlbReceipt R,rptTransactions AL,tlbReceiptSplit S,Fund F, Units U,GLobalData G where G.PropertyID=U.PropertyID AND G.isAgentToSubmit=true AND R.TransactionID= S.RptHeader AND TYPE IN(3,4,23) " & _
'            "AND R.amount>R.OSamount AND S.FundID=F.FundID AND R.BankCODE='" & szSelectedBankAccount & "' and F.FundCode<>'TENANTDEPOSIT'  AND (R.RentSumStatement='' OR R.RentSumStatement='" & idToinclude & "'  or isnull(R.RentSumStatement)) AND ClientID ='" & szSelectedClient & "' " & _
'            "AND R.UnitID=U.UnitNumber and AL.Deleteflag=false and AL.FromTran=R.TransactionID AND  F.FundCode in (" & ListOfFunds & ") AND " & whereProperty & " R.RDate >#" & Format(txtLastStatementDate1.text, "dd/mmm/yyyy") & "# AND R.RDate <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "#"
'            rsReceipt.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'            If Not rsReceipt.EOF Then
'                    GetTenantReceiptsPreview = GetTenantReceiptsPreview - IIf(IsNull(rsReceipt.Fields.Item("Dr").Value), 0, rsReceipt.Fields.Item("Dr").Value)
'                     'result is 175
'            End If
'            rsReceipt.Close
'            Set rsReceipt = Nothing
'    adoConn.Close
'    Set adoConn = Nothing
'End Function
'Private Function GetPaymentsonAccount() As Double
'    Dim rsPayment As New ADODB.Recordset
'    Dim adoConn As New ADODB.Connection
'    adoConn.Open getConnectionString
'    Dim szSQL As String
'    szSQL = "Select  SUM(S.Amount) as AMT from tlbPayment P,tlbPaymentSplit S,Fund F,Supplier SP where  SP.SupplierID=P.SAGEACCOUNTNUMBER AND " & _
'            "P.TransactionID=S.PayHeader AND P.PDate >#" & Format(txtLastStatementDate1.text, "dd/mmm/yyyy") & "# AND P.PDate <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "# AND P.TYPE  " & _
'            "IN(9) AND S.FundID=F.FundID AND  P.BankCODE='" & szSelectedBankAccount & "' and  F.FundCode in (" & _
'             ListOfFunds & ") AND SP.SupplierID ='" & szSelectedClient & "' AND (P.RentSumStatement='' OR isnull(P.RentSumStatement)) AND SP.Type in ('CLIENT','LLORD') "
'    rsPayment.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'    If Not rsPayment.EOF Then
'        GetPaymentsonAccount = IIf(IsNull(rsPayment.Fields.Item("AMT").Value), 0, rsPayment.Fields.Item("AMT").Value)
'    End If
'    rsPayment.Close
'    Set rsPayment = Nothing
'    adoConn.Close
'    Set adoConn = Nothing
'
'End Function




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
''    'frmRetentionDetails.Top = Me.Top + 600
    Dim rCount As Integer
''    Dim selRow As Integer
''    Dim iIncDec As Integer
''    For rCount = 1 To flxClients.Rows - 1
''         If flxClients.TextMatrix(rCount, 0) = "X" Then
''             iIncDec = iIncDec + 1
''             selRow = rCount
''         End If
''    Next
''    If iIncDec <> 1 Then
''       MsgBox "Please select a client.", vbInformation + vbOKOnly, "Client Selection"
''       Exit Sub
''    End If
''    frmRetentionDetails.szClienyIDForRetention = flxClients.TextMatrix(selRow, 1)
''    'frmRetentionDetails.ConfigflxRetensionDetailsMainMain
''    frmRetentionDetails.Top = Me.Top + 500
''    frmRetentionDetails.Show
''    If flxRetensionDetails.TextMatrix(1, 2) <> "" Then
''            Dim iCol As Integer
''            Dim iRow As Integer
''
''            'We are setting clientID in this form because form is listing the retention details for the client with cumulative balance
''            frmRetentionDetails.ConfigflxRetensionDetails
''            frmRetentionDetails.flxRetensionDetails.Cols = flxRetensionDetails.Cols
''            frmRetentionDetails.flxRetensionDetails.Rows = flxRetensionDetails.Rows
''            For iRow = 1 To flxRetensionDetails.Rows - 1
''                For iCol = 1 To flxRetensionDetails.Cols - 1
''                     frmRetentionDetails.flxRetensionDetails.TextMatrix(iRow, iCol) = flxRetensionDetails.TextMatrix(iRow, iCol)
''                Next
''            Next
''    End If
''    frmRetentionDetails.ZOrder 0
   LoadForm frmRetentionAdd
   
   For rCount = 1 To flxClients.Rows - 1
         If flxClients.TextMatrix(rCount, 0) = "X" Then
            frmRetentionAdd.txtClientList.text = flxClients.TextMatrix(rCount, 2)
            frmRetentionAdd.txtClientList.Tag = flxClients.TextMatrix(rCount, 1)
         End If
    Next
'   flxClients
   
End Sub

Private Sub cmdCalculateAvailableFund_Click()
    Dim rCount As Integer
    Dim iIncDec As Integer
    For rCount = 1 To flxClients.Rows - 1
         If flxClients.TextMatrix(rCount, 0) = "X" Then
             szSelectedClient = flxClients.TextMatrix(rCount, 1)
             iIncDec = iIncDec + 1
            ' selRow = rCount
         End If
    Next
    If iIncDec <> 1 Then
       MsgBox "Please select a client.", vbInformation + vbOKOnly, "Client Selection"
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
    Dim adoConn As New ADODB.Connection
    Dim szStatmentID As String
    Dim rsRentSummaryStatement As New ADODB.Recordset
    adoConn.Open getConnectionString
    Dim dblLasClosingBalance As Double
    Dim szSQL As String
    'Before writing this table you need to delete this table
    adoConn.Execute "Delete from  RentSummaryStatementPreview"
    
    'new Coding
    szStatmentID = GetLastStatementID + 1
    szSQL = "Select StatementClosingBal from RentSummaryStatement where StatementNo=" & GetLastStatementNoByClient & " AND ClientIDLandlordID='" & szSelectedClient & "'"
    rsRentSummaryStatement.Open szSQL, adoConn, adOpenKeyset, adLockReadOnly
    If Not rsRentSummaryStatement.EOF Then
        dblLasClosingBalance = rsRentSummaryStatement!StatementClosingBal
    End If
    rsRentSummaryStatement.Close
    Set rsRentSummaryStatement = Nothing
     adoConn.Close
    Set adoConn = Nothing
    'txtAvailableFunds.text = Format(getAvailablefunds(dblLasClosingBalance)
    Call MarkAllTransactionsWithCSIDPreview(szStatmentID)
    txtAvailableFunds.text = Format(getAvailablefundsPreview(dblLasClosingBalance, szStatmentID), "0.00")
    txtRentPayable.text = txtAvailableFunds.text
    
End Sub


Private Sub cmdClose_Click(Index As Integer)
         Unload Me
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

Private Function GetControlAccountForPayableString(adoConn As ADODB.Connection, szSelectedPayableTypeID As String) As Boolean
    Dim rsPayableTypes As New ADODB.Recordset
    rsPayableTypes.Open "Select * from  PayableTypes where ID=" & szSelectedPayableTypeID & "", adoConn, adOpenStatic, adLockReadOnly
    If Not rsPayableTypes.EOF Then
            GetControlAccountForPayableString = rsPayableTypes!PayNCAmt 'PayNCAmt
    End If
    rsPayableTypes.Close
    Set rsPayableTypes = Nothing

End Function

'Private Sub cmdFinalizeStatement_Click()
'        Dim connn As New ADODB.Connection
''        If validateIsFinalised = False Then
''          ' Exit Sub
''        End If
'        Dim szSelectedClient As String
'        connn.Open getConnectionString
'        Dim szCurrentStatementIDID As String
'        Dim rCount, iIncDec, selRow As Integer
'               For rCount = 1 To frmRentPayable.flxPayFees.Rows - 1
'                    If frmRentPayable.flxPayFees.TextMatrix(rCount, 0) = "X" Then
'                        iIncDec = iIncDec + 1
'                        selRow = rCount
'                    End If
'               Next
'               szCurrentStatementIDID = frmRentPayable.flxPayFees.TextMatrix(selRow, 2)
'               szCurrentStatementID = Replace(szCurrentStatementIDID, "CS", "")
'        bEditMode = True
'        whichFieldToCheck = "RentSumStatement"
'        If bEditMode = True Then
'            If PIvalidation = False Then
'                MsgBox "Please select a statement at header level", vbInformation, "Warning"
'                Exit Sub
'            End If
'        End If
'        If txtLastStatementDate1.text = "" Then
'               txtLastStatementDate1.Locked = False
'               If szStatementNo = 1 Then GoTo XX ' you can keep empty for the first statement date else you need to must enter date
'               MsgBox "Please enter last statement date", vbInformation, "Warning!"
'               Exit Sub
'        Else
'               If DateDiff("d", txtStatementDate1.text, txtLastStatementDate1.text) >= 0 And bEditMode = False Then
'                   MsgBox "A statement already exists for this date. Please enter a date after the 'Last Statement Date'", vbInformation, "Statement Date!!!"
'                   Exit Sub
'               End If
'        End If
'
''        If Val(txtRentPayable.text) > Val(txtAvailableFunds.text) Then
''                MsgBox "Rent Payable amount cannot be greater than the Available funds", vbInformation, "Warning!"
''                Exit Sub
''        End If
'        If Trim(txtStatementDate1.text) = "" Then
'              MsgBox "Please enter statment ", vbInformation, "Statement Date!!!"
'              FocusControl txtStatementDate1
'                Exit Sub
'        End If
'         For rCount = 1 To flxClients.Rows - 1
'            If flxClients.TextMatrix(rCount, 0) = "X" Then
'              szSelectedClient = flxClients.TextMatrix(rCount, 1)
'            End If
'         Next rCount
'
'
'XX:
''    For rCount = 1 To flxClients.Rows - 1
''         If flxClients.TextMatrix(rCount, 0) = "X" Then
''             iIncDec = iIncDec + 1
''             selRow = rCount
''         End If
''    Next
''    If iIncDec <> 1 Then
''       MsgBox "Please select a client.", vbInformation + vbOKOnly, "Client Selection"
''       Exit Sub
''    End If
'
'
'    iIncDec = 0
'    If ListOfProperties = "" Then
'         MsgBox "Please select a Property", vbInformation, "Property!!!"
'         flxProperties.SetFocus
'         Exit Sub
'    End If
'
'    'Validation for Property
'    For rCount = 1 To flxProperties.Rows - 1
'         If flxProperties.TextMatrix(rCount, 0) = "X" Then
'             iIncDec = iIncDec + 1
'             selRow = rCount
'         End If
'    Next
'    If iIncDec < 1 Then
'       MsgBox "Please select a property.", vbInformation + vbOKOnly, "Property Selection"
'       Exit Sub
'    End If
'    If ListOfProperties = "" Then
'         MsgBox "Please select a Property", vbInformation, "Property!!!"
'         flxProperties.SetFocus
'         Exit Sub
'    End If
'    If ListOfFunds = "" Then
'         MsgBox "Please select a fund", vbInformation, "Fund!!!"
'         flxProperties.SetFocus
'         Exit Sub
'    End If
'    For rCount = 1 To flxBankAccounts.Rows - 1
'         If flxBankAccounts.TextMatrix(rCount, 0) = "X" Then
'            szSelectedBankAccount = flxBankAccounts.TextMatrix(rCount, 2)
'            Exit For
'         End If
'    Next
'    Dim szSelectedPropertyID As String
'    If szSelectedBankAccount = "" Then
'        MsgBox "Please select a Bank account", vbInformation, "Warning "
'        Exit Sub
'    End If
''    If isAnyTransactionAvailable = False Then
''                MsgBox "There are no transactions for the statement period selected", vbInformation, "Warning!"
''                Exit Sub
''    End If
'
'    Dim szStatmentID As String
'    Dim szReportGenID As String
'    szReportGenID = ReportGenID
'    For rCount = 1 To frmRentPayable.flxPayFees.Rows - 1
'         If frmRentPayable.flxPayFees.TextMatrix(rCount, 0) = "X" Then
'            szSelectedStatement = frmRentPayable.flxPayFees.TextMatrix(rCount, 2)
'            Exit For
'         End If
'    Next
'    szStatmentID = Replace(szSelectedStatement, "CS", "")
'    If szStatmentID = "" Then Exit Sub
'    If MsgBox("Are you sure, you want to save this statement?", vbYesNo, "Please confirm") = vbYes Then
'
'        whichFieldToCheck = "RentSumStatement"
'        If boolConsolidatedStatement = 1 Then
'            Call GenerateSummaryStatementFinalized(szStatmentID, szReportGenID)  'Write into SummaryStatement table in this function
'            Call MarkAllTransactionsWithSS(szStatmentID)
'        Else        'For each property write a statement for non consolidated
''            For rCount = 1 To flxProperties.Rows - 1
''                 If flxProperties.TextMatrix(rCount, 0) = "X" Then
''                     szSelectedPropertyID = flxProperties.TextMatrix(rCount, 1)
''                     Call GenerateSummaryStatementNonConsolidated(szStatmentID, szSelectedPropertyID, szReportGenID)
''                     szStatmentID = CLng(szStatmentID) + 1
''                 End If
''             Next
'        End If
'
'
'        'run TestReportForRentSummary.rpt
'
'        Dim reportApp As New CRAXDRT.Application
'        Dim Report As CRAXDRT.Report
'
'        Set Report = reportApp.OpenReport(App.Path & szReportPath & "\RentSummaryStatement.rpt")
'        Report.EnableParameterPrompting = False
'        Report.DiscardSavedData
'        Report.ParameterFields(1).AddCurrentValue CLng(szReportGenID)
'        Load frmReport
'        frmReport.LoadReportViewer Report
''         Dim szCurrentStatementIDID As String
''            For rCount = 1 To frmRentPayable.flxPayFees.Rows - 1
''                 If frmRentPayable.flxPayFees.TextMatrix(rCount, 0) = "X" Then
''                     iIncDec = iIncDec + 1
''                     selRow = rCount
''                 End If
''            Next
''            szCurrentStatementIDID = frmRentPayable.flxPayFees.TextMatrix(selRow, 2)
''            If selRow > 0 Then
''                Call printClientStatement(szCurrentStatementIDID, selRow)
''            End If
'            Call frmRentPayable.loadflxPayFees("")
'    End If
'    If bEditDone Then
'       connn.Execute "Update rentSummarystatement set isfinalized=1 where statementID=" & szCurrentStatementID & ""
'       'msgbox""
'    End If
'    connn.Close
'End Sub
Private Function validateIsFinalised() As Boolean
     Dim iIncDec As Long
    iIncDec = 0
    Dim rCount As Integer
    Dim selRow As Integer
    Dim isitPlus As Boolean
    For rCount = 1 To frmRentPayable.flxPayFees.Rows - 1
         If frmRentPayable.flxPayFees.TextMatrix(rCount, 0) = "X" Then
             If frmRentPayable.flxPayFees.TextMatrix(rCount, 1) = "+" Then
                isitPlus = True
             Else
                isitPlus = False
             End If
             iIncDec = iIncDec + 1
             selRow = rCount
         End If
    Next
    If iIncDec < 1 Then
       MsgBox "Please select one statement only.", vbInformation + vbOKOnly, "statement Selection"
       Exit Function
    End If
'    If szCurrentStatementID = "" Then
'         Exit Function
'    End If
    'flxPayFees.TextMatrix(i, 3) is the statement ID by client
    '66) It should only be possible to modify a statement provided a rent payable
    ' PI has not been generated against the statement and a subsequent statement has not been produced.
    If isitPlus = True And frmRentPayable.flxPayFees.TextMatrix(selRow, 29) = "" Then
        MsgBox "This statement cannot be finalised, because a Rent Payable invoice " & frmRentPayable.flxPayFees.TextMatrix(selRow, 29) & " has not been generated against it.", vbInformation + vbOKOnly, "statement Selection"
        Exit Function
    ElseIf isitPlus = False Then 'when you selected"-" it wont let you modify
        MsgBox "Please select one statement to modify Rent Summary Statement", vbInformation + vbOKOnly, "statement Selection"
        Exit Function
    End If
    validateIsFinalised = True
End Function
Private Sub SelectLastStatementForthisClient() 'this is not by client
    Dim intmaxStatementNo As Integer
    Dim adoConn As New ADODB.Connection
    Dim rsRentSummaryStatement As New ADODB.Recordset
    adoConn.Open getConnectionString
    Dim szSQL As String
    Dim LastStatementID As String
    'Rent Summary Statement Opening Balance (=Closing balance of previous statement)
    'This is by client
    'Get ID by Client max ID from RentSummaryStatement
    szSQL = "Select max(StatementID) as IDbyCL from RentSummaryStatement where ClientIDLandlordID='" & szSelectedClient & "'"
    rsRentSummaryStatement.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
    If Not rsRentSummaryStatement.EOF Then
        LastStatementID = IIf(IsNull(rsRentSummaryStatement!IDbyCL), 0, rsRentSummaryStatement!IDbyCL)
    End If
    rsRentSummaryStatement.Close
    Set rsRentSummaryStatement = Nothing
    adoConn.Close
    Set adoConn = Nothing
    Dim iCount As Integer
    For iCount = frmRentPayable.flxPayFees.Rows - 1 To 1 Step -1
         If Replace(frmRentPayable.flxPayFees.TextMatrix(iCount, 2), "CS", "") = LastStatementID Then
                frmRentPayable.flxPayFees.TextMatrix(iCount, 0) = "X"
                Exit For
         End If
    Next
    
End Sub
Private Function isClientCondolidated(szSelectedClient As String) As Boolean
    Dim adoConn As New ADODB.Connection
    Dim rsRentSummaryStatement As New ADODB.Recordset
    adoConn.Open getConnectionString
    Dim szSQL As String
    Dim LastStatementID As String
    szSQL = "Select * from client where ConsolidatedStatement<>1 and clientID='" & szSelectedClient & "'"
    rsRentSummaryStatement.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
    If Not rsRentSummaryStatement.EOF Then
        isClientCondolidated = True
    End If
    rsRentSummaryStatement.Close
    Set rsRentSummaryStatement = Nothing
    adoConn.Close
    Set adoConn = Nothing
End Function

Private Sub cmdFix_Click()
    Dim adoConn As New ADODB.Connection
    If MsgBox("Do you wish to run the fix?", vbYesNo, "Yes/No") = vbNo Then Exit Sub
            adoConn.Open getConnectionString
            adoConn.Execute "Update RetentionDetails set StatementID=null ,SlNumber=null  where StatementID=148"
            MsgBox "Completeted"
            adoConn.Close
End Sub

Private Sub cmdOKInouts_Click()
    'validaton
    'at least one bank , one fund, one property ,one payable type is selected.
    'this procedure shall make visible only that
    If Val(txtRentPayable.text) > Val(txtAvailableFunds.text) Then
        MsgBox "Rent Payable amount cannot be greater than the Available funds", vbInformation, "Warning!"
        Exit Sub
    End If
    If txtLastStatementDate1.text = "" And szStatementNo = 1 Then
        txtLastStatementDate1.Locked = False
        txtLastStatementDate1.text = "01/01/2000"
        MsgBox "Please enter last statement date", vbInformation, "Warning!"
        FocusControl txtLastStatementDate1
        Exit Sub
    End If
    'for the first time they can modify the last statement date
    'if szStatementNo is 1 that means it is first statement for the client
'    If bEditMode = True Then
            If txtLastStatementDate1.text = "" Then
                    txtLastStatementDate1.Locked = False
                    If szStatementNo = 1 Then GoTo XX ' you can keep empty for the first statement date else you need to must enter date
                    MsgBox "Please enter last statement date", vbInformation, "Warning!"
                    
                    Exit Sub
             Else
                    If DateDiff("d", txtStatementDate1.text, txtLastStatementDate1.text) >= 0 And bEditMode = False Then
                        MsgBox "A statement already exists for this date. Please enter a date after the 'Last Statement Date'", vbInformation, "Statement Date!!!"
                        Exit Sub
                    End If
             End If
XX:
    'Validation for Client
    Dim rCount As Integer
    Dim selRow As Integer
    Dim iIncDec As Long
    Dim szSelectedClient As String
    For rCount = 1 To flxClients.Rows - 1
         If flxClients.TextMatrix(rCount, 0) = "X" Then
             szSelectedClient = flxClients.TextMatrix(rCount, 1)
             iIncDec = iIncDec + 1
             selRow = rCount
         End If
    Next
    If iIncDec <> 1 Then
       MsgBox "Please select a client.", vbInformation + vbOKOnly, "Client Selection"
       Exit Sub
    End If
    If isClientCondolidated(szSelectedClient) Then
        MsgBox "Please enable consolidated statements for this client on the client record", vbInformation, "Warning"
        Exit Sub
    End If
    
    iIncDec = 0
    If ListOfProperties = "" Then
         MsgBox "Please select a Property", vbInformation, "Property!!!"
         flxProperties.SetFocus
         Exit Sub
    End If
    If ListOfFunds = "" Then
         MsgBox "Please select a fund", vbInformation, "fund!!!"
         flxInFunds.SetFocus
         Exit Sub
    End If
    'Validation for Property
    iIncDec = 0
    For rCount = 1 To flxProperties.Rows - 1
         If flxProperties.TextMatrix(rCount, 0) = "X" Then
             iIncDec = iIncDec + 1
             selRow = rCount
         End If
    Next
'    If iIncDec <> 1 And boolConsolidatedStatement = 0 Then
'       MsgBox "Please select only one property.", vbInformation + vbOKOnly, "Property Selection"
'       Exit Sub
'    End If
Rem by anol 2023-04-20
'   If ComparePreviousStatementFinalized = True Then
'         MsgBox "There is previous client statement that is showing as not having been finalised as at the statement date selected. " & _
'                "Please finalise this previous client statement or select a statement date on or after your previous statement was finalised", vbInformation, "Warning!"
'        Exit Sub
'   End If
   If GetSupplierOSAmountAllTime <> 0 Then
        If MsgBox("You have outstanding supplier balances to pay. Do you wish to pay them before previewing your client statement?", vbYesNo, "Supplier Os Balance") = vbYes Then
                LoadForm frmPurchaseExpense
                frmPurchaseExpense.tabPurExp.Tab = 1
                frmPurchaseExpense.txtClientIDPurPay.text = szSelectedClient
                frmPurchaseExpense.txtBankCode.text = ""
                frmPurchaseExpense.txtBankAc.text = ""
                Exit Sub
        Else
            'proceed
        End If
    End If
    If PreviousStatementFinalized = True Then
        MsgBox "There are previous client statements that have not been finalised for this client. " & _
                "You must finalise all previous client statements before proceeding.", vbInformation, "Warning!"
        Exit Sub
    End If
    If GetClientOSAmount > 0 Then
        MsgBox "There is previous client statement that has not been finalised at the statement date selected. Please select a statement date on or after your last statement was finalised", vbInformation, "Client Os Balance"
        Exit Sub
'        If MsgBox("There is previous client statement that has not been finalised at the statement date selected. Please select a statement date on or after your last statement was finalised", vbInformation, "Client Os Balance") = vbYes Then
'                LoadForm frmPurchaseExpense
'                frmPurchaseExpense.tabPurExp.Tab = 1
'                frmPurchaseExpense.txtClientIDPurPay.text = szSelectedClient
'                frmPurchaseExpense.txtBankCode.text = ""
'                frmPurchaseExpense.txtBankAc.text = ""
'                Exit Sub
'        Else
'            'proceed
'        End If
    End If
    If GetAgentOSAmount > 0 Then
        If MsgBox("You have outstanding Managing Agent balances to pay. Do you wish to pay them before previewing your client statement?", vbYesNo, "Managing Agent Os Balance") = vbYes Then
                LoadForm frmPurchaseExpense
                frmPurchaseExpense.tabPurExp.Tab = 1
                frmPurchaseExpense.txtClientIDPurPay.text = szSelectedClient
                frmPurchaseExpense.txtBankCode.text = ""
                frmPurchaseExpense.txtBankAc.text = ""
                Exit Sub
        Else
            'proceed
        End If
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
    
        MsgBox "Please select a Bank account", vbInformation, "Warning!"
        Exit Sub
    End If
    
    'You have management fees to generate. Do you wish to generate them before producing your client statement
    'check if management fee has been generated or not
    'If Feestogenerate Then   new update on 20230205
    If GeneratePIPreview Then
        
        If chkShowDue.Value = 0 Then
            'Now delete mangagement fee preview because user dont want to see it in the report
            Dim adocon As New ADODB.Connection
            adocon.Open getConnectionString
            adocon.Execute "Delete from tblPurInvPreview"
            adocon.Execute "Delete from ManagementFeePreview"
            adocon.Close
            If MsgBox("You have management fees to generate. Do you wish to generate  " & _
                "them before producing your client statement?", vbYesNo, "Please confirm") = vbYes Then
                        LoadForm frmManagementFees
                        Exit Sub
            Else
                'proceed
            End If
        Else
                'Show Management Fees preview here
               'Call showManagementFeeReport
        End If
    End If
    'check if management fee has been paid or not
    'rem by anol 08-08-2023 this check should only be on produce
''    If FeeshasBeenPaid Then
'''        If chkShowDue.Value = 1 Then
'''        Else
''            If MsgBox("You have management fees to pay. Do you wish to pay  " & _
''                "them before producing your client statement?", vbYesNo, "Please confirm") = vbYes Then
''                LoadForm frmPurchaseExpense
''                frmPurchaseExpense.tabPurExp.Tab = 1
''                frmPurchaseExpense.txtClientIDPurPay.text = szSelectedClient
''                frmPurchaseExpense.txtBankCode.text = ""
''                frmPurchaseExpense.txtBankAc.text = ""
''                Exit Sub
''            Else
''                'proceed
''            End If
'''       End If
''    End If
    
'    If ListOfPayableTypes = "" Then
'        MsgBox "Please select a Payable Type", vbInformation, "Warning!"
'        Exit Sub
'    End If
'        If isAnyTransactionAvailable = False Then
'                MsgBox "There are no transactions for the statement period selected", vbInformation, "Warning!"
'                Exit Sub
'        End If
            Dim szStatmentID As String
               'Dim iIncDec As Long
   iIncDec = 0
   'Dim rCount As Integer
'   Dim selRow As Integer
   'Dim iIncDec As Long
   
   
   iIncDec = 0
   'Dim rCount As Integer
   'Dim selRow As Integer
   Dim adoConn As New ADODB.Connection
   For rCount = 1 To frmRentPayable.flxPayFees.Rows - 1
        If frmRentPayable.flxPayFees.TextMatrix(rCount, 0) = "X" Then
            iIncDec = iIncDec + 1
            selRow = rCount
        End If
   Next
  

   
           '
           
           whichFieldToCheck = "ClientStatementPrevID"
            
'           If frmRentPayableNew.Caption = "Modify Rent Summary Statement" Then
'                szStatmentID = frmRentPayable.flxPayFees.TextMatrix(selRow, 2)
'                szStatmentID = Replace(szStatmentID, "CS", "")
'           Else
                szStatmentID = GetLastStatementID + 1
'           End If
        adoConn.Open getConnectionString
        Dim rsRentPayable As New ADODB.Recordset
        Dim szSQL As String
        szSQL = "Select P.* from tlbPayment P,tlbPaymentSplit S,Fund F,Supplier SP,PayTransactions AL,tlbPayment AP,tblPurINV V where " & _
        "P.PDate <=#" & Format(txtLastStatementDate1.text, "dd/mmm/yyyy") & "# AND P.BankCODE='" & szSelectedBankAccount & "' AND AL.FromTran=P.TransactionID  AND AL.ToTran=AP.TransactionID AND " & _
        "SP.SupplierID=P.SageaccountNumber and  P.Amount>P.OSAmount  AND  P.TransactionID=S.PayHeader  AND P.TYPE IN(7,8,9) AND S.FundID=F.FundID AND F.FundCode in (" & _
         ListOfFunds & ") AND  isnull(S.ClientStatementID) AND AP.PI=V.MY_ID AND V.isRentPayable=true  AND P.ClientID ='" & szSelectedClient & "' "
         
        rsRentPayable.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
        If Not rsRentPayable.EOF Then
            If MsgBox("There is Rent Payable Paid relating to the previous Statement for '" & _
                    szSelectedClient & "' that has not been finalised. Do you wish to re-finalise this Statement?", vbYesNo, "Warning") = vbYes Then
                    'Select last statement for this client then load finalized form
                    Call SelectLastStatementForthisClient
                    frmRentPayable.cmdPreViewGenDmds_Click
                Exit Sub
            End If
        End If
        rsRentPayable.Close
             
            Call MarkAllTransactionsWithCSIDPreview(szStatmentID)
            Call GenerateCSPreview(szStatmentID)
            Call PrintCSlineByLineNewPreview(szStatmentID) 'New update from 2023/01/23 which saves time
            Debug.Print time & " End"

End Sub


Private Function showManagementFeeReport()
    On Error GoTo Err
     Sleep (100)
     Dim rCount As Integer
     Dim szPropertySelectionALL As String
      For rCount = 1 To flxProperties.Rows - 1
                If flxProperties.TextMatrix(rCount, 0) = "X" Then
                    szPropertySelectionALL = szPropertySelectionALL + "," + flxProperties.TextMatrix(rCount, 1)
                    
                End If
        Next
        
                Dim reportApp As New CRAXDRT.Application
                Dim Report As CRAXDRT.Report
                Dim rep As frmReport
                Set Report = reportApp.OpenReport(App.Path & szReportPath & "\" & "ManagementFeeDue.rpt")
                Report.Database.Tables(1).ConnectionProperties.Item("Database Password") = accessDBPws
                Report.DiscardSavedData
                Report.EnableParameterPrompting = False
                
                Report.ParameterFields(1).AddCurrentValue szSelectedClient
                Report.ParameterFields(2).AddCurrentValue Replace(szPropertySelectionALL, ",", "   ")
'                Report.ParameterFields(3).AddCurrentValue warning1
'                Report.ParameterFields(4).AddCurrentValue warning2
'                Report.ParameterFields(5).AddCurrentValue warning3
                
                Set rep = New frmReport
                Load rep
                rep.LoadReportViewer Report
                Exit Function
Err:
                MsgBox Err.description
End Function
Private Sub printCSlinebyLinePreview(ByVal CSID As String)
    Dim iIncDec As Long
    iIncDec = 0
    Dim rCount As Integer
    Dim selRow As Integer
    'Dim CSID As String
    Dim adoConn As New ADODB.Connection
   'Here I need to pass csID which was update in previous sub routine

    Dim reportApp As New CRAXDRT.Application
    Dim Report As CRAXDRT.Report
    adoConn.Open getConnectionString
    Dim rsDemandSplit As New ADODB.Recordset
    Dim rsReceived As New ADODB.Recordset
    Dim rsRentSummaryStatement As New ADODB.Recordset
    Dim dblReceivedAmt As Double
    Dim dblCrReceivedAmt As Double
    Dim dblOSAmount As Double
    Dim szListofFunds As String
    Dim szTypeOfDemanddesc As String
    Dim dateFrom As String
    Dim rsDemandSplitAmt As New ADODB.Recordset
    Dim dblDemandSplitamt As Double
    Debug.Print time & " start1"
    Dim DateTO As String
    Dim szSelectedClient As String
    For rCount = 1 To flxClients.Rows - 1
         If flxClients.TextMatrix(rCount, 0) = "X" Then
             szSelectedClient = flxClients.TextMatrix(rCount, 1)
             Exit For
         End If
    Next
    If ListOfFunds = "" Then
                   ' MsgBox "Please select a fund"
                    Exit Sub
     End If
    'Exit Sub
    adoConn.Execute "Update DemandSplitRecords DS,DemandRecords D,Units U,Property P  set ReportNetAmountS=0,ReportVATAmountS=0,ReportReceivedAmountS= 0,ReportDateFromS=Null," & _
                " ReportCreditAmountS=0, ReportCsShowFlag= '',ReportDateTOS =null,ReportDemandTypeDescS= '',reportOSAmountS=0 where D.DemandID=DS.DemandID  and " & _
                "U.UnitNumber=D.UnitNumber AND P.PropertyID=U.PropertyID AND P.ClientID='" & _
                szSelectedClient & "'"
     Debug.Print time & " start ss1"
    Dim strDueDate As String
   
    'for type 1- SI
    rsDemandSplit.Open "Select D.DemandId,D.TransactionType,SplitID, sum(Amount) as NAmt,sum(VATAmount)as TVAT from  DemandSplitRecords DS,DemandRecords D,Units U,Property P where D.DemandID=DS.DemandID " & _
                " and D.TransactionType=1 AND U.UnitNumber=D.UnitNumber AND P.PropertyID=U.PropertyID   AND P.ClientID='" & szSelectedClient & "' group by  " & _
                " D.DemandId,D.TransactionType,SplitID", adoConn, adOpenStatic, adLockReadOnly
'
    Dim rsDemandSplit1 As New ADODB.Recordset
    Dim dblCrSumAmt As Double
    Dim rsDemandSplitCredit As New ADODB.Recordset
    'for type 2- CR Note
    Dim iCount As Integer
'    rsDemandSplitCredit.Open "Select R.TransactionID,D.DemandId,D.SplitID, Sum(RS.Amount) as NAmt from  tlbReceipt R,tlbReceiptSplit RS, DemandSplitRecords D where D.SplitID=Rs.SplitID AND RS.rptHeader=R.transactionID " & _
'                " AND R.DemandRef=D.DemandID AND Type=2 AND R.OSAmount=0 AND R.ClientID='" & _
'                 szSelectedClient & "' group by R.TransactionID,D.DemandId,D.SplitID", adoConn, adOpenStatic, adLockReadOnly
    'issue
    rsDemandSplitCredit.Open "Select R.TransactionID,SPLITID,D.DemandId from  DemandSplitRecords D,DemandRecords DR,tlbReceipt R where  D.DemandID=DR.DemandID AND " & _
                " R.DemandRef=D.DemandID AND Type=2 AND R.OSAmount=0 AND R.ClientID='" & _
                 szSelectedClient & "' ", adoConn, adOpenStatic, adLockReadOnly
    Debug.Print time & " start ss2"
    While Not rsDemandSplitCredit.EOF
                 dateFrom = ""
                 DateTO = ""
                 strDueDate = ""
'                 FromTran = 1461
                 rsDemandSplit1.Open "Select DateFrom,DateTo,Type,DueDate from  DemandSplitRecords D,DemandTypes DT where DT.ID=D.TypeOfDemand AND DemandId=" & _
                 rsDemandSplitCredit("DEMANDID").Value & " order by D.SPlitID ", adoConn, adOpenStatic, adLockReadOnly
                 If Not rsDemandSplit1.EOF Then
                        dateFrom = rsDemandSplit1("DateFrom").Value
                        DateTO = rsDemandSplit1("DateTo").Value
                        szTypeOfDemanddesc = rsDemandSplit1("Type").Value
                        strDueDate = rsDemandSplit1("DueDate").Value
                 End If
                 rsDemandSplit1.Close
                 '*************************
                 rsDemandSplitAmt.Open "Select Count(amount) as CNT from DemandSplitRecords DS where  DS.DemandId=" & _
                            rsDemandSplitCredit("DEMANDID").Value & "", adoConn, adOpenStatic, adLockReadOnly
                            
                 If Not rsDemandSplitAmt.EOF Then
                             iCount = rsDemandSplitAmt("CNT").Value
                 End If
                 rsDemandSplitAmt.Close
                 
                 
                 rsDemandSplitAmt.Open "Select amount from DemandSplitRecords DS where  DS.SPLITID=" & rsDemandSplitCredit("SplitID").Value & "  and DS.DemandId=" & _
                            rsDemandSplitCredit("DEMANDID").Value & "", adoConn, adOpenStatic, adLockReadOnly
                            
                 If Not rsDemandSplitAmt.EOF Then
                             dblDemandSplitamt = rsDemandSplitAmt("amount").Value
                 End If
                 dblOSAmount = dblDemandSplitamt
                 rsDemandSplitAmt.Close
                  
                       'when you are allocating SC with a SI then splitID you are getting that are split ID of a SI . so you cannot calculate anything here with SC SplitID
                       'Ideally there is only one split in the SC
                If iCount = 1 Then
                    rsDemandSplit1.Open "Select Sum(RS.Amount) as Amt from  RptTransactionsSplit D,tlbReceiptSplit RS,tlbReceipt RC where RC.TransactionID=RS.RptHeader AND " & _
                    "RS.RptTransactionsIDSplit=D.TransactionID AND D.Allocdate <=#" & Format(txtStatementDate1.text, "dd MMM yyyy") & "#  AND D.Allocdate > #" & Format(txtLastStatementDate1.text, "dd MMM yyyy") & "# AND FromTran=" & _
                    rsDemandSplitCredit("TransactionID").Value & " and deleteflag=false  ", adoConn, adOpenStatic, adLockReadOnly
                    If Not rsDemandSplit1.EOF Then
                           dblCrSumAmt = IIf(IsNull(rsDemandSplit1("Amt").Value), 0, rsDemandSplit1("Amt").Value)
                           dblOSAmount = dblOSAmount - dblCrSumAmt
                    End If
                    rsDemandSplit1.Close
                    If dblCrSumAmt > 0 Then
                            adoConn.Execute "Update DemandSplitRecords DS,Fund F set ReportCsShowFlag= '1' where DS.SageDepartment=F.FundID and F.FundCode in ( " & ListOfFunds & ") AND DemandId=" & _
                        rsDemandSplitCredit("DEMANDID").Value & " "
                    End If
                 Else
                        rsDemandSplit1.Open "Select Sum(RS.Amount) as Amt from  RptTransactionsSplit D,tlbReceiptSplit RS,tlbReceipt RC where RC.TransactionID=RS.RptHeader AND " & _
                        "RS.RptTransactionsIDSplit=D.TransactionID   AND FromTran=" & _
                        rsDemandSplitCredit("TransactionID").Value & " and SPlitIDofSi=" & rsDemandSplitCredit("SplitID").Value & " and deleteflag=false  group by D.SPlitIDofSi ", adoConn, adOpenStatic, adLockReadOnly
                        If Not rsDemandSplit1.EOF Then
                               dblCrSumAmt = rsDemandSplit1("Amt").Value
                               dblOSAmount = dblOSAmount - dblCrSumAmt
                        End If
                        rsDemandSplit1.Close
                        If dblCrSumAmt > 0 Then
                                adoConn.Execute "Update DemandSplitRecords DS,Fund F set ReportCsShowFlag= '1' where DS.SageDepartment=F.FundID and F.FundCode in ( " & ListOfFunds & ") AND DemandId=" & _
                            rsDemandSplitCredit("DEMANDID").Value & " "
                        End If
                 End If

                adoConn.Execute "Update DemandSplitRecords DS  set ReportNetAmountS=" & -dblDemandSplitamt & ",ReportVATAmountS=0,ReportCreditAmountS=" & _
                                -dblCrSumAmt & ",ReportReceivedAmountS= 0,ReportDateFromS=#" & dateFrom & "#," & _
                                " reportOSAmountS=" & -dblOSAmount & ",ReportDateTOS=#" & DateTO & "#,ReportDemandTypeDescS= '" & szTypeOfDemanddesc & "' where DemandId=" & _
                                rsDemandSplitCredit("DEMANDID").Value & " AND DS.SplitID =" & rsDemandSplitCredit("SplitID").Value & ""
            rsDemandSplitCredit.MoveNext
    Wend
    rsDemandSplitCredit.Close
    Debug.Print time & " start ss3"
    'for type 1- SI
    While Not rsDemandSplit.EOF
                'Income Side
                 dblReceivedAmt = 0
                 dblCrReceivedAmt = 0
                 dblOSAmount = 0
                 If rsDemandSplit("DemandId").Value = 128 Then
                    Debug.Print rsDemandSplit("transactionType").Value
                 End If
                 If rsDemandSplit("DEMANDID").Value = 204 Then
                    Debug.Print ""
                 End If
                 If rsDemandSplit("DEMANDID").Value = 26 Then
                    Debug.Print ""
                 End If
                 'SR SA income
                 rsReceived.Open "Select DemandId,sum(R.Amount) as ReceivedAmt from  tlbReceiptSplit R,tlbReceipt RC, RptTransactionsSplit T,DemandRecords D,tlbReceipt RL,FUND F where " & _
                 " T.Deleteflag=False AND RL.transactionID=T.FromTran AND T.transactionID=R.RptTransactionsIDSplit AND RC.transactionID=T.ToTran AND R.FundID=F.FundID and  " & _
                 "F.FundCode in ( " & ListOfFunds & ")AND R.ClientStatementPrevID=" & CSID & " AND RL.Type in(3,4) and RC.DemandRef=D.DemandID and D.DemandId=" & rsDemandSplit("DEMANDID").Value & " " & _
                 "AND  R.SPLITID=" & rsDemandSplit("SplitID").Value & " group by DemandID", adoConn, adOpenStatic, adLockReadOnly
                 dblReceivedAmt = 0
                 If Not rsReceived.EOF Then
                         dblReceivedAmt = IIf(IsNull(rsReceived("ReceivedAmt").Value), 0, rsReceived("ReceivedAmt").Value)
                 End If
                 rsReceived.Close
                 'SC income
                 rsReceived.Open "Select DemandId,sum(R.Amount) as ReceivedAmt from  tlbReceiptSplit R,tlbReceipt RC, RptTransactionsSplit T,DemandRecords D,tlbReceipt RL ,FUND F where T.Deleteflag=False " & _
                 "AND RL.transactionID=T.FromTran AND T.transactionID=R.RptTransactionsIDSplit AND RC.transactionID=T.ToTran AND R.FundID=F.FundID and F.FundCode in ( " & ListOfFunds & ") " & _
                 "AND  R.ClientStatementPrevID=" & CSID & " " & _
                 "AND RL.Type in(2) and RC.DemandRef=D.DemandID and D.DemandId=" & rsDemandSplit("DEMANDID").Value & " AND  R.SPLITID=" & rsDemandSplit("SplitID").Value & " group by DemandID", adoConn, adOpenStatic, adLockReadOnly
                 If Not rsReceived.EOF Then
                         dblCrReceivedAmt = IIf(IsNull(rsReceived("ReceivedAmt").Value), 0, rsReceived("ReceivedAmt").Value)
                 End If
                 rsReceived.Close
                 'OS amount writing code for collectiong the osamount 'osamount report field what ever be the CSID
                 rsReceived.Open "Select DemandId,sum(R.Amount) as ReceivedAmt from  tlbReceiptSplit R,tlbReceipt RC, RptTransactionsSplit T,DemandRecords D,tlbReceipt RL,Fund F where T.Deleteflag=False " & _
                 "AND RL.transactionID=T.FromTran AND T.transactionID=R.RptTransactionsIDSplit AND RC.transactionID=T.ToTran AND R.FundID=F.FundID " & _
                 "AND RL.RDate <=#" & Format(txtStatementDate1.text, "dd MMM yyyy") & "# AND  R.SPLITID=T.SplitIDofSI  AND  R.SPLITID=" & rsDemandSplit("SplitID").Value & "  " & _
                 "AND RL.Type in(2,3,4) and RC.DemandRef=D.DemandID and D.DemandId=" & rsDemandSplit("DEMANDID").Value & " group by DemandID", adoConn, adOpenStatic, adLockReadOnly
                 
                 If Not rsReceived.EOF Then
                         dblOSAmount = rsDemandSplit("NAmt").Value + rsDemandSplit("TVAT").Value - IIf(IsNull(rsReceived("ReceivedAmt").Value), 0, rsReceived("ReceivedAmt").Value)
                 End If
                 rsReceived.Close
                 
                 
                 dateFrom = ""
                 DateTO = ""
                 
                 rsDemandSplit1.Open "Select DateFrom,DateTo,Type from  DemandSplitRecords D,DemandTypes DT where DT.ID=D.TypeOfDemand AND DemandId=" & _
                        rsDemandSplit("DEMANDID").Value & " AND  D.SPLITID=" & rsDemandSplit("SplitID").Value & "  order by D.SPlitID ", adoConn, adOpenStatic, adLockReadOnly
                 If Not rsDemandSplit1.EOF Then
                        dateFrom = rsDemandSplit1("DateFrom").Value
                        DateTO = rsDemandSplit1("DateTo").Value
                        szTypeOfDemanddesc = rsDemandSplit1("Type").Value
                 End If
                 rsDemandSplit1.Close

                        
                  If dblReceivedAmt > 0 Then
                        adoConn.Execute "Update DemandSplitRecords DS set ReportNetAmountS= " & rsDemandSplit("NAmt").Value & ",ReportVATAmountS= " & _
                        rsDemandSplit("TVAT").Value & ", ReportOSAmountS =  " & dblOSAmount & ",ReportReceivedAmountS= " & dblReceivedAmt & ",ReportDateFromS=#" & Format(dateFrom, "dd MMM yyyy") _
                        & "#,ReportDateTOS =#" & Format(DateTO, "dd MMM yyyy") & "#,ReportDemandTypeDescS= '" & szTypeOfDemanddesc & "' where DemandId=" & _
                        rsDemandSplit("DEMANDID").Value & " AND SPLITID=" & rsDemandSplit("SplitID").Value & " "
                        
                        adoConn.Execute "Update DemandSplitRecords DS,Fund F set ReportCsShowFlag= '1' where DS.SageDepartment=F.FundID and F.FundCode in ( " & ListOfFunds & ") AND DemandId=" & _
                        rsDemandSplit("DEMANDID").Value & " "
                        dblReceivedAmt = 0
                  End If
                  
                  If dblCrReceivedAmt > 0 Then
                        adoConn.Execute "Update DemandSplitRecords DS set ReportNetAmountS= " & rsDemandSplit("NAmt").Value & ",ReportVATAmountS= " & _
                        rsDemandSplit("TVAT").Value & ", ReportOSAmountS =  " & dblOSAmount & ", ReportCreditAmountS= " & -dblCrReceivedAmt & ",ReportDateFromS=#" & Format(dateFrom, "dd MMM yyyy") _
                        & "#,ReportDateTOS =#" & Format(DateTO, "dd MMM yyyy") & "#,ReportDemandTypeDescS= '" & szTypeOfDemanddesc & "' where DemandId=" & _
                        rsDemandSplit("DEMANDID").Value & " AND SPLITID=" & rsDemandSplit("SplitID").Value & " "
                        dblCrReceivedAmt = 0
                        adoConn.Execute "Update DemandSplitRecords DS,Fund F set ReportCsShowFlag= '1' where DS.SageDepartment=F.FundID and F.FundCode in ( " & ListOfFunds & ") AND DemandId=" & _
                        rsDemandSplit("DEMANDID").Value & " "
                  End If
                    
                    
            rsDemandSplit.MoveNext
    Wend
    rsDemandSplit.Close
  Debug.Print time & " start ss4"
    'Now here you take care of those line split where invoice is partially paid/receipt . but some line is not which is not paid/receipt need to show the os amount
     rsDemandSplit.Open "Select D.DemandId,D.TransactionType,SplitID, sum(Amount) as NAmt,sum(VATAmount)as TVAT from  DemandSplitRecords DS,DemandRecords D,Units U,Property P where D.DemandID=DS.DemandID " & _
                " and U.UnitNumber=D.UnitNumber AND P.PropertyID=U.PropertyID   AND P.ClientID='" & szSelectedClient & "' AND ReportCsShowFlag= '1' AND ReportNetAmountS=0 group by  " & _
                " D.DemandId,D.TransactionType,SplitID", adoConn, adOpenStatic, adLockReadOnly

       While Not rsDemandSplit.EOF
                 dblOSAmount = 0
                 dblDemandSplitamt = 0
                 If rsDemandSplit("DEMANDID").Value = 159 Then
                    Debug.Print ""
                 End If
                 rsDemandSplitAmt.Open "Select amount from DemandSplitRecords DS where  DS.SPLITID=" & rsDemandSplit("SplitID").Value & "  and DS.DemandId=" & _
                            rsDemandSplit("DEMANDID").Value & "", adoConn, adOpenStatic, adLockReadOnly
                            
                 If Not rsDemandSplitAmt.EOF Then
                             dblDemandSplitamt = rsDemandSplitAmt("amount").Value
                 End If
                 rsDemandSplitAmt.Close
                 'when you find os amount, its not needed which CSID you filter
                 rsReceived.Open "Select DemandId,sum(R.Amount) as ReceivedAmt from  tlbReceiptSplit R,tlbReceipt RC, RptTransactionsSplit T,DemandRecords D,tlbReceipt RL,Fund F where T.Deleteflag=False " & _
                 "AND RL.transactionID=T.FromTran AND T.transactionID=R.RptTransactionsIDSplit AND RC.transactionID=T.ToTran AND R.FundID=F.FundID and F.FundCode in ( " & ListOfFunds & ") " & _
                 "AND RL.RDate <=#" & Format(txtStatementDate1.text, "dd MMM yyyy") & "# AND  R.SPLITID=" & rsDemandSplit("SplitID").Value & "  " & _
                 "AND RL.Type in(2,3,4) and RC.DemandRef=D.DemandID and D.DemandId=" & rsDemandSplit("DEMANDID").Value & " group by DemandID", adoConn, adOpenStatic, adLockReadOnly
                 dblOSAmount = dblDemandSplitamt
                 If Not rsReceived.EOF Then
                             dblOSAmount = dblOSAmount - IIf(IsNull(rsReceived("ReceivedAmt").Value), 0, rsReceived("ReceivedAmt").Value)
                 End If
                 rsReceived.Close
                 rsDemandSplit1.Open "Select DateFrom,DateTo,Type from  DemandSplitRecords D,DemandTypes DT where DT.ID=D.TypeOfDemand AND DemandId=" & _
                        rsDemandSplit("DEMANDID").Value & " AND  D.SPLITID=" & rsDemandSplit("SplitID").Value & "  order by D.SPlitID ", adoConn, adOpenStatic, adLockReadOnly
                 If Not rsDemandSplit1.EOF Then
                        dateFrom = rsDemandSplit1("DateFrom").Value
                        DateTO = rsDemandSplit1("DateTo").Value
                        szTypeOfDemanddesc = rsDemandSplit1("Type").Value
                 End If
                 rsDemandSplit1.Close

                 adoConn.Execute "Update DemandSplitRecords DS set ReportNetAmountS= " & rsDemandSplit("NAmt").Value & ",ReportVATAmountS= " & _
                        rsDemandSplit("TVAT").Value & ", ReportOSAmountS =  " & dblOSAmount & ",ReportReceivedAmountS= " & dblReceivedAmt & ",ReportDateFromS=#" & Format(dateFrom, "dd MMM yyyy") _
                        & "#,ReportDateTOS =#" & Format(DateTO, "dd MMM yyyy") & "#,ReportDemandTypeDescS= '" & szTypeOfDemanddesc & "' where DemandId=" & _
                        rsDemandSplit("DEMANDID").Value & " AND ReportCsShowFlag= '1'AND SPLITID=" & rsDemandSplit("SplitID").Value & " "

            rsDemandSplit.MoveNext
       Wend
     rsDemandSplit.Close
    ''' Expense side ***********************************************************************
    Dim rstblPurInv As New ADODB.Recordset
    Dim rsPayment As New ADODB.Recordset
    Dim rsInv As New ADODB.Recordset
    Dim dblPayAmount As Double
    Dim dblinvNetamount As Double
    Dim dblinvVATamount As Double
    Dim szPDesc As String
    Dim szNominalCode As String
    'IN this slection you are only including type PI 6 , so you need to also include type 24 anyway
    rstblPurInv.Open "Select P.TransactionID,PI.MY_ID from tblPurInv PI,tlbPayment P where P.Type=6 AND PI.MY_ID=P.PI AND CL_ID='" & szSelectedClient & "'", adoConn, adOpenStatic, adLockReadOnly
    While Not rstblPurInv.EOF
        If rstblPurInv("My_ID").Value = "22013111271401633183" Then
            Debug.Print ""
            'There is two problem 1)PayTransactionsSplit is not writing PPR and second for PPR you nee to start from fromtran
        End If
         dblPayAmount = 0
         rsPayment.Open "Select sum(T.PaymentAmount) as amt from tlbPayment P,PayTransactionsSplit T,tlbPaymentSplit S,Fund F where S.payHeader=P.TransactionID " & _
                        "AND S.ClientStatementPrevID=" & CSID & " AND S.PayTransactionIDSplit=T.TransactionID AND S.FundID=F.FundID and F.FundCode in ( " & ListOfFunds & ")" & _
                        " AND T.FromTran=P.TransactionID and T.Deleteflag=false and P.amount>P.osamount AND T.ToTran=" & _
                        rstblPurInv("TransactionID").Value & "", adoConn, adOpenStatic, adLockReadOnly
         If Not rsPayment.EOF Then
                dblPayAmount = IIf(IsNull(rsPayment("amt").Value), 0, rsPayment("amt").Value)
         End If
         rsPayment.Close
         rsPayment.Open "Select sum(NET_AMOUNT) as amt, sum(VAT) as VAT1 from tblPurInvSRec S where ParentID='" & rstblPurInv("My_ID").Value & "'", adoConn, adOpenStatic, adLockReadOnly
         If Not rsPayment.EOF Then
                dblinvNetamount = IIf(IsNull(rsPayment("amt").Value), 0, rsPayment("amt").Value)
                dblinvVATamount = IIf(IsNull(rsPayment("VAT1").Value), 0, rsPayment("VAT1").Value)
         End If
         rsPayment.Close
         szPDesc = ""
         rsPayment.Open "Select Description,NOMINAL_CODE from tlbPaymentSplit P where P.PayHeader=" & _
                    rstblPurInv("TransactionID").Value & "", adoConn, adOpenStatic, adLockReadOnly
         If Not rsPayment.EOF Then
              szPDesc = IIf(IsNull(rsPayment("Description").Value), 0, rsPayment("Description").Value)
              szPDesc = Replace(szPDesc, "'", "")
              szNominalCode = IIf(IsNull(rsPayment("NOMINAL_CODE").Value), 0, rsPayment("NOMINAL_CODE").Value)
         End If
         rsPayment.Close
         
         adoConn.Execute "Update tblPurInv P set ReportPaymentAmount= " & dblPayAmount & ", ReportPayDescription='" & _
                szPDesc & "',ReportNominalCode='" & szNominalCode & "',ReportInvNetAmount= " & dblinvNetamount & ",ReportINVVATAmount= " & _
                dblinvVATamount & " where P.My_ID='" & rstblPurInv("My_ID").Value & "'"
                
         rstblPurInv.MoveNext
    Wend
    rstblPurInv.Close
    
    'For type 24 PPR we are adding this segment
     rstblPurInv.Open "Select P.TransactionID,PI.MY_ID from tblPurInv PI,tlbPayment P where P.Type=7 AND PI.MY_ID=P.PI AND CL_ID='" & szSelectedClient & "'", adoConn, adOpenStatic, adLockReadOnly
    While Not rstblPurInv.EOF
         dblPayAmount = 0
         rsPayment.Open "Select sum(T.PaymentAmount) as amt from tlbPayment P,PayTransactionsSplit T,tlbPaymentSplit S,Fund F where S.payHeader=P.TransactionID " & _
                        "AND S.ClientStatementPrevID=" & CSID & " AND S.PayTransactionIDSplit=T.TransactionID AND S.FundID=F.FundID and F.FundCode in ( " & ListOfFunds & ")" & _
                        " AND T.TOTran=P.TransactionID and T.Deleteflag=false and P.amount>P.osamount AND T.FromTran=" & _
                        rstblPurInv("TransactionID").Value & "", adoConn, adOpenStatic, adLockReadOnly
         If Not rsPayment.EOF Then
                dblPayAmount = IIf(IsNull(rsPayment("amt").Value), 0, rsPayment("amt").Value)
         End If
         rsPayment.Close
'         rsPayment.Open "Select sum(NET_AMOUNT) as amt, sum(VAT) as VAT1 from tblPurInvSRec S where ParentID='" & rstblPurInv("My_ID").Value & "'", adoConn, adOpenStatic, adLockReadOnly
'         If Not rsPayment.EOF Then
'                dblinvNetamount = IIf(IsNull(rsPayment("amt").Value), 0, rsPayment("amt").Value)
'                dblinvVATamount = IIf(IsNull(rsPayment("VAT1").Value), 0, rsPayment("VAT1").Value)
'         End If
'         rsPayment.Close
         szPDesc = ""
         rsPayment.Open "Select Description,NOMINAL_CODE from tlbPaymentSplit P where P.PayHeader=" & _
                    rstblPurInv("TransactionID").Value & "", adoConn, adOpenStatic, adLockReadOnly
         If Not rsPayment.EOF Then
              szPDesc = IIf(IsNull(rsPayment("Description").Value), 0, rsPayment("Description").Value)
              szPDesc = Replace(szPDesc, "'", "")
             ' szNominalCode = IIf(IsNull(rsPayment("NOMINAL_CODE").Value), 0, rsPayment("NOMINAL_CODE").Value)
         End If
         rsPayment.Close
         Debug.Print rstblPurInv("My_ID").Value
         rsPayment.Open "Select NOMINAL_CODE from tblPurInvSRec P where P.ParentID='" & _
                    rstblPurInv("My_ID").Value & "'", adoConn, adOpenStatic, adLockReadOnly
         If Not rsPayment.EOF Then
              szNominalCode = IIf(IsNull(rsPayment("NOMINAL_CODE").Value), 0, rsPayment("NOMINAL_CODE").Value)
         End If
         rsPayment.Close
         'YOu need to update sc with this on payment side
         'Need to do the same on SI side
         adoConn.Execute "Update tblPurInv P set ReportPaymentAmount= " & -dblPayAmount & ", ReportPayDescription='" & _
                szPDesc & "',ReportNominalCode='" & szNominalCode & "',ReportInvNetAmount= " & -dblPayAmount & ",ReportINVVATAmount= " & _
                dblinvVATamount & " where P.My_ID='" & rstblPurInv("My_ID").Value & "'"

         rstblPurInv.MoveNext
    Wend
    rstblPurInv.Close
    adoConn.Close
 
Debug.Print time & " end"
    Set Report = reportApp.OpenReport(App.Path & szReportPath & "\ClientStatementPreviewSplit.rpt")
    Dim dblPercentage As Double
    Report.EnableParameterPrompting = False
    Report.DiscardSavedData
    Report.ParameterFields(1).AddCurrentValue CInt(1)
    Report.ParameterFields(2).AddCurrentValue szSelectedClient 'client ID
    Report.ParameterFields(3).AddCurrentValue CDate(txtStatementDate1.text)  'statement date
    Report.ParameterFields(4).AddCurrentValue CDate(txtLastStatementDate1.text)  'Previuos statement date
    Report.ParameterFields(5).AddCurrentValue 100 '100 Percent
    Report.ParameterFields(6).AddCurrentValue "0" '0 is for detail record so print address for passing client ID in parameter 2
    adoConn.Open getConnectionString
    Report.ParameterFields(7).AddCurrentValue findClientaddress(adoConn, szSelectedClient)
    Report.ParameterFields(8).AddCurrentValue IIf(chkShowDue.Value = 1, True, False)
    adoConn.Close
    Load frmReport
    frmReport.LoadReportViewer Report
End Sub
Private Sub PrintCSlineByLineNewPreview(ByVal CSID As String) ' This is new method of previewing CS statement 24/01/2023 which is faster
 'first we are marking receipts and payments
 'Create new preview table and insert data into it. pass CSID at CS Preview ID
    'On Error GoTo ERR
   Dim iIncDec As Long
   iIncDec = 0
   Dim rCount As Integer
   Dim selRow As Integer
   Dim adoConn As New ADODB.Connection

    For rCount = 1 To flxClients.Rows - 1
         If flxClients.TextMatrix(rCount, 0) = "X" Then
            szSelectedClient = flxClients.TextMatrix(rCount, 1)
            Exit For
         End If
    Next
        
    szCurrentStatementID = CSID
    CSID = Replace(szCurrentStatementID, "CS", "")
    Dim reportApp As New CRAXDRT.Application
    Dim Report As CRAXDRT.Report
    
    Dim rsDemandSplit As New ADODB.Recordset
    Dim rsReceived As New ADODB.Recordset
    Dim rsRentSummaryStatement As New ADODB.Recordset
    Dim dblReceivedAmt As Double
    Dim dblCrReceivedAmt As Double
    Dim dblOSAmount As Double
    Dim szListofFunds As String
    Dim szTypeOfDemanddesc As String
    Dim dateFrom As String
    
    Dim DateTO As String
    If szCurrentStatementID = "" Then
        MsgBox "Please select a statement", vbInformation, "Warning"
        Exit Sub
    End If

    
   Dim strDueDate As String
   Dim rsDemandSplitAmt As New ADODB.Recordset
   Dim iCount As Integer
   Dim dblDemandSplitamt As Double
   Dim szSQL As String
   Dim SQLforInsert As String
   Dim adoOsamount As New ADODB.Recordset
   adoConn.Open getConnectionString
   adoConn.Execute "Delete from ReportClientStatementDemandsPreview"
   adoConn.Execute "Delete from ReportClientStatementPurchasesPreview"
   'Here I shall explain the business logic Anol 2023-08-10
   'First we are marking SR
   'Show all the SI that allocation with SR
   'insert into a report table and print
   
'Type 1 SI
'Unitnumber has been replaces with ''
   SQLforInsert = " Select " & szCurrentStatementID & " as StatementID,TransactionID,ClientID,PropertyID,DemandID,SplitID,Sageaccountnumber,'',Type as DemandTypeDesc,TypeOfDemand,DueDate,D.DateFrom,D.DateTo,switch(Type=1,D.Amount,Type=2,-D.Amount)as NETAmount," & _
                    "switch(Type=1,D.VATAmount,Type=2,-D.VATAmount),round(switch(Type=1,NetAmounts+VATAmounts),2) as ReceivedAmountS,switch(Type=2,-NetAmounts-VATAmounts) as CreditAmount,T.OSAmount from " & _
                    "(Select D.DemandID,T.TransactionID,T.sageaccountnumber,U.Unitnumber,U.PropertyID,T.ClientID,D.TotalAmount,D.SplitID,T.Type,D.TypeOfDemand,D.DueDate,D.DateFrom,D.DateTO,D.Amount,T.OSAmount,D.VATAmount" & _
                    " from DemandSplitRecords D INNER JOIN  ((tlbReceipt T INNER JOIN ( SELECT Distinct A.ToTran as TrxId FROM" & _
                    " RptTransactionsSplit A, tlbReceipt R,  tlbReceiptSplit RS  where A.FromTran=R.TransactionID AND" & _
                    " RS.rptHeader=R.TransactionID and RS.RptTransactionsIDSplit=A.transactionID AND  R.ClientID='" & szSelectedClient & "' AND Deleteflag=false  AND " & _
                    " RS.ClientStatementPrevID=" & szCurrentStatementID & "   group by  A.ToTran ) B ON   B.trxID=T.TransactionID) INNER JOIN units U ON T.UnitID=U.UnitNumber)" & _
                    " on T.Demandref=D.DemandID) X LEFT JOIN  (SELECT A.ToTran as FromTran, A.SplitIDofSI," & _
                    " A.FundID, Sum(A.NetAmount) AS NetAmountS, Sum(A.VATAMOUNT) AS VATAMOUNTS FROM RptTransactionsSplit  A" & _
                    " , tlbReceipt R,  tlbReceiptSplit RS  where A.FromTran=R.TransactionID AND RS.rptHeader=R.TransactionID" & _
                    " and RS.RptTransactionsIDSplit=A.transactionID AND  R.ClientID='" & szSelectedClient & "' AND Deleteflag=false  AND  RS.ClientStatementPrevID=" & szCurrentStatementID & "" & _
                    "   group by  A.ToTran, A.SplitIDofSI, A.fundID  Union SELECT" & _
                    " A.FromTran, A.SplitIDofSI, A.FundID, Sum(A.NetAmount) AS NetAmountS, Sum(A.VATAMOUNT) AS VATAMOUNTS FROM RptTransactionsSplit" & _
                    " A, tlbReceipt R,  tlbReceiptSplit RS  where A.FromTran=R.TransactionID AND RS.rptHeader=R.TransactionID" & _
                    " and RS.RptTransactionsIDSplit=A.transactionID AND  R.ClientID='" & szSelectedClient & "' AND A.Deleteflag=false    and" & _
                    " A.Allocdate <=# " & Format(txtStatementDate1.text, "DD MMM yyyy") & " # And A.Allocdate ># " & Format(txtLastStatementDate1, "DD MMM yyyy") & " #  group" & _
                    " by  A.FromTran, A.SplitIDofSI, A.FundID) Y ON X.TransactionID=Y.FromTran  AND X.splitID=Y.SplitIDofSI"
             adoConn.Execute "Insert into ReportClientStatementDemandsPreview(StatementID,SITrxID,ClientID,PropertyID,DemandID,SplitID,SageAccountNumber,UnitNumber" & _
            ",TransactionType,TypeOfDemand,DueDate,DateFrom,DateTo,NetAmount,VATAmount,ReceivedAmountS,CreditAmount,OSAmount)" & _
            SQLforInsert

           SQLforInsert = " Select OSAmount,M.Netamount,vatamount, ReceiptAMOUNTS,SITrxID,SplitID from  ReportClientStatementDemandsPreview M INNER JOIN  (SELECT A.ToTran as FromTran," & _
           "A.SplitIDofSI,Sum(A.NetAmount+A.VATAMOUNT) AS ReceiptAMOUNTS FROM RptTransactionsSplit  A , tlbReceipt R,  tlbReceiptSplit RS  where " & _
           "R.RDate <=#" & Format(txtStatementDate1.text, "DD MMM yyyy") & "# AND " & _
           "A.FromTran=R.TransactionID AND RS.rptHeader=R.TransactionID and RS.RptTransactionsIDSplit=A.transactionID AND  R.ClientID='" & szSelectedClient & "' AND Deleteflag=false " & _
           " group by  A.ToTran, A.SplitIDofSI )  AS X ON M.SITrxID = X.FromTran ANd X.SplitIDofSI=M.SplitID"
            adoOsamount.Open SQLforInsert, adoConn, adOpenStatic, adLockReadOnly
            If Not adoOsamount.EOF Then
                adoOsamount.MoveFirst
            End If
            While Not adoOsamount.EOF
                adoConn.Execute "Update  ReportClientStatementDemandsPreview SET Osamount =" & adoOsamount("NetAmount").Value + adoOsamount("vatamount").Value - adoOsamount("ReceiptAMOUNTS").Value & "" & _
                                " where SITrxID=" & adoOsamount("SITrxID").Value & " and  SplitID=" & adoOsamount("SplitID").Value & ""
                adoOsamount.MoveNext
            Wend
            'adoOsamount.UpdateBatch
            adoOsamount.Close
           adoConn.Execute "update ReportClientStatementDemandsPreview set CreditAmount=0  where CreditAmount is null"
           adoConn.Execute "update ReportClientStatementDemandsPreview set OSAmount=0  where OSAmount is null"
           adoConn.Execute "update ReportClientStatementDemandsPreview set ReceivedAmountS=0  where ReceivedAmountS is null"
           
'''           'Sleep (100)
'''           'Type 2 SC
'''           'Here D.[FromTran])=[R].[TransactionID]  R tlbReceipt represents Credit side
'''           'Taking there some by Sum(D.ReceiptAmount) AS Amt where D is allocation table with date range criteria
'''           'So now  have collected all allocated amount for credit notes
'''           'Now I inner join them with DemandsplitRecords DS to get split level Credit note details
'''           'Finally inserting them into a report table
'''            SQLforInsert = "select " & szCurrentStatementID & " as StatementID,R.TransactionID,ClientID,U.PropertyID,D.DemandID,SplitID,D.Sageaccountnumber,D.Unitnumber, TransactionType,TypeOfDemand,DueDate, DS.DateFrom,DS.DateTo,  " & _
'''               "switch(D.transactionType=1,X.Amt,D.transactionType=2,x.amt)as NETAmount,switch(D.transactionType=1,DS.VATAmount,D.transactionType=2,DS.VATAmount) as VAT, " & _
'''                "switch(D.transactionType=2,DS.Amount+DS.VATAmount) as CreditAmount,0 as DS1  from (DemandsplitRecords DS Inner join  " & _
'''               "(SELECT R.DemandRef,RC.SlNumber, Sum(D.ReceiptAmount) AS Amt, R.TransactionID, D.SPlitIDofSi FROM RptTransactionsSplit AS D, tlbReceiptSplit AS RS, tlbReceipt AS RC, tlbReceipt AS R  " & _
'''               "WHERE (((RC.TransactionID)=[RS].[RptHeader]) AND R.Amount>R.OsAmount AND ((RS.RptTransactionsIDSplit)=[D].[TransactionID]) AND ((D.[FromTran])=[R].[TransactionID]) AND ((D.[deleteflag])=False) AND ((R.Type)=2)  " & _
'''               "AND ((R.ClientID)='" & szSelectedClient & "')) and D.Allocdate <=# " & Format(txtStatementDate1.text, "DD MMM yyyy") & " #  And D.Allocdate ># " & Format(txtLastStatementDate1, "DD MMM yyyy") & "#GROUP BY  R.DemandRef,  " & _
'''                "RC.SlNumber, R.TransactionID, D.SPlitIDofSi)X ON X.DemandRef=DS.DemandID and DS.SplitID=X.SPlitIDofSi), DemandRecords D,UNITS U,Property P where  P.PropertyID=U.PropertyID AND  " & _
'''                "D.DemandID=DS.DemandID AND D.UnitNumber=U.UnitNumber and D.exclCRNtoCS=false ; "
'''            adoconn.Execute "Insert into ReportClientStatementDemandsPreview(StatementID,SITrxID,ClientID,PropertyID,DemandID,SplitID,SageAccountNumber,UnitNumber," & _
'''            "TransactionType,TypeOfDemand,DueDate,DateFrom,DateTo,NetAmount,VATAmount,CreditAmount,osAmount)" & _
'''            SQLforInsert
'''            adoconn.Execute "update ReportClientStatementDemandsPreview D set ReceivedAmountS=0,D.OSAmount =-(NetAmount+VATAmount)+CreditAmount where TransactionType=2"
            
            'Credit note needs to show relavent SI. I shall be using UnitNumber field from ReportClientStatementDemandsPreview for showing SI number
'             adoConn.Execute "update ReportClientStatementDemandsPreview D,RptTransactionsSplit T, tlbReceipt AS R,DemandRecords DR set D.UnitNumber ='/ SI'& R.slnumber where D.TransactionType=2 " & _
'                    " AND T.ToTran=R.TransactionID and R.DemandRef=Dr.DemandID and D.SITrxID=T.FromTran"
             
   'writing code for inserting SRR  23    by anol 2023-08-20
            SQLforInsert = "select " & szCurrentStatementID & " as StatementID,TransactionID, R.clientID,PropertyID,R.RDate,SageAccountNumber," & _
                        "23,1,-X.Amt,0,-X.Amt,0,0  from ( (SELECT R.TransactionID,R.UNITID as PropertyID,R.clientID,R.RDate,SageAccountNumber,slnumber,R.NominalCode,Sum(RS.Amount) AS Amt FROM " & _
                        "tlbReceiptSplit AS RS,  tlbReceipt AS R  WHERE R.Amount>R.OsAmount AND ((R.Type)=23)  AND  " & _
                        "(R.ClientID)='" & szSelectedClient & "' AND RS.rptHeader=R.TransactionID AND RS.ClientStatementPrevID=" & szCurrentStatementID & " " & _
                        " GROUP BY R.TransactionID,R.clientID,R.UNITID,R.RDate,slnumber,SageAccountNumber,R.NominalCode)X )"
                        
               adoConn.Execute "Insert into ReportClientStatementDemandsPreview(StatementID,SITrxID,ClientID,PropertyID,DueDate,SageAccountNumber," & _
            "TransactionType,SplitID,NetAmount,VATAmount,ReceivedAmountS,CreditAmount,osAmount)" & _
            SQLforInsert
'
'                adoconn.Execute "update ReportClientStatementDemandsPreview D,RptTransactionsSplit T, tlbReceipt AS R set D.UnitNumber ='SRR'& R.slnumber where D.TransactionType=23 " & _
'                    " AND T.ToTran=R.TransactionID  and D.SITrxID=T.toTran and T.DeleteFlag=false "

    adoConn.Execute "update ReportClientStatementDemandsPreview D,RptTransactionsSplit T, tlbReceipt AS R, DemandSplitRecords AS RS set D.UnitNumber ='SRR'& R.slnumber,  " & _
                    "      D.DateFrom=RS.DateFrom, D.DateTo=RS.DateTo  where D.TransactionType=23  AND R.Demandref=RS.DemandID " & _
                    "     AND T.fromTran=R.TransactionID  and D.SITrxID=T.toTran and T.DeleteFlag=false   "
                    
            'Neet to check OS amounts here for acceptance

            'Insert expense for PI Type 6
            SQLforInsert = "  Select StatementID, '6' as type,D.ParentID as MY_ID,SplitID,TransactionID,T.ClientID, PropertID," & _
                            " T.Pdate,SageAccountNumber,NOMINAL_CODE,D.Net_Amount,D.VAT,PaidAmounts from (Select  " & szCurrentStatementID & _
                             " as StatementID, D.ParentID as MY_ID,D.TRAN_ID as SplitID, D.ParentID,T.ClientID, D.TRANS as PropertID" & _
                            " , T.Pdate,T.SageAccountNumber,NOMINAL_CODE,T.TransactionID,D.Net_Amount,D.VAT from tblPurInvSRec D INNER JOIN  (Select" & _
                            " * from tlbPayment T INNER JOIN (SELECT Distinct A.ToTran as TrxId" & _
                            " FROM PayTransactionsSplit A, tlbPayment R,  tlbPaymentSplit RS  where A.FromTran=R.TransactionID" & _
                            " AND RS.PayHeader=R.TransactionID and RS.PayTransactionIDSplit=A.transactionID AND  R.ClientID='" & szSelectedClient & "'  AND Deleteflag=false  AND" & _
                            "  RS.ClientStatementPrevID=" & szCurrentStatementID & " group by  A.ToTran)X ON T.TransactionID=X.TrxId)N ON" & _
                            " N.PI=D.ParentID ) Y LEFT JOIN  (SELECT A.ToTran as FromTran, A.SplitIDofPI," & _
                            "  Sum(A.NetAmount+A.VATAMOUNT) AS PaidAmounts FROM PayTransactionsSplit  A" & _
                            " , tlbPayment R, tlbPaymentSplit RS  where A.FromTran=R.TransactionID AND RS.PayHeader=R.TransactionID and" & _
                            " RS.PayTransactionIDSplit=A.transactionID AND  R.ClientID='" & szSelectedClient & "'  AND Deleteflag=false  AND  RS.ClientStatementPrevID=" & szCurrentStatementID & "" & _
                            "  group by  A.ToTran, A.SplitIDofPI, A.fundID  Union SELECT A.FromTran," & _
                            " A.SplitIDofPI,  Sum(A.NetAmount+A.VATAMOUNT) AS PaidAmounts FROM PayTransactionsSplit A," & _
                            " tlbPayment R,  tlbPaymentSplit RS  where A.FromTran=R.TransactionID AND RS.PayHeader=R.TransactionID and" & _
                            " RS.PayTransactionIDSplit=A.transactionID AND  R.ClientID='" & szSelectedClient & "' AND A.Deleteflag=false and A.Allocdate <=#" & Format(txtStatementDate1.text, "DD MMM yyyy") & "#" & _
                            " And A.Allocdate ># " & Format(txtLastStatementDate1, "DD MMM yyyy") & "#  group by  A.FromTran, A.SplitIDofPI," & _
                            " A.FundID )Z  ON Y.TransactionID=Z.FromTran AND Y.SplitID=cstr(Z.SplitIDofPI)"
           adoConn.Execute "Insert into ReportClientStatementPurchasesPreview(StatementID,Type,MY_ID,SplitID,TransactionID,ClientID,PropertyID,TranDate,SupplierID,NOMINAL_CODE,Netamount,VatAmount,PaymentAmount)" & _
                    SQLforInsert
                    'Payment description and  payment ref needs to be included
                    
           adoConn.Execute "update ReportClientStatementPurchasesPreview SET CreditAmount=0  where CreditAmount is null"
           adoConn.Execute "update ReportClientStatementPurchasesPreview SET OSAmount=0  where OSAmount is null"
           'This line is not correct OS amount should be taken from actual PI
           adoConn.Execute "update ReportClientStatementPurchasesPreview R,tlbPayment P SET R.OSAmount=P.OSAmount where P.PI=R.MY_ID " ' I cant fully remember were it date based?
           
                    'There is  no os amount calculation for expense side calculations
                    'Also no credit note is coming in this section
                    
           'Insert code for type 7 PI Credit note


'        SQLforInsert = "select " & szCurrentStatementID & " as StatementID,'7' as type,INV.MY_ID,TRAN_ID,TransactionID,ClientID,P.PropertyID,INV.TRAN_DATE,INV.SUPP_AC,DS.NOMINAL_CODE," & _
'                "-DS.Net_amount, -DS.VAT,0, -X.Amt,0  from (tblPurInvSRec DS Inner join " & _
'                "(SELECT P.PI,P.TransactionID,Sum(D.PaymentAmount) AS Amt, D.SPlitIDofPI FROM PayTransactionsSplit AS D,  tlbPayment AS P  WHERE P.Amount>P.OsAmount AND ((D.[FromTran])=[P].[TransactionID]) AND " & _
'                "((D.[deleteflag])=False) AND ((P.Type)=7)  AND ((P.ClientID)='" & szSelectedClient & "') and D.Allocdate <=# " & Format(txtStatementDate1.text, "DD MMM yyyy") & " #  And D.Allocdate ># " & _
'                Format(txtLastStatementDate1, "DD MMM yyyy") & "# GROUP BY  P.PI,   P.TransactionID,D.SPlitIDofPI)X " & _
'                "ON X.PI=DS.ParentID and DS.TRAN_ID=cstr(X.SPlitIDofPI)), tblPurInv INV,Property P where  P.PropertyID=INV.PropertyID AND  INV.My_ID=DS.ParentID "
'2023-07-08 [previous one is the backup before modification . now add a where clause in PPR with bank account filter
                '*********** this part is for showing credit notes*****rem 2023-08-20
''                SQLforInsert = "select " & szCurrentStatementID & " as StatementID,'7' as type,INV.MY_ID,TRAN_ID,TransactionID,ClientID,P.PropertyID,INV.TRAN_DATE,INV.SUPP_AC,DS.NOMINAL_CODE," & _
''                "-DS.Net_amount, -DS.VAT,0, -X.Amt,0  from (tblPurInvSRec DS Inner join " & _
''                "(SELECT P.PI,P.TransactionID,Sum(D.PaymentAmount) AS Amt, D.SPlitIDofPI FROM PayTransactionsSplit AS D,  tlbPayment AS P ,  tlbPayment AS Q  " & _
''                " WHERE P.Amount>P.OsAmount AND ((D.[FromTran])=[P].[TransactionID]) AND ((D.[TOTran])=[Q].[TransactionID]) AND D.NominalCODE='" & szSelectedBankAccount & "' AND " & _
''                "((D.[deleteflag])=False) AND ((P.Type)=7)  AND ((P.ClientID)='" & szSelectedClient & "') and D.Allocdate <=# " & Format(txtStatementDate1.text, "DD MMM yyyy") & " #  And D.Allocdate ># " & _
''                Format(txtLastStatementDate1, "DD MMM yyyy") & "# GROUP BY  P.PI,   P.TransactionID,D.SPlitIDofPI)X " & _
''                "ON X.PI=DS.ParentID and DS.TRAN_ID=cstr(X.SPlitIDofPI)), tblPurInv INV,Property P where  P.PropertyID=INV.PropertyID AND  INV.My_ID=DS.ParentID "
''                    adoConn.Execute "Insert into ReportClientStatementPurchasesPreview(StatementID,Type,MY_ID,SplitID,TransactionID,ClientID,PropertyID,TranDate,SupplierID,NOMINAL_CODE," & _
''                            "NetAmount,VATAmount,PaymentAmount,CreditAmount,osAmount)" & _
''                            SQLforInsert
               adoConn.Execute "update ReportClientStatementPurchasesPreview a,tblPurInv b,tlbTransactionTypes C  " & _
                            "SET PaymentRef=MID(CONSTANT,4,len(CONSTANT))& b.slnumber  where b.MY_ID=a.MY_ID and C.TYPE_ID=a.Type"
               '**************************************
        
     '*************************Now insert Management fee when OSAmount >0 here  date 13/08/2023****************
        
              If chkShowDue.Value = 1 Then
              'Do not enter those management fee which are in the snapshot
                    SQLforInsert = "select " & szCurrentStatementID & " as StatementID,P.TransactionType as Type, P.MY_ID,S.TRAN_ID,C.TransactionID,C.ClientID,C.UnitID,C.PDate,C.SageAccountNumber," & _
                                 "NOMINAL_CODE,'PI'& P.slnumber,S.NET_AMOUNT,VAT, 0,0,S.NET_AMOUNT+VAT " & _
                                 "from tblPurInv P,tblPurInvSRec S,tlbPayment C, Fund F Where C.PI=P.MY_ID and P.MY_ID= " & _
                        "S.parentID and C.osAmount>0 and F.FundID=S.dept_ID and F.FundCode in (" & ListOfFunds & ") and C.ClientID='" & szSelectedClient & "' and P.TransactionType=6 AND P.isManagementFee=true "
                    '1st part we are selecting managment fee preview produced in this ready made area
                    Dim rsCheck As New ADODB.Recordset
                    Dim rscheck1 As New ADODB.Recordset
                    Dim key
                    rsCheck.Open SQLforInsert, adoConn, adOpenStatic, adLockReadOnly
                    Dim myIDDIct As New Dictionary
                    While Not rsCheck.EOF
                                rscheck1.Open "Select * from ClientStatementPurchasesSnapshot where MY_ID='" & rsCheck("My_ID").Value & "' and (Netamount+vatamount)=osamount", adoConn, adOpenStatic, adLockReadOnly
                                If Not rscheck1.EOF Then
                                    myIDDIct.Add rsCheck("My_ID").Value, rsCheck("My_ID").Value
                                End If
                                rscheck1.Close
                            rsCheck.MoveNext
                    Wend
                    rsCheck.Close
                    
                    adoConn.Execute "Insert into ReportClientStatementPurchasesPreview(StatementID,Type,MY_ID,SplitID,TransactionID,ClientID,PropertyID,TranDate,SupplierID,NOMINAL_CODE,PaymentRef," & _
                            "NetAmount,VATAmount,PaymentAmount,CreditAmount,osAmount)" & _
                            SQLforInsert
                            
                    'Deleting manange fee from the preview table that matches with snapshot and os is full
                    For Each key In myIDDIct.Keys
                            adoConn.Execute "Delete R.* from ReportClientStatementPurchasesPreview where MY_ID='" & key & "'"
                    Next key
                        
              End If
                 
             If chkExcludeSupOS.Value = 1 Then 'this is actually include is true
                 'Type:6
                 'This records do not have allocation record and they are fully unallocated
                 'Select PI SPlit lines those have the selected fund and osAmount=amount in tlbpursplit table
                ' Dim PurchaseLedgerControl As String
                 
                 
                 'PurchaseLedgerControl = GetNominalCodeForControlAccount(adoConn, "Purchase Ledger Control", szSelectedClient)
                    SQLforInsert = "select " & szCurrentStatementID & " as StatementID,P.TransactionType as Type, P.MY_ID,S.TRAN_ID,C.TransactionID,C.ClientID,C.UnitID,C.PDate,C.SageAccountNumber," & _
                                 "NOMINAL_CODE,'PI'& P.slnumber,S.NET_AMOUNT,VAT, 0,0,S.NET_AMOUNT+VAT " & _
                                 "from tblPurInv P,tblPurInvSRec S,tlbPayment C, Fund F Where C.PI=P.MY_ID and P.MY_ID= " & _
                        "S.parentID and osamount=amount and F.FundID=S.dept_ID and F.FundCode in (" & ListOfFunds & ") and C.ClientID='" & szSelectedClient & "' and P.TransactionType=6 AND P.isManagementFee=false AND P.isRentPayable=false"


                  adoConn.Execute "Insert into ReportClientStatementPurchasesPreview(StatementID,Type,MY_ID,SplitID,TransactionID,ClientID,PropertyID,TranDate,SupplierID,NOMINAL_CODE,PaymentRef," & _
                            "NetAmount,VATAmount,PaymentAmount,CreditAmount,osAmount)" & _
                            SQLforInsert
                            
                  'Type:7
                    'PurchaseLedgerControl = GetNominalCodeForControlAccount(adoConn, "Purchase Ledger Control", szSelectedClient)
                    SQLforInsert = "select " & szCurrentStatementID & " as StatementID,P.TransactionType as Type, P.MY_ID,S.TRAN_ID,C.TransactionID,C.ClientID,C.UnitID,C.PDate,C.SageAccountNumber," & _
                                 "NOMINAL_CODE,-S.NET_AMOUNT,-VAT, 0,0,-(S.NET_AMOUNT+VAT) " & _
                                 "from tblPurInv P,tblPurInvSRec S,tlbPayment C, Fund F Where C.PI=P.MY_ID and P.MY_ID= " & _
                        "S.parentID and osamount=amount and F.FundID=S.dept_ID and F.FundCode in (" & ListOfFunds & ") and C.ClientID='" & szSelectedClient & "' and P.TransactionType=7"


                  adoConn.Execute "Insert into ReportClientStatementPurchasesPreview(StatementID,Type,MY_ID,SplitID,TransactionID,ClientID,PropertyID,TranDate,SupplierID,NOMINAL_CODE," & _
                            "NetAmount,VATAmount,PaymentAmount,CreditAmount,osAmount)" & _
                            SQLforInsert
                 '****************end********************************************************
              End If

        adoConn.Execute "Update   tlbPaymentSplit PS,tlbPayment P,ReportClientStatementPurchasesPreview R  set R.PaymentDescription=PS.Description where   Ps.PayHeader=p.transactionID " & _
        "AND PS.SplitID=R.SplitID and  P.type=R.Type and P.transactionID=R.TransactionID"
        adoConn.Execute "Update  Supplier S, ReportClientStatementPurchasesPreview R,GlobalData G set R.VATAmount= Round((R.NetAmount * 20/120),2) where R.PropertyID=G.PropertyID and " & _
                            " R.SupplierID= S.SupplierID and S.OptedtoTax=true "
        adoConn.Execute "Update  Supplier S, ReportClientStatementPurchasesPreview R,GlobalData G set R.NetAmount= Round((R.NetAmount * 100/120),2) where R.PropertyID=G.PropertyID and " & _
                            " R.SupplierID= S.SupplierID and S.OptedtoTax=true "
        adoConn.Execute "update ReportClientStatementPurchasesPreview a,tlbPayment b,tlbTransactionTypes C  " & _
                            "SET PaymentRef=MID(CONSTANT,4,len(CONSTANT))& b.slnumber,a.PaymentDescription=b.Details  where b.TransactionID=a.TransactionID and C.TYPE_ID=a.Type and a.Type=24 "
           
           'Insert code for type 24 PI Purchase Payment Refund
        'for type 24 there is no corresponding entry in tlbPurInv. so need to remove that relationship
        'TRAN_ID is the split ID in tblPurInvSRec table
'         SQLforInsert = "select " & szCurrentStatementID & " as StatementID,'24' as type,TransactionID, 1 as TRAN_ID,slnumber,P.clientID,PropertyID,P.PDate,SageAccountNumber,P.NominalCode," & _
'                        "-X.Amt,0,-X.Amt,0,0  from ( (SELECT P.TransactionID,P.UNITID as PropertyID,P.clientID,P.PDate,SageAccountNumber,slnumber,P.NominalCode,Sum(D.Amount) AS Amt FROM " & _
'                        "tlbPaymentSplit AS D,  tlbPayment AS P  WHERE P.Amount>P.OsAmount AND ((P.Type)=24)  AND  " & _
'                        "(P.ClientID)='" & szSelectedClient & "' AND D.PayHeader=P.TransactionID AND D.ClientStatementPrevID=" & szCurrentStatementID & " " & _
'                        " GROUP BY P.TransactionID,P.clientID,P.UNITID,P.PDate,slnumber,SageAccountNumber,P.NominalCode)X )"

         SQLforInsert = "select " & szCurrentStatementID & " as StatementID,'24' as type,P.TransactionID, 1 as TRAN_ID,P.slnumber,P.Details,P.clientID,PropertyID,P.PDate,P.SageAccountNumber,PD.NOMINAL_CODE," & _
                        "-X.Amt,0,-X.Amt,0,0  from ( (SELECT P.TransactionID,P.UNITID as PropertyID,P.Details,P.clientID,P.PDate,P.SageAccountNumber,P.slnumber,PD.NOMINAL_CODE,Sum(D.Amount) AS Amt FROM " & _
                        "tlbPaymentSplit AS D,  tlbPayment AS P,  tlbPayment AS Q , PayTransactions PS,tblPurInvSRec PD  WHERE PD.ParentID=Q.PI and P.TransactionID=PS.totran and ps.fromtran=Q.transactionID and P.Amount>P.OsAmount AND ((P.Type)=24)  AND  " & _
                        "(P.ClientID)='" & szSelectedClient & "' AND ps.DeleteFlag=false AND D.PayHeader=P.TransactionID AND D.ClientStatementPrevID=" & szCurrentStatementID & " " & _
                        " GROUP BY P.TransactionID,P.clientID,P.UNITID,P.PDate,P.slnumber,P.SageAccountNumber,PD.NOMINAL_CODE,P.Details)X )"
                        

                    adoConn.Execute "Insert into ReportClientStatementPurchasesPreview(StatementID,Type,MY_ID,SplitID,TransactionID,PaymentDescription,ClientID,PropertyID,TranDate,SupplierID,NOMINAL_CODE," & _
                            "NetAmount,VATAmount,PaymentAmount,CreditAmount,osAmount)" & _
                            SQLforInsert
                            
'                               adoconn.Execute "update ReportClientStatementPurchasesPreview a,tblPurInv b,tlbTransactionTypes C  " & _
'                            "SET PaymentRef=MID(CONSTANT,4,len(CONSTANT))& b.slnumber  where b.MY_ID=a.MY_ID and C.TYPE_ID=a.Type"

  adoConn.Execute "update ReportClientStatementPurchasesPreview a,tlbPayment b,tlbTransactionTypes C  " & _
                            "SET PaymentRef=MID(CONSTANT,4,len(CONSTANT))& b.slnumber  where cstr(b.transactionID)=a.MY_ID and C.TYPE_ID=a.Type"


    'code for report  template add 2023-06-13
   Dim rsStatementTemplate As New ADODB.Recordset
   Dim strReportName As String
   rsStatementTemplate.Open "Select CSPreviewTemplate from client where ClientID='" & szSelectedClient & "'", adoConn, adOpenStatic, adLockReadOnly
   If Not rsStatementTemplate.EOF Then
        strReportName = IIf(IsNull(rsStatementTemplate("CSPreviewTemplate").Value), "", rsStatementTemplate("CSPreviewTemplate").Value)
   End If
   rsStatementTemplate.Close
   
                            
    adoConn.Close
    Sleep (100)
   'Set Report = reportApp.OpenReport(App.Path & szReportPath & "\ClientStatementPreviewSplitNew.rpt")
   If strReportName = "" Then
        Set Report = reportApp.OpenReport(App.Path & szReportPath & "\ClientStatementPreviewSplitNew.rpt")
   Else
        Set Report = reportApp.OpenReport(App.Path & szReportPath & "\" & strReportName & "")
   End If
    Dim dblPercentage As Double
    Report.EnableParameterPrompting = False
    Report.DiscardSavedData
    Report.ParameterFields(1).AddCurrentValue CInt(CSID)
    Report.ParameterFields(2).AddCurrentValue szSelectedClient 'client ID
    Report.ParameterFields(3).AddCurrentValue CDate(txtStatementDate1.text)  'statement date
    Report.ParameterFields(4).AddCurrentValue CDate(txtLastStatementDate1.text)  'Previuos statement date
    Report.ParameterFields(5).AddCurrentValue 100 '100 Percent
    Report.ParameterFields(6).AddCurrentValue "0" '0 is for detail record so print address for passing client ID in parameter 2
    adoConn.Open getConnectionString
    Report.ParameterFields(7).AddCurrentValue findClientaddress(adoConn, szSelectedClient)
    Report.ParameterFields(8).AddCurrentValue 0
    Report.ParameterFields(9).AddCurrentValue -GetSupplierOSAmount 'Send supplier OS Amount at parameter 9
    Report.ParameterFields(10).AddCurrentValue CDbl(GetTotalExpenditure) 'Send supplier OS Amount at parameter 10
    
    adoConn.Close
    Load frmReport
    frmReport.LoadReportViewer Report
    Exit Sub
Err:
    MsgBox Err.description
End Sub
Private Function GetTotalExpenditure() As Currency  'This function works on OS column on the expediture @ CS
            Dim rsPayment As New ADODB.Recordset
            Dim szSQL As String
            Dim adoConn As New ADODB.Connection

            adoConn.Open getConnectionString
            Dim whereProperty As String

            szSQL = "Select  SUM(P.OSAmount) as AMT from ReportClientStatementPurchasesPreview P,tblPurInv S where " & _
                    " S.MY_ID=P.MY_ID and isManagementFee=false"
            rsPayment.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
            If Not rsPayment.EOF And chkExcludeSupOS.Value = 1 Then
                GetTotalExpenditure = IIf(IsNull(rsPayment.Fields.Item("AMT").Value), 0, rsPayment.Fields.Item("AMT").Value)
            End If
            rsPayment.Close
            
            szSQL = "Select  SUM(P.OSAmount) as AMT from ReportClientStatementPurchasesPreview P,tblPurInv S where " & _
                    " S.MY_ID=P.MY_ID and isManagementFee=true"
            rsPayment.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
            If Not rsPayment.EOF And chkShowDue.Value = 1 Then
                GetTotalExpenditure = GetTotalExpenditure + IIf(IsNull(rsPayment.Fields.Item("AMT").Value), 0, rsPayment.Fields.Item("AMT").Value)
            End If
            adoConn.Close
            Set adoConn = Nothing
End Function
Private Sub printClientStatementPreview(ByVal CSID As String)
   Dim iIncDec As Long
   iIncDec = 0
   Dim rCount As Integer
   Dim adoConn As New ADODB.Connection

    Dim reportApp As New CRAXDRT.Application
    Dim Report As CRAXDRT.Report
    adoConn.Open getConnectionString
    Dim rsDemandSplit As New ADODB.Recordset
    Dim rsReceived As New ADODB.Recordset
    Dim rsRentSummaryStatement As New ADODB.Recordset
    Dim dblReceivedAmt As Double
    Dim dblOSAmount As Double
    Dim szListofFunds As String
    Dim szTypeOfDemanddesc As String
    Dim dateFrom As String
    Dim strDueDate As String
    Dim strWhere As String
    
    Dim DateTO As String
    Dim szSelectedClient As String
    For rCount = 1 To flxClients.Rows - 1
         If flxClients.TextMatrix(rCount, 0) = "X" Then
             szSelectedClient = flxClients.TextMatrix(rCount, 1)
             Exit For
         End If
    Next
                    
    adoConn.Execute "Update DemandSplitRecords DS,DemandRecords D,Units U,Property P  set ReportNetAmount=0,ReportVATAmount=0,ReportCreditAmount=0,ReportReceivedAmount= 0,ReportDateFrom=Null," & _
                " reportOSAmount=0,ReportDateTO =null,ReportDemandTypeDesc= '' where D.DemandID=DS.DemandID  and U.UnitNumber=D.UnitNumber AND P.PropertyID=U.PropertyID AND P.ClientID='" & _
                szSelectedClient & "' "
     
                    
    rsDemandSplit.Open "Select D.DemandId,D.TransactionType,sum(Amount) as NAmt,sum(VATAmount)as TVAT from  DemandSplitRecords DS,DemandRecords D,Units U,Property P where D.DemandID=DS.DemandID " & _
                "and U.UnitNumber=D.UnitNumber AND TransactionType=1 AND P.PropertyID=U.PropertyID AND P.ClientID='" & _
                szSelectedClient & "' group by D.DemandId,D.TransactionType", adoConn, adOpenStatic, adLockReadOnly
    Dim rsDemandSplitCredit As New ADODB.Recordset
    Dim rsDemandSplit1 As New ADODB.Recordset
    Dim dblCrReceivedAmt As Double
    Dim dblCrSumAmt As Double
    
    
    rsDemandSplitCredit.Open "Select D.DemandId,sum(R.Amount) as NAmt from  tlbReceipt R,tlbReceiptSplit RS,DemandRecords D where  RS.ClientStatementPrevID=" & CSID & " AND R.DemandRef=D.DemandID " & _
                "AND R.TransactionID=RS.rptHeader AND Type=2 AND R.OSAmount=0 AND R.ClientID='" & _
                 szSelectedClient & "' group by D.DemandId", adoConn, adOpenStatic, adLockReadOnly
'    While Not rsDemandSplitCredit.EOF
'                 dateFrom = ""
'                 DateTO = ""
'                 strDueDate = ""
'                 szTypeOfDemanddesc = ""
'                 rsDemandSplit1.Open "Select DateFrom,DateTo,Type,DueDate from  DemandSplitRecords D,DemandTypes DT,Fund F where F.fundID=D.SageDepartment AND DT.ID=D.TypeOfDemand AND DemandId=" & rsDemandSplitCredit("DEMANDID").Value & " order by D.SPlitID ", adoConn, adOpenStatic, adLockReadOnly
'                 While Not rsDemandSplit1.EOF
'                        dateFrom = rsDemandSplit1("DateFrom").Value
'                        DateTO = rsDemandSplit1("DateTo").Value
'                        szTypeOfDemanddesc = rsDemandSplit1("Type").Value
'                        strDueDate = rsDemandSplit1("DueDate").Value
'                        rsDemandSplit1.MoveNext
'                 Wend
'                 rsDemandSplit1.Close
'
'                adoConn.Execute "Update DemandRecords DS  set ReportNetAmount=" & -rsDemandSplitCredit("NAmt").Value & ",ReportVATAmount=0,ReportCreditAmount=" & -rsDemandSplitCredit("NAmt").Value & ",ReportReceivedAmount= 0,ReportDateFrom=#" & dateFrom & "#," & _
'                " reportOSAmount=0,ReportDateTO =#" & DateTO & "#,ReportDemandTypeDesc= '" & szTypeOfDemanddesc & "' where TransactionType=2 AND DemandId=" & rsDemandSplitCredit("DEMANDID").Value & " "
'            rsDemandSplitCredit.MoveNext
'    Wend
'    rsDemandSplitCredit.Close
    Dim rsDemandSplitAmt As New ADODB.Recordset
    Dim iCount As Integer
    Dim dblDemandSplitamt As Double
    
    While Not rsDemandSplitCredit.EOF
                 dateFrom = ""
                 DateTO = ""
                 strDueDate = ""
'                 FromTran = 1461
                 rsDemandSplit1.Open "Select DateFrom,DateTo,Type,DueDate from  DemandSplitRecords D,DemandTypes DT where DT.ID=D.TypeOfDemand AND DemandId=" & _
                 rsDemandSplitCredit("DEMANDID").Value & " order by D.SPlitID ", adoConn, adOpenStatic, adLockReadOnly
                 If Not rsDemandSplit1.EOF Then
                        dateFrom = rsDemandSplit1("DateFrom").Value
                        DateTO = rsDemandSplit1("DateTo").Value
                        szTypeOfDemanddesc = rsDemandSplit1("Type").Value
                        strDueDate = rsDemandSplit1("DueDate").Value
                 End If
                 rsDemandSplit1.Close
                 '*************************
                 rsDemandSplitAmt.Open "Select Count(amount) as CNT from DemandSplitRecords DS where  DS.DemandId=" & _
                            rsDemandSplitCredit("DEMANDID").Value & "", adoConn, adOpenStatic, adLockReadOnly
                            
                 If Not rsDemandSplitAmt.EOF Then
                             iCount = rsDemandSplitAmt("CNT").Value
                 End If
                 rsDemandSplitAmt.Close
                 
                 
                 rsDemandSplitAmt.Open "Select amount from DemandSplitRecords DS where  DS.SPLITID=" & rsDemandSplitCredit("SplitID").Value & "  and DS.DemandId=" & _
                            rsDemandSplitCredit("DEMANDID").Value & "", adoConn, adOpenStatic, adLockReadOnly
                            
                 If Not rsDemandSplitAmt.EOF Then
                             dblDemandSplitamt = rsDemandSplitAmt("amount").Value
                 End If
                 dblOSAmount = dblDemandSplitamt
                 rsDemandSplitAmt.Close
                  
                       'when you are allocating SC with a SI then splitID you are getting that are split ID of a SI . so you cannot caluculat anything here with SC SplitID
                       'Ideally there is only one split in the SC
                If iCount = 1 Then
                    rsDemandSplit1.Open "Select Sum(RS.Amount) as Amt from  RptTransactionsSplit D,tlbReceiptSplit RS,tlbReceipt RC where RC.TransactionID=RS.RptHeader AND " & _
                    "RS.RptTransactionsIDSplit=D.TransactionID AND D.Allocdate <=#" & Format(txtStatementDate1.text, "dd MMM yyyy") & "#  AND D.Allocdate > #" & Format(txtLastStatementDate1.text, "dd MMM yyyy") & "# AND FromTran=" & _
                    rsDemandSplitCredit("TransactionID").Value & " and deleteflag=false  ", adoConn, adOpenStatic, adLockReadOnly
                    If Not rsDemandSplit1.EOF Then
                           dblCrSumAmt = IIf(IsNull(rsDemandSplit1("Amt").Value), 0, rsDemandSplit1("Amt").Value)
                           dblOSAmount = dblOSAmount - dblCrSumAmt
                    End If
                    rsDemandSplit1.Close
                    If dblCrSumAmt > 0 Then
                            adoConn.Execute "Update DemandSplitRecords DS,Fund F set ReportCsShowFlag= '1' where DS.SageDepartment=F.FundID and F.FundCode in ( " & ListOfFunds & ") AND DemandId=" & _
                        rsDemandSplitCredit("DEMANDID").Value & " "
                    End If
                 Else
                        rsDemandSplit1.Open "Select Sum(RS.Amount) as Amt from  RptTransactionsSplit D,tlbReceiptSplit RS,tlbReceipt RC where RC.TransactionID=RS.RptHeader AND " & _
                        "RS.RptTransactionsIDSplit=D.TransactionID   AND FromTran=" & _
                        rsDemandSplitCredit("TransactionID").Value & " and SPlitIDofSi=" & rsDemandSplitCredit("SplitID").Value & " and deleteflag=false  group by D.SPlitIDofSi ", adoConn, adOpenStatic, adLockReadOnly
                        If Not rsDemandSplit1.EOF Then
                               dblCrSumAmt = rsDemandSplit1("Amt").Value
                               dblOSAmount = dblOSAmount - dblCrSumAmt
                        End If
                        rsDemandSplit1.Close
                        If dblCrSumAmt > 0 Then
                                adoConn.Execute "Update DemandSplitRecords DS,Fund F set ReportCsShowFlag= '1' where DS.SageDepartment=F.FundID and F.FundCode in ( " & ListOfFunds & ") AND DemandId=" & _
                            rsDemandSplitCredit("DEMANDID").Value & " "
                        End If
                 End If

                adoConn.Execute "Update DemandSplitRecords DS  set ReportNetAmountS=" & -dblDemandSplitamt & ",ReportVATAmountS=0,ReportCreditAmountS=" & _
                                -dblCrSumAmt & ",ReportReceivedAmountS= 0,ReportDateFromS=#" & dateFrom & "#," & _
                                " reportOSAmountS=" & -dblOSAmount & ",ReportDateTOS=#" & DateTO & "#,ReportDemandTypeDescS= '" & szTypeOfDemanddesc & "' where DemandId=" & _
                                rsDemandSplitCredit("DEMANDID").Value & " AND DS.SplitID =" & rsDemandSplitCredit("SplitID").Value & ""
            rsDemandSplitCredit.MoveNext
    Wend
    rsDemandSplitCredit.Close
    While Not rsDemandSplit.EOF
                
                 If rsDemandSplit("DEMANDID").Value = 26 Then
                    Debug.Print ""
                 End If
                 dateFrom = ""
                 DateTO = ""
                 strDueDate = ""
                 szTypeOfDemanddesc = ""
                 rsDemandSplit1.Open "Select DateFrom,DateTo,Type,DueDate from  DemandSplitRecords D,DemandTypes DT,Fund F where " & _
                        "F.FundID=D.SageDepartment AND DT.ID=D.TypeOfDemand and F.FundCode in ( " & ListOfFunds & ") AND DemandId=" & _
                        rsDemandSplit("DEMANDID").Value & " order by D.SPlitID ", adoConn, adOpenStatic, adLockReadOnly
                 While Not rsDemandSplit1.EOF
                        dateFrom = rsDemandSplit1("DateFrom").Value
                        DateTO = rsDemandSplit1("DateTo").Value
                        szTypeOfDemanddesc = szTypeOfDemanddesc + " " + rsDemandSplit1("Type").Value
                        strDueDate = rsDemandSplit1("DueDate").Value
                        rsDemandSplit1.MoveNext
                 Wend
                 rsDemandSplit1.Close
                 
                 strWhere = ""
                 If chkExcludeReceipt.Value = 1 Then
                        strWhere = "AND RL.RDate> # " & Format(strDueDate, "dd MMM yyyy") & "# "
                 End If
                 dblReceivedAmt = 0
                 dblCrReceivedAmt = 0
                 dblOSAmount = 0
                 'RL is receipts side RL.RDate,R.DueDate,
                 rsReceived.Open "Select DemandId, sum(R.Amount) as ReceivedAmt from  tlbReceiptSplit R,tlbReceipt RC, RptTransactionsSplit T,DemandRecords D,tlbReceipt RL,Fund F where T.Deleteflag=False " & _
                 "AND RL.transactionID=T.FromTran AND T.transactionID=R.RptTransactionsIDSplit AND RC.transactionID=T.ToTran AND R.FundID=F.FundID " & _
                 "AND R.ClientStatementID=" & CSID & " AND RL.Type in(3,4) and RC.DemandRef=D.DemandID " & strWhere & _
                 "AND D.DemandId=" & rsDemandSplit("DEMANDID").Value & " group by DemandID", adoConn, adOpenStatic, adLockReadOnly
                ' dblReceivedAmt = 0 'AND R.FundID in ( " & szListofFunds & ")'AND RC.Type in(3,4,23)'sum(switch(RC.Type=3,R.Amount,RC.Type=4,R.Amount,RC.Type=23,-R.Amount))
                 If Not rsReceived.EOF Then
                             dblReceivedAmt = IIf(IsNull(rsReceived("ReceivedAmt").Value), 0, rsReceived("ReceivedAmt").Value)
                 End If
                 rsReceived.Close
                 rsReceived.Open "Select DemandId,sum(R.Amount) as ReceivedAmt from  tlbReceiptSplit R,tlbReceipt RC, RptTransactionsSplit T,DemandRecords D,tlbReceipt RL,Fund F where T.Deleteflag=False " & _
                 "AND RL.transactionID=T.FromTran AND T.transactionID=R.RptTransactionsIDSplit AND RC.transactionID=T.ToTran AND R.FundID=F.FundID " & _
                 "AND R.ClientStatementID=" & CSID & " AND RL.Type in(2) and RC.DemandRef=D.DemandID " & _
                 "AND D.DemandId=" & rsDemandSplit("DEMANDID").Value & " group by DemandID", adoConn, adOpenStatic, adLockReadOnly
                 'AND R.FundID in ( " & szListofFunds & ")'AND RC.Type in(3,4,23)'sum(switch(RC.Type=3,R.Amount,RC.Type=4,R.Amount,RC.Type=23,-R.Amount))
                 If Not rsReceived.EOF Then
                             dblCrReceivedAmt = IIf(IsNull(rsReceived("ReceivedAmt").Value), 0, rsReceived("ReceivedAmt").Value)
                 End If
                 rsReceived.Close
                 'writing code for collectiong the osamount .when we calculate osamount we exclude fund . because that may be paid without concerning what you have selected, no CSId required
                 rsReceived.Open "Select DemandId,sum(R.Amount) as ReceivedAmt from  tlbReceiptSplit R,tlbReceipt RC, RptTransactionsSplit T,DemandRecords D,tlbReceipt RL,Fund F where T.Deleteflag=False " & _
                 "AND RL.transactionID=T.FromTran AND T.transactionID=R.RptTransactionsIDSplit AND RC.transactionID=T.ToTran AND R.FundID=F.FundID AND " & _
                 "RL.RDate <=#" & Format(txtStatementDate1.text, "dd MMM yyyy") & "# " & _
                 "AND RL.Type in(2,3,4) and RC.DemandRef=D.DemandID and D.DemandId=" & rsDemandSplit("DEMANDID").Value & " group by DemandID", adoConn, adOpenStatic, adLockReadOnly
                 
                 If Not rsReceived.EOF Then
                             dblOSAmount = rsDemandSplit("NAmt").Value + rsDemandSplit("TVAT").Value - IIf(IsNull(rsReceived("ReceivedAmt").Value), 0, rsReceived("ReceivedAmt").Value)
                 End If
                 rsReceived.Close
                ' ReportOSAmount =  " & dblOSAmount& "
                 
                
                 If dateFrom = "" Then
                        'MsgBox "Date from in the demand split is empty", vbInformation, "Warning"
                       ' Exit Sub
                 End If
                  If dblReceivedAmt > 0 Then
                        adoConn.Execute "Update DemandRecords set ReportNetAmount= " & rsDemandSplit("NAmt").Value & ",ReportVATAmount= " & _
                        rsDemandSplit("TVAT").Value & ", ReportOSAmount =  " & dblOSAmount & ",ReportReceivedAmount= " & dblReceivedAmt & ",ReportDateFrom=#" & Format(dateFrom, "dd MMM yyyy") _
                        & "#,ReportDateTO =#" & Format(DateTO, "dd MMM yyyy") & "#,ReportDemandTypeDesc= '" & szTypeOfDemanddesc & "' where DemandId=" & rsDemandSplit("DEMANDID").Value & ""
                        dblReceivedAmt = 0
                  End If
                  If dblCrReceivedAmt > 0 Then
                        adoConn.Execute "Update DemandRecords set ReportNetAmount= " & rsDemandSplit("NAmt").Value & ",ReportVATAmount= " & _
                        rsDemandSplit("TVAT").Value & ", ReportOSAmount =  " & dblOSAmount & ", ReportCreditAmount= " & dblCrReceivedAmt & ",ReportDateFrom=#" & Format(dateFrom, "dd MMM yyyy") _
                        & "#,ReportDateTO =#" & Format(DateTO, "dd MMM yyyy") & "#,ReportDemandTypeDesc= '" & szTypeOfDemanddesc & "' where DemandId=" & rsDemandSplit("DEMANDID").Value & ""
                        dblCrReceivedAmt = 0
                  End If
                 
            rsDemandSplit.MoveNext
    Wend
    rsDemandSplit.Close
    Dim rstblPurInv As New ADODB.Recordset
    Dim rsPayment As New ADODB.Recordset
    Dim rsInv As New ADODB.Recordset
    Dim dblPayAmount As Double
    Dim dblinvNetamount As Double
    Dim dblinvVATamount As Double
    Dim szPDesc As String
    Dim szNominalCode As String
    rstblPurInv.Open "Select P.TransactionID,PI.MY_ID from tblPurInv PI,tlbPayment P where P.Type=6 AND PI.MY_ID=P.PI AND CL_ID='" & szSelectedClient & "'", adoConn, adOpenStatic, adLockReadOnly
    While Not rstblPurInv.EOF
         dblPayAmount = 0
         rsPayment.Open "Select sum(T.PaymentAmount) as amt from tlbPayment P,PayTransactionsSplit T,tlbPaymentSplit S,Fund F where S.payHeader=P.TransactionID " & _
                        "AND S.ClientStatementPrevID=" & CSID & " AND S.PayTransactionIDSplit=T.TransactionID AND S.FundID=F.FundID and F.FundCode in ( " & ListOfFunds & ")" & _
                        " AND T.TOTran=P.TransactionID and T.Deleteflag=false and P.amount>P.osamount AND T.FromTran=" & _
                        rstblPurInv("TransactionID").Value & "", adoConn, adOpenStatic, adLockReadOnly
         If Not rsPayment.EOF Then
                dblPayAmount = IIf(IsNull(rsPayment("amt").Value), 0, rsPayment("amt").Value)
         End If
         rsPayment.Close
         rsPayment.Open "Select sum(NET_AMOUNT) as amt, sum(VAT) as VAT1 from tblPurInvSRec S where ParentID='" & rstblPurInv("My_ID").Value & "'", adoConn, adOpenStatic, adLockReadOnly
         If Not rsPayment.EOF Then
                dblinvNetamount = IIf(IsNull(rsPayment("amt").Value), 0, rsPayment("amt").Value)
                dblinvVATamount = IIf(IsNull(rsPayment("VAT1").Value), 0, rsPayment("VAT1").Value)
         End If
         rsPayment.Close
         szPDesc = ""
         rsPayment.Open "Select Description,NOMINAL_CODE from tlbPaymentSplit P where P.PayHeader=" & _
                    rstblPurInv("TransactionID").Value & "", adoConn, adOpenStatic, adLockReadOnly
         If Not rsPayment.EOF Then
              szPDesc = IIf(IsNull(rsPayment("Description").Value), 0, rsPayment("Description").Value)
              szPDesc = Replace(szPDesc, "'", "")
              szNominalCode = IIf(IsNull(rsPayment("NOMINAL_CODE").Value), 0, rsPayment("NOMINAL_CODE").Value)
         End If
         rsPayment.Close
         
         adoConn.Execute "Update tblPurInv P set ReportPaymentAmount= " & dblPayAmount & ", ReportPayDescription='" & _
                szPDesc & "',ReportNominalCode='" & szNominalCode & "',ReportInvNetAmount= " & dblinvNetamount & ",ReportINVVATAmount= " & _
                dblinvVATamount & " where P.My_ID='" & rstblPurInv("My_ID").Value & "'"
                
         rstblPurInv.MoveNext
    Wend
    rstblPurInv.Close
    'adoconn.Close
    'For type 24 PPR we are adding this segment
    rstblPurInv.Open "Select P.TransactionID,PI.MY_ID from tblPurInv PI,tlbPayment P where P.Type=7 AND PI.MY_ID=P.PI AND CL_ID='" & szSelectedClient & "'", adoConn, adOpenStatic, adLockReadOnly
    While Not rstblPurInv.EOF
         dblPayAmount = 0
         rsPayment.Open "Select sum(T.PaymentAmount) as amt from tlbPayment P,PayTransactionsSplit T,tlbPaymentSplit S,Fund F where S.payHeader=P.TransactionID " & _
                        "AND S.ClientStatementPrevID=" & CSID & " AND S.PayTransactionIDSplit=T.TransactionID AND S.FundID=F.FundID and F.FundCode in ( " & ListOfFunds & ")" & _
                        " AND T.TOTran=P.TransactionID and T.Deleteflag=false and P.amount>P.osamount AND T.FromTran=" & _
                        rstblPurInv("TransactionID").Value & "", adoConn, adOpenStatic, adLockReadOnly
         If Not rsPayment.EOF Then
                dblPayAmount = IIf(IsNull(rsPayment("amt").Value), 0, rsPayment("amt").Value)
         End If
         rsPayment.Close
'         rsPayment.Open "Select sum(NET_AMOUNT) as amt, sum(VAT) as VAT1 from tblPurInvSRec S where ParentID='" & rstblPurInv("My_ID").Value & "'", adoConn, adOpenStatic, adLockReadOnly
'         If Not rsPayment.EOF Then
'                dblinvNetamount = IIf(IsNull(rsPayment("amt").Value), 0, rsPayment("amt").Value)
'                dblinvVATamount = IIf(IsNull(rsPayment("VAT1").Value), 0, rsPayment("VAT1").Value)
'         End If
'         rsPayment.Close
         szPDesc = ""
         rsPayment.Open "Select Description,NOMINAL_CODE from tlbPaymentSplit P where P.PayHeader=" & _
                    rstblPurInv("TransactionID").Value & "", adoConn, adOpenStatic, adLockReadOnly
         If Not rsPayment.EOF Then
              szPDesc = IIf(IsNull(rsPayment("Description").Value), 0, rsPayment("Description").Value)
              szPDesc = Replace(szPDesc, "'", "")
             ' szNominalCode = IIf(IsNull(rsPayment("NOMINAL_CODE").Value), 0, rsPayment("NOMINAL_CODE").Value)
         End If
         rsPayment.Close
         rsPayment.Open "Select NOMINAL_CODE from tblPurInvSRec P where P.My_ID='" & _
                    rstblPurInv("My_ID").Value & "'", adoConn, adOpenStatic, adLockReadOnly
         If Not rsPayment.EOF Then
              szNominalCode = IIf(IsNull(rsPayment("NOMINAL_CODE").Value), 0, rsPayment("NOMINAL_CODE").Value)
         End If
         rsPayment.Close
         'YOu need to update sc with this on payment side
         'Need to do the same on SI side
         adoConn.Execute "Update tblPurInv P set ReportPaymentAmount= " & -dblPayAmount & ", ReportPayDescription='" & _
                szPDesc & "',ReportNominalCode='" & szNominalCode & "',ReportInvNetAmount= " & -dblPayAmount & ",ReportINVVATAmount= " & _
                dblinvVATamount & " where P.My_ID='" & rstblPurInv("My_ID").Value & "'"

         rstblPurInv.MoveNext
    Wend
   adoConn.Close
    Sleep (100)
    Set Report = reportApp.OpenReport(App.Path & szReportPath & "\ClientStatementPreview.rpt")
    Dim dblPercentage As Double
    Report.EnableParameterPrompting = False
    Report.DiscardSavedData
    Report.ParameterFields(1).AddCurrentValue CInt(1)
    Report.ParameterFields(2).AddCurrentValue szSelectedClient 'client ID
    Report.ParameterFields(3).AddCurrentValue CDate(txtStatementDate1.text)  'statement date
    Report.ParameterFields(4).AddCurrentValue CDate(txtLastStatementDate1.text)  'Previuos statement date
    Report.ParameterFields(5).AddCurrentValue 100 '100 Percent
    Report.ParameterFields(6).AddCurrentValue "0" '0 is for detail record so print address for passing client ID in parameter 2
    Report.ParameterFields(8).AddCurrentValue IIf(chkShowDue.Value = 1, True, False)
    adoConn.Open getConnectionString
    Report.ParameterFields(7).AddCurrentValue findClientaddress(adoConn, szSelectedClient)
    adoConn.Close
    Load frmReport
    frmReport.LoadReportViewer Report
   
End Sub
Private Function findClientaddress(adoConn As ADODB.Connection, ClientID As String) As String
    Dim rsClient As New ADODB.Recordset
    rsClient.Open "Select ClientAddressLine1,ClientAddressLine2,ClientAddressLine3,ClientAddressLine4,Client.ClientPostCode from Client where ClientID='" & ClientID & "'", adoConn, adOpenStatic, adLockReadOnly
    If Not rsClient.EOF Then
            findClientaddress = rsClient("ClientAddressLine1").Value + Chr(10) + Chr(13) + rsClient("ClientAddressLine2").Value + Chr(10) + Chr(13) + rsClient("ClientAddressLine3").Value + Chr(10) + Chr(13) + rsClient("ClientAddressLine4").Value + Chr(10) + Chr(13) + rsClient("ClientPostCode").Value + Chr(10) + Chr(13)
    End If
    rsClient.Close
End Function
Private Function findLandlordAddress(adoConn As ADODB.Connection, landLordID As String) As String
    Dim rsClient As New ADODB.Recordset
    rsClient.Open "Select LandlordAddressLine1,LandlordAddressLine2,LandlordAddressLine3,LandlordAddressLine4,LandlordPostCode from Landlord where LandlordID='" & landLordID & "'", adoConn, adOpenStatic, adLockReadOnly
    If Not rsClient.EOF Then
            findLandlordAddress = rsClient("LandlordAddressLine1").Value + Chr(10) + Chr(13) + rsClient("LandlordAddressLine2").Value + Chr(10) + Chr(13) + rsClient("LandlordAddressLine3").Value + Chr(10) + Chr(13) + rsClient("LandlordAddressLine4").Value + Chr(10) + Chr(13) + rsClient("LandlordPostCode").Value + Chr(10) + Chr(13)
    End If
    rsClient.Close
End Function

Private Sub cmdPrintThis_Click()
   Dim iIncDec As Long
   iIncDec = 0
   Dim rCount As Integer
   Dim selRow As Integer
   Dim adoConn As New ADODB.Connection
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

    szCurrentStatementID = frmRentPayable.flxPayFees.TextMatrix(selRow, 2)
    'run TestReportForRentSummary.rpt
    Dim reportApp As New CRAXDRT.Application
    Dim Report As CRAXDRT.Report
    'Dim adoconn As New ADODB.Connection
    adoConn.Open getConnectionString
    Dim rsDemandSplit As New ADODB.Recordset
    Dim rsReceived As New ADODB.Recordset
    Dim rsRentSummaryStatement As New ADODB.Recordset
    Dim dblReceivedAmt As Double
    Dim szListofFunds As String
    Dim szTypeOfDemanddesc As String
    Dim dateFrom As String
    
    Dim DateTO As String
    If szCurrentStatementID = "" Then
        MsgBox "Please select a statement", vbInformation, "Warning"
        Exit Sub
    End If
    szCurrentStatementID = Replace(szCurrentStatementID, "CS", "")
    rsRentSummaryStatement.Open "Select ListOfFundId  from RentSummaryStatement where statementID=" & szCurrentStatementID & "", adoConn, adOpenStatic, adLockReadOnly ' group by D.DemandId", adoconn, adOpenStatic, adLockReadOnly
    If Not rsRentSummaryStatement.EOF Then
            szListofFunds = rsRentSummaryStatement("ListOfFundId").Value
    End If
    
    
    rsRentSummaryStatement.Close
    
    rsDemandSplit.Open "Select D.DemandId,sum(Amount) as NAmt,sum(VATAmount)as TVAT from  DemandSplitRecords DS,DemandRecords D,Units U,Property P where D.DemandID=DS.DemandID " & _
                " and U.UnitNumber=D.UnitNumber AND P.PropertyID=U.PropertyID   AND P.ClientID='" & frmRentPayable.flxPayFees.TextMatrix(selRow, 4) & "' group by D.DemandId", adoConn, adOpenStatic, adLockReadOnly
    
    Dim rsDemandSplit1 As New ADODB.Recordset
     
    While Not rsDemandSplit.EOF
            rsReceived.Open "Select DemandId,sum(R.Amount) as ReceivedAmt from  tlbReceiptSplit R,tlbReceipt RC, RptTransactionsSplit T,DemandRecords D,tlbReceipt RL where T.Deleteflag=False " & _
                 "AND RL.transactionID=T.FromTran AND T.transactionID=R.RptTransactionsIDSplit AND RC.transactionID=T.ToTran AND R.FundID in ( " & szListofFunds & ")  " & _
                 "AND RL.Type in(3,4,23) and RC.DemandRef=D.DemandID and D.DemandId=" & rsDemandSplit("DEMANDID").Value & " group by DemandID", adoConn, adOpenStatic, adLockReadOnly
                 dblReceivedAmt = 0 'AND R.FundID in ( " & szListofFunds & ")'AND RC.Type in(3,4,23)'sum(switch(RC.Type=3,R.Amount,RC.Type=4,R.Amount,RC.Type=23,-R.Amount))
                 If Not rsReceived.EOF Then
                        dblReceivedAmt = IIf(IsNull(rsReceived("ReceivedAmt").Value), 0, rsReceived("ReceivedAmt").Value)
                 End If
                 rsReceived.Close
                 dateFrom = ""
                 DateTO = ""
                 rsDemandSplit1.Open "Select DateFrom,DateTo,Type from  DemandSplitRecords D,DemandTypes DT where DT.ID=D.TypeOfDemand AND DemandId=" & rsDemandSplit("DEMANDID").Value & " order by D.SPlitID ", adoConn, adOpenStatic, adLockReadOnly
                 If Not rsDemandSplit1.EOF Then
                        dateFrom = rsDemandSplit1("DateFrom").Value
                        DateTO = rsDemandSplit1("DateTo").Value
                        szTypeOfDemanddesc = rsDemandSplit1("Type").Value
                 End If
                 rsDemandSplit1.Close
                 If dateFrom = "" Then
                            MsgBox "Date from in the demand split is empty", vbInformation, "Warning"
                        Exit Sub
                 End If
                 
            adoConn.Execute "Update DemandRecords set ReportNetAmount= " & rsDemandSplit("NAmt").Value & ",ReportVATAmount= " & _
                    rsDemandSplit("TVAT").Value & ",ReportReceivedAmount= " & dblReceivedAmt & ",ReportDateFrom=#" & Format(dateFrom, "dd MMM yyyy") _
                    & "#,ReportDateTO =#" & Format(DateTO, "dd MMM yyyy") & "#,ReportDemandTypeDesc= '" & szTypeOfDemanddesc & "' where DemandId=" & rsDemandSplit("DEMANDID").Value & ""
                    
                    
            rsDemandSplit.MoveNext
    Wend
    rsDemandSplit.Close
    Dim rstblPurInv As New ADODB.Recordset
    Dim rsPayment As New ADODB.Recordset
    Dim rsInv As New ADODB.Recordset
    Dim dblPayAmount As Double
    Dim dblinvNetamount As Double
    Dim dblinvVATamount As Double
    Dim szPDesc As String
    Dim szNominalCode As String
    rstblPurInv.Open "Select P.TransactionID,PI.MY_ID from tblPurInv PI,tlbPayment P where PI.MY_ID=P.PI AND CL_ID='" & frmRentPayable.flxPayFees.TextMatrix(selRow, 4) & "'", adoConn, adOpenStatic, adLockReadOnly
    While Not rstblPurInv.EOF
         dblPayAmount = 0
         rsPayment.Open "Select sum(T.PaymentAmount) as amt from tlbPayment P,PayTransactionsSplit T,tlbPaymentSplit S where S.payHeader=P.TransactionID and " & _
                " S.PayTransactionIDSplit=T.TransactionID AND S.FundID in ( " & szListofFunds & ") AND T.FromTran=P.TransactionID and T.Deleteflag=false and P.amount>P.osamount AND T.ToTran=" & _
                    rstblPurInv("TransactionID").Value & "", adoConn, adOpenStatic, adLockReadOnly
         If Not rsPayment.EOF Then
                dblPayAmount = IIf(IsNull(rsPayment("amt").Value), 0, rsPayment("amt").Value)
         End If
         rsPayment.Close
         rsPayment.Open "Select sum(NET_AMOUNT) as amt, sum(VAT) as VAT1 from tblPurInvSRec S where ParentID='" & rstblPurInv("My_ID").Value & "'", adoConn, adOpenStatic, adLockReadOnly
         If Not rsPayment.EOF Then
                dblinvNetamount = IIf(IsNull(rsPayment("amt").Value), 0, rsPayment("amt").Value)
                dblinvVATamount = IIf(IsNull(rsPayment("VAT1").Value), 0, rsPayment("VAT1").Value)
         End If
         rsPayment.Close
         szPDesc = ""
         rsPayment.Open "Select Description,NOMINAL_CODE from tlbPaymentSplit P where P.PayHeader=" & _
                    rstblPurInv("TransactionID").Value & "", adoConn, adOpenStatic, adLockReadOnly
         If Not rsPayment.EOF Then
              szPDesc = IIf(IsNull(rsPayment("Description").Value), 0, rsPayment("Description").Value)
              szPDesc = Replace(szPDesc, "'", "")
              szNominalCode = IIf(IsNull(rsPayment("NOMINAL_CODE").Value), 0, rsPayment("NOMINAL_CODE").Value)
         End If
         rsPayment.Close
         
         adoConn.Execute "Update tblPurInv P set ReportPaymentAmount= " & dblPayAmount & ", ReportPayDescription='" & _
                szPDesc & "',ReportNominalCode='" & szNominalCode & "',ReportInvNetAmount= " & dblinvNetamount & ",ReportINVVATAmount= " & _
                dblinvVATamount & " where P.My_ID='" & rstblPurInv("My_ID").Value & "'"
                
         rstblPurInv.MoveNext
    Wend
    rstblPurInv.Close
    adoConn.Close
  
    
    Set Report = reportApp.OpenReport(App.Path & szReportPath & "\ClientStatement.rpt")
    Dim dblPercentage As Double
    Report.EnableParameterPrompting = False
    Report.DiscardSavedData
    'Report.ParameterFields(1).AddCurrentValue CInt(Right(szCurrentStatementID, Len(szCurrentStatementID) - 2))
     Report.ParameterFields(1).AddCurrentValue CInt(Trim(Replace(szCurrentStatementID, "CS", "")))
    
'     Report.ParameterFields(4).AddCurrentValue CDate(IIf(frmRentPayable.flxPayFees.TextMatrix(selRow, 6) = "", "01-01-1900", frmRentPayable.flxPayFees.TextMatrix(selRow, 6))) 'Previuos statement date
     Dim selRowTemp As Integer
     selRowTemp = selRow
     If frmRentPayable.flxPayFees.TextMatrix(selRow, 1) = "+" Or frmRentPayable.flxPayFees.TextMatrix(selRow, 1) = ">" Then
             Report.ParameterFields(2).AddCurrentValue frmRentPayable.flxPayFees.TextMatrix(selRow, 4) 'client ID
            Report.ParameterFields(3).AddCurrentValue CDate(frmRentPayable.flxPayFees.TextMatrix(selRow, 7)) 'statement date
            Report.ParameterFields(4).AddCurrentValue CDate(IIf(frmRentPayable.flxPayFees.TextMatrix(selRow, 6) = "", "01-01-1900", frmRentPayable.flxPayFees.TextMatrix(selRow, 6))) 'Previuos statement date
            Report.ParameterFields(5).AddCurrentValue 100 '100 Percent
            Report.ParameterFields(6).AddCurrentValue "0" '0 is for detail record so print address for passing client ID in parameter 2
            adoConn.Open getConnectionString
            Report.ParameterFields(7).AddCurrentValue findClientaddress(adoConn, frmRentPayable.flxPayFees.TextMatrix(selRow, 4))
            adoConn.Close
      Else
            dblPercentage = Replace(frmRentPayable.flxPayFees.TextMatrix(selRow, 9), "%", "") 'Take Percenatge from Grid
            Do
                selRowTemp = selRowTemp - 1
            Loop Until (frmRentPayable.flxPayFees.TextMatrix(selRowTemp, 1) = "+" Or frmRentPayable.flxPayFees.TextMatrix(selRowTemp, 1) = ">")
            Report.ParameterFields(2).AddCurrentValue frmRentPayable.flxPayFees.TextMatrix(selRowTemp, 4)  'client ID
            Report.ParameterFields(3).AddCurrentValue CDate(frmRentPayable.flxPayFees.TextMatrix(selRowTemp, 7)) 'statement date
            Report.ParameterFields(4).AddCurrentValue CDate(IIf(frmRentPayable.flxPayFees.TextMatrix(selRowTemp, 6) = "", "01-01-1900", frmRentPayable.flxPayFees.TextMatrix(selRowTemp, 6))) 'Previuos statement date
            Report.ParameterFields(5).AddCurrentValue dblPercentage
            Report.ParameterFields(6).AddCurrentValue "1" '1 is for detail record so print address for passing clientlnadlord ID
            'Report.ParameterFields(7).AddCurrentValue frmRentPayable.flxPayFees.TextMatrix(selRow, 3) 'take clientlnadlord ID from grid
            adoConn.Open getConnectionString
            Report.ParameterFields(7).AddCurrentValue findLandlordAddress(adoConn, frmRentPayable.flxPayFees.TextMatrix(selRow, 3)) 'take clientlnadlord ID from grid
            adoConn.Close
      End If
'        If frmRentPayable.flxPayFees.TextMatrix(selRow, 1) = "-" Then
'                 Report.ParameterFields(5).AddCurrentValue 100
'        Else
'                 dblPercentage = StrDigitVal(frmRentPayable.flxPayFees.TextMatrix(selRow, 9))
'                 Report.ParameterFields(5).AddCurrentValue dblPercentage
'        End If
        
       'Report.ParameterFields(4).AddCurrentValue CDate(IIf(frmRentPayable.flxPayFees.TextMatrix(selRow, 6) = "", "01-01-1900", frmRentPayable.flxPayFees.TextMatrix(selRow, 6)))
      
        'Report.ParameterFields(3).AddCurrentValue frmRentPayable.flxPayFees.TextMatrix(selRow, 3)
    
    '               Report.ParameterFields(1).AddCurrentValue CStr(txtLLID.text)
    '               Report.ParameterFields(2).AddCurrentValue CDate(txtFromDate.text)
    '               Report.ParameterFields(3).AddCurrentValue CDate(txtToDate.text)
    '                Report.ParameterFields(4).AddCurrentValue cboCategory.text
    Load frmReport
    frmReport.LoadReportViewer Report
End Sub
'Private

'Function ListOfPayableTypesForDBSave() As String
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
   'ListOfProperties = ("''," ' This shall always include No Property
    'ListOfProperties = "''," ' This shall always include No Property
   For i = 1 To flxProperties.Rows - 1
      If flxProperties.TextMatrix(i, 0) = "X" Then
         ListOfProperties = ListOfProperties & " '" & flxProperties.TextMatrix(i, 1) & "', "
      End If
   Next i
   If Len(ListOfProperties) > 0 Then ListOfProperties = Left(ListOfProperties, Len(ListOfProperties) - 2)
End Function
'Private Function ListOfPropertiesIncludingNull() As String
'   Dim i As Integer
'   'ListOfProperties = ("''," ' This shall always include No Property
'    ListOfPropertiesIncludingNull = "(isnull(S.PropertyID) OR ''), " ' This shall always include No Property
'   For i = 1 To flxProperties.Rows - 1
'      If flxProperties.TextMatrix(i, 0) = "X" Then
'         ListOfPropertiesIncludingNull = ListOfPropertiesIncludingNull & " '" & flxProperties.TextMatrix(i, 1) & "', "
'      End If
'   Next i
'   If Len(ListOfPropertiesIncludingNull) > 0 Then ListOfPropertiesIncludingNull = Left(ListOfPropertiesIncludingNull, Len(ListOfPropertiesIncludingNull) - 2)
'End Function
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
    Dim adoConn As New ADODB.Connection
    adoConn.Open getConnectionString
    PurchaseLedgerControl = GetNominalCodeForControlAccount(adoConn, "Purchase Ledger Control", szSelectedClient)
    Dim rsNLposting As New ADODB.Recordset
    rsNLposting.Open "Select sum(AMOUNT) as dr from NLPosting where ACCOUNT_NUMBER ='" & _
                    szSelectedBankAccount & "'  AND NOMINAL_CODE='" & PurchaseLedgerControl & "' AND ClientID='" & _
                    szSelectedClient & "' ", adoConn, adOpenStatic, adLockReadOnly
    
    If Not rsNLposting.EOF Then
        dblAmt = IIf(IsNull(rsNLposting.Fields.Item("Dr").Value), 0, rsNLposting.Fields.Item("Dr").Value)
    End If
    rsNLposting.Close
    Set rsNLposting = Nothing
    adoConn.Close
    Set adoConn = Nothing
    GetBalanceSupplier = dblAmt
End Function
Private Function GetBalanceAgent() As Double
    Dim ManagementFeesControl As String
    Dim dblAmt As Double
    Dim adoConn As New ADODB.Connection
    adoConn.Open getConnectionString
    ManagementFeesControl = GetNominalCodeForControlAccount(adoConn, "Managing Agents control Account (B/S)", szSelectedClient)
    Dim rsNLposting As New ADODB.Recordset
    rsNLposting.Open "Select sum(AMOUNT) as dr from NLPosting where ACCOUNT_NUMBER ='" & _
                    szSelectedBankAccount & "'  AND NOMINAL_CODE='" & ManagementFeesControl & "' AND ClientID='" & _
                    szSelectedClient & "' ", adoConn, adOpenStatic, adLockReadOnly
    
    If Not rsNLposting.EOF Then
        dblAmt = IIf(IsNull(rsNLposting.Fields.Item("Dr").Value), 0, rsNLposting.Fields.Item("Dr").Value)
    End If
    rsNLposting.Close
    Set rsNLposting = Nothing
    adoConn.Close
    Set adoConn = Nothing
    GetBalanceAgent = dblAmt
End Function

'Private Function isAnyTransactionAvailable() As Boolean
'        Dim adoConn As New ADODB.Connection
'        Dim whereProperty As String
'        Dim szSQL As String
'        Dim rsReceipt As New ADODB.Recordset
'        Dim dblAmt As Double
'        adoConn.Open getConnectionString
'        'We cannot filter by fund here because we are not counting tlbPaymentSplit in this SQL
'        whereProperty = "(P.UNITID IN (" & ListOfProperties & ") OR isnull(P.UNITID) or P.UNITID='' ) AND "
'        szSQL = "Select   SUM(SWITCH(P.TYPE=24,S.Amount,P.TYPE=8,-S.Amount,P.TYPE=9,-S.Amount)) as AMT  from tlbPayment P,tlbPaymentSplit S,Supplier SP where " & _
'                " S.Payheader=P.TransactionID and P.Type in (8,9,24)  AND SP.SupplierID=P.SageaccountNumber AND " & whereProperty & " " & _
'                " P.PDate >#" & Format(txtLastStatementDate1.text, "dd/mmm/yyyy") & "# AND P.PDate <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "# AND " & _
'                " P.ClientID ='" & szSelectedClient & "'"
'
'                '" R.RDate >=#" & Format(txtLastStatementDate1.text, "dd/mmm/yyyy") & "# AND R.RDate <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "#" & _
'        rsReceipt.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
'        rsReceipt.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'        If Not rsReceipt.EOF Then
'             dblAmt = IIf(IsNull(rsReceipt.Fields.Item("amt").Value), 0, rsReceipt.Fields.Item("amt").Value)
'        End If
'        rsReceipt.Close
'        whereProperty = "(U.propertyID IN (" & ListOfProperties & ") OR isnull(R.UNITID) or R.UNITID='' ) AND "
'        szSQL = "Select  sum(SWITCH(R.TYPE =3,RS.Amount,R.TYPE =4,RS.Amount,R.TYPE =23,-RS.Amount))   as DR from tlbReceipt R,TlbReceiptSplit RS,Fund F, Units U " & _
'                "where RS.rptHeader=R.transactionID  and R.Type in (3,4,23)  AND R.UnitID=U.UnitNumber AND " & whereProperty & " F.FundCode IN(" & ListOfFunds & ") AND  RS.FundID=F.FundID AND ClientID ='" & _
'                 szSelectedClient & "' AND R.RDate >#" & Format(txtLastStatementDate1.text, "dd/mmm/yyyy") & "# AND R.RDate <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "#"
'
'        rsReceipt.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'        If Not rsReceipt.EOF Then
'             dblAmt = dblAmt + IIf(IsNull(rsReceipt.Fields.Item("Dr").Value), 0, rsReceipt.Fields.Item("Dr").Value)
'        End If
'        rsReceipt.Close
'
'        whereProperty = "(B.PROPERTYID IN (" & ListOfProperties & ") OR isnull(B.PROPERTYID) or B.PROPERTYID='' ) AND "
'        szSQL = "Select  SUM(B.NET_AMOUNT) as DR from tlbBankPayment B,Fund F where TransactionType " & _
'                "IN(11,12) AND " & whereProperty & " B.DEPT_ID=cstr(F.FundID) and " & _
'                "B.TRAN_DATE >#" & Format(txtLastStatementDate1.text, "dd/mmm/yyyy") & "# AND B.TRAN_DATE <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "# " & _
'                "AND B.ClientID ='" & szSelectedClient & "'"
'
'        rsReceipt.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'        If Not rsReceipt.EOF Then
'             dblAmt = dblAmt + IIf(IsNull(rsReceipt.Fields.Item("Dr").Value), 0, rsReceipt.Fields.Item("Dr").Value)
'        End If
'        rsReceipt.Close
'        If dblAmt > 0 Then
'            isAnyTransactionAvailable = True
'        End If
'        adoConn.Close
'        Set adoConn = Nothing
'End Function
'Private Function GetRentDeposit() As Double
'    Dim szSQL As String
''    Dim szSQL1 As String
'    Dim szSQL2 As String
''    Dim szSQL3 As String
'    Dim rsPayment As New ADODB.Recordset
'    Dim rsReceipt1 As New ADODB.Recordset
'    Dim rsReceipt2 As New ADODB.Recordset
'    Dim rsReceipt3 As New ADODB.Recordset
'    Dim rsReceipt As New ADODB.Recordset
'    Dim adoConn As New ADODB.Connection
'    Dim dblAmt, dblamt1, dblamt2, dblamt3 As Double
'    adoConn.Open getConnectionString
'    'tlbBankPayment
'    'BANK_AC
'    'TRAN_TYPE
'    'DEPT_ID
'    'propertyID
'    'clientID
'    'NET_AMOUNT
'    Dim whereProperty  As String
'    If boolConsolidatedStatement = 1 Then
'            whereProperty = "(U.PROPERTYID IN (" & ListOfProperties & ") OR isnull(U.PROPERTYID) or U.PROPERTYID='' ) AND "
'    Else
'            whereProperty = "U.PROPERTYID  in (" & ListOfProperties & ") AND "
'    End If
'
'    'From unit Id i Ned to build a relation with the selected properties
'    szSQL = "Select  SUM(SWITCH(TYPE=1,R.Amount,TYPE=2,R.Amount,TYPE=3,-R.Amount,TYPE=4,-R.Amount,TYPE=23,-R.Amount)) as DR from tlbReceipt R,Fund F, Units U " & _
'            "where R.UnitID=U.UnitNumber AND " & whereProperty & "  TYPE IN(1,2,3,4,23) AND R.FundID=F.FundID and F.FundCode='RENTDEPOSIT' AND ClientID ='" & _
'             szSelectedClient & "' AND  R.RDATE >#" & Format(txtLastStatementDate1.text, "dd/mmm/yyyy") & "# AND  R.RDATE <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "#"
''    szSQL1 = "Select  SUM(R.Amount)  as DR from tlbReceipt R,Fund F where TYPE IN(3,4,23) AND R.FundID=F.FundID and F.FundCode='RENTDEPOSIT' AND ClientID ='" & _
''              szSelectedClient & "'"
'    'PropertyID
'    If boolConsolidatedStatement = 1 Then
'            whereProperty = "(B.PROPERTYID IN (" & ListOfProperties & ") OR isnull(B.PROPERTYID) or B.PROPERTYID='' ) AND "
'    Else
'            whereProperty = "B.PROPERTYID  in (" & ListOfProperties & ") AND "
'    End If
'
'
'    szSQL2 = "Select  SUM(SWITCH(TransactionType=11,B.NET_AMOUNT,TransactionType=12,-B.NET_AMOUNT)) as DR from tlbBankPayment B,Fund F where TransactionType " & _
'            " IN(11,12) AND " & whereProperty & " B.DEPT_ID=cstr(F.FundID) and " & _
'            "F.FundCode='RENTDEPOSIT' AND B.ClientID ='" & szSelectedClient & "' AND B.TRAN_DATE >#" & Format(txtLastStatementDate1.text, "dd/mmm/yyyy") & "# AND  B.TRAN_DATE<=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "#"
''    szSQL3 = "Select  SUM(B.NET_AMOUNT)  as DR from tlbBankPayment B,Fund F where TransactionType IN(12) AND B.DEPT_ID= cstr(F.FundID) and " & _
''            "F.FundCode='RENTDEPOSIT' AND B.ClientID ='" & szSelectedClient & "'"
'
'   ' vat Deducting  on reciept
''            szSQL = "Select  SUM(AL.VatAmount) as DR from tlbReceipt R,rptTransactions AL,tlbReceiptSplit S,Fund F, Units B,GLobalData G where G.PropertyID=B.PropertyID AND G.isAgentToSubmit=true AND R.TransactionID= S.RptHeader AND TYPE IN(3,4,23) " & _
''            "AND S.FundID=F.FundID AND R.BankCODE='" & szSelectedBankAccount & "' and F.FundCode<>'TENANTDEPOSIT'  AND (R.RentSumStatement='' or isnull(R.RentSumStatement)) AND ClientID ='" & szSelectedClient & "' " & _
''            "AND R.UnitID=B.UnitNumber and AL.Deleteflag=false and AL.FromTran=R.TransactionID AND  F.FundCode in (" & ListOfFunds & ") AND " & whereProperty & " R.RDate >=#" & Format(txtLastStatementDate1.text, "dd/mmm/yyyy") & "# AND R.RDate <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "#"
''            rsReceipt.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
''            If Not rsReceipt.EOF Then
''                    dblAmt = dblAmt - IIf(IsNull(rsReceipt.Fields.Item("Dr").Value), 0, rsReceipt.Fields.Item("Dr").Value)
''                     'result is 175
''            End If
''            rsReceipt.Close
''            Set rsReceipt = Nothing
'
'
'
'    rsReceipt.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'    If Not rsReceipt.EOF Then
'        dblAmt = IIf(IsNull(rsReceipt.Fields.Item("Dr").Value), 0, rsReceipt.Fields.Item("Dr").Value)
'    End If
'    rsReceipt.Close
''    rsReceipt1.Open szSQL1, adoconn, adOpenStatic, adLockReadOnly
''    If Not rsReceipt1.EOF Then
''         dblamt1 = IIf(IsNull(rsReceipt1.Fields.Item("Dr").Value), 0, rsReceipt1.Fields.Item("Dr").Value)
''    End If
''    rsReceipt1.Close
'
'    rsReceipt2.Open szSQL2, adoConn, adOpenStatic, adLockReadOnly
'    If Not rsReceipt2.EOF Then
'        dblamt2 = IIf(IsNull(rsReceipt2.Fields.Item("Dr").Value), 0, rsReceipt2.Fields.Item("Dr").Value)
'    End If
'    rsReceipt2.Close
'
''    rsReceipt3.Open szSQL3, adoconn, adOpenStatic, adLockReadOnly
''    If Not rsReceipt3.EOF Then
''        dblamt3 = IIf(IsNull(rsReceipt3.Fields.Item("Dr").Value), 0, rsReceipt3.Fields.Item("Dr").Value)
''    End If
''    rsReceipt3.Close
''    dblamt2 = 0
''    dblamt1 = 0
''    dblAmt = 0
'    GetRentDeposit = dblAmt + dblamt2
'    adoConn.Close
'    Set adoConn = Nothing
'End Function
Private Function GetBalance(szType As String) As Double
    'szType is the suppplier type from supplier table
   Dim szSQL   As String
   Dim szSqlPI As String
   Dim szSQLSI As String
   Dim i       As Integer
   Dim iSI     As Integer
   Dim iPI     As Integer
   Dim iIndex  As Integer
   Dim adoConn As New ADODB.Connection
   Dim adoPayDr As New ADODB.Recordset, adoPayCr As New ADODB.Recordset
   Dim adoRptDr As New ADODB.Recordset, adoRptCr As New ADODB.Recordset
   adoConn.Open getConnectionString
   Dim szaClientBal(1, 1) As String

   szSQL = "SELECT  SUM(P.Amount) AS Dr " & _
           "FROM tlbPayment AS P, Client C, Supplier S " & _
           "WHERE (P.Type = 6 OR P.Type = 24) AND C.ClientID=S.SupplierID AND P.SageAccountNumber = C.ClientID " & _
           "and  C.ClientID='" & szSelectedClient & "' AND S.Type='" & szType & "'  "

   adoPayDr.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

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

   adoPayCr.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   While Not adoPayCr.EOF
         szaClientBal(1, iIndex) = IIf(IsNull(adoPayCr.Fields.Item("Cr").Value), 0, adoPayCr.Fields.Item("Cr").Value) 'adoPayCr.Fields.Item("Cr").Value
         adoPayCr.MoveNext
   Wend

   adoPayCr.Close
   GetBalance = szaClientBal(1, 0)
   Set adoRptDr = Nothing
   Set adoRptCr = Nothing
   adoConn.Close
   Set adoConn = Nothing
End Function
Private Function GetManagingAgentACBalance() As Double
    Dim adoConn As New ADODB.Connection
    Dim rsAccrualsControlBalance As New ADODB.Recordset
    adoConn.Open getConnectionString
    Dim szSQL As String
'    szSQL = "Select Sum(AMOUNT) as SumAmount from NLPOSTING where NOMINAL_CODE='" & _
'            NominalCode & "' AND ClientID='" & szSelectedClient & "' AND PROPERTY_ID in (" & ListOfProperties & ")"
    rsAccrualsControlBalance.Open szSQL, adoConn, adOpenKeyset, adLockReadOnly
    If Not rsAccrualsControlBalance.EOF Then
        GetManagingAgentACBalance = rsAccrualsControlBalance("SumAmount").Value
    End If
    rsAccrualsControlBalance.Close
    Set rsAccrualsControlBalance = Nothing
End Function
Private Function GetAccrualsControlBalance() As Double
    'include no property when  calculating accruals
    Dim adoConn As New ADODB.Connection
    Dim rsAccrualsControlBalance As New ADODB.Recordset
    adoConn.Open getConnectionString
    Dim szSQL As String
    Dim AccrualsCode As String
    If ListOfProperties = "'" Then Exit Function
    AccrualsCode = GetNominalCodeForControlAccount(adoConn, "Accruals Control Account (B/S)", szSelectedClient)
    
    szSQL = "Select Sum(AMOUNT) as SumAmount from NLPOSTING where NOMINAL_CODE='" & AccrualsCode & "' AND ClientID='" & szSelectedClient & "' AND PROPERTY_ID in (" & ListOfProperties & ")"
    rsAccrualsControlBalance.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
    If Not rsAccrualsControlBalance.EOF Then
        GetAccrualsControlBalance = IIf(IsNull(rsAccrualsControlBalance("SumAmount").Value), 0, rsAccrualsControlBalance("SumAmount").Value)
    End If
    rsAccrualsControlBalance.Close
    Set rsAccrualsControlBalance = Nothing
End Function
Private Function GetAccrualsControlBalanceNonConsolidated(szPropertyID) As Double
    'include no property when  calculating accruals
    Dim adoConn As New ADODB.Connection
    Dim rsAccrualsControlBalance As New ADODB.Recordset
    adoConn.Open getConnectionString
    Dim szSQL As String
    Dim AccrualsCode As String
    If ListOfProperties = "'" Then Exit Function
    AccrualsCode = GetNominalCodeForControlAccount(adoConn, "Accruals Control Account (B/S)", szSelectedClient)
    
    szSQL = "Select Sum(AMOUNT) as SumAmount from NLPOSTING where NOMINAL_CODE='" & AccrualsCode & "' AND ClientID='" & szSelectedClient & "' AND PROPERTY_ID ='" & szPropertyID & "'"
    rsAccrualsControlBalance.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
    If Not rsAccrualsControlBalance.EOF Then
        GetAccrualsControlBalanceNonConsolidated = IIf(IsNull(rsAccrualsControlBalance("SumAmount").Value), 0, rsAccrualsControlBalance("SumAmount").Value)
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
    Dim adoConn As New ADODB.Connection
    Dim rsRentSummaryStatement As New ADODB.Recordset
    adoConn.Open getConnectionString
    Dim szSQL As String
    'Rent Summary Statement Opening Balance (=Closing balance of previous statement)
    'This is by client
    'Get ID by Client max ID from RentSummaryStatement
    szSQL = "Select max(StatementNo) as IDbyCL from RentSummaryStatement where ClientIDLandlordID='" & szSelectedClient & "'"
    rsRentSummaryStatement.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
    If Not rsRentSummaryStatement.EOF Then
          GetLastStatementNoByClient = IIf(IsNull(rsRentSummaryStatement!IDbyCL), 0, rsRentSummaryStatement!IDbyCL)
    Else
          GetLastStatementNoByClient = "1"
    End If
    rsRentSummaryStatement.Close
    Set rsRentSummaryStatement = Nothing
    adoConn.Close
    Set adoConn = Nothing
End Function
Private Function GetLastStatementID() As Long 'this is not by client
    Dim intmaxStatementNo As Integer
    Dim adoConn As New ADODB.Connection
    Dim rsRentSummaryStatement As New ADODB.Recordset
    adoConn.Open getConnectionString
    Dim szSQL As String
    'Rent Summary Statement Opening Balance (=Closing balance of previous statement)
    'This is by client
    'Get ID by Client max ID from RentSummaryStatement
    szSQL = "Select max(StatementID) as IDbyCL from RentSummaryStatement"
    rsRentSummaryStatement.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
    If Not rsRentSummaryStatement.EOF Then
        GetLastStatementID = IIf(IsNull(rsRentSummaryStatement!IDbyCL), 0, rsRentSummaryStatement!IDbyCL)
    End If
    rsRentSummaryStatement.Close
    Set rsRentSummaryStatement = Nothing
    adoConn.Close
    Set adoConn = Nothing
End Function
Private Function isLastStatementFinalized() As Long 'this is not by client
    Dim intmaxStatementNo As Integer
    Dim adoConn As New ADODB.Connection
    Dim rsRentSummaryStatement As New ADODB.Recordset
    adoConn.Open getConnectionString
    Dim szSQL As String
    'Rent Summary Statement Opening Balance (=Closing balance of previous statement)
    'This is by client
    'Get ID by Client max ID from RentSummaryStatement
    szSQL = "Select max(StatementID) as IDbyCL from RentSummaryStatement where clientID =" & szSelectedClient & ""
    rsRentSummaryStatement.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
    If Not rsRentSummaryStatement.EOF Then
        isLastStatementFinalized = IIf(IsNull(rsRentSummaryStatement!IDbyCL), 0, rsRentSummaryStatement!IDbyCL)
    End If
'    If isLastStatementFinalized = 0 Then
'        'if last statement is zero that means you are checking it is always true=fine to move forward
'    Else
'        'if this is 1 or more than 2 that means you are selecting the last statement
'
'    End If
    rsRentSummaryStatement.Close
    Set rsRentSummaryStatement = Nothing
    adoConn.Close
    Set adoConn = Nothing
End Function
Private Function ReportGenID() As Long 'this is not by for one report one ID one time one print
    Dim intmaxStatementNo As Integer
    Dim adoConn As New ADODB.Connection
    Dim rsRentSummaryStatement As New ADODB.Recordset
    adoConn.Open getConnectionString
    Dim szSQL As String
    szSQL = "Select max(ReportGenID)+1 as IDbyCL from RentSummaryStatement"
    rsRentSummaryStatement.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
    If Not rsRentSummaryStatement.EOF Then
        ReportGenID = IIf(IsNull(rsRentSummaryStatement!IDbyCL), 0, rsRentSummaryStatement!IDbyCL)
    End If
    rsRentSummaryStatement.Close
    Set rsRentSummaryStatement = Nothing
    adoConn.Close
    Set adoConn = Nothing
End Function
Private Function GetLastStatementDateByClient(ByVal statementID As Long) As String
    Dim intmaxStatementNo As Integer
    Dim adoConn As New ADODB.Connection
    Dim rsRentSummaryStatement As New ADODB.Recordset
    adoConn.Open getConnectionString
    Dim szSQL As String
    'Rent Summary Statement Opening Balance (=Closing balance of previous statement)
    'This is by client
    'Get ID by Client max ID from RentSummaryStatement
    szSQL = "Select StatementDate from RentSummaryStatement where ClientIDLandlordID='" & szSelectedClient & "' AND  statementID<" & statementID & " order by StatementNo Desc"
    rsRentSummaryStatement.Open szSQL, adoConn, adOpenKeyset, adLockReadOnly
    If Not rsRentSummaryStatement.EOF Then
        GetLastStatementDateByClient = rsRentSummaryStatement!StatementDate
    End If
    rsRentSummaryStatement.Close
    Set rsRentSummaryStatement = Nothing
    adoConn.Close
    Set adoConn = Nothing
End Function
Private Sub cmdPostRP_Click()
   If MsgBox("Do you want to post to the history?", vbQuestion + vbYesNo, "Post") = vbNo Then Exit Sub
   Dim adoConn As New ADODB.Connection
   Dim adoRst As ADODB.Recordset
   Dim sSQLQuery As String

   adoConn.Open getConnectionString
   Set adoRst = New ADODB.Recordset

   sSQLQuery = "UPDATE tblPurInv " & _
               "SET tblPurInv.HISTORY = TRUE " & _
               "WHERE tblPurInv.TTP = " & CByte(TransactionTakePlace("TTP", "RENT PAYABLE", adoConn)) & " AND " & _
                  "tblPurInv.HISTORY = FALSE AND tblPurInv.UPDATE_SAGE = TRUE"
   adoRst.Open sSQLQuery, adoConn, adOpenStatic, adLockReadOnly
   
   Set adoRst = Nothing
   Set adoConn = Nothing
End Sub






Private Function PIvalidation() As Boolean
     Dim iIncDec As Long
    iIncDec = 0
    Dim rCount As Integer
    Dim selRow As Integer
    Dim isitPlus As Boolean
    For rCount = 1 To frmRentPayable.flxPayFees.Rows - 1
         If frmRentPayable.flxPayFees.TextMatrix(rCount, 0) = "X" Then
             If frmRentPayable.flxPayFees.TextMatrix(rCount, 1) = "+" Then
                isitPlus = True
             Else
                isitPlus = False
             End If
             iIncDec = iIncDec + 1
             selRow = rCount
         End If
    Next
    If iIncDec < 1 Then
       MsgBox "Please select one statement only.", vbInformation + vbOKOnly, "statement Selection"
       Exit Function
    End If
    If isitPlus = True Then
        PIvalidation = True
    Else
        PIvalidation = False
    End If
    
    'PIvalidation = True
End Function
Private Function Feestogenerate() As Boolean
        Dim lSlNumber As Long
        Dim adoConn As New ADODB.Connection
        Dim adoPIHeader As New ADODB.Recordset
        Dim adoPISplit As New ADODB.Recordset
        Dim szSQL As String
        Dim szSQLManagingAgent  As String
        Dim szMYID As String
        Dim szFundID As String
        Dim szSelectedPayableTypeID As String
        Dim szSQL1 As String
        Dim rsfixedMethod As New ADODB.Recordset
        Dim rsfixedMethodDetails As New ADODB.Recordset
        Dim j As Integer
        Dim percnetageOramount As Double
        Dim dblGrandTotal As Double
        Dim dtNextDue As Date
        Dim dtFDD As Date
        Dim dblFeqID As Integer
        Dim dblTotalAmount As Double
        Dim lngMgtFeeSL As Long
        Dim iCountPI As Integer
        Dim iClientCount As Integer
        Dim iPropertyCount As Integer
        Dim strLastChargeDate  As String
        Dim strFundName As String
        Dim strStopDate  As String
        Dim dblCapAmount As Double
        Dim rsManagingAgent As New ADODB.Recordset
        Dim rsGlobalData As New ADODB.Recordset
        Dim szManagingAgent() As String
        Dim iManagingAgentCount As Integer
        Dim bControlACForPayable As Boolean
        Dim FinalControlACForPayable As String
        Dim szTemp
        Dim dblNoOfDaysToSendMFB4Due As Integer
        Dim dtNDDInitial As Date
        Dim strFromDate As String
        Dim strToDate As String
        Dim rsFromandToDate As New ADODB.Recordset
        Dim szSQLFrom As String
        Dim dtNDD As Date
        
        Dim iCount As Long
        Dim iCount1 As Long
        Dim rCount As Long
        Dim dblFundId As Integer
        Dim dblDemandTypeId As Integer
        Dim i As Integer
        Dim lT_ID As Long
    
        Dim strSelectedDemandType As String
        Dim szSQL5 As String
        Dim rstSet As New ADODB.Recordset
        Dim szPropertySelectionALL As String
        Dim szPropertySelection1 As String
    
        
        
 '      ************************************Write tblPurInv **************************************
     
        adoConn.Open getConnectionString
        adoConn.Execute "Delete from tblPurInvPreview"
        adoConn.Execute "Delete from tblPurInvSRecPreview"
        adoConn.Close
      For iClientCount = 1 To flxClients.Rows - 1
            If flxClients.TextMatrix(iClientCount, 0) = "X" Then
                        szSelectedClient = flxClients.TextMatrix(iClientCount, 1)
                        If szSelectedClient = "" Then Exit Function
           
            For iPropertyCount = 1 To flxProperties.Rows - 1
                    If iPropertyCount >= flxProperties.Rows Then Exit For
                    If flxProperties.TextMatrix(iPropertyCount, 0) = "X" And szSelectedClient = flxProperties.TextMatrix(iPropertyCount, 3) Then
                            szPropertySelection1 = flxProperties.TextMatrix(iPropertyCount, 1)
            If adoConn.State = 0 Then
                adoConn.Open getConnectionString
            End If
            Dim rsCharge As New ADODB.Recordset
            szTemp = ""
            'iPropertyCount = 3
            szSQLManagingAgent = "SELECT DISTINCT agr.ManagingAgentID " & _
              "FROM tlbAgreement agr, ClientProAgr CPA, ChargeTypes C,Fund  F,SECONDARYCODE SC " & _
              "WHERE agr.CPA_ID = CPA.CPA_ID And F.FundID=agr.fund AND  SC.CODE=agr.CHARGE_METHOD AND  " & _
              "CPA.ClientID = '" & szSelectedClient & "' And C.ID = agr.CHARGE_TYPE And " & _
              "CPA.PropertyID = '" & szPropertySelection1 & "'"
              rsManagingAgent.Open szSQLManagingAgent, adoConn, adOpenDynamic, adLockOptimistic
              szTemp = SQL2String(rsManagingAgent, 0)
              rsManagingAgent.Close
              If Len(szTemp) > 0 Then
                    szManagingAgent = Split(szTemp, ",")
              Else
                    adoConn.Close
                    GoTo EndOfAgreement
              End If
              
              
             szSQL5 = "SELECT MAX(ManagementFeeSL) AS x FROM tblPurInv;"
             rstSet.Open szSQL5, adoConn, adOpenStatic, adLockReadOnly
             lngMgtFeeSL = Val(IIf(IsNull(rstSet.Fields.Item(0).Value), 0, rstSet.Fields.Item(0).Value)) + 1
             rstSet.Close
             Set rstSet = Nothing
             adoConn.Close
             
            For iManagingAgentCount = 0 To UBound(szManagingAgent) 'this for shall end after the creation of the PI preview
              
            adoConn.Open getConnectionString
            'For each managing agent I am creating PI
            lSlNumber = SlNumber("PI", "tblPurInv", adoConn)
             
            szSQL = "SELECT agr.EachPeriod,agr.Capamount, agr.StopDate, CPA.agreementStartDate,CPA.agreementEndDate,agr.CHARGE_METHOD," & _
                "cpa.agreementEndDate as agreementEndD,agr.LastChargeDate,agr.TotalAmount,agr.Amount,agr.Fund,F.FundName, " & _
                "agr.NtDueDate,agr.FDD,agr.Frequency as FrequencyID,(Select FC.Frequency from Frequencies FC where  FC.ID=agr.Frequency) as FrequencyName,ManagingAgentID  " & _
              "FROM tlbAgreement agr, ClientProAgr CPA, ChargeTypes C,Fund  F,SECONDARYCODE SC,tlbPayable P " & _
              "WHERE P.CPA_ID = CPA.CPA_ID AND agr.CPA_ID = CPA.CPA_ID AND  P.ClientID=CPA.ClientID AND  P.PAY_FUND=agr.fund  And F.FundID=agr.fund AND  SC.CODE=agr.CHARGE_METHOD " & _
              "AND P.PAY_FUND=agr.Fund AND CPA.ClientID = '" & szSelectedClient & "' And C.ID = agr.CHARGE_TYPE And " & _
              "CPA.PropertyID = '" & szPropertySelection1 & "' AND agr.ManagingAgentID='" & Trim(szManagingAgent(iManagingAgentCount)) & "'"
        Debug.Print szPropertySelection1
        Debug.Print Trim(szManagingAgent(iManagingAgentCount))
        Debug.Print iPropertyCount
            rsCharge.Open szSQL, adoConn, adOpenDynamic, adLockOptimistic
            If rsCharge.EOF Then
                rsCharge.Close
                Set rsCharge = Nothing
               
                GoTo EndOfAgreement
            End If

                            
            i = 1
                'Dim rsGlobalData As New ADODB.Recordset
                Dim VAT_ID As String
                Dim VAT_CODE As String
                Dim VAT_RATE As Double
                
                 rsGlobalData.Open "Select vatOptionEnabled,V.VAT_ID,V.VAT_CODE,V.VAT_RATE from (Globaldata G LEFT JOIN tlbVatCode V ON G.vatRate=V.VAT_ID) where PropertyID='" & _
                                    szPropertySelection1 & "' AND vatOptionEnabled=true", adoConn, adOpenStatic, adLockReadOnly
                 If Not rsGlobalData.EOF Then
                          VAT_ID = IIf(IsNull(rsGlobalData.Fields("VAT_ID").Value), "-1", rsGlobalData.Fields("VAT_ID").Value)
                          VAT_RATE = IIf(IsNull(rsGlobalData.Fields("VAT_RATE").Value), "0.00", rsGlobalData.Fields("VAT_RATE").Value)
                          VAT_CODE = IIf(IsNull(rsGlobalData.Fields("VAT_CODE").Value), "", rsGlobalData.Fields("VAT_CODE").Value)
                 Else
                          VAT_ID = -1
                          VAT_RATE = 0
                          VAT_CODE = ""
                 End If
                 rsGlobalData.Close
                 Set rsGlobalData = Nothing
                 
                 Dim rsGlobalData1 As New ADODB.Recordset
                 Dim bolVatOptionEnabled As Boolean
                 Dim bolOptedTotax As String
                 Dim strManagingAgentID As String
                 rsGlobalData1.Open "Select vatOptionEnabled from (Globaldata G LEFT JOIN tlbVatCode V ON G.vatRate=V.VAT_ID) where PropertyID='" & _
                                    szPropertySelection1 & "' ", adoConn, adOpenStatic, adLockReadOnly
                                    
                If Not rsGlobalData1.EOF Then
                        bolVatOptionEnabled = IIf(IsNull(rsGlobalData1("vatOptionEnabled").Value), False, rsGlobalData1("vatOptionEnabled").Value)
                        strManagingAgentID = rsCharge("ManagingAgentID").Value
                End If
                rsGlobalData1.Close
                
              
              
           
      
            szSQL = "SELECT MAX(TRANSACTIONID) AS TID FROM tlbPayment;"
            adoPIHeader.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
            lT_ID = CLng(IIf(IsNull(adoPIHeader!TID), 1, adoPIHeader!TID + 1))
            adoPIHeader.Close
             j = 1
                 rsGlobalData1.Open "SELECT optedTotax,* FROM Supplier where supplierID='" & strManagingAgentID & "'", adoConn, adOpenStatic, adLockReadOnly
                If Not rsGlobalData1.EOF Then
                        bolOptedTotax = rsGlobalData1("optedTotax").Value
                Else
                        bolOptedTotax = False
                End If
                rsGlobalData1.Close
           
   
   
            While Not rsCharge.EOF
                  dblTotalAmount = rsCharge.Fields.Item("TotalAmount").Value
                  dblCapAmount = rsCharge.Fields.Item("CapAmount").Value
                  dblFundId = rsCharge.Fields.Item("Fund").Value
                 ' dblDemandTypeId = rsCharge.Fields.Item("DEMAND_TYPE").Value
                  
                  strFundName = rsCharge.Fields.Item("FundName").Value
                  If Not IsNull(rsCharge.Fields.Item("NtDueDate").Value) Then
                      dtNextDue = rsCharge.Fields.Item("NtDueDate").Value
                      dtNDDInitial = rsCharge.Fields.Item("NtDueDate").Value
                  End If
                  If Not IsNull(rsCharge.Fields.Item("FDD").Value) Then
                      dtFDD = rsCharge.Fields.Item("FDD").Value
                  End If
                  If Not IsNull(rsCharge.Fields.Item("FrequencyID").Value) Then
                    If rsCharge.Fields.Item("FrequencyID").Value <> "" Then
                            dblFeqID = rsCharge.Fields.Item("FrequencyID").Value
                      End If
                  End If
                szMYID = UniqueID()
                strLastChargeDate = IIf(IsNull(rsCharge("LastChargeDate").Value), "", rsCharge("LastChargeDate").Value)
                If strLastChargeDate = "" Then
                       rsCharge.Close
                       
                       GoTo EndOfAgreement
                End If
                rsGlobalData.Open "Select NoOfDaysToSendMFB4Due from globaldata where PropertyID='" & szPropertySelection1 & "'", adoConn, adOpenStatic, adLockReadOnly
                If Not rsGlobalData.EOF Then
                    dblNoOfDaysToSendMFB4Due = IIf(IsNull(rsGlobalData!NoOfDaysToSendMFB4Due), 0, rsGlobalData!NoOfDaysToSendMFB4Due)
                End If
                rsGlobalData.Close
               
                If DateDiff("d", Date, rsCharge("NtDueDate").Value) > dblNoOfDaysToSendMFB4Due Then
                        GoTo EndOfChargeType
                End If
            
                           
                                'validations
                                If rsCharge.Fields.Item("CHARGE_METHOD").Value = "RE_FIX" Then   'when working on the fixed procedure only 1 line of setup is done then (fixed basis)
                                                If rsCharge("agreementEndD").Value < rsCharge("FDD").Value Then
                                                       
                                                        GoTo EndOfAgreement
                                                End If
                                                If DateDiff("d", txtStatementDate1.text, rsCharge.Fields.Item("agreementStartDate").Value) > 0 Then
                                                        rsCharge.Close
                                                        Set rsCharge = Nothing
                                                       
                                                        GoTo EndOfChargeType
                                                End If

                            '      ************************************Write tblPurInvSRec **************************************
                                                                                                                                   
                                                    dblTotalAmount = IIf(IsNull(rsCharge("EachPeriod").Value), 0, rsCharge("EachPeriod").Value) 'rsfixedMethodDetails.Fields.Item("Amt").Value
        '                                            dblGrandTotal = dblGrandTotal + dblTotalAmount
                                                    dblFundId = rsCharge.Fields.Item("Fund").Value

                                                    If dblCapAmount > 0 Then
                                                           'make a condition in the split as well so that amount doesnot exeed cap amount
                                                           If dblTotalAmount > dblCapAmount Then
                                                                dblTotalAmount = dblCapAmount
                                                           End If
                                                    End If

                                                                 txtComparenextDueDate1 = DateAdd("d", 1, dtNextDue)
                                                                dtNDD = NextDueDate1(CInt(dblFeqID), txtComparenextDueDate1, szPropertySelection1)
                                                                txtComparenextDueDate1 = DateAdd("d", 1, dtNDD)
                                                                dtFDD = NextDueDate1(CInt(dblFeqID), txtComparenextDueDate1, szPropertySelection1)
                                                                strFromDate = Format(dtNextDue, "dd/MM/yyyy")
                                                                strToDate = Format(DateAdd("d", -1, dtNDD), "dd/MM/yyyy")
                                                    
                                                    szSQL = "SELECT * FROM tblPurInvSRecPreview"
                                                   ' adoPISplit.Close
                                                    adoPISplit.Open szSQL, adoConn, adOpenDynamic, adLockOptimistic
                                                    'Add New Records. At least there is only one split line
                                                       With adoPISplit
                                                           .AddNew
                                                           .Fields.Item("MY_ID").Value = UniqueID()
                                                           .Fields.Item("ParentID").Value = szMYID
                                                           .Fields.Item("TRAN_ID").Value = j
                                                          
                                                          'If chkAssignProperty.Value = 0 Then
                                                                .Fields.Item("TRANS").Value = szPropertySelection1  ' If you select One property then you can write a value here
                                                           ' Else
                                                             '    .Fields.Item("TRANS").Value = ""
                                                            'End If
                                                           .Fields.Item("UNIT_ID").Value = ""
                                                           .Fields.Item("NOMINAL_CODE").Value = FinalControlACForPayable
                                                           .Fields.Item("DEPT_ID").Value = dblFundId
                                                          ' .Fields.Item("JOB_ID").Value = flxPI.TextMatrix(iRow, 9)            'Job No
                                                           .Fields.Item("RecoverablePt").Value = 0
                                                                
                                                           .Fields.Item("description").Value = "Management Fees for " & strFundName & " (" & strFromDate & " - " & strToDate & ")"
                                                           If bolVatOptionEnabled = True And bolOptedTotax = True Then
                                                                .Fields.Item("NET_AMOUNT").Value = dblTotalAmount
                                                                .Fields.Item("TAX_CODE").Value = VAT_CODE
                                                                .Fields.Item("VAT").Value = Round(dblTotalAmount * (VAT_RATE / 100), 2) 'VAT_RATE
                                                                .Fields.Item("TOTAL_AMOUNT").Value = .Fields.Item("VAT").Value + dblTotalAmount
                                                                 dblTotalAmount = .Fields.Item("TOTAL_AMOUNT").Value
                                                           ElseIf bolVatOptionEnabled = True And bolOptedTotax = False Then

                                                                     .Fields.Item("NET_AMOUNT").Value = dblTotalAmount
                                                                    .Fields.Item("TAX_CODE").Value = Null 'VAT_CODE
                                                                    .Fields.Item("VAT").Value = 0 'Format(dblTotalAmount * (VAT_RATE / 100), "0.00") 'VAT_RATE
                                                                    .Fields.Item("TOTAL_AMOUNT").Value = dblTotalAmount
                                                                     dblTotalAmount = .Fields.Item("TOTAL_AMOUNT").Value
                                                                                                 
                                                          ElseIf bolVatOptionEnabled = False And bolOptedTotax = True Then 'bolVatOptionEnabled means global data and bolOptedTotax is supplier table
                                                                rsGlobalData.Open "Select V.VAT_ID,V.VAT_CODE,V.VAT_RATE from  Supplier S,tlbVatCode V where  (S.VATCode)=cstr(V.VAT_ID)  AND  SupplierID='" & _
                                                                                   strManagingAgentID & "' ", adoConn, adOpenStatic, adLockReadOnly
                                                                If Not rsGlobalData.EOF Then
                                                                         VAT_ID = IIf(IsNull(rsGlobalData.Fields("VAT_ID").Value), "-1", rsGlobalData.Fields("VAT_ID").Value)
                                                                         VAT_RATE = IIf(IsNull(rsGlobalData.Fields("VAT_RATE").Value), "0.00", rsGlobalData.Fields("VAT_RATE").Value)
                                                                         VAT_CODE = IIf(IsNull(rsGlobalData.Fields("VAT_CODE").Value), "", rsGlobalData.Fields("VAT_CODE").Value)
                                                                Else
                                                                         VAT_ID = -1
                                                                         VAT_RATE = 0
                                                                         VAT_CODE = ""
                                                                End If
                                                                rsGlobalData.Close

                                                                 'done modification on 15-10-2021
                                                                    .Fields.Item("NET_AMOUNT").Value = dblTotalAmount
                                                                    .Fields.Item("TAX_CODE").Value = Null 'VAT_CODE
                                                                    .Fields.Item("VAT").Value = 0 'Format(dblTotalAmount * (VAT_RATE / 100), "0.00") ' Format(dblTotalAmount * (VAT_RATE / 100), "0.00") 'VAT_RATE' + Format(dblTotalAmount * (VAT_RATE / 100), "0.00")
                                                                    .Fields.Item("NET_AMOUNT").Value = dblTotalAmount + Round(dblTotalAmount * (VAT_RATE / 100), 2)
                                                                    .Fields.Item("TOTAL_AMOUNT").Value = dblTotalAmount + Round(dblTotalAmount * (VAT_RATE / 100), 2)
                                                                     dblTotalAmount = .Fields.Item("TOTAL_AMOUNT").Value
                                                                                                 
                                                                                                 
                                                          ElseIf bolVatOptionEnabled = False And bolOptedTotax = False Then

                                                                 .Fields.Item("NET_AMOUNT").Value = dblTotalAmount
                                                                .Fields.Item("TAX_CODE").Value = Null 'VAT_CODE
                                                                .Fields.Item("VAT").Value = 0 'Format(dblTotalAmount * (VAT_RATE / 100), "0.00") 'VAT_RATE
                                                                .Fields.Item("TOTAL_AMOUNT").Value = dblTotalAmount
                                                                 dblTotalAmount = .Fields.Item("TOTAL_AMOUNT").Value
                        
                                                           End If
                                                           
                                                           .Update
                                                       End With
                                                    adoPISplit.Close
                                                   If rsCharge.Fields.Item("CHARGE_METHOD").Value = "RE_FIX" Then

                                                     End If
                                                    dblGrandTotal = dblGrandTotal + dblTotalAmount
                                                                       
                                                                       
                                End If 'end of rsCharge.Fields.Item("CHARGE_METHOD").Value = "RE_FIX"
                                If rsCharge.Fields.Item("CHARGE_METHOD").Value = "RE_ED" Then ' receipt basis

                                                If rsCharge("agreementEndD").Value < rsCharge("FDD").Value Then
                                                      
                                                        GoTo EndOfAgreement
                                                End If
                                                If DateDiff("d", txtStatementDate1.text, rsCharge.Fields.Item("agreementStartDate").Value) > 0 Then
                                                        rsCharge.Close
                                                        Set rsCharge = Nothing
                                                        
                                                        GoTo EndOfChargeType
                                                        
                                                End If

                                                    
'                                            szSQL1 = "Select  sum(SWITCH(R.TYPE =3,RS.Amount,R.TYPE =4,RS.Amount,R.TYPE =23,-RS.Amount))  as Amt from tlbReceipt R,tlbReceiptsplit RS,tlbReceipt R1, " & _
'                                            "rptTransactionsSPlit AL, DemandSplitRecords DS, Units U where R1.DemandRef=DS.DemandID and AL.TOTRAN=R1.TransactionID AND RS.SPLITID=DS.SPLITID AND AL.deleteflag=false AND " & _
'                                            "AL.TransactionID= RS.RptTransactionsIDSplit AND  R.TransactionID=RS.RptHeader AND R.RDate<=#" & Format(txtStatementDate1.text, "dd MMM yyyy") & "# " & _
'                                            "AND R.RDate>#" & Format(strLastChargeDate, "dd MMM yyyy") & "# and R.Type in (3,4,23)  AND U.UnitNumber=R.UnitID AND U.PropertyID='" & _
'                                            szPropertySelection1 & "' and R.ISMGTFEE=false AND Rs.FundID=" & dblFundId & " AND DS.TypeOfDemand=" & rsCharge("DEMAND_TYPE").Value & " "
                                            szSQL1 = "Select  sum(SWITCH(R.TYPE =3,RS.Amount,R.TYPE =4,RS.Amount,R.TYPE =23,-RS.Amount))  as Amt from tlbReceipt R,tlbReceiptsplit RS, " & _
                                            "rptTransactionsSPlit AL, Units U where AL.deleteflag=false AND " & _
                                            "AL.TransactionID= RS.RptTransactionsIDSplit AND  R.TransactionID=RS.RptHeader AND R.RDate<=#" & Format(txtStatementDate1.text, "dd MMM yyyy") & "# " & _
                                            "AND R.Type in (3,4,23)  AND U.UnitNumber=R.UnitID AND U.PropertyID='" & _
                                            szPropertySelection1 & "' and RS.ISMGTFEES=false AND Rs.FundID=" & dblFundId & ""
                                     
                                                                                           'by anol 20211024
                                                     If rsfixedMethod.State = 1 Then
                                                        rsfixedMethod.Close
                                                     End If
                                                    rsfixedMethod.Open szSQL1, adoConn, adOpenStatic, adLockReadOnly
                                                    'Here type 3 is for reciept type . I have not written for the credit yet need to understand the principle
                                                    
                                                    If rsfixedMethod.EOF Then
                                                        rsfixedMethod.Close
                                                        Set rsfixedMethod = Nothing
                                                        GoTo EndOfChargeType
                                                    End If
                                                    percnetageOramount = IIf(IsNull(rsCharge("amount").Value), 0, rsCharge("amount").Value)
                                                    
                            '      ************************************Write tblPurInvSRec **************************************

'                                     szSQL = "Select  sum(SWITCH(R.TYPE =3,RS.Amount,R.TYPE =4,RS.Amount,R.TYPE =23,-RS.Amount))  as Amt from tlbReceipt R,tlbReceiptsplit RS,tlbReceipt R1, " & _
'                                     "rptTransactionsSPlit AL, DemandSplitRecords DS, Units U where R1.DemandRef=DS.DemandID and AL.TOTRAN=R1.TransactionID AND RS.SPLITID=DS.SPLITID AND AL.deleteflag=false AND " & _
'                                     "AL.TransactionID= RS.RptTransactionsIDSplit AND  R.TransactionID=RS.RptHeader AND R.RDate<=#" & Format(txtStatementDate1.text, "dd MMM yyyy") & "# " & _
'                                     "AND R.RDate>#" & Format(strLastChargeDate, "dd MMM yyyy") & "# and R.Type in (3,4,23)  AND U.UnitNumber=R.UnitID AND U.PropertyID='" & _
'                                     szPropertySelection1 & "' and R.ISMGTFEE=false AND Rs.FundID=" & dblFundId & " AND DS.TypeOfDemand=" & rsCharge("DEMAND_TYPE").Value & ""

                                      szSQL = "Select  sum(SWITCH(R.TYPE =3,RS.Amount,R.TYPE =4,RS.Amount,R.TYPE =23,-RS.Amount))  as Amt from tlbReceipt R,tlbReceiptsplit RS, " & _
                                     "rptTransactionsSPlit AL, Units U where AL.deleteflag=false AND " & _
                                     "AL.TransactionID= RS.RptTransactionsIDSplit AND  R.TransactionID=RS.RptHeader AND R.RDate<=#" & Format(txtStatementDate1.text, "dd MMM yyyy") & "# " & _
                                     "AND  R.Type in (3,4,23)  AND U.UnitNumber=R.UnitID AND U.PropertyID='" & _
                                     szPropertySelection1 & "' and RS.ISMGTFEES=false AND Rs.FundID=" & dblFundId & ""
                                                  'modified on 20211103
                                                        'need to consider the selected property in where clause
                                                        'modified on 20211103
'                                                        szSQLFrom = "Select  min(DS.DateFrom) as  DateFromMin ,max(DS.DateTO) as DateTOMAX   from tlbReceipt R,tlbReceiptsplit RS,tlbReceipt R1, " & _
'                                                        "rptTransactionsSPlit AL, DemandSplitRecords DS, Units U where R1.DemandRef=DS.DemandID and AL.TOTRAN=R1.TransactionID AND RS.SPLITID=DS.SPLITID AND AL.deleteflag=false AND " & _
'                                                        "AL.TransactionID= RS.RptTransactionsIDSplit AND  R.TransactionID=RS.RptHeader AND R.RDate<=#" & Format(txtStatementDate1.text, "dd MMM yyyy") & "# " & _
'                                                        "AND R.RDate>#" & Format(strLastChargeDate, "dd MMM yyyy") & "# and R.Type in (3,4,23)  AND U.UnitNumber=R.UnitID AND U.PropertyID='" & _
'                                                        szPropertySelection1 & "' and R.ISMGTFEE=false AND Rs.FundID=" & dblFundId & " AND DS.TypeOfDemand=" & rsCharge("DEMAND_TYPE").Value & ""
                                                        szSQLFrom = "Select  min(DS.DateFrom) as  DateFromMin ,max(DS.DateTO) as DateTOMAX   from tlbReceipt R,tlbReceiptsplit RS, " & _
                                                        "rptTransactionsSPlit AL, DemandSplitRecords DS, Units U where  AL.deleteflag=false AND " & _
                                                        "AL.TransactionID= RS.RptTransactionsIDSplit AND  R.TransactionID=RS.RptHeader AND R.RDate<=#" & Format(txtStatementDate1.text, "dd MMM yyyy") & "# " & _
                                                        "AND R.Type in (3,4,23)  AND U.UnitNumber=R.UnitID AND U.PropertyID='" & _
                                                        szPropertySelection1 & "' AND RS.ISMGTFEES=false AND Rs.FundID=" & dblFundId & " "
                                                        
                                                                                                                 
                                                        rsFromandToDate.Open szSQLFrom, adoConn, adOpenStatic, adLockReadOnly
                                                        If Not rsFromandToDate.EOF Then
                                                                strFromDate = Format(rsFromandToDate("DateFromMin").Value, "dd/MM/yyyy")
                                                                strToDate = Format(rsFromandToDate("DateTOMAX").Value, "dd/MM/yyyy")
                                                        Else
                                                                strFromDate = ""
                                                                strToDate = ""
                                                        End If
                                                        rsFromandToDate.Close
                                                        
                                                         'Need to take only allocated transactions
                                                        If rsfixedMethodDetails.State = 1 Then
                                                            rsfixedMethodDetails.Close
                                                        End If
                                                                 rsfixedMethodDetails.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
                                                                  If Not rsfixedMethodDetails.EOF Then 'we rare using while because itr
                                                                                 dblTotalAmount = IIf(IsNull(rsfixedMethodDetails.Fields.Item("Amt").Value), 0, rsfixedMethodDetails.Fields.Item("Amt").Value)
                                                                                 If dblTotalAmount <= 0 Then
                                                                                        rsfixedMethod.Close
                                                                                        GoTo EndOfChargeType
                                                                                 End If
                                                                                 'dblFundId = rsfixedMethodDetails.Fields.Item("FundID").Value
                                                                                 If rsCharge.Fields.Item("CHARGE_METHOD").Value = "RE_ED" Then
                                                                                       dblTotalAmount = dblTotalAmount * (percnetageOramount / 100)
                                                                                        dblTotalAmount = Round(dblTotalAmount, 2)
                                                                                 End If
                                                                                 
                                                                                 If dblCapAmount > 0 Then
                                                                                        If dblTotalAmount > dblCapAmount Then
                                                                                             dblTotalAmount = dblCapAmount
                                                                                        End If
                                                                                 End If
                                                                                 szSQL = "SELECT * FROM tblPurInvSRecPreview"
                                                                                 adoPISplit.Open szSQL, adoConn, adOpenDynamic, adLockOptimistic
                                                                                 'Add New Records. At least there is only one split line
                                                                                    With adoPISplit
                                                                                        .AddNew
                                                                                        .Fields.Item("MY_ID").Value = UniqueID()
                                                                                        .Fields.Item("ParentID").Value = szMYID
                                                                                        .Fields.Item("TRAN_ID").Value = j
                                                                                       ' .Fields.Item("TRANS").Value = szPropertySelection1  ' If you select One property then you can write a value here
                                                                                         'If chkAssignProperty.Value = 0 Then
                                                                                              .Fields.Item("TRANS").Value = szPropertySelection1  ' If you select One property then you can write a value here
                                                                                         'Else
                                                                                          '    .Fields.Item("TRANS").Value = ""
                                                                                         'End If
                                                                                        .Fields.Item("UNIT_ID").Value = ""
                                                                                        .Fields.Item("NOMINAL_CODE").Value = FinalControlACForPayable
                                                                                        .Fields.Item("DEPT_ID").Value = dblFundId
                                                                                       ' .Fields.Item("JOB_ID").Value = flxPI.TextMatrix(iRow, 9)            'Job No
                                                                                        .Fields.Item("RecoverablePt").Value = 0
                                                                                        '' (Current Charge date)" '"MFee" + szPropertySelection1 + Format(lngMgtFeeSL, "0000") '"Management Fee"
                                                                                        .Fields.Item("description").Value = "Management Fees for " & strFundName & " (" & strFromDate & " - " & strToDate & ")"

                                                                                        If bolVatOptionEnabled = True And bolOptedTotax = True Then
                                                                                                dblTotalAmount = dblTotalAmount * Round((100 / (100 + VAT_RATE)), 2)
                                                                                                .Fields.Item("NET_AMOUNT").Value = dblTotalAmount
                                                                                                .Fields.Item("TAX_CODE").Value = VAT_CODE
                                                                                                .Fields.Item("VAT").Value = Round(dblTotalAmount * (VAT_RATE / 100), 2) 'VAT_RATE
                                                                                                .Fields.Item("TOTAL_AMOUNT").Value = .Fields.Item("VAT").Value + dblTotalAmount
                                                                                                 dblTotalAmount = .Fields.Item("TOTAL_AMOUNT").Value
                                                                                           ElseIf bolVatOptionEnabled = True And bolOptedTotax = False Then 'bolVatOptionEnabled=global data
                                                                                                'Modified by anol 2021-10-15
                                                                                                 dblTotalAmount = dblTotalAmount * Round((100 / (100 + VAT_RATE)), 2)
                                                                                                 .Fields.Item("NET_AMOUNT").Value = dblTotalAmount
                                                                                                .Fields.Item("TAX_CODE").Value = Null 'VAT_CODE
                                                                                                .Fields.Item("VAT").Value = 0 'Format(dblTotalAmount * (VAT_RATE / 100), "0.00") 'VAT_RATE
                                                                                                .Fields.Item("TOTAL_AMOUNT").Value = dblTotalAmount
                                                                                                 dblTotalAmount = .Fields.Item("TOTAL_AMOUNT").Value
                                                                                          ElseIf bolVatOptionEnabled = False And bolOptedTotax = True Then 'bolVatOptionEnabled=global data
                                                                                                rsGlobalData.Open "Select V.VAT_ID,V.VAT_CODE,V.VAT_RATE from  Supplier S,tlbVatCode V where  (S.VATCode)=cstr(V.VAT_ID)  AND  SupplierID='" & _
                                                                                                                   strManagingAgentID & "' ", adoConn, adOpenStatic, adLockReadOnly
                                                                                                If Not rsGlobalData.EOF Then
                                                                                                         VAT_ID = IIf(IsNull(rsGlobalData.Fields("VAT_ID").Value), "-1", rsGlobalData.Fields("VAT_ID").Value)
                                                                                                         VAT_RATE = IIf(IsNull(rsGlobalData.Fields("VAT_RATE").Value), "0.00", rsGlobalData.Fields("VAT_RATE").Value)
                                                                                                         VAT_CODE = IIf(IsNull(rsGlobalData.Fields("VAT_CODE").Value), "", rsGlobalData.Fields("VAT_CODE").Value)
                                                                                                Else
                                                                                                         VAT_ID = -1
                                                                                                         VAT_RATE = 0
                                                                                                         VAT_CODE = ""
                                                                                                End If
                                                                                                rsGlobalData.Close
                                                                                                'done modification on 15-10-2021
                                                                                                .Fields.Item("NET_AMOUNT").Value = dblTotalAmount
                                                                                                .Fields.Item("TAX_CODE").Value = Null ' VAT_CODE
                                                                                                .Fields.Item("VAT").Value = 0 'Format(dblTotalAmount * (VAT_RATE / 100), "0.00") ' Format(dblTotalAmount * (VAT_RATE / 100), "0.00") 'VAT_RATE' + Format(dblTotalAmount * (VAT_RATE / 100), "0.00")
                                                                                                .Fields.Item("NET_AMOUNT").Value = dblTotalAmount + Round(dblTotalAmount * (VAT_RATE / 100), 2)
                                                                                                .Fields.Item("TOTAL_AMOUNT").Value = dblTotalAmount + Round(dblTotalAmount * (VAT_RATE / 100), 2)
                                                                                                 dblTotalAmount = .Fields.Item("TOTAL_AMOUNT").Value
                                                                                          ElseIf bolVatOptionEnabled = False And bolOptedTotax = False Then
                                                                                                .Fields.Item("NET_AMOUNT").Value = dblTotalAmount
                                                                                                .Fields.Item("TAX_CODE").Value = Null 'VAT_CODE
                                                                                                .Fields.Item("VAT").Value = 0 'Format(dblTotalAmount * (VAT_RATE / 100), "0.00") 'VAT_RATE
                                                                                                .Fields.Item("TOTAL_AMOUNT").Value = dblTotalAmount
                                                                                                 dblTotalAmount = .Fields.Item("TOTAL_AMOUNT").Value
                                                                                           End If
                                                                                        .Update
                                                                                    End With
                                                                                 adoPISplit.Close
                                                                                 dblGrandTotal = dblGrandTotal + dblTotalAmount
                                                            End If
                                                            rsfixedMethod.Close
                                End If 'end of  If rsCharge.Fields.Item("CHARGE_METHOD").Value = "RE_ED" Then
                                If rsCharge.Fields.Item("CHARGE_METHOD").Value = "RE_ABL" Then

                                                If rsCharge("agreementEndD").Value < rsCharge("FDD").Value Then
                                                       
                                                        GoTo EndOfAgreement
                                                End If
                                                If DateDiff("d", txtStatementDate1.text, rsCharge.Fields.Item("agreementStartDate").Value) > 0 Then
                                                        rsCharge.Close
                                                        Set rsCharge = Nothing
                                                       
                                                        GoTo EndOfChargeType
                                                        
                                                End If
                                                 szSQL1 = "Select sum(SWITCH(TYPE =1,RS.Amount,TYPE =2,-RS.Amount,TYPE =24,RS.Amount)) as Amt from tlbReceipt R,tlbReceiptsplit RS,Units U where " & _
                                                              "R.TransactionID=RS.RptHeader AND RDate<=#" & Format(txtStatementDate1.text, " dd MMM yyyy") & "# AND RDate>#" & Format(strLastChargeDate, " dd MMM yyyy") & "#  and Type in (1,2,24) " & _
                                                              "AND U.UnitNumber=R.UnitID AND U.PropertyID='" & szPropertySelection1 & "' and ISMGTFEE=false AND Rs.FundID=" & dblFundId & ""
                                                              'need to consider the selected property in where clause
                                                            '  rsfixedMethod.Close
                                                    rsfixedMethod.Open szSQL1, adoConn, adOpenStatic, adLockReadOnly
                                                    'Here type 3 is for recipt type . I have not written for the credit yet need to understand the principle
                                                    
                                                    If rsfixedMethod.EOF Then
                                                        rsfixedMethod.Close
                                                        Set rsfixedMethod = Nothing
                                                        GoTo EndOfChargeType
                                                    End If
                                                    percnetageOramount = IIf(IsNull(rsCharge("amount").Value), 0, rsCharge("amount").Value)
                                                    
                            '      ************************************Write tblPurInvSRec **************************************
                            
                                                                 szSQL = "Select sum(SWITCH(TYPE =3,RS.Amount,TYPE =4,-RS.Amount,TYPE =23,RS.Amount)) as Amt,rs.FundID from tlbReceipt R,tlbReceiptsplit RS,Units U where " & _
                                                                         "R.TransactionID=RS.RptHeader and Type in (3,4,23) AND RDate<=#" & Format(txtStatementDate1.text, " dd MMM yyyy") & "# AND RDate>#" & Format(strLastChargeDate, " dd MMM yyyy") & "#  AND  " & _
                                                                         "U.UnitNumber=R.UnitID and ISMGTFEE=false AND U.PropertyID='" & szPropertySelection1 & "' AND RS.FundID=" & dblFundId & " group by RS.FundID"
                                                                 'need to consider the selected property in where clause
                                                                 If rsfixedMethodDetails.State = 1 Then
                                                                     rsfixedMethodDetails.Close
                                                                 End If
                                                                 rsfixedMethodDetails.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
                                                                  If Not rsfixedMethodDetails.EOF Then 'we rare using while because itr
                                                                                 dblTotalAmount = rsfixedMethodDetails.Fields.Item("Amt").Value
                                                                                 If dblTotalAmount < 0 Then
                                                                                        GoTo EndOfChargeType
                                                                                 End If
                                                                                 dblFundId = rsfixedMethodDetails.Fields.Item("FundID").Value
                                                                                 If rsCharge.Fields.Item("CHARGE_METHOD").Value = "RE_ED" Then
                                                                                       dblTotalAmount = dblTotalAmount * percnetageOramount / 100
                                                                                 End If
                                                                                 
                                                                                 If dblCapAmount > 0 Then
                                                                                        If dblTotalAmount > dblCapAmount Then
                                                                                             dblTotalAmount = dblCapAmount
                                                                                        End If
                                                                                 End If
                                                                                 szSQL = "SELECT * FROM tblPurInvSRecPreview"
                                                                                 adoPISplit.Open szSQL, adoConn, adOpenDynamic, adLockOptimistic
                                                                                 'Add New Records. At least there is only one split line
                                                                                    With adoPISplit
                                                                                        .AddNew
                                                                                        .Fields.Item("MY_ID").Value = UniqueID()
                                                                                        .Fields.Item("ParentID").Value = szMYID
                                                                                        .Fields.Item("TRAN_ID").Value = j
                                                                                       ' .Fields.Item("TRANS").Value = szPropertySelection1  ' If you select One property then you can write a value here
                                                                                         'If chkAssignProperty.Value = 0 Then
                                                                                              .Fields.Item("TRANS").Value = szPropertySelection1  ' If you select One property then you can write a value here
                                                                                         'Else
                                                                                          '    .Fields.Item("TRANS").Value = ""
                                                                                         'End If
                                                                                        .Fields.Item("UNIT_ID").Value = ""
                                                                                        .Fields.Item("NOMINAL_CODE").Value = FinalControlACForPayable
                                                                                        .Fields.Item("DEPT_ID").Value = dblFundId
                                                                                       ' .Fields.Item("JOB_ID").Value = flxPI.TextMatrix(iRow, 9)            'Job No
                                                                                        .Fields.Item("RecoverablePt").Value = 0
                                                                                        '' (Current Charge date)" '"MFee" + szPropertySelection1 + Format(lngMgtFeeSL, "0000") '"Management Fee"
                                                                                        
                                                                                        .Fields.Item("description").Value = "Management Fees for " & strFundName & " " & DateAdd("d", 1, CDate(strFromDate)) & " - " & strToDate & ""
                                                                                        .Fields.Item("NET_AMOUNT").Value = dblTotalAmount
                                                                                        .Fields.Item("TAX_CODE").Value = VAT_CODE
                                                                                        .Fields.Item("VAT").Value = Round(dblTotalAmount * (VAT_RATE / 100), 2) 'VAT_RATE
                                                                                        .Fields.Item("TOTAL_AMOUNT").Value = .Fields.Item("VAT").Value + dblTotalAmount
                                                                                         dblTotalAmount = .Fields.Item("TOTAL_AMOUNT").Value
                                                                                        .Update
                                                                                    End With
                                                                                 adoPISplit.Close
                                                                                
                                                                                 dblGrandTotal = dblGrandTotal + dblTotalAmount
                                                            End If
                                                            rsfixedMethod.Close
                                End If 'end of  If rsCharge.Fields.Item("CHARGE_METHOD").Value = "RE_ABL" Then

                                szSQL = "SELECT * FROM tblPurInvPreview"

                                    dblTotalAmount = Round(dblTotalAmount, 2)
                                    If dblTotalAmount = 0 Then

                                           GoTo EndOfChargeType
                                    End If
                                    With adoPIHeader
                                            .Open szSQL, adoConn, adOpenDynamic, adLockPessimistic
                                            .AddNew
                                            .Fields.Item("MY_ID").Value = szMYID
                                            .Fields.Item("SlNumber").Value = lSlNumber
                                            .Fields.Item("SUPP_AC").Value = Trim(szManagingAgent(iManagingAgentCount))
                                            .Fields.Item("TRAN_DATE").Value = Format(txtStatementDate1.text, "DD MMMM YYYY")
                                            .Fields.Item("TransactionType").Value = 6
                                            .Fields.Item("INV_NO").Value = szPropertySelection1 + "-" + "MFee" + "-" + CStr(lngMgtFeeSL)
                                             lngMgtFeeSL = lngMgtFeeSL + 1
                                            .Fields.Item("TOTAL_AMOUNT").Value = dblTotalAmount
                                            .Fields.Item("History").Value = False
                                            .Fields.Item("TrfPayment").Value = False
                                            .Fields.Item("PropertyID").Value = ""
                                            .Fields.Item("CL_ID").Value = szSelectedClient
                                            .Fields.Item("NLPost").Value = False
                                            .Fields.Item("DueDate").Value = Format(dtNDDInitial, "DD MMMM YYYY")
                                            .Fields.Item("PostingDate").Value = Format(txtStatementDate1.text, "DD MMMM YYYY")
                                            .Fields.Item("ReportFromDatePreview").Value = IIf(strFromDate = "", Null, strFromDate)
                                            .Fields.Item("ReportToDatePreview").Value = IIf(strToDate = "", Null, strToDate)
                                            .Update
                                            iCountPI = iCountPI + 1
                                            lSlNumber = lSlNumber + 1
                                   End With
                                   adoPIHeader.Close
                                   If iCountPI > 0 Then
                                        Feestogenerate = True
                                        adoConn.Close
                                        Set adoConn = Nothing
                                        Exit Function
                                   End If
                                    
                                
EndOfChargeType:
            rsCharge.MoveNext
            j = j + 1
            
      Wend
                        rsCharge.Close
                        Set rsCharge = Nothing
                        adoConn.Close
                        Set adoConn = Nothing
EndOfOneManagingAgentforOneAgreement:
           Next iManagingAgentCount
         'end if for 'X' in grid seletion
EndOfAgreement:
          End If

          Next iPropertyCount
          End If 'end if only for selected client
  Next iClientCount
            
       
End Function
Private Function PreviousStatementFinalized() As Boolean
    Dim adoConn As New ADODB.Connection
    adoConn.Open getConnectionString
    Dim rstblPurInv As New ADODB.Recordset
    rstblPurInv.Open "Select * from RentSummaryStatement where  ClientIDLandlordID='" & szSelectedClient & "' and (isfinalized=0 or isnull(isfinalized))", adoConn, adOpenStatic, adLockReadOnly
    If Not rstblPurInv.EOF Then
            PreviousStatementFinalized = True
            rstblPurInv.Close
            adoConn.Close
            Exit Function
    End If
    rstblPurInv.Close
    adoConn.Close
    Set adoConn = Nothing
End Function
Private Function ComparePreviousStatementFinalized() As Boolean
    Dim adoConn As New ADODB.Connection
    adoConn.Open getConnectionString
    Dim rstblPurInv As New ADODB.Recordset
    rstblPurInv.Open "Select * from RentSummaryStatement where  ClientIDLandlordID='" & szSelectedClient & "' and isfinalized=1 order by StatementDate DESC", adoConn, adOpenStatic, adLockReadOnly
    If Not rstblPurInv.EOF Then
             If DateDiff("d", rstblPurInv("DateFinalized").Value, txtStatementDate1.text) < 0 Then
                    ComparePreviousStatementFinalized = True
                    rstblPurInv.Close
                    adoConn.Close
                    Exit Function
             End If
    End If
    rstblPurInv.Close
    adoConn.Close
    Set adoConn = Nothing
End Function

Private Function FeeshasBeenPaid() As Boolean
    Dim adoConn As New ADODB.Connection
    adoConn.Open getConnectionString
    Dim rstblPurInv As New ADODB.Recordset
    rstblPurInv.Open "Select * from tblPurInv V,tlbPayment P  where V.MY_ID=P.PI AND isManagementFee=true and OSAmount>0 and TRAN_DATE<= #" & _
    Format(txtStatementDate1.text, "dd/MMM/yyyy") & " # AND TRAN_DATE>#" & Format(txtLastStatementDate1.text, "dd/MMM/yyyy") & "# and ClientID='" & szSelectedClient & "'", adoConn, adOpenStatic, adLockReadOnly
    If Not rstblPurInv.EOF Then
            FeeshasBeenPaid = True
            rstblPurInv.Close
            adoConn.Close
            Exit Function
    End If
    rstblPurInv.Close
    adoConn.Close
    Set adoConn = Nothing
End Function
Private Function GetLastStatementNoByClientNew(szSelectedClient As String) As String
    Dim intmaxStatementNo As Integer
    Dim adoConn As New ADODB.Connection
    Dim rsRentSummaryStatement As New ADODB.Recordset
    
    adoConn.Open getConnectionString
    Dim szSQL As String
    'Rent Summary Statement Opening Balance (=Closing balance of previous statement)
    'This is by client
    'Get ID by Client max ID from RentSummaryStatement
    szSQL = "Select StatementNo as IDbyCL,isfinalized from RentSummaryStatement where ClientIDLandlordID='" & szSelectedClient & "'  order by StatementNo"
    rsRentSummaryStatement.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
    If Not rsRentSummaryStatement.EOF Then
        GetLastStatementNoByClientNew = IIf(IsNull(rsRentSummaryStatement!isfinalized), "", rsRentSummaryStatement!isfinalized)
    End If
    rsRentSummaryStatement.Close
    Set rsRentSummaryStatement = Nothing
    
    
    adoConn.Close
    Set adoConn = Nothing
End Function

Private Sub cmdSave_Click()
    Dim rCount As Integer
    Dim selRow As Integer
    Dim szSelectedClient As String
    Dim szSelectedPropertyID As String
    Dim iIncDec As Long
       Dim szStatmentID As String
    Dim szStatmentNo As String
    Dim szReportGenID As String
    Dim szFinalized As String
    Dim szSQL As String
    'Backup code start
'    If ComparePreviousStatementFinalized = True Then
'         MsgBox "There is previous client statement that is showing as not having been finalised as at the statement date selected. " & _
'                "Please finalise this previous client statement or select a statement date on or after your previous statement was finalised", vbInformation, "Warning!"
'        Exit Sub
'   End If
Rem by anol 2023-04-20
    If PreviousStatementFinalized = True Then
        MsgBox "There are previous client statements that have not been finalised for this client. " & _
                "You must finalise all previous client statements before proceeding.", vbInformation, "Warning!"
        Exit Sub
    End If
    If MsgBox("Do you wish to take a backup?", vbYesNo + vbQuestion, "Data Backup") = vbYes Then
      If BackupDB Then
         MsgBox "Backup successful.", vbInformation, "The Backup has been successful"
      Else
         MsgBox "Backup failed. Please try again", vbInformation, "Warning"
      End If
   End If
   'Backup code End
    If bEditMode = True Then
        If PIvalidation = False Then
            Exit Sub
        End If
    End If
    If txtLastStatementDate1.text = "" Then
           txtLastStatementDate1.Locked = False
           If szStatementNo = 1 Then GoTo XX ' you can keep empty for the first statement date else you need to must enter date
           MsgBox "Please enter last statement date", vbInformation, "Warning!"
           
           Exit Sub
    Else
           If DateDiff("d", txtStatementDate1.text, txtLastStatementDate1.text) >= 0 And bEditMode = False Then
               MsgBox "A statement already exists for this date. Please enter a date after the 'Last Statement Date'", vbInformation, "Statement Date!!!"
               Exit Sub
           End If
    
    End If
   
   
    If Val(txtRentPayable.text) > Val(txtAvailableFunds.text) Then
        MsgBox "Rent Payable amount cannot be greater than the available funds", vbInformation, "Warning!"
        Exit Sub
    End If
    If Trim(txtStatementDate1.text) = "" Then
              MsgBox "Please enter a statement date ", vbInformation, "Statement Date!!!"
              FocusControl txtStatementDate1
        Exit Sub
    End If
    
    
XX:
    For rCount = 1 To flxClients.Rows - 1
         If flxClients.TextMatrix(rCount, 0) = "X" Then
             szSelectedClient = flxClients.TextMatrix(rCount, 1)
             iIncDec = iIncDec + 1
             selRow = rCount
         End If
    Next
    If iIncDec <> 1 Then
       MsgBox "Please select a client.", vbInformation + vbOKOnly, "Client Selection"
       FocusControl flxClients
       Exit Sub
    End If
    If isClientCondolidated(szSelectedClient) Then
        MsgBox "Please enable consolidated statements for this client on the client record", vbInformation, "Warning"
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
       MsgBox "Please select a Property.", vbInformation + vbOKOnly, "Property Selection"
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
'    If isAnyTransactionAvailable = False Then
'        MsgBox "There are no transactions for the statement period selected", vbInformation, "Warning!"
'               ' Exit Sub
'    End If
'   If PreviousStatementFinalized = True Then
'        MsgBox "There is a previous client statement that has not been finalised. ", vbInformation, "Previous client statement"
'        Exit Sub
'    End If
 Dim rsRentSummaryStatement As New ADODB.Recordset
    szReportGenID = ReportGenID
    szStatmentID = GetLastStatementID + 1
    szStatmentNo = GetLastStatementNoByClient + 1
    If szStatmentNo = 1 Then
                'no message no warning
    ElseIf szStatmentNo > 1 Then
                    Dim adoConn As New ADODB.Connection
                    adoConn.Open getConnectionString
                    szSQL = "Select isfinalized from RentSummaryStatement where StatementNo=" & (szStatmentNo - 1) & " AND ClientIDLandlordID='" & _
                    szSelectedClient & "'"
                    rsRentSummaryStatement.Open szSQL, adoConn, adOpenKeyset, adLockReadOnly
                    If Not rsRentSummaryStatement.EOF Then
                       szFinalized = IIf(IsNull(rsRentSummaryStatement("isfinalized").Value), "", rsRentSummaryStatement("isfinalized").Value)
                    End If
                    rsRentSummaryStatement.Close
                    Set rsRentSummaryStatement = Nothing
                    If szFinalized <> "1" Then
                        adoConn.Close
                        MsgBox "It is not possible to produce a new client statement, unless the previous statement has been finalised", vbInformation, "Warning"
                        Exit Sub
                    End If
                    adoConn.Close
    End If
'    If GetLastStatementNoByClientNew(szSelectedClient) <> "1" Then
'            MsgBox "You cannot create a new statement. Please finalise previous statement for this client.", vbInformation, "Warning!"
'            Exit Sub
'    End If

    If GetSupplierOSAmountAllTime <> 0 Then
        If MsgBox("You have outstanding supplier balances to pay. Do you wish to pay them before producing your client statement?", vbYesNo, "Supplier Os Balance") = vbYes Then
                LoadForm frmPurchaseExpense
                frmPurchaseExpense.tabPurExp.Tab = 1
                frmPurchaseExpense.txtClientIDPurPay.text = szSelectedClient
                frmPurchaseExpense.txtBankCode.text = ""
                frmPurchaseExpense.txtBankAc.text = ""
                Exit Sub
        Else
            'proceed
        End If
    End If
    If GetClientOSAmount > 0 Then
        MsgBox "There is previous client statement that has not been finalised at the statement date selected. Please select a statement date on or after your last statement was finalised", vbInformation, "Client Os Balance"
        Exit Sub
'        If MsgBox("You have outstanding Client balances to pay. Do you wish to pay them before produce your client statement?", vbYesNo, "Client Os Balance") = vbYes Then
'                LoadForm frmPurchaseExpense
'                frmPurchaseExpense.tabPurExp.Tab = 1
'                frmPurchaseExpense.txtClientIDPurPay.text = szSelectedClient
'                frmPurchaseExpense.txtBankCode.text = ""
'                frmPurchaseExpense.txtBankAc.text = ""
'                Exit Sub
'        Else
'            'proceed
'        End If
    End If
    If chkShowDue.Value = 1 Then 'Code added by anol 2023/08/14
        If GetAgentOSAmount > 0 Then
            If MsgBox("You have outstanding Managing Agent balances to pay. Do you wish to pay them before produce your client statement?", vbYesNo, "Managing Agent Os Balance") = vbYes Then
                    LoadForm frmPurchaseExpense
                    frmPurchaseExpense.tabPurExp.Tab = 1
                    frmPurchaseExpense.txtClientIDPurPay.text = szSelectedClient
                    frmPurchaseExpense.txtBankCode.text = ""
                    frmPurchaseExpense.txtBankAc.text = ""
                    Exit Sub
            Else
                'proceed
            End If
        End If
    End If
    
    'You have management fees to generate. Do you wish to generate them before producing your client statement
    'check if management fee has been generated or not
    If Feestogenerate Then
            If MsgBox("You have management fees to generate. Do you wish to generate  " & _
                "them before producing your client statement?", vbYesNo, "Please confirm") = vbYes Then
'                 frmManagementFeeSelection.Caption = "Management Fee Preview"
'                 frmManagementFeeSelection.szCallingFrom = "ManagementFee Preview"
'                 LoadForm frmManagementFeeSelection
                 LoadForm frmManagementFees
                 Exit Sub
            Else
                'proceed
            End If
    End If
    'check if management fee has been paid or not rem 2023-08-14
    If FeeshasBeenPaid Then
        If MsgBox("You have management fees to pay. Do you wish to pay  " & _
                "them before producing your client statement?", vbYesNo, "Please confirm") = vbYes Then
                LoadForm frmPurchaseExpense
                frmPurchaseExpense.tabPurExp.Tab = 1
                 frmPurchaseExpense.txtClientIDPurPay.text = szSelectedClient
                 frmPurchaseExpense.txtBankCode.text = ""
                frmPurchaseExpense.txtBankAc.text = ""
                 Exit Sub
            Else
                'proceed
            End If
    End If
    'check if we have finalized the last/previous statement or not
 
    
    
   
    If MsgBox("Are you sure, you wish to produce this statement?", vbYesNo, "Please confirm") = vbYes Then

       ' whichFieldToCheck = "RentSumStatementS"
        If boolConsolidatedStatement = 1 Then
            Call MarkAllTransactionsWithCSID(szStatmentID)
            Call ProduceClientStatement(szStatmentID, szReportGenID)  'Write into SummaryStatement table in this function
            
        Else        'For each property write a statement for non consolidated
'            For rCount = 1 To flxProperties.Rows - 1
'                 If flxProperties.TextMatrix(rCount, 0) = "X" Then
'                     szSelectedPropertyID = flxProperties.TextMatrix(rCount, 1)
'                     Call GenerateSummaryStatementNonConsolidated(szStatmentID, szSelectedPropertyID, szReportGenID)
'                     szStatmentID = CLng(szStatmentID) + 1
'                 End If
'             Next
        End If
        
    
        'run TestReportForRentSummary.rpt
       
        Dim reportApp As New CRAXDRT.Application
        Dim Report As CRAXDRT.Report
      
        Set Report = reportApp.OpenReport(App.Path & szReportPath & "\RentSummaryStatement.rpt")
        Report.EnableParameterPrompting = False
        Report.DiscardSavedData
        Report.ParameterFields(1).AddCurrentValue CLng(szReportGenID)
        Load frmReport
        frmReport.LoadReportViewer Report
'         Dim szCurrentStatementIDID As String
'            For rCount = 1 To frmRentPayable.flxPayFees.Rows - 1
'                 If frmRentPayable.flxPayFees.TextMatrix(rCount, 0) = "X" Then
'                     iIncDec = iIncDec + 1
'                     selRow = rCount
'                 End If
'            Next
'            szCurrentStatementIDID = frmRentPayable.flxPayFees.TextMatrix(selRow, 2)
'            If selRow > 0 Then
'                Call printClientStatement(szCurrentStatementIDID, selRow)
'            End If
            Call frmRentPayable.loadflxPayFees("")
    End If
'    Frame1(6)  .Visible = False
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

Private Sub cmdViewAllRetention_Click()
     LoadForm frmRetentionMaster
End Sub

Private Sub Command1_Click()
        'This is recalculate rent summary sub procedure where we are clearing all the flags when we press recalculate button
        Dim adoConn As New ADODB.Connection
        adoConn.Open getConnectionString
        adoConn.Execute "Delete from RentSummaryStatement"
        adoConn.Execute "Delete from RentSummaryStatementPreview"
        adoConn.Execute "Update tlbBankPayment Set RentSumStatement=''"
        adoConn.Execute "Update tlbPayment Set RentSumStatement=''"
        adoConn.Execute "Update tlbReceipt Set RentSumStatement=''"
        Call frmRentPayable.loadflxPayFees("")
'    adoconn.Execute "Update tlbBankPayment Set RentSumStatement='' where RentSumStatement='" & szCurrentStatementID & "'"
'    adoconn.Execute "Update tlbPayment Set RentSumStatement='' where RentSumStatement='" & szCurrentStatementID & "'"
'    adoconn.Execute "Update tlbReceipt Set RentSumStatement='' where RentSumStatement='" & szCurrentStatementID & "'"
    adoConn.Close
    Set adoConn = Nothing
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
                chkInFunds.Value = 1
                chkInFunds_Click
                Exit For
            End If
    Next
    If hasSelBankAccounts = False Then
        Call ConfigFlxProperties
    End If
End Sub



Private Sub flxBankAccounts_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call flxBankAccounts_Click
        FocusControl flxProperties
    End If
    
End Sub

Private Sub flxClients_Click()
        txtAvailableFunds.text = "0.00"
        txtRentPayable.text = "0.00"
'     SelectOnly1RowFlxGrid flxClients, flxClients.row, 0
     'SelectOnly1RowFlxGrid flxBankAccounts, flxBankAccounts.row, 0
     SelectOnly1RowFlxGrid flxClients, flxClients.row, 0
     'Select1RowFlxGrid flxClients, flxClients.row, 0
     szSelectedClient = flxClients.TextMatrix(flxClients.row, 1)
     'Auto select Properties bases on whatever you select at client
     Dim iRow As Integer
     Dim rCount As Integer
   Dim adoConn1 As New ADODB.Connection
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
    If szSelectedClient <> "" And szSelectedClient <> "ClientID" Then
        PurchaseLedgerControl = GetNominalCodeForControlAccount(adoConn1, "Client/Landlord Control Account (B/S)", szSelectedClient)
        If (PurchaseLedgerControl = "") Then
            MsgBox "Please setup Rent and Other Amounts Payable control accounts for '" & szSelectedClient & "' "
            Exit Sub
        End If
    
    
        PurchaseLedgerControl = GetNominalCodeForControlAccount(adoConn1, "Rent & Other Amounts Payable (P&L)", szSelectedClient)
        If (PurchaseLedgerControl = "") Then
            MsgBox "Please setup Rent and Other Amounts Payable  control accounts for '" & szSelectedClientName & "' "
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
        Dim adoConn As New ADODB.Connection
        Dim adoRst As New ADODB.Recordset
        adoConn.Open getConnectionString
        Dim rsConsolidatedStatement As New ADODB.Recordset
        'adoConn.Open getConnectionString
        rsConsolidatedStatement.Open "Select * from client where clientID ='" & szSelectedClient & "'", adoConn, adOpenStatic, adLockReadOnly
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
        adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
        iRow = 1

        While Not adoRst.EOF 'While Not adoRst.EOF
            flxProperties.AddItem ""
            flxProperties.TextMatrix(iRow, 0) = ""
            flxProperties.TextMatrix(iRow, 1) = adoRst("PROPERTYID").Value
            flxProperties.TextMatrix(iRow, 2) = adoRst("PROPERTYNAME").Value
            flxProperties.TextMatrix(iRow, 3) = adoRst("ClientID").Value
            flxProperties.RowHeight(iRow) = 280
'            If iRow > 1 Then
'                flxProperties.AddItem ""
'            End If
            iRow = iRow + 1
    
            adoRst.MoveNext
        'End If
        Wend
        If flxProperties.TextMatrix(iRow, 1) = "" Then
            flxProperties.RemoveItem iRow
        End If
       adoConn.Close
       Set adoConn = Nothing
       If boolConsolidatedStatement = 1 Then
            chkAllProperties.Value = 1
       End If
End Sub





Private Sub flxClients_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call flxClients_Click
        FocusControl flxBankAccounts
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
'Public Sub SelectOnly1RowFlxGrid(conFlxGrid As Control, iNewRow As Integer, Optional iColID As Integer = 0)
'   Dim iRow       As Integer
'   Dim iCol       As Integer
'   Dim iColPaint  As Integer
'
'   iColPaint = IIf(iColID = 0, 1, 0)
'
'   For iRow = conFlxGrid.Rows - 1 To 1 Step -1
'      If conFlxGrid.TextMatrix(iRow, iColID) = "X" Then
'         If iRow = iNewRow And conFlxGrid.TextMatrix(iRow, iColID) = "X" Then
'                conFlxGrid.TextMatrix(iRow, iColID) = ""
'                conFlxGrid.TextMatrix(iRow, iColID) = ""
'                conFlxGrid.row = iRow
'                For iCol = iColPaint To conFlxGrid.Cols - 1
'                   conFlxGrid.col = iCol
'                   conFlxGrid.CellBackColor = vbWhite
'                Next iCol
'                Exit Sub
'         End If
''         conFlxGrid.TextMatrix(iRow, iColID) = ""
''         conFlxGrid.row = iRow
''         For iCol = iColPaint To conFlxGrid.Cols - 1
''            conFlxGrid.col = iColvj
''            conFlxGrid.CellBackColor = vbWhite
''         Next iCol
'      End If
'   Next iRow
'
'   conFlxGrid.TextMatrix(iNewRow, iColID) = "X"
'   conFlxGrid.row = iNewRow
'
'   For iCol = conFlxGrid.Cols - 1 To iColPaint Step -1
'      conFlxGrid.col = iCol
'      conFlxGrid.CellBackColor = RGB(174, 179, 233)
'   Next iCol
'End Sub
Public Sub SelectOnly1RowFlxGrid(conFlxGrid As Control, iNewRow As Integer, Optional iColID As Integer = 0)
   Dim iRow       As Integer
   Dim iCol       As Integer
   Dim iColPaint  As Integer

   iColPaint = IIf(iColID = 0, 1, 0)
   
   For iRow = 1 To conFlxGrid.Rows - 1
      If conFlxGrid.TextMatrix(iRow, iColID) = "X" Then
         If iRow = iNewRow And conFlxGrid.TextMatrix(iRow, iColID) = "X" Then Exit Sub
         conFlxGrid.TextMatrix(iRow, iColID) = ""
         conFlxGrid.row = iRow
         For iCol = iColPaint To conFlxGrid.Cols - 1
            conFlxGrid.col = iCol
            conFlxGrid.CellBackColor = vbWhite
         Next iCol
      End If
   Next iRow

   conFlxGrid.TextMatrix(iNewRow, iColID) = "X"
   conFlxGrid.row = iNewRow

   For iCol = iColPaint To conFlxGrid.Cols - 1
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
Public Function GetFNCGlobalDataPropertyWise(szPropertyID As String, Conn As ADODB.Connection) As Boolean
   'gets the global data from the global data table and puts the payment dates and
   'VAT rate, base rate, number of days to send demands before due, and price per sq foot for
   'service charge and puts then to global variables for when needed by program later.

   'This procedure will be called when program is opened and when the global data is
   'changed.

   Dim i As Integer, iDateSet As Integer
   Dim rst As ADODB.Recordset
   Dim SQLStr As String
   
   Set rst = New ADODB.Recordset
   
   SQLStr = "SELECT * FROM GlobalData WHERE PropertyID = '" & szPropertyID & "';"
   rst.Open SQLStr, Conn, adOpenDynamic, adLockPessimistic

   If rst.EOF Then
        MsgBox "Please enter global data for this property", vbInformation, "Warning"
        rst.Close
        Exit Function
   End If
   rst.Close
   SQLStr = "SELECT PropertyID, MDueDate1,MDueDate2,MDueDate3,MDueDate4,MDueDate5,MDueDate6,MDueDate7," & _
            "MDueDate8,MDueDate9,MDueDate10,MDueDate11,MDueDate12,QDueDate1,QDueDate2,QDueDate3,QDueDate4,HYDueDate1,HYDueDate2,YDueDate " & _
            "FROM GlobalData WHERE GlobalData.PropertyID = '" & szPropertyID & "';"
   
   rst.Open SQLStr, Conn, adOpenDynamic, adLockPessimistic

   If rst.EOF Then
       ShowMsgInTaskBar "You Need to Enter the Global Data.", , "N"
       rst.Close
       Set rst = Nothing
       GetFNCGlobalDataPropertyWise = False
       Exit Function
   End If
   szGDYearly = rst!YDueDate
   szGDHalfYearly1 = rst!HYDueDate1
   szGDHalfYearly2 = rst!HYDueDate2
   szGDQuarterly1 = rst!QDueDate1
   szGDQuarterly2 = rst!QDueDate2
   szGDQuarterly3 = rst!QDueDate3
   szGDQuarterly4 = rst!QDueDate4

   szaMonthlyGD(0) = rst!MDueDate1
   szaMonthlyGD(1) = rst!MDueDate2
   szaMonthlyGD(2) = rst!MDueDate3
   szaMonthlyGD(3) = rst!MDueDate4
   szaMonthlyGD(4) = rst!MDueDate5
   szaMonthlyGD(5) = rst!MDueDate6
   szaMonthlyGD(6) = rst!MDueDate7
   szaMonthlyGD(7) = rst!MDueDate8
   szaMonthlyGD(8) = rst!MDueDate9
   szaMonthlyGD(9) = rst!MDueDate10
   szaMonthlyGD(10) = rst!MDueDate11
   szaMonthlyGD(11) = rst!MDueDate12

   rst.Close
   Set rst = Nothing
   GetFNCGlobalDataPropertyWise = True
End Function
Public Function NextDueDate1(FrequencyID As Integer, dtStartDate As Date, PROPERTY_ID As String) As Date
   If PROPERTY_ID = "" Then
       MsgBox "You must select a property!", vbOKOnly + vbCritical, "No property Selected"
       Exit Function
   End If

   Dim adoConn As New ADODB.Connection

   adoConn.Open getConnectionString

   If Not GetFNCGlobalDataPropertyWise(PROPERTY_ID, adoConn) Then Exit Function

   adoConn.Close
   Set adoConn = Nothing

  
      Select Case FrequencyID
            Case 1:                              'Weekly in advance
               NextDueDate1 = dtStartDate
            Case 2:                              'Weekly in arrears
              NextDueDate1 = DateAdd("d", 7, dtStartDate)
            Case 3:                              'Fortnightly in advance
               NextDueDate1 = dtStartDate
            Case 4:                              'Fortnightly in arrears
               NextDueDate1 = DateAdd("d", 14, dtStartDate)
            Case 5:                              'Monthly in advance
               NextDueDate1 = NextPayingDate(dtStartDate, InAdv, Pay_Monthly)
            Case 6:                              'Monthly in arrears
               NextDueDate1 = NextPayingDate(dtStartDate, InArr, Pay_Monthly)
            Case 7:                              'Quarterly in advance
               NextDueDate1 = NextPayingDate(dtStartDate, InAdv, Pay_Quarterly)
            Case 8:                              'Quarterly in arrears
               NextDueDate1 = NextPayingDate(dtStartDate, InArr, Pay_Quarterly)
            Case 9:                              'Half yearly in advance
               NextDueDate1 = NextPayingDate(dtStartDate, InAdv, Pay_Half_Yearly)
            Case 10:                              'Half yearly in arrears
               NextDueDate1 = NextPayingDate(dtStartDate, InArr, Pay_Half_Yearly)
            Case 11:                             'yearly in advance
               NextDueDate1 = NextPayingDate(dtStartDate, InAdv, Pay_Yearly)
            Case 12:                             'yearly in arrears
               NextDueDate1 = NextPayingDate(dtStartDate, InArr, Pay_Yearly)
            Case 13:                             'Daily
               NextDueDate1 = ""
            Case 14:                             '4 Weekly in advance
               NextDueDate1 = dtStartDate
            Case 15:                             '4 Weekly in arrears
               NextDueDate1 = DateAdd("d", 28, dtStartDate)
            Case 16:                             '4 Monrhly in advance
               NextDueDate1 = dtStartDate
            Case 17:                             '4 Monrhly in arrears
               NextDueDate1 = DateAdd("m", 4, dtStartDate)
      End Select

     
End Function

Private Sub Form_Activate()
    Dim adoConn As New ADODB.Connection
    Dim iRow As Integer
    Dim szSelectedFunds1  As String
    Dim szSelectedProperties1 As String
    Dim szSelectedBankAC1 As String
    If bEditMode = True Then
        If szCurrentStatementID = "" Then Exit Sub
        adoConn.Open getConnectionString
        Dim szSQL As String
        Dim rsRentSummaryStatement As New ADODB.Recordset
'        Dim adoconn As New ADODB.Connection
'        adoconn.Open getConnectionString
        szSQL = "Select StatementNo,ClientIDLandlordID,ListOfFundId,ListOfinputProperties,BankCode,Retentions,PayableAmount,AvailableFunds,PreviousStatementDate," & _
                "StatementDate from RentSummaryStatement where StatementID=" & Replace(szCurrentStatementID, "CS", "") & ""
        rsRentSummaryStatement.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
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
                    If IsNull(rsRentSummaryStatement("PreviousStatementDate").Value) = False Then
                        txtLastStatementDate1.text = Format(rsRentSummaryStatement("PreviousStatementDate").Value, "dd/MM/yyyy")
                    Else
                         txtLastStatementDate1.text = "01/01/2000"
                    End If
                     If IsNull(rsRentSummaryStatement("StatementDate").Value) = False Then
                        txtStatementDate1.text = Format(rsRentSummaryStatement("StatementDate").Value, "dd/MM/yyyy")
                    Else
                         txtStatementDate1.text = ""
                    End If
                    szStatementNo = rsRentSummaryStatement("StatementNo").Value
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
        adoConn.Close
        Set adoConn = Nothing
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
    Dim adoConn As New ADODB.Connection
    Me.BackColor = MODULEBACKCOLOR
    chkExcludeSupOS.BackColor = MODULEBACKCOLOR
    Frame1(6).BackColor = MODULEBACKCOLOR
    chkShowDue.BackColor = MODULEBACKCOLOR
    Frame1(12).BackColor = MODULEBACKCOLOR
    Frame1(8).BackColor = MODULEBACKCOLOR
    Frame1(9).BackColor = MODULEBACKCOLOR
    Frame1(10).BackColor = MODULEBACKCOLOR
    chkAllProperties.BackColor = MODULEBACKCOLOR
    Label2.BackColor = MODULEBACKCOLOR
    Label3.BackColor = MODULEBACKCOLOR
    Label1.BackColor = MODULEBACKCOLOR
    'Frame1(13).BackColor = MODULEBACKCOLOR
'    Frame1(7).BackColor = MODULEBACKCOLOR
'    Frame1(14).BackColor = MODULEBACKCOLOR
    chkConsolidatedCS.BackColor = MODULEBACKCOLOR
    chkExcludeReceipt.BackColor = MODULEBACKCOLOR
    chkExcludeSupOS.Value = 0
    chkShowDue.Value = 0
    chkExcludeSupOS.Enabled = False
    chkShowDue.Enabled = False
    If Len(txtLastStatementDate1.text) < 10 Then txtLastStatementDate1.text = Format("01/01/2000", "dd/mm/yyyy")
    If Len(txtStatementDate1.text) < 10 Then txtStatementDate1.text = Format(Date, "dd/mm/yyyy")
    SelTxtInCtrl txtLastStatementDate1
    chkInFunds.BackColor = MODULEBACKCOLOR
    Call ConfigflxRetensionDetails
    Dim szSelectedFunds1 As String
    Dim szSelectedProperties1 As String
    Dim szSelectedBankAC1 As String
  '  txtStatementDate1.text = Format(Now, "dd/mm/yyyy")
    Me.Width = 14985
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
'        Call loadfrmRentPayable.flxPayFees
    End If
    Call WheelHook(Me.hWnd)
End Sub
Private Sub LoadLaststatementdate()
    Dim szSQL As String
    Dim rsRentSummaryStatement As New ADODB.Recordset
    Dim adoConn As New ADODB.Connection
    adoConn.Open getConnectionString
    szSQL = "Select StatementDate,isfinalized from RentSummaryStatement where StatementNo=" & GetLastStatementNoByClient & " AND ClientIDLandlordID='" & szSelectedClient & "'"
    rsRentSummaryStatement.Open szSQL, adoConn, adOpenKeyset, adLockReadOnly
    If Not rsRentSummaryStatement.EOF Then
        txtLastStatementDate1.text = rsRentSummaryStatement!StatementDate
        szLastStatementDate = rsRentSummaryStatement!StatementDate
         'txtStatementDate1.text = rsRentSummaryStatement!StatementDate
'          txtStatementDate1.text = Format(Now, "dd/mm/yyyy") 'always now
        If IsNull(rsRentSummaryStatement!isfinalized) = False Then
            If rsRentSummaryStatement!isfinalized = "1" Then
                    txtLastStatementDate1.Locked = True
             Else
                    txtLastStatementDate1.Locked = False
             End If
        End If
    Else
'        txtLastStatementDate1.text = Format(Date, "dd/mm/yyyy")
'            txtStatementDate1.text = Format(Now, "dd/mm/yyyy")
            txtLastStatementDate1.Locked = False
    End If
    
    rsRentSummaryStatement.Close
    Set rsRentSummaryStatement = Nothing
    adoConn.Close
    Set adoConn = Nothing
End Sub

'Private Sub txtAvailableFund1_KeyPress(KeyAscii As Integer)
'    DigitTextKeyPress txtAvailableFund1, KeyAscii
'End Sub

Private Sub txtAvailableFunds_KeyPress(KeyAscii As Integer)
     DigitTextKeyPress txtAvailableFunds, KeyAscii
End Sub

Private Sub txtClientSearch_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        FocusControl flxClients
    End If
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
'   If Len(txtLastStatementDate1.text) < 10 Then txtLastStatementDate1.text = Format(Date, "dd/mm/yyyy")
   SelTxtInCtrl txtLastStatementDate1
End Sub
Private Sub txtLastStatementDate1_Change()
    TextBoxChangeDate txtLastStatementDate1
End Sub

Private Function NextID(adoConn As ADODB.Connection) As Long
   Dim szSQL As String
   Dim adoRst As New ADODB.Recordset
   szSQL = "SELECT MAX(Cint(StatementID))+1 AS Ref FROM RentSummaryStatement;"
   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
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
   flxBankAccounts.ColWidth(3) = 4200       'Name
   flxBankAccounts.ColWidth(4) = 0          'ClientID
   flxBankAccounts.Rows = 2
End Sub
Private Sub LoadflxInFundsONEditInput(szSelectedClient As String)
   Dim adoRst   As New ADODB.Recordset
   Dim szSQL       As String
   Dim iRow As Integer
   Dim conClient As New ADODB.Connection
   
   ConfigFlxInFunds
   conClient.Open getConnectionString
    Dim rsFundMatrix As New ADODB.Recordset
    Dim iSel As Integer
    szSQL = "SELECT Distinct FundID, FundName, FundCode,CategoryCode FROM Fund LEFT JOIN tlbPayable PB ON F.FundID=agr.fund where PB.clientID='" & _
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
   MsgBox Err.description & "::" & Err.Number

   Set adoRst = Nothing
End Sub
Private Sub LoadflxInFunds()
   Dim adoRst   As New ADODB.Recordset
   Dim szSQL       As String
   Dim iRow As Integer
   Dim conClient As New ADODB.Connection
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
    Dim rsFundMatrix As New ADODB.Recordset
    Dim iSel As Integer
    szSQL = "SELECT Distinct FundID, FundName, FundCode,CategoryCode FROM Fund LEFT JOIN tlbPayable PB ON PB.Pay_fund=Fund.FUNDID where PB.clientID='" & _
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
      flxInFunds.TextMatrix(iRow, 0) = "X"
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
    chkInFunds.Value = 1
NoRes:
   adoRst.Close
   Set adoRst = Nothing
   Exit Sub

ErrorHandler:
   MsgBox Err.description & "::" & Err.Number

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
      .ColWidth(2) = 4000 'Label2(2).Left - Label2(1).Left 'Property Name
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
   Dim conClient As New ADODB.Connection
   Dim rstClient   As New ADODB.Recordset
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
   MsgBox Err.description & "::" & Err.Number

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
    flxClients.ColWidth(2) = 4200
    flxClients.ColWidth(3) = 0
    
End Sub
Private Sub LoadFlxClients()
   Call ConfigFlxClients
   Dim szSQL As String, r As Integer
   Dim adoConn As New ADODB.Connection
   Dim adoRst As New ADODB.Recordset

'   connect to database
   adoConn.Open getConnectionString

   szSQL = "SELECT SupplierID, SupplierName " & _
           "FROM Supplier where type in ('Client') order by SupplierID;"
   adoRst.Open szSQL, adoConn, adOpenDynamic, adLockPessimistic

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
   Dim adoRstFreq As ADODB.Recordset
   Dim adoConn As ADODB.Connection
   Dim strSQLTitles As String

   Set adoConn = New ADODB.Connection
   Set adoRstFreq = New ADODB.Recordset
   adoConn.Open getConnectionString
   strSQLTitles = "SELECT * FROM FREQUENCIES;"
   adoRstFreq.Open strSQLTitles, adoConn, adOpenStatic, adLockReadOnly

   ReDim szaFreq(adoRstFreq.RecordCount) As String

   While Not adoRstFreq.EOF
      szaFreq(adoRstFreq.Fields("ID").Value) = adoRstFreq.Fields("CALDAYS").Value
      adoRstFreq.MoveNext
   Wend
   adoRstFreq.Close
   adoConn.Close
   Set adoRstFreq = Nothing
   Set adoConn = Nothing
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
Private Function GeneratePIPreview() As Boolean
        Dim percentageOramount1 As Double
        Dim ManagementFeePreview As Long
        Dim lSlNumber As Long
        Dim adoConn As New ADODB.Connection
        Dim adoPIHeader As New ADODB.Recordset
        Dim adoPISplit As New ADODB.Recordset
        Dim szSQL As String
        Dim szSQLManagingAgent  As String
        Dim szMYID As String
        Dim szFundID As String
        Dim szSelectedPayableTypeID As String
        Dim szSelectedClient As String
        Dim szPropertySelection1 As String
        Dim szSQL1 As String
        Dim rsfixedMethod As New ADODB.Recordset
        Dim rsfixedMethodDetails As New ADODB.Recordset
        Dim j As Integer
        Dim vatFixedBasis As Double
        'Dim percnetageOramount As Double
        Dim dblGrandTotal As Double
        Dim dtNextDue As Date
        Dim dtFDD As Date
        Dim dblFeqID As Integer
        Dim dblTotalAmount As Double
        Dim lngMgtFeeSL As Long
        Dim iCountPI As Integer
        Dim iClientCount As Integer
        Dim iPropertyCount As Integer
        Dim strLastChargeDate  As String
        Dim strFundName As String
        Dim strStopDate  As String
        Dim dblCapAmount As Double
        Dim rsManagingAgent As New ADODB.Recordset
        Dim rsGlobalData As New ADODB.Recordset
        Dim szManagingAgent() As String
        Dim iManagingAgentCount As Integer
        Dim bControlACForPayable As Boolean
        Dim FinalControlACForPayable As String
        Dim szTemp
        Dim dblNoOfDaysToSendMFB4Due As Integer
        Dim dtNDDInitial As Date
        Dim strFromDate As String
        Dim strToDate As String
        Dim rsFromandToDate As New ADODB.Recordset
        Dim szSQLFrom As String
        Dim percnetageOramount As Double
        Dim SQLforInsert As String
        Dim dtNDD As Date
        
        Dim iCount As Long
        Dim iCount1 As Long
        Dim rCount As Long
        Dim dblFundId As Integer
        Dim dblDemandTypeId As Integer
        Dim i As Integer
        Dim lT_ID As Long
        Dim dicWarningProp As New Dictionary
        Dim dicWarningAgreement As New Dictionary
         Dim dicWarningFinPeriod As New Dictionary
      '  Dim strSelectedChargeType As String
         Dim strSelectedFundID As String
        Dim szSQL5 As String
        Dim rstSet As New ADODB.Recordset
         Dim szPropertySelectionALL As String
         Dim warning1 As String
         Dim warning2 As String
         Dim warning3 As String
         Dim iManagementFeePreview As Long
         
        If txtStatementDate1.text = "" Then
            'MsgBox "Please enter invoice date", vbInformation, "Warning "
            FocusControl txtStatementDate1
            Exit Function
        End If
        strSelectedFundID = ""
        For iCount = 1 To flxInFunds.Rows - 1
            If flxInFunds.TextMatrix(iCount, 2) <> "" Then
                   iCount1 = iCount1 + 1
                   If flxInFunds.TextMatrix(iCount, 0) = "X" Then
                        strSelectedFundID = strSelectedFundID + "" + flxInFunds.TextMatrix(iCount, 1) + ","
                   End If
            End If
        Next
        If Len(strSelectedFundID) > 0 Then
                strSelectedFundID = Left(strSelectedFundID, Len(strSelectedFundID) - 1)
        End If

        
        For rCount = 1 To flxClients.Rows - 1
            If flxClients.TextMatrix(rCount, 0) = "X" Then
               szSelectedClient = flxClients.TextMatrix(rCount, 1)
               Exit For
            End If
        Next
       
'        For rCount = 1 To flxProperties.Rows - 1
'                If flxProperties.TextMatrix(rCount, 0) = "X" Then
'                    szPropertySelectionALL = szPropertySelectionALL + "," + flxProperties.TextMatrix(rCount, 1)
'
'                End If
'        Next

        For rCount = 1 To flxProperties.Rows - 1
                If flxProperties.TextMatrix(rCount, 0) = "X" Then
                    If szPropertySelectionALL = "" Then
                        szPropertySelectionALL = "'" + flxProperties.TextMatrix(rCount, 1) + "'"
                    Else
                        szPropertySelectionALL = szPropertySelectionALL + ",'" + flxProperties.TextMatrix(rCount, 1) & "'"
                    End If

                End If
        Next


        For rCount = 1 To flxProperties.Rows - 1
            If flxProperties.TextMatrix(rCount, 0) = "X" Then
               szPropertySelection1 = flxProperties.TextMatrix(rCount, 1)
               Exit For
            End If
        Next
         Debug.Print time & " 1 start"
'         " I don't have to check  setup  of Management Fee is there or not
''        For rCount = 1 To flxProperties.Rows - 1
''                If flxProperties.TextMatrix(rCount, 0) = "X" Then
''                    szPropertySelection1 = flxProperties.TextMatrix(rCount, 1)
''                    If checksetupDONE(szPropertySelection1) = True Then
''                        ' MsgBox "Please enter a valid setup for the property: " & szPropertySelection1, vbInformation, "Client agreement Fees and charges setup"
''                          dicWarningAgreement.Add rCount, szPropertySelection1
''                    End If
'''                    If hasChargeTypeinlist(szPropertySelection1) = False Then
'''                              MsgBox "Please setup a demand type for the selected property: " & szPropertySelection1, vbInformation, "Warning"
'''                    End If
''                End If
''        Next
         Debug.Print time & " 2 End"
        
        If szPropertySelection1 = "" Then
            MsgBox "Please select a Property.", vbInformation, "Warning"
            Exit Function
        End If

        If strSelectedFundID = "" Then
            MsgBox "Please select a fund", vbInformation, "Warning"
            Exit Function
        End If
'        If chkAssignProperty.Value = 1 Then
'            If MsgBox("Are you sure you wish to generate your management fees without assigning a property", vbYesNo, "Confirm") = vbNo Then
'                Exit Sub
'            End If
'        End If
        
        
 '      ************************************Write tblPurInv **************************************
        
        adoConn.Open getConnectionString
        adoConn.Execute "Delete from tblPurInvPreview"
        adoConn.Execute "Delete from tblPurInvSRecPreview"
        adoConn.Execute "Delete from ManagementFeePreview"
       lSlNumber = SlNumber("PI", "tblPurInv", adoConn)
        adoConn.Close
  
    For iClientCount = 1 To flxClients.Rows - 1
            If flxClients.TextMatrix(iClientCount, 0) = "X" Then 'also check financial year is correct period and control account is there and it is set in a variable before looping
                        szSelectedClient = flxClients.TextMatrix(iClientCount, 1)
                        If szSelectedClient = "" Then Exit Function
                        adoConn.Open getConnectionString
                        
                        If FinalControlACForPayable = "" Then
                                  FinalControlACForPayable = GetNominalCodeForControlAccount(adoConn, "Management Fee Payable (P&L)", szSelectedClient)
                        End If
                            'if still control acount is empty generate a warning message and exit this sub procedure
                        If FinalControlACForPayable = "" Then
                                MsgBox "Control Account is not set for this Payable type for Client:" & szSelectedClient, vbInformation, "Warning"
                                GoTo EndOfAgreement
                        End If
                        If IsPeriodStatus(txtStatementDate1.text, szSelectedClient, adoConn) = 0 Then
                            MsgBox "The posting date cannot fall within a closed financial period for the client :" & szSelectedClient, vbInformation, "Warning"
                            adoConn.Close
                            'FocusControl txtStatementDate1
                            GoTo EndOfAgreement
                        ElseIf IsPeriodStatus(txtStatementDate1.text, szSelectedClient, adoConn) = 9 Then
                            'MsgBox "The posting date does not fall in any existing financial period :" & szSelectedClient, vbInformation, "Warning"
                            dicWarningFinPeriod.Add iClientCount, szSelectedClient
                            adoConn.Close
                            'FocusControl txtStatementDate1
                            GoTo EndOfAgreement
                        End If
                        adoConn.Close
            
            For iPropertyCount = 1 To flxProperties.Rows - 1
                    If iPropertyCount >= flxProperties.Rows Then Exit For
                    If flxProperties.TextMatrix(iPropertyCount, 0) = "X" And szSelectedClient = flxProperties.TextMatrix(iPropertyCount, 3) Then
                            Debug.Print iPropertyCount
                            szPropertySelection1 = flxProperties.TextMatrix(iPropertyCount, 1)
                            If checkGlobalDataEntered(szPropertySelection1) = True Then
                                dicWarningProp.Add iPropertyCount, szPropertySelection1
                                 'MsgBox "Please enter a valid Client Global Settings setup for the property: " & szPropertySelection1, vbInformation, "Client Global Data setup"
                                 'Exit For
                                GoTo EndOfAgreement
                            End If
                            If adoConn.State = 0 Then
                                    adoConn.Open getConnectionString
                            End If
'                            adoconn.BeginTrans
                            
                            
                            
            Dim rsCharge As New ADODB.Recordset
            szTemp = ""
            szSQLManagingAgent = "SELECT DISTINCT agr.ManagingAgentID " & _
              "FROM tlbAgreement agr, ClientProAgr CPA, ChargeTypes C,Fund  F,SECONDARYCODE SC " & _
              "WHERE agr.CPA_ID = CPA.CPA_ID And F.FundID=agr.fund AND  SC.CODE=agr.CHARGE_METHOD AND " & _
              "CPA.ClientID = '" & szSelectedClient & "' And C.ID = agr.CHARGE_TYPE And " & _
              "CPA.PropertyID = '" & szPropertySelection1 & "' ANd F.FundID IN(" & strSelectedFundID & ")"
              rsManagingAgent.Open szSQLManagingAgent, adoConn, adOpenDynamic, adLockOptimistic
              szTemp = SQL2String(rsManagingAgent, 0)
              rsManagingAgent.Close
              If Len(szTemp) > 0 Then
                    szManagingAgent = Split(szTemp, ",")
              Else
                    adoConn.Close
                    GoTo EndOfAgreement
              End If
              
              
             szSQL5 = "SELECT MAX(ManagementFeeSL) AS x FROM tblPurInv;"
             rstSet.Open szSQL5, adoConn, adOpenStatic, adLockReadOnly
             lngMgtFeeSL = Val(IIf(IsNull(rstSet.Fields.Item(0).Value), 0, rstSet.Fields.Item(0).Value)) + 1
             rstSet.Close
             Set rstSet = Nothing
             
             adoConn.Close
             
            For iManagingAgentCount = 0 To UBound(szManagingAgent) 'this for shall end after the creation of the PI preview
              
            adoConn.Open getConnectionString
            'For each managing agent I am creating PI
            'lSlNumber = SlNumber("PI", "tblPurInv", adoconn)
             
            szSQL = "SELECT agr.EachPeriod,agr.Capamount, agr.StopDate, CPA.agreementStartDate,CPA.agreementEndDate,agr.CHARGE_METHOD," & _
                "cpa.agreementEndDate as agreementEndD,agr.LastChargeDate,agr.TotalAmount,agr.Amount,agr.Fund,F.FundName, " & _
                "agr.NtDueDate,agr.FDD,agr.Frequency as FrequencyID,(Select FC.Frequency from Frequencies FC where  FC.ID=agr.Frequency) as FrequencyName,ManagingAgentID " & _
                "FROM tlbAgreement agr, ClientProAgr CPA, ChargeTypes C,Fund  F,SECONDARYCODE SC  " & _
                "WHERE agr.CPA_ID = CPA.CPA_ID AND  F.FundID=agr.fund AND SC.CODE=agr.CHARGE_METHOD AND " & _
                "CPA.ClientID = '" & szSelectedClient & "' And C.ID = agr.CHARGE_TYPE And " & _
                "CPA.PropertyID = '" & szPropertySelection1 & "' AND F.FundID in (" & strSelectedFundID & ") AND agr.ManagingAgentID='" & Trim(szManagingAgent(iManagingAgentCount)) & "'"
              
              
            '    szSQL = "SELECT agr.TotalAmount,agr.Fund,agr.Frequency ,agr.NtDueDate,agr.FDD " & _
            '            "FROM tlbAgreement agr, ClientProAgr CPA, DemandTypes D, ChargeTypes C,Fund  F,FREQUENCIES FC,SECONDARYCODE SC " & _
            '            "WHERE agr.CPA_ID = CPA.CPA_ID And cstr(F.FundID)=agr.fund AND FC.ID=agr.Frequency AND SC.CODE=agr.CHARGE_METHOD AND " & _
            '            "CPA.ClientID = '" & szSelectedClient & "' And D.ID = agr.DEMAND_TYPE And C.ID = agr.CHARGE_TYPE And " & _
            '            "CPA.PropertyID = '" & szPropertySelection1 & "'"
            rsCharge.Open szSQL, adoConn, adOpenDynamic, adLockOptimistic
            If rsCharge.EOF Then
                rsCharge.Close
                Set rsCharge = Nothing
                'MsgBox "Please setup an agreement for this property:" & szPropertySelection1, vbInformation, "Client agreement Fees and charge setup"
                dicWarningProp.Add iPropertyCount, szPropertySelection1
                GoTo EndOfAgreement
            End If

                            
            i = 1
                'Dim rsGlobalData As New ADODB.Recordset
                Dim VAT_ID As String
                Dim VAT_CODE As String
                Dim VAT_RATE As Double
                
                 rsGlobalData.Open "Select vatOptionEnabled,V.VAT_ID,V.VAT_CODE,V.VAT_RATE from (Globaldata G LEFT JOIN tlbVatCode V ON G.vatRate=V.VAT_ID) where PropertyID='" & _
                                    szPropertySelection1 & "' AND vatOptionEnabled=true", adoConn, adOpenStatic, adLockReadOnly
                 If Not rsGlobalData.EOF Then
                          VAT_ID = IIf(IsNull(rsGlobalData.Fields("VAT_ID").Value), "-1", rsGlobalData.Fields("VAT_ID").Value)
                          VAT_RATE = IIf(IsNull(rsGlobalData.Fields("VAT_RATE").Value), "0.00", rsGlobalData.Fields("VAT_RATE").Value)
                          VAT_CODE = IIf(IsNull(rsGlobalData.Fields("VAT_CODE").Value), "", rsGlobalData.Fields("VAT_CODE").Value)
                 Else
                          VAT_ID = -1
                          VAT_RATE = 0
                          VAT_CODE = ""
                 End If
                 rsGlobalData.Close
                 Set rsGlobalData = Nothing
                 
                 Dim rsGlobalData1 As New ADODB.Recordset
                 Dim bolVatOptionEnabled As Boolean
                 Dim bolOptedTotax As String
                 Dim strManagingAgentID As String
                 rsGlobalData1.Open "Select vatOptionEnabled from (Globaldata G LEFT JOIN tlbVatCode V ON G.vatRate=V.VAT_ID) where PropertyID='" & _
                                    szPropertySelection1 & "' ", adoConn, adOpenStatic, adLockReadOnly
                                    
                If Not rsGlobalData1.EOF Then
                        bolVatOptionEnabled = IIf(IsNull(rsGlobalData1("vatOptionEnabled").Value), False, rsGlobalData1("vatOptionEnabled").Value)
                        strManagingAgentID = rsCharge("ManagingAgentID").Value
                End If
                rsGlobalData1.Close
                
              
              
           
      
            szSQL = "SELECT MAX(TRANSACTIONID) AS TID FROM tlbPayment;"
            adoPIHeader.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
            lT_ID = CLng(IIf(IsNull(adoPIHeader!TID), 1, adoPIHeader!TID + 1))
            adoPIHeader.Close
             j = 1
                 rsGlobalData1.Open "SELECT optedTotax,* FROM Supplier where supplierID='" & strManagingAgentID & "'", adoConn, adOpenStatic, adLockReadOnly
                If Not rsGlobalData1.EOF Then
                        bolOptedTotax = rsGlobalData1("optedTotax").Value
                Else
                        bolOptedTotax = False
                End If
                rsGlobalData1.Close
           
   
            'Dim percnetageOramount As Double
            While Not rsCharge.EOF
                  dblTotalAmount = rsCharge.Fields.Item("TotalAmount").Value
                  dblCapAmount = rsCharge.Fields.Item("CapAmount").Value
                  dblFundId = rsCharge.Fields.Item("Fund").Value
                  'dblDemandTypeId = rsCharge.Fields.Item("DEMAND_TYPE").Value
                  
                  strFundName = rsCharge.Fields.Item("FundName").Value
                  If Not IsNull(rsCharge.Fields.Item("NtDueDate").Value) Then
                      dtNextDue = rsCharge.Fields.Item("NtDueDate").Value
                      dtNDDInitial = rsCharge.Fields.Item("NtDueDate").Value
                  End If
                  If Not IsNull(rsCharge.Fields.Item("FDD").Value) Then
                      dtFDD = rsCharge.Fields.Item("FDD").Value
                  End If
                  If Not IsNull(rsCharge.Fields.Item("FrequencyID").Value) Then
                    If rsCharge.Fields.Item("FrequencyID").Value <> "" Then
                            dblFeqID = rsCharge.Fields.Item("FrequencyID").Value
                      End If
                  End If
                szMYID = UniqueID()
'                strLastChargeDate = IIf(IsNull(rsCharge("LastChargeDate").Value), "", rsCharge("LastChargeDate").Value)
'                If strLastChargeDate = "" Then
'                       rsCharge.Close
'                       MsgBox "Please enter a last charge date for Property:" & szPropertySelection1, vbInformation, "Warning"
'                       GoTo EndOfAgreement
'                End If
                rsGlobalData.Open "Select NoOfDaysToSendMFB4Due from globaldata where PropertyID='" & szPropertySelection1 & "'", adoConn, adOpenStatic, adLockReadOnly
                If Not rsGlobalData.EOF Then
                    dblNoOfDaysToSendMFB4Due = IIf(IsNull(rsGlobalData!NoOfDaysToSendMFB4Due), 0, rsGlobalData!NoOfDaysToSendMFB4Due)
                End If
                rsGlobalData.Close
               
                If DateDiff("d", Date, rsCharge("NtDueDate").Value) > dblNoOfDaysToSendMFB4Due Then
                        GoTo EndOfChargeType
                End If
                 strStopDate = rsCharge("StopDate").Value
                If strStopDate = "" Then
                   ' Exit Sub
                Else
                    If IsDate(strStopDate) Then
                            If DateDiff("d", strStopDate, txtStatementDate1.text) >= 0 Then
'                                MsgBox "It is not possible to generate fees after the stop date.:" & szPropertySelection1, vbInformation, "Warning"
                            End If
                    Else
                            MsgBox "Stop date date format is not correct.:" & szPropertySelection1, vbInformation, "Warning"
                    End If
                    GoTo EndOfChargeType
                End If
            
                           
                                'validations
                                If rsCharge.Fields.Item("CHARGE_METHOD").Value = "RE_FIX" Then   'when working on the fixed procedure only 1 line of setup is done then (fixed basis)
                                                    'This validation is not valid for fixed method
'                                                dblTotalAmount = IIf(IsNull(rsfixedMethod("amt").Value), 0, rsfixedMethod("amt").Value)
'                                                If dblTotalAmount = 0 Then
'                                                        MsgBox "There are no fees to generate for property:" & szPropertySelection1, vbInformation, "Warning"
'                                                        GoTo EndOfAgreement
'                                                End If
                                                If rsCharge("agreementEndD").Value < rsCharge("FDD").Value Then
                                                        MsgBox "agreement End Date is greatar than following due date for the property:" & szPropertySelection1
                                                        GoTo EndOfAgreement
                                                End If
                                                If DateDiff("d", txtStatementDate1.text, rsCharge.Fields.Item("agreementStartDate").Value) > 0 Then
                                                        rsCharge.Close
                                                        Set rsCharge = Nothing
                                                        MsgBox "Charge date cannot be before Agreement Start date for property:" & szPropertySelection1, vbInformation, "Warning"
                                                        GoTo EndOfChargeType
                                                End If
'                                                 percnetageOramount = IIf(IsNull(rsCharge("amount").Value), 0, rsCharge("amount").Value)
                            '      ************************************Write tblPurInvSRec **************************************
                                                                                                                                   
                                                    dblTotalAmount = IIf(IsNull(rsCharge("EachPeriod").Value), 0, rsCharge("EachPeriod").Value) 'rsfixedMethodDetails.Fields.Item("Amt").Value
        '                                            dblGrandTotal = dblGrandTotal + dblTotalAmount
                                                    dblFundId = rsCharge.Fields.Item("Fund").Value
'                                                                                            If rsCharge.Fields.Item("CHARGE_METHOD").Value = "RE_ED" Then
'                                                                                                  dblTotalAmount = dblTotalAmount * percnetageOramount / 100
'                                                                                            End If
                                                    'dblGrandTotal = 50
                                                    If dblCapAmount > 0 Then
                                                           'make a condition in the split as well so that amount doesnot exeed cap amount
                                                           If dblTotalAmount > dblCapAmount Then
                                                                dblTotalAmount = dblCapAmount
                                                           End If
                                                    End If
'                                                    Dim rsNextdueDate As New ADODB.Recordset
'                                                    szSQL = "SELECT min(agr.NtDueDate),agr.FDD,agr.Frequency as FrequencyID,ManagingAgentID,agr.DEMAND_TYPE  " & _
'                                                            "FROM tlbAgreement agr, ClientProAgr CPA, DemandTypes D, ChargeTypes C,Fund  F,SECONDARYCODE SC " & _
'                                                            "WHERE agr.CPA_ID = CPA.CPA_ID And cstr(F.FundID)=agr.fund AND  SC.CODE=agr.CHARGE_METHOD AND " & _
'                                                            "CPA.ClientID = '" & szSelectedClient & "' And D.ID = agr.DEMAND_TYPE And C.ID = agr.CHARGE_TYPE And " & _
'                                                            "CPA.PropertyID = '" & szPropertySelection1 & "' AND DEMAND_TYPE IN(" & strSelectedFundID & ") " & _
'                                                            "AND agr.ManagingAgentID='" & Trim(szManagingAgent(iManagingAgentCount)) & "'"
'                                                    rsNextdueDate.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'                                                    If Not rsNextdueDate.EOF Then
                                                                 txtComparenextDueDate1 = DateAdd("d", 1, dtNextDue)
                                                                dtNDD = NextDueDate1(CInt(dblFeqID), txtComparenextDueDate1, szPropertySelection1)
                                                                txtComparenextDueDate1 = DateAdd("d", 1, dtNDD)
                                                                dtFDD = NextDueDate1(CInt(dblFeqID), txtComparenextDueDate1, szPropertySelection1)
                                                                strFromDate = Format(dtNextDue, "dd/MM/yyyy")
                                                                strToDate = Format(DateAdd("d", -1, dtNDD), "dd/MM/yyyy")
'                                                    End If
'                                                    rsNextdueDate.Close
                                                    
                                                    szSQL = "SELECT * FROM tblPurInvSRecPreview"
                                                   ' adoPISplit.Close
                                                    adoPISplit.Open szSQL, adoConn, adOpenDynamic, adLockOptimistic
                                                    'Add New Records. At least there is only one split line
                                                       With adoPISplit
                                                           .AddNew
                                                           .Fields.Item("MY_ID").Value = UniqueID()
                                                           .Fields.Item("ParentID").Value = szMYID
                                                           .Fields.Item("TRAN_ID").Value = j
                                                          ' .Fields.Item("TRANS").Value = szPropertySelection1  ' If you select One property then you can write a value here
                                                          'If chkAssignProperty.Value = 0 Then
                                                                 .Fields.Item("TRANS").Value = szPropertySelection1  ' If you select One property then you can write a value here
                                                           ' Else
                                                                 .Fields.Item("TRANS").Value = ""
                                                            'End If
                                                           .Fields.Item("UNIT_ID").Value = ""
                                                           .Fields.Item("NOMINAL_CODE").Value = FinalControlACForPayable
                                                           .Fields.Item("DEPT_ID").Value = dblFundId
                                                          ' .Fields.Item("JOB_ID").Value = flxPI.TextMatrix(iRow, 9)            'Job No
                                                           .Fields.Item("RecoverablePt").Value = 0
                                                           '' (Current Charge date)" '"MFee" + szPropertySelection1 + Format(lngMgtFeeSL, "0000") '"Management Fee"
                                                                
                                                           .Fields.Item("description").Value = "Management Fees for " & strFundName & " (" & strFromDate & " - " & strToDate & ")"
                                                           If bolVatOptionEnabled = True And bolOptedTotax = True Then
                                                                .Fields.Item("NET_AMOUNT").Value = dblTotalAmount
                                                                .Fields.Item("TAX_CODE").Value = VAT_CODE
                                                                .Fields.Item("VAT").Value = Round(dblTotalAmount * (VAT_RATE / 100), 2) 'VAT_RATE
                                                                 vatFixedBasis = .Fields.Item("VAT").Value
                                                                .Fields.Item("TOTAL_AMOUNT").Value = .Fields.Item("VAT").Value + dblTotalAmount
                                                                 dblTotalAmount = .Fields.Item("TOTAL_AMOUNT").Value
                                                           ElseIf bolVatOptionEnabled = True And bolOptedTotax = False Then
                                                                    .Fields.Item("NET_AMOUNT").Value = dblTotalAmount
                                                                    .Fields.Item("TAX_CODE").Value = Null 'VAT_CODE
                                                                    .Fields.Item("VAT").Value = 0 'Format(dblTotalAmount * (VAT_RATE / 100), "0.00") 'VAT_RATE
                                                                    .Fields.Item("TOTAL_AMOUNT").Value = dblTotalAmount
                                                                     dblTotalAmount = .Fields.Item("TOTAL_AMOUNT").Value
                                                          ElseIf bolVatOptionEnabled = False And bolOptedTotax = True Then 'bolVatOptionEnabled means global data and bolOptedTotax is supplier table
                                                                rsGlobalData.Open "Select V.VAT_ID,V.VAT_CODE,V.VAT_RATE from  Supplier S,tlbVatCode V where  (S.VATCode)=cstr(V.VAT_ID)  AND  SupplierID='" & _
                                                                                   strManagingAgentID & "' ", adoConn, adOpenStatic, adLockReadOnly
                                                                If Not rsGlobalData.EOF Then
                                                                         VAT_ID = IIf(IsNull(rsGlobalData.Fields("VAT_ID").Value), "-1", rsGlobalData.Fields("VAT_ID").Value)
                                                                         VAT_RATE = IIf(IsNull(rsGlobalData.Fields("VAT_RATE").Value), "0.00", rsGlobalData.Fields("VAT_RATE").Value)
                                                                         VAT_CODE = IIf(IsNull(rsGlobalData.Fields("VAT_CODE").Value), "", rsGlobalData.Fields("VAT_CODE").Value)
                                                                Else
                                                                         VAT_ID = -1
                                                                         VAT_RATE = 0
                                                                         VAT_CODE = ""
                                                                End If
                                                                rsGlobalData.Close

                                                                 'done modification on 15-10-2021
                                                                    '.Fields.Item("NET_AMOUNT").Value = dblTotalAmount
                                                                    .Fields.Item("TAX_CODE").Value = Null 'VAT_CODE
                                                                    .Fields.Item("VAT").Value = 0 'Format(dblTotalAmount * (VAT_RATE / 100), "0.00") ' Format(dblTotalAmount * (VAT_RATE / 100), "0.00") 'VAT_RATE' + Format(dblTotalAmount * (VAT_RATE / 100), "0.00")
                                                                    .Fields.Item("NET_AMOUNT").Value = dblTotalAmount + Round(dblTotalAmount * (VAT_RATE / 100), 2)
                                                                    .Fields.Item("TOTAL_AMOUNT").Value = dblTotalAmount + Round(dblTotalAmount * (VAT_RATE / 100), 2)
                                                                     dblTotalAmount = .Fields.Item("TOTAL_AMOUNT").Value
                                                                     
                                                          ElseIf bolVatOptionEnabled = False And bolOptedTotax = False Then
                                                                .Fields.Item("NET_AMOUNT").Value = dblTotalAmount
                                                                .Fields.Item("TAX_CODE").Value = Null 'VAT_CODE
                                                                .Fields.Item("VAT").Value = 0 'Format(dblTotalAmount * (VAT_RATE / 100), "0.00") 'VAT_RATE
                                                                .Fields.Item("TOTAL_AMOUNT").Value = dblTotalAmount
                                                                 dblTotalAmount = .Fields.Item("TOTAL_AMOUNT").Value
                                                          End If
                                                          .Update
                                                       End With
                                                    adoPISplit.Close
                                                   If rsCharge.Fields.Item("CHARGE_METHOD").Value = "RE_FIX" Then  'In preview mode we are not updating any FDD
                                                                                     'Updating FDDand next due date
'                                                                                        txtComparenextDueDate1 = DateAdd("d", 1, dtNextDue)
'                                                                                        dtNDD = NextDueDate1(CInt(dblFeqID), txtComparenextDueDate1, szPropertySelection1)
'                                                                                        txtComparenextDueDate1 = DateAdd("d", 1, dtNDD)
'                                                                                        dtFDD = NextDueDate1(CInt(dblFeqID), txtComparenextDueDate1, szPropertySelection1)
'                                                                                        strFromDate = Format(dtNextDue, "dd/MM/yyyy")
'                                                                                        strToDate = Format(DateAdd("d", -1, dtNDD), "dd/MM/yyyy") ' Format(DateAdd("d", -1, dtFDD), "dd/MMM/yyyy")
'                                                                                      szSQL = "Update tlbAgreement agr, ClientProAgr CPA, DemandTypes D, ChargeTypes C,Fund  F,FREQUENCIES FC,SECONDARYCODE SC " & _
'                                                                                     " Set NtDueDate=#" & dtNDD & "# ,FDD=#" & dtFDD & "#,lastchargeDate=#" & txtStatementDate1 & "# " & _
'                                                                                     "WHERE agr.CPA_ID = CPA.CPA_ID And cstr(F.FundID)=agr.fund AND FC.ID=agr.Frequency AND SC.CODE=agr.CHARGE_METHOD AND " & _
'                                                                                     "CPA.ClientID = '" & szSelectedClient & "' And D.ID = agr.DEMAND_TYPE And C.ID = agr.CHARGE_TYPE And " & _
'                                                                                     "CPA.PropertyID = '" & szPropertySelection1 & "' and AGREEMENT_ID=" & rsCharge.Fields.Item("AGREEMENT_ID").Value & ";"
'                                                                                     adoConnTransactions.Execute szSQL
                                                                                     
                                                                                     'adoConnTransactions.Execute szSQL
                                                   End If
                                                   dblGrandTotal = dblGrandTotal + dblTotalAmount
'

                                                 SQLforInsert = "'" & szMYID & "','" & Format(txtStatementDate1.text, "dd MMM yyyy") & "','" & szPropertySelection1 & "','Fixed Basis'," & dblTotalAmount - vatFixedBasis & "," & vatFixedBasis & "," & dblTotalAmount & ")"
                                                 adoConn.Execute "Insert into ManagementFeePreview(PI_ActualID,ChargeDate,PropertyID,ChargingMethod," & _
                                                "MgtFeeAmt,VAT,MgtFeeAmtTotal) values (" & _
                                                SQLforInsert
                                                vatFixedBasis = 0
                                                                       
                                                                       
                                End If 'end of rsCharge.Fields.Item("CHARGE_METHOD").Value = "RE_FIX"
                                 If rsCharge.Fields.Item("CHARGE_METHOD").Value = "RE_ABL" Then 'RE_ABL what does it mean?????
'                                                If IIf(IsNull(rsfixedMethod("amt").Value), 0, rsfixedMethod("amt").Value) = 0 Then
'                                                        rsfixedMethod.Close
'                                                        GoTo EndOfChargeType
'                                                End If
                                                If rsCharge("agreementEndD").Value < rsCharge("FDD").Value Then
                                                        MsgBox "agreement End Date is greatar than following due date for the property:" & szPropertySelection1
                                                        GoTo EndOfAgreement
                                                End If
                                                If DateDiff("d", txtStatementDate1.text, rsCharge.Fields.Item("agreementStartDate").Value) > 0 Then
                                                        rsCharge.Close
                                                        Set rsCharge = Nothing
                                                        MsgBox "Charge date cannot be before Agreement Start date for property:" & szPropertySelection1, vbInformation, "Warning"
                                                        GoTo EndOfChargeType
                                                        
                                                End If
                                                 szSQL1 = "Select sum(SWITCH(TYPE =1,RS.Amount,TYPE =2,-RS.Amount,TYPE =24,RS.Amount)) as Amt from tlbReceipt R,tlbReceiptsplit RS,Units U where " & _
                                                              "R.TransactionID=RS.RptHeader AND RDate<=#" & Format(txtStatementDate1.text, " dd MMM yyyy") & "# AND RDate>#" & Format(strLastChargeDate, " dd MMM yyyy") & "#  and Type in (1,2,24) " & _
                                                              "AND U.UnitNumber=R.UnitID AND U.PropertyID='" & szPropertySelection1 & "' and ISMGTFEES=false AND Rs.FundID=" & dblFundId & ""
                                                              'need to consider the selected property in where clause
                                                            '  rsfixedMethod.Close
                                                    rsfixedMethod.Open szSQL1, adoConn, adOpenStatic, adLockReadOnly
                                                    'Here type 3 is for recipt type . I have not written for the credit yet need to understand the principle
                                                    
                                                    If rsfixedMethod.EOF Then
                                                        rsfixedMethod.Close
                                                        Set rsfixedMethod = Nothing
                                                        GoTo EndOfChargeType
                                                    End If
                                                    percnetageOramount = IIf(IsNull(rsCharge("amount").Value), 0, rsCharge("amount").Value)
                                                    
                            '      ************************************Write tblPurInvSRec **************************************
                            
                                                                 szSQL = "Select sum(SWITCH(TYPE =3,RS.Amount,TYPE =4,-RS.Amount,TYPE =23,RS.Amount)) as Amt,rs.FundID from tlbReceipt R,tlbReceiptsplit RS,Units U where " & _
                                                                         "R.TransactionID=RS.RptHeader and Type in (3,4,23) AND RDate<=#" & Format(txtStatementDate1.text, " dd MMM yyyy") & "# AND RDate>#" & Format(strLastChargeDate, " dd MMM yyyy") & "#  AND  " & _
                                                                         "U.UnitNumber=R.UnitID and ISMGTFEES=false AND U.PropertyID='" & szPropertySelection1 & "' AND RS.FundID=" & dblFundId & " group by RS.FundID"
                                                                 'need to consider the selected property in where clause
                                                                 If rsfixedMethodDetails.State = 1 Then
                                                                     rsfixedMethodDetails.Close
                                                                 End If
                                                                 rsfixedMethodDetails.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
                                                                  If Not rsfixedMethodDetails.EOF Then 'we rare using while because itr
                                                                                 dblTotalAmount = rsfixedMethodDetails.Fields.Item("Amt").Value
                                                                                 If dblTotalAmount < 0 Then
                                                                                        GoTo EndOfChargeType
                                                                                 End If
                                                                                 dblFundId = rsfixedMethodDetails.Fields.Item("FundID").Value
                                                                                 If rsCharge.Fields.Item("CHARGE_METHOD").Value = "RE_ED" Then
                                                                                       dblTotalAmount = dblTotalAmount * percnetageOramount / 100
                                                                                 End If
                                                                                 
                                                                                 If dblCapAmount > 0 Then
                                                                                        If dblTotalAmount > dblCapAmount Then
                                                                                             dblTotalAmount = dblCapAmount
                                                                                        End If
                                                                                 End If
                                                                                 szSQL = "SELECT * FROM tblPurInvSRecPreview"
                                                                                 adoPISplit.Open szSQL, adoConn, adOpenDynamic, adLockOptimistic
                                                                                 'Add New Records. At least there is only one split line
                                                                                    With adoPISplit
                                                                                        .AddNew
                                                                                        .Fields.Item("MY_ID").Value = UniqueID()
                                                                                        .Fields.Item("ParentID").Value = szMYID
                                                                                        .Fields.Item("TRAN_ID").Value = j
                                                                                       ' .Fields.Item("TRANS").Value = szPropertySelection1  ' If you select One property then you can write a value here
                                                                                         'If chkAssignProperty.Value = 0 Then
                                                                                              .Fields.Item("TRANS").Value = szPropertySelection1  ' If you select One property then you can write a value here
                                                                                         'Else
                                                                                              .Fields.Item("TRANS").Value = ""
                                                                                         'End If
                                                                                        .Fields.Item("UNIT_ID").Value = ""
                                                                                        .Fields.Item("NOMINAL_CODE").Value = FinalControlACForPayable
                                                                                        .Fields.Item("DEPT_ID").Value = dblFundId
                                                                                       ' .Fields.Item("JOB_ID").Value = flxPI.TextMatrix(iRow, 9)            'Job No
                                                                                        .Fields.Item("RecoverablePt").Value = 0
                                                                                        '' (Current Charge date)" '"MFee" + szPropertySelection1 + Format(lngMgtFeeSL, "0000") '"Management Fee"
                                                                                        
                                                                                        .Fields.Item("description").Value = "Management Fees for " & strFundName & " " & DateAdd("d", 1, CDate(strFromDate)) & " - " & strToDate & ""
                                                                                        .Fields.Item("NET_AMOUNT").Value = dblTotalAmount
                                                                                        .Fields.Item("TAX_CODE").Value = VAT_CODE
                                                                                        .Fields.Item("VAT").Value = Round(dblTotalAmount * (VAT_RATE / 100), 2) 'VAT_RATE
                                                                                        .Fields.Item("TOTAL_AMOUNT").Value = .Fields.Item("VAT").Value + dblTotalAmount
                                                                                         dblTotalAmount = .Fields.Item("TOTAL_AMOUNT").Value
                                                                                        .Update
                                                                                    End With
                                                                                 adoPISplit.Close
                                                                                
                                                                                 dblGrandTotal = dblGrandTotal + dblTotalAmount
                                                            End If
                                                            rsfixedMethod.Close
                                End If 'end of  If rsCharge.Fields.Item("CHARGE_METHOD").Value = "RE_ABL" Then
                                
                               '****************************************************     receipt basis   ******************
                                Dim descriptionAndDate As String
                                If rsCharge.Fields.Item("CHARGE_METHOD").Value = "RE_ED" Then ' receipt basis
'                                                If IIf(IsNull(rsfixedMethod("amt").Value), 0, rsfixedMethod("amt").Value) = 0 Then
'                                                        rsfixedMethod.Close
'                                                        GoTo EndOfChargeType
'                                                End If
                                                If rsCharge("agreementEndD").Value < rsCharge("FDD").Value Then
                                                        MsgBox "agreement End Date is greatar than following due date for the property:" & szPropertySelection1
                                                        GoTo EndOfAgreement
                                                End If
                                                If DateDiff("d", txtStatementDate1.text, rsCharge.Fields.Item("agreementStartDate").Value) > 0 Then
                                                        rsCharge.Close
                                                        Set rsCharge = Nothing
                                                        MsgBox "Charge date cannot be before Agreement Start date for property:" & szPropertySelection1, vbInformation, "Warning"
                                                        GoTo EndOfChargeType
                                                        
                                                End If
                                
'                                            add report table and insert data into it
                                                    
                                            szSQL1 = "Select  sum(SWITCH(R.TYPE =3,RS.Amount,R.TYPE =4,RS.Amount,R.TYPE =23,-RS.Amount))  as Amt from tlbReceipt R,tlbReceiptsplit RS, " & _
                                            "rptTransactionsSPlit AL, Units U where AL.deleteflag=false AND " & _
                                            "AL.TransactionID= RS.RptTransactionsIDSplit AND  R.TransactionID=RS.RptHeader AND R.RDate<=#" & Format(txtStatementDate1.text, "dd MMM yyyy") & "# " & _
                                            "AND R.Type in (3,4,23)  AND U.UnitNumber=R.UnitID AND U.PropertyID='" & _
                                            szPropertySelection1 & "' and RS.ISMGTFEES=false AND Rs.FundID=" & dblFundId & ""
                                     
                                                    'here tlbReceipt R is reciopt table and RS,tlbReceipt R1 is invoice table
                                                    'need to consider the selected property in where clause
                                                    'Need to take only allocated transactions
                                                            '  rsfixedMethod.Close
                                                        'newly added this line
                                                        ' "Units U, (Select distinct fromtran from rptTransactions where deleteflag=false) as A  where A.Fromtran= R.TransactionID AND " & _
                                                        'by anol 20211024
                                                        If rsfixedMethod.State = 1 Then
                                                                rsfixedMethod.Close
                                                        End If
                                                    rsfixedMethod.Open szSQL1, adoConn, adOpenStatic, adLockReadOnly
                                                    'Here type 3 is for reciept type . I have not written for the credit yet need to understand the principle
                                                    
                                                    If rsfixedMethod.EOF Then
                                                        rsfixedMethod.Close
                                                        Set rsfixedMethod = Nothing
                                                        GoTo EndOfChargeType
                                                    End If
                                                    percnetageOramount = IIf(IsNull(rsCharge("amount").Value), 0, rsCharge("amount").Value)
                                                    percentageOramount1 = percnetageOramount
                                                    
                            '      ************************************Write tblPurInvSRec **************************************

'                                     szSQL = "Select  sum(SWITCH(R.TYPE =3,RS.Amount,R.TYPE =4,RS.Amount,R.TYPE =23,-RS.Amount))  as Amt from tlbReceipt R,tlbReceiptsplit RS, " & _
'                                     "rptTransactionsSPlit AL, Units U where AL.deleteflag=false AND " & _
'                                     "AL.TransactionID= RS.RptTransactionsIDSplit AND  R.TransactionID=RS.RptHeader AND R.RDate<=#" & Format(txtStatementDate1.text, "dd MMM yyyy") & "# " & _
'                                     "AND  R.Type in (3,4,23)  AND U.UnitNumber=R.UnitID AND U.PropertyID='" & _
'                                     szPropertySelection1 & "' and RS.ISMGTFEES=false AND Rs.FundID=" & dblFundId & ""
                                     
                                     'Modification done by anol 07-08-2023
                                     
                                     szSQL = "Select  (SWITCH(R.TYPE =3,AL.NetAmount,R.TYPE =4,AL.NetAmount,R.TYPE =23,-AL.NetAmount))  as Amt," & _
                                     "(SWITCH(R.TYPE =3,R.Amount,R.TYPE =4,R.Amount,R.TYPE =23,-R.Amount)) as Orig from tlbReceipt R,tlbReceiptsplit RS, " & _
                                     "rptTransactionsSPlit AL, Units U where AL.deleteflag=false AND " & _
                                     "AL.TransactionID= RS.RptTransactionsIDSplit AND  R.TransactionID=RS.RptHeader AND R.RDate<=#" & Format(txtStatementDate1.text, "dd MMM yyyy") & "# " & _
                                     "AND  R.Type in (3,4,23)  AND U.UnitNumber=R.UnitID AND U.PropertyID='" & _
                                     szPropertySelection1 & "' and RS.ISMGTFEES=false AND Rs.FundID=" & dblFundId & ""
                                     
                                                        'need to consider the selected property in where clause
                                                        'modified on 20211103
                                                        
'                                                        SQLforInsert = "Select " & szMYID & " as MYID ,'" & Format(txtStatementDate1.text, "dd MMM yyyy") & _
'                                                "' as ChargeDate,R.SlNumber, R.Type,R.SageAccountNumber,R.Ref,U.PropertyID,RS.FundID,RptAmtType,ExtRef,Rdate,R.TransactionID,RS.SplitID, " & _
'                                                "(SWITCH(R.TYPE =3,RS.Amount,R.TYPE =4,RS.Amount,R.TYPE =23,-RS.Amount)) as Amt from tlbReceipt R," & _
'                                                "tlbReceiptsplit RS, rptTransactionsSplit AL, Units U where AL.deleteflag=false AND " & _
'                                                "AL.TransactionID= RS.RptTransactionsIDSplit AND R.TransactionID=RS.RptHeader AND R.RDate<=#" & Format(txtStatementDate1.text, "dd MMM yyyy") & "# " & _
'                                                "and R.Type in (3,4,23)  AND U.UnitNumber=R.UnitID  AND U.PropertyID='" & _
'                                                 szPropertySelection1 & "' and RS.ISMGTFEES=false AND Rs.FundID=" & dblFundId & " "
                                                   'modified on 07-08-2023
                                                 SQLforInsert = "Select " & szMYID & " as MYID ,'" & Format(txtStatementDate1.text, "dd MMM yyyy") & _
                                                "' as ChargeDate,R.SlNumber, R.Type,R.SageAccountNumber,R.Ref,U.PropertyID,RS.FundID,RptAmtType,ExtRef,Rdate,R.TransactionID,RS.SplitID, " & _
                                                "(SWITCH(R.TYPE =3,AL.NETAmount,R.TYPE =4,AL.NETAmount,R.TYPE =23,-AL.NETAmount)) as Amt from tlbReceipt R," & _
                                                "tlbReceiptsplit RS, rptTransactionsSplit AL, Units U where AL.deleteflag=false AND " & _
                                                "AL.TransactionID= RS.RptTransactionsIDSplit AND R.TransactionID=RS.RptHeader AND R.RDate<=#" & Format(txtStatementDate1.text, "dd MMM yyyy") & "# " & _
                                                "and R.Type in (3,4,23)  AND U.UnitNumber=R.UnitID  AND U.PropertyID='" & _
                                                 szPropertySelection1 & "' and RS.ISMGTFEES=false AND Rs.FundID=" & dblFundId & " "
                                                 
                                            adoConn.Execute "Insert into ManagementFeePreview(PI_ActualID,ChargeDate,SRSlNumber,ReceiptType,SageAccountNumber,ReceiptTypeDescription," & _
                                                    "PropertyID,FundID,RptAmtType,ExtRef,ReceiptDate,ReceiptTransactionID,ReceiptSplitID," & _
                                                "ReceiptAmount)" & _
                                                SQLforInsert
                                                
                                                
'                                                        szSQLFrom = "Select  min(DS.DateFrom) as  DateFromMin ,max(DS.DateTO) as DateTOMAX   from tlbReceipt R,tlbReceiptsplit RS, " & _
'                                                        "rptTransactionsSPlit AL, DemandSplitRecords DS, Units U where  AL.deleteflag=false AND " & _
'                                                        "AL.TransactionID= RS.RptTransactionsIDSplit AND  R.TransactionID=RS.RptHeader AND R.RDate<=#" & Format(txtStatementDate1.text, "dd MMM yyyy") & "# " & _
'                                                        "AND R.Type in (3,4,23)  AND U.UnitNumber=R.UnitID AND U.PropertyID='" & _
'                                                        szPropertySelection1 & "' AND RS.ISMGTFEES=false AND Rs.FundID=" & dblFundId & " "
  'SQL Modification done on 07-08-2023.. max and min date are not specific for one property it is in between all selected property
                                                        szSQLFrom = "Select  min(DS.DateFrom) as  DateFromMin ,max(DS.DateTO) as DateTOMAX   from  tlbReceipt R, tlbReceipt R1,tlbReceiptsplit RS,  " & _
                                                        "rptTransactionsSPlit AL, DemandSplitRecords DS, Units U,DemandRecords DR where  DS.DemandID=DR.DemandID and DS.DemandID= R1.DemandRef AND R1.transactionID=AL.ToTran AND AL.deleteflag=false AND " & _
                                                        "AL.TransactionID= RS.RptTransactionsIDSplit AND  R.TransactionID=RS.RptHeader AND R.RDate<=#" & Format(txtStatementDate1.text, "dd MMM yyyy") & "# " & _
                                                        "AND R.Type in (3,4,23)  AND U.UnitNumber=DR.UnitNumber AND U.PropertyID='" & _
                                                        szPropertySelection1 & "' AND RS.ISMGTFEES=false AND Rs.FundID=" & dblFundId & " "
                                                        
                                                        rsFromandToDate.Open szSQLFrom, adoConn, adOpenStatic, adLockReadOnly
                                                        If Not rsFromandToDate.EOF Then
                                                                strFromDate = Format(rsFromandToDate("DateFromMin").Value, "dd/MM/yyyy")
                                                                strToDate = Format(rsFromandToDate("DateTOMAX").Value, "dd/MM/yyyy")
                                                        Else
                                                                strFromDate = Null
                                                                strToDate = Null
                                                        End If
                                                        rsFromandToDate.Close
                                                        Dim dblOriginalAmount As Currency
                                                         'Need to take only allocated transactions
                                                        If rsfixedMethodDetails.State = 1 Then
                                                            rsfixedMethodDetails.Close
                                                        End If
                                                                 rsfixedMethodDetails.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
                                                                  While Not rsfixedMethodDetails.EOF  'we rare using while because itr
                                                                     dblOriginalAmount = IIf(IsNull(rsfixedMethodDetails.Fields.Item("Orig").Value), 0, rsfixedMethodDetails.Fields.Item("Orig").Value)
                                                                     dblTotalAmount = IIf(IsNull(rsfixedMethodDetails.Fields.Item("Amt").Value), 0, rsfixedMethodDetails.Fields.Item("Amt").Value)
                                                                     If dblTotalAmount <= 0 Then
                                                                           rsfixedMethod.Clone
                                                                           GoTo EndOfChargeType
                                                                     End If
                                                                    'dblFundId = rsfixedMethodDetails.Fields.Item("FundID").Value
                                                                    If rsCharge.Fields.Item("CHARGE_METHOD").Value = "RE_ED" Then
                                                                          dblTotalAmount = dblTotalAmount * (percnetageOramount / 100)
                                                                           dblTotalAmount = Round(dblTotalAmount, 2)
                                                                    End If
                                                                    
                                                                    If dblCapAmount > 0 Then
                                                                           If dblTotalAmount > dblCapAmount Then
                                                                                dblTotalAmount = dblCapAmount
                                                                           End If
                                                                    End If
                                                                    szSQL = "SELECT * FROM tblPurInvSRecPreview"
                                                                    adoPISplit.Open szSQL, adoConn, adOpenDynamic, adLockOptimistic
                                                                    'Add New Records. At least there is only one split line
                                                                       With adoPISplit
                                                                           .AddNew
                                                                           .Fields.Item("MY_ID").Value = UniqueID()
                                                                           .Fields.Item("ParentID").Value = szMYID
                                                                           .Fields.Item("TRAN_ID").Value = j
                                                                          ' .Fields.Item("TRANS").Value = szPropertySelection1  ' If you select One property then you can write a value here
                                                                            'If chkAssignProperty.Value = 0 Then
                                                                                 .Fields.Item("TRANS").Value = szPropertySelection1  ' If you select One property then you can write a value here
                                                                            'Else
                                                                                 .Fields.Item("TRANS").Value = ""
                                                                            'End If
                                                                           .Fields.Item("UNIT_ID").Value = ""
                                                                           .Fields.Item("NOMINAL_CODE").Value = FinalControlACForPayable
                                                                           .Fields.Item("DEPT_ID").Value = dblFundId
                                                                          ' .Fields.Item("JOB_ID").Value = flxPI.TextMatrix(iRow, 9)            'Job No
                                                                           .Fields.Item("RecoverablePt").Value = 0
                                                                           '' (Current Charge date)" '"MFee" + szPropertySelection1 + Format(lngMgtFeeSL, "0000") '"Management Fee"
                                                                           .Fields.Item("description").Value = "Management Fees for " & strFundName & " (" & strFromDate & " - " & strToDate & ")"
                                                                           descriptionAndDate = "Management Fees for " & strFundName & " (" & strFromDate & " - " & strToDate & ")"
                                                                           'Original Part
'                                                                                        .Fields.Item("NET_AMOUNT").Value = dblTotalAmount
'                                                                                        .Fields.Item("TAX_CODE").Value = VAT_CODE
'                                                                                        .Fields.Item("VAT").Value = Format(dblTotalAmount * (VAT_RATE / 100), "0.00") 'VAT_RATE
'                                                                                        .Fields.Item("TOTAL_AMOUNT").Value = .Fields.Item("VAT").Value + dblTotalAmount
'                                                                                         dblTotalAmount = .Fields.Item("TOTAL_AMOUNT").Value
                                                                           If bolVatOptionEnabled = True And bolOptedTotax = True Then
                                                                                   'dblTotalAmount = dblTotalAmount * Round((100 / (100 + VAT_RATE)), 2)
                                                                                   .Fields.Item("NET_AMOUNT").Value = dblTotalAmount
                                                                                   .Fields.Item("TAX_CODE").Value = VAT_CODE
                                                                                   .Fields.Item("VAT").Value = Round(dblTotalAmount * (VAT_RATE / 100), 2) 'VAT_RATE
                                                                                   .Fields.Item("TOTAL_AMOUNT").Value = .Fields.Item("VAT").Value + dblTotalAmount
                                                                                    dblTotalAmount = .Fields.Item("TOTAL_AMOUNT").Value
                                                                                    adoConn.Execute "Update ManagementFeePreview set AgrPercentage=" & percnetageOramount & ",VATPercentage=" & VAT_RATE & " where MgtFeeAmtTotal is null"
                                                                              ElseIf bolVatOptionEnabled = True And bolOptedTotax = False Then 'bolVatOptionEnabled=global data
                                                                                   'Modified by anol 2021-10-15
                                                                                    dblTotalAmount = dblTotalAmount * Round((100 / (100 + VAT_RATE)), 2)
                                                                                    .Fields.Item("NET_AMOUNT").Value = dblTotalAmount
                                                                                   .Fields.Item("TAX_CODE").Value = Null 'VAT_CODE
                                                                                   .Fields.Item("VAT").Value = 0 'Format(dblTotalAmount * (VAT_RATE / 100), "0.00") 'VAT_RATE
                                                                                   .Fields.Item("TOTAL_AMOUNT").Value = dblTotalAmount
                                                                                    dblTotalAmount = .Fields.Item("TOTAL_AMOUNT").Value
                                                                                    adoConn.Execute "Update ManagementFeePreview set AgrPercentage=" & percentageOramount1 & ",VATPercentage=0 where MgtFeeAmtTotal is null"
                                                                             ElseIf bolVatOptionEnabled = False And bolOptedTotax = True Then 'bolVatOptionEnabled=global data
                                                                                   rsGlobalData.Open "Select V.VAT_ID,V.VAT_CODE,V.VAT_RATE from  Supplier S,tlbVatCode V where  (S.VATCode)=cstr(V.VAT_ID)  AND  SupplierID='" & _
                                                                                                      strManagingAgentID & "' ", adoConn, adOpenStatic, adLockReadOnly
                                                                                   If Not rsGlobalData.EOF Then
                                                                                            VAT_ID = IIf(IsNull(rsGlobalData.Fields("VAT_ID").Value), "-1", rsGlobalData.Fields("VAT_ID").Value)
                                                                                            VAT_RATE = IIf(IsNull(rsGlobalData.Fields("VAT_RATE").Value), "0.00", rsGlobalData.Fields("VAT_RATE").Value)
                                                                                            VAT_CODE = IIf(IsNull(rsGlobalData.Fields("VAT_CODE").Value), "", rsGlobalData.Fields("VAT_CODE").Value)
                                                                                   Else
                                                                                            VAT_ID = -1
                                                                                            VAT_RATE = 0
                                                                                            VAT_CODE = ""
                                                                                   End If
                                                                                   rsGlobalData.Close
                                                                                   'done modification on 15-10-2021
                                                                                   .Fields.Item("NET_AMOUNT").Value = dblTotalAmount
                                                                                   .Fields.Item("TAX_CODE").Value = Null ' VAT_CODE
                                                                                   .Fields.Item("VAT").Value = 0 'Format(dblTotalAmount * (VAT_RATE / 100), "0.00") ' Format(dblTotalAmount * (VAT_RATE / 100), "0.00") 'VAT_RATE' + Format(dblTotalAmount * (VAT_RATE / 100), "0.00")
                                                                                   .Fields.Item("NET_AMOUNT").Value = dblTotalAmount + Round(dblTotalAmount * (VAT_RATE / 100), 2)
                                                                                   .Fields.Item("TOTAL_AMOUNT").Value = dblTotalAmount + Round(dblTotalAmount * (VAT_RATE / 100), 2)
                                                                                    dblTotalAmount = .Fields.Item("TOTAL_AMOUNT").Value
                                                                                    adoConn.Execute "Update ManagementFeePreview set AgrPercentage=" & percentageOramount1 & ",VATPercentage=" & VAT_RATE & " where MgtFeeAmtTotal is null"
                                                                             ElseIf bolVatOptionEnabled = False And bolOptedTotax = False Then
                                                                                   .Fields.Item("NET_AMOUNT").Value = dblTotalAmount
                                                                                   .Fields.Item("TAX_CODE").Value = Null 'VAT_CODE
                                                                                   .Fields.Item("VAT").Value = 0 'Format(dblTotalAmount * (VAT_RATE / 100), "0.00") 'VAT_RATE
                                                                                   .Fields.Item("TOTAL_AMOUNT").Value = dblTotalAmount
                                                                                    dblTotalAmount = .Fields.Item("TOTAL_AMOUNT").Value
                                                                                    adoConn.Execute "Update ManagementFeePreview set AgrPercentage=" & percentageOramount1 & ",VATPercentage=0 where MgtFeeAmtTotal is null"
                                                                              End If
'                                                                              adoConn.Execute "Update ManagementFeePreview set MgtFeeAmt=round(AgrPercentage*ReceiptAmount/100,3),ChargingMethod='Receipt Basis' where ChargingMethod is null AND MgtFeeAmtTotal is null"
'                                                                              adoConn.Execute "Update ManagementFeePreview set MgtFeeAmtTotal=round(MgtFeeAmt*(1+VATPercentage/100),3),VAT=round(MgtFeeAmt*(VATPercentage/100),3) where ChargingMethod='Receipt Basis' AND MgtFeeAmtTotal is null"
                                                            'Modified by anol 07-08-2023
                                                                                adoConn.Execute "Update ManagementFeePreview set MgtFeeAmt=round(AgrPercentage*ReceiptAmount/100,3),ChargingMethod='Receipt Basis' where ChargingMethod is null AND MgtFeeAmtTotal is null"
                                                                                adoConn.Execute "Update ManagementFeePreview set MgtFeeAmtTotal=round(MgtFeeAmt*(1+VATPercentage/100),3),VAT=round(MgtFeeAmt*(VATPercentage/100),3),ReceiptAmount=" & dblOriginalAmount & " where ChargingMethod='Receipt Basis' AND MgtFeeAmtTotal is null"
                                                                              .Update
                                                                       End With
                                                                    adoPISplit.Close
                                                                    dblGrandTotal = dblGrandTotal + dblTotalAmount
                                                                    rsfixedMethodDetails.MoveNext
                                                           Wend
                                                            rsfixedMethod.Close
                                End If 'end of  If rsCharge.Fields.Item("CHARGE_METHOD").Value = "RE_ED" Then
                               
                               ' dblTotalAmount = dblGrandTotal
                                szSQL = "SELECT * FROM tblPurInvPreview"
                                    'dblTotalAmount = Format(dblGrandTotal, "00000.00")
                                    dblTotalAmount = Round(dblTotalAmount, 2)
                                    If dblTotalAmount = 0 Then
'                                           If adoConn.State = 1 Then
'                                                adoConn.Close
'                                           End If
                                           GoTo EndOfChargeType
                                    End If
                                    
                                    With adoPIHeader
                                            .Open szSQL, adoConn, adOpenDynamic, adLockPessimistic
                                            .AddNew
                                            .Fields.Item("MY_ID").Value = szMYID
                                            .Fields.Item("SlNumber").Value = lSlNumber
                                            .Fields.Item("SUPP_AC").Value = Trim(szManagingAgent(iManagingAgentCount))
                                            .Fields.Item("TRAN_DATE").Value = Format(txtStatementDate1.text, "DD MMMM YYYY")
                                            .Fields.Item("TransactionType").Value = 6
                                            .Fields.Item("INV_NO").Value = szPropertySelection1 + "-" + "MFee" + "-" + CStr(lngMgtFeeSL)
                                             lngMgtFeeSL = lngMgtFeeSL + 1
                                            .Fields.Item("TOTAL_AMOUNT").Value = dblTotalAmount
                                            .Fields.Item("History").Value = False
                                            .Fields.Item("TrfPayment").Value = False
                                            .Fields.Item("PropertyID").Value = ""
                                            .Fields.Item("CL_ID").Value = szSelectedClient
                                            .Fields.Item("NLPost").Value = False
                                            .Fields.Item("DueDate").Value = Format(dtNDDInitial, "DD MMMM YYYY")
                                            .Fields.Item("PostingDate").Value = Format(txtStatementDate1.text, "DD MMMM YYYY")
                                            .Fields.Item("ReportFromDatePreview").Value = IIf(strFromDate = "", Null, strFromDate)
                                            .Fields.Item("ReportToDatePreview").Value = IIf(strToDate = "", Null, strToDate)
                                            .Fields.Item("DescriptionANDDates").Value = descriptionAndDate
                                            .Update
                                            iCountPI = iCountPI + 1
                                            lSlNumber = lSlNumber + 1
                                   End With
                                   adoPIHeader.Close
                                    
                                
EndOfChargeType:
            rsCharge.MoveNext
            j = j + 1
            
      Wend


                        rsCharge.Close
                        Set rsCharge = Nothing
                        'adoConn.CommitTrans
                        Dim rsManagementFeePreview As New ADODB.Recordset
                        Dim previousSL As String
                        Dim SL As Integer
                        SL = 1
                        'Exit Sub
                        rsManagementFeePreview.Open "select * from ManagementFeePreview order by PI_ActualID", adoConn, adOpenDynamic, adLockOptimistic
                        While Not rsManagementFeePreview.EOF
                            iManagementFeePreview = ManagementFeePreview + 1
                            If previousSL <> rsManagementFeePreview("PI_ActualID").Value Then
                                    previousSL = rsManagementFeePreview("PI_ActualID").Value
                                    SL = 1
                            End If
                            rsManagementFeePreview!ReceiptSLNumber = SL
                            rsManagementFeePreview.Update
                            'adoConn.Execute "Update ManagementFeePreview set ReceiptSLNumber=" & SL & " where PI_ActualID='" & rsManagementFeePreview("PI_ActualID").Value & "'"
                            SL = SL + 1
                            rsManagementFeePreview.MoveNext
                        Wend
                        adoConn.Close
                        Set adoConn = Nothing
EndOfOneManagingAgentforOneAgreement:
           Next iManagingAgentCount
        End If 'end if for 'X' in grid selection
EndOfAgreement:
            Next iPropertyCount
            End If
            Next iClientCount
           Dim StrPropertyCol As String
           If dicWarningProp.Count > 0 Then
                Dim oProp
                For Each oProp In dicWarningProp.Items
                    StrPropertyCol = StrPropertyCol & oProp & ", "
                Next
           End If
           If Len(StrPropertyCol) > 0 Then
                StrPropertyCol = Left(StrPropertyCol, Len(StrPropertyCol) - 2)
                MsgBox "Please enter a valid Client Global Settings setup for the property: " & vbCrLf & StrPropertyCol, vbInformation, "Client Global Data setup"
                warning1 = "Please enter a valid Client Global Settings setup for the property: " & vbCrLf & StrPropertyCol
           End If
           '
           
           StrPropertyCol = ""
           If dicWarningAgreement.Count > 0 Then
                Dim oWar
                For Each oWar In dicWarningAgreement.Items
                    StrPropertyCol = StrPropertyCol & oWar & ", "
                Next
           End If
           If Len(StrPropertyCol) > 0 Then
                StrPropertyCol = Left(StrPropertyCol, Len(StrPropertyCol) - 2)
                MsgBox "Please enter a valid setup for the property: " & vbCrLf & StrPropertyCol, vbInformation, "Client agreement Fees and charges setup"
                warning2 = "Please enter a valid setup for the property: " & vbCrLf & StrPropertyCol
           End If
           StrPropertyCol = ""
           If dicWarningFinPeriod.Count > 0 Then
                Dim oWarning
                For Each oWarning In dicWarningFinPeriod.Items
                    StrPropertyCol = StrPropertyCol & oWarning & ", "
                Next
           End If
           If Len(StrPropertyCol) > 0 Then
                StrPropertyCol = Left(StrPropertyCol, Len(StrPropertyCol) - 2)
                 'MsgBox "The posting date does not fall in any existing financial period :" & szSelectedClient, vbInformation, "Warning"
                'MsgBox "The posting date does not fall in any existing financial period for following Client(s): " & vbCrLf & StrPropertyCol, vbInformation, "Warning"
                warning3 = "The posting date does not fall in any existing financial period for following Client(s): " & vbCrLf & StrPropertyCol
           End If
          
           
            'MsgBox iCountPI & " Management Fee Invoice(s) Preview generated.", vbInformation, "Generate Fee Preview"
            
            
            
          If iManagementFeePreview > 0 Then
                        
                  GeneratePIPreview = True
          End If
   
End Function
Private Function checksetupDONE(szProperty As String) As Boolean
        Dim adoConn As New ADODB.Connection
        Dim rsCharge As New ADODB.Recordset
        Dim szSQL  As String
        adoConn.Open getConnectionString
        szSQL = "SELECT agr.CHARGE_METHOD,agr.LastChargeDate, agr.TotalAmount,agr.Amount,agr.Fund, agr.NtDueDate,agr.FDD,(Select FC.Frequency from Frequencies FC where  FC.ID=agr.Frequency) as Frequency  " & _
                                      "FROM tlbAgreement agr, ClientProAgr CPA,  ChargeTypes C,Fund  F,SECONDARYCODE SC " & _
                                      "WHERE agr.CPA_ID = CPA.CPA_ID And F.FundID=agr.fund AND  SC.CODE=agr.CHARGE_METHOD AND " & _
                                      "C.ID = agr.CHARGE_TYPE And " & _
                                      "CPA.PropertyID = '" & szProperty & "' "
        rsCharge.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
        If rsCharge.EOF Then
            checksetupDONE = True
        End If
        rsCharge.Close
        
        
        adoConn.Close
End Function
Private Function checkGlobalDataEntered(szProperty As String) As Boolean
        Dim adoConn As New ADODB.Connection
        Dim rsCharge As New ADODB.Recordset
        Dim szSQL  As String
        adoConn.Open getConnectionString
        szSQL = "SELECT *,YDueDate from globaldata where PropertyID = '" & szProperty & "' "
        rsCharge.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
        If Not rsCharge.EOF Then
            If IsNull(rsCharge("YDueDate").Value) = True Then
                checkGlobalDataEntered = True
            End If
        End If
        rsCharge.Close
        
        adoConn.Close
End Function
'Private Function hasChargeTypeinlist(szProperty1 As String) As Boolean
'    Dim iCount As Integer
'    For iCount = 1 To flxDemandTypes.Rows - 1
'         If flxDemandTypes.TextMatrix(iCount, 2) <> "" Then
'                'If flxDemandTypes.TextMatrix(iCount, 0) = "X" Then
'                     If flxDemandTypes.TextMatrix(iCount, 1) = szProperty1 Then
'                        hasChargeTypeinlist = True
'                     End If
'                'End If
'         End If
'     Next
'End Function
