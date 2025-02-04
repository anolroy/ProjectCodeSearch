VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmRentPayableModification 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Modification of Summary Statement"
   ClientHeight    =   11535
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   18945
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
   Icon            =   "frmRentPayableModification.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   11535
   ScaleWidth      =   18945
   Begin VB.TextBox txtComparenextDueDate1 
      Height          =   270
      Left            =   15480
      TabIndex        =   24
      Top             =   630
      Visible         =   0   'False
      Width           =   1545
   End
   Begin VB.Frame Frame1 
      Height          =   10815
      Index           =   6
      Left            =   40
      TabIndex        =   0
      Top             =   -45
      Width           =   14055
      Begin VB.CheckBox chkShowDue 
         Caption         =   "Incl. Mngt Fees Due"
         Height          =   210
         Left            =   10800
         TabIndex        =   33
         Top             =   9540
         UseMaskColor    =   -1  'True
         Value           =   1  'Checked
         Width           =   1995
      End
      Begin VB.CheckBox chkExcludeSupOS 
         Caption         =   "Incl. Supplier OS"
         Height          =   210
         Left            =   8910
         TabIndex        =   32
         Top             =   9540
         UseMaskColor    =   -1  'True
         Width           =   1590
      End
      Begin VB.CommandButton cmdAddRetention 
         Caption         =   "Add Retention"
         Height          =   375
         Left            =   270
         TabIndex        =   31
         Top             =   9765
         Width           =   1950
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
         Left            =   2565
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   30
         Text            =   "0.00"
         Top             =   9765
         Width           =   1125
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
         Left            =   3645
         MaxLength       =   10
         TabIndex        =   28
         Text            =   "0.00"
         Top             =   9765
         Width           =   1260
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
         Left            =   4995
         MaxLength       =   10
         TabIndex        =   27
         Text            =   "0.00"
         Top             =   9765
         Width           =   1350
      End
      Begin VB.CommandButton cmdFinalizeStatement 
         Caption         =   "&Finalise Statement"
         Height          =   375
         Left            =   11655
         TabIndex        =   23
         Top             =   9855
         Width           =   1665
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
         TabIndex        =   16
         Top             =   4950
         Width           =   6690
         Begin VB.CheckBox chkAllProperties 
            Caption         =   "All Properties"
            Height          =   255
            Left            =   135
            TabIndex        =   17
            Top             =   240
            Width           =   2025
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxProperties 
            Height          =   3810
            Left            =   135
            TabIndex        =   18
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
         Caption         =   "Funds:"
         Height          =   4455
         Index           =   10
         Left            =   6840
         TabIndex        =   13
         Top             =   4950
         Width           =   7095
         Begin VB.CheckBox chkInFunds 
            Caption         =   "All Funds"
            Height          =   255
            Left            =   180
            TabIndex        =   14
            Top             =   270
            Width           =   1095
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxInFunds 
            Height          =   3765
            Left            =   120
            TabIndex        =   15
            Top             =   570
            Width           =   6900
            _ExtentX        =   12171
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
         Left            =   6840
         TabIndex        =   11
         Top             =   765
         Width           =   7050
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxBankAccounts 
            Height          =   3810
            Left            =   90
            TabIndex        =   12
            Top             =   270
            Width           =   6855
            _ExtentX        =   12091
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
         Left            =   6165
         TabIndex        =   10
         Top             =   10800
         Visible         =   0   'False
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
         TabIndex        =   9
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
         TabIndex        =   8
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
         TabIndex        =   7
         Top             =   450
         Width           =   1575
      End
      Begin VB.Frame Frame1 
         Caption         =   "Clients:"
         Height          =   3780
         Index           =   12
         Left            =   90
         TabIndex        =   5
         Top             =   1170
         Width           =   6690
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxClients 
            Height          =   3495
            Left            =   135
            TabIndex        =   6
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
         Caption         =   "Fix AV Fund"
         Height          =   420
         Left            =   7065
         TabIndex        =   4
         Top             =   10800
         Visible         =   0   'False
         Width           =   2400
      End
      Begin VB.CommandButton cmdCalculateAvailableFund 
         Caption         =   "Calculate Available Fund"
         Height          =   375
         Left            =   8235
         TabIndex        =   3
         Top             =   10485
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Modify Statement"
         Height          =   375
         Left            =   9675
         TabIndex        =   2
         Top             =   9855
         Width           =   1845
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
         Height          =   375
         Index           =   0
         Left            =   12375
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   10305
         Width           =   1200
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Current Retention"
         Height          =   210
         Left            =   2295
         TabIndex        =   29
         Top             =   9450
         Width           =   1395
      End
      Begin VB.Label Label2 
         Caption         =   "Rent Payable"
         Height          =   240
         Left            =   5085
         TabIndex        =   26
         Top             =   9450
         Width           =   1320
      End
      Begin VB.Label Label1 
         Caption         =   "Available Fund"
         Height          =   240
         Left            =   3735
         TabIndex        =   25
         Top             =   9450
         Width           =   1320
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
Attribute VB_Name = "frmRentPayableModification"
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
Dim szLastStatementDate As String
Dim szStatementNo As Long
Dim bEditDone As Boolean
Dim szSelectedBankAC1 As String
Dim szSelectedBankACName1 As String
'Dim bolInclSupplierOS As Boolean
'Dim InclMngtFeesDue As Boolean

Private Sub MarkAllTransactionsWithSSFinalized(strSSID As String)
    Dim adoConn As New ADODB.Connection
    adoConn.Open getConnectionString
    Dim szSQL As String
    Dim whereProperty As String
    'We are considering all properties as 1
    whereProperty = "AND (P.UnitID in (" & ListOfProperties & ") OR isnull(P.UnitID) OR P.UnitID='' ) "
    szSQL = "Update tlbPayment P,tlbPaymentSplit S,Fund F,Supplier SP,PayTransactions AL,tlbPayment AP,tblPurINV V SET ClientStatementID=" & strSSID & "  where " & _
            "P.PDate <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "# AND P.BankCODE='" & szSelectedBankAccount & "' AND AL.FromTran=P.TransactionID  AND AL.ToTran=AP.TransactionID AND " & _
            "SP.SupplierID=P.SageaccountNumber and  P.Amount>P.OSAmount  AND  P.TransactionID=S.PayHeader  AND P.TYPE IN(7,8,9) AND S.FundID=F.FundID AND F.FundCode in (" & _
             ListOfFunds & ") AND  isnull(S.ClientStatementID) AND AP.PI=V.MY_ID AND V.isRentPayable=true  AND P.ClientID ='" & szSelectedClient & "' " & whereProperty & ""
    adoConn.Execute szSQL
    ' AND  S.PropertyID in (" & ListOfProperties & ") we cannot add this because while marking we need to make for all properties in a client as a whole set. We cannot pay rent payable twice
    'Paying rent payable twice is not supported in this design
    adoConn.Close
    Set adoConn = Nothing
End Sub
Private Sub MarkAllTransactionsWithSS(strSSID As String)
    Dim adoConn As New ADODB.Connection
    adoConn.Open getConnectionString
    Dim szSQL As String
    Dim whereProperty As String
    'We are considering all properties as 1
    'whereProperty = "AND (P.UnitID in (" & ListOfProperties & ") OR isnull(P.UnitID) OR P.UnitID='' ) "
    
    szSQL = "Update tlbReceipt R,tlbReceiptSplit S,Fund F SET ClientStatementID=" & strSSID & "  where R.TransactionID= S.RptHeader AND TYPE IN(3,4,23) " & _
            "AND S.FundID=F.FundID AND R.BankCODE='" & szSelectedBankAccount & "' and R.amount>R.OSamount AND F.FundCode<>'TENANTDEPOSIT'  AND isnull(S.ClientStatementID) " & _
            "AND R.RDate <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "# " & _
            "AND F.FundCode in (" & ListOfFunds & ") AND  S.PropertyID in (" & ListOfProperties & ")  AND  ClientID ='" & szSelectedClient & "'"
    adoConn.Execute szSQL
    
    'ONE SRR do not have property ID in it. and stackholder wants that to be in CS. SO I need to consider empty Property ID for SSR , include that for marking 2023-08-20
    szSQL = "Update tlbReceipt R,tlbReceiptSplit S,Fund F SET ClientStatementID='" & strSSID & "'  where R.TransactionID= S.RptHeader AND TYPE IN(3,4,23) " & _
            "AND S.FundID=F.FundID AND R.BankCODE='" & szSelectedBankAccount & "' AND F.FundCode<>'TENANTDEPOSIT' and S.amount>S.OSamount  AND  isnull(S.ClientStatementID) " & _
            "AND R.RDate <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "# " & _
            "AND F.FundCode in (" & ListOfFunds & ") AND  (isnull(S.PropertyID) OR  S.PropertyID='')  AND ClientID ='" & szSelectedClient & "'"
    adoConn.Execute szSQL
    
    'UnitID in the tlbpayment table means  propertyID
    whereProperty = "AND (P.UnitID in (" & ListOfProperties & ") OR isnull(P.UnitID) OR P.UnitID='' ) "
    szSQL = "Update tlbPayment P,tlbPaymentSplit S,Fund F,Supplier SP SET ClientStatementID=" & strSSID & "  where " & _
            "P.PDate >#" & Format(txtLastStatementDate1.text, "dd/mmm/yyyy") & "# AND P.PDate <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "# AND P.BankCODE='" & szSelectedBankAccount & "' AND " & _
            "SP.SupplierID=P.SageaccountNumber and  P.Amount>P.OSAmount  AND  P.TransactionID=S.PayHeader and P.amount>P.OSamount AND P.TYPE IN(24,8,9) AND S.FundID=F.FundID AND F.FundCode in (" & _
             ListOfFunds & ") AND  isnull(S.ClientStatementID) AND P.ClientID ='" & szSelectedClient & "' " & whereProperty & ""
    adoConn.Execute szSQL
    
    szSQL = "Update tlbBankPayment B, Fund F  SET RentSumStatement='" & strSSID & "'  where B.DEPT_ID=F.FundID " & _
            "AND B.TRAN_DATE >#" & Format(txtLastStatementDate1.text, "dd/mmm/yyyy") & "# AND B.TRAN_DATE <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "# " & _
            "AND  B.PropertyID in (" & ListOfProperties & ") and BANK_AC='" & szSelectedBankAccount & "' AND F.FundCode in (" & ListOfFunds & ") AND (B.RentSumStatement='' OR isnull(B.RentSumStatement)) and clientID='" & szSelectedClient & "' "
            
    adoConn.Execute szSQL
    
    adoConn.Close
    Set adoConn = Nothing
End Sub

Private Function getAvailablefundsModified(dblLasClosingBalance As Double, ByVal trtoinclude As Long) As Double 'this one I am using while finalize
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
    'Total reciept

    szSQL = "Select  SUM(switch(type=3,S.Amount,type=4,S.Amount,type=23,-S.Amount)) as DR from tlbReceipt R,tlbReceiptSplit S where R.TransactionID= S.RptHeader " & _
    "AND ClientID ='" & szSelectedClient & "' and S.ClientStatementID=" & trtoinclude & ""
    rsReceipt.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
    If Not rsReceipt.EOF Then
            dblAmt = IIf(IsNull(rsReceipt.Fields.Item("Dr").Value), 0, rsReceipt.Fields.Item("Dr").Value)
             'result is 175
    End If
    rsReceipt.Close
    Set rsReceipt = Nothing
    
    
    getAvailablefundsModified = dblLasClosingBalance + dblAmt
    'Vat calculation of B)take the  VAT amount from the allocation table and deduct it
            szSQL = "Select  SUM(AL.VatAmount) as DR from tlbReceipt R,rptTransactionsSplit AL,tlbReceiptSplit S,Fund F, Units U,GlobalData G where G.PropertyID=U.PropertyID " & _
            "AND G.isAgentToSubmit=true AND R.TransactionID= S.RptHeader AND AL.Deleteflag=False AND S.FundID=F.FundID AND R.BankCODE='" & szSelectedBankAccount & "' " & _
            "AND S.ClientStatementID=" & trtoinclude & " AND ClientID ='" & szSelectedClient & "' AND R.UnitID=U.UnitNumber " & _
            "and AL.Deleteflag=false and AL.FromTran=R.TransactionID AND R.RDate <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "#"
            rsReceipt.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
            If Not rsReceipt.EOF Then
                    dblAmt = IIf(IsNull(rsReceipt.Fields.Item("Dr").Value), 0, rsReceipt.Fields.Item("Dr").Value)
                     'result is 175
            End If
            rsReceipt.Close
            Set rsReceipt = Nothing
            getAvailablefundsModified = getAvailablefundsModified - dblAmt
'    End If
 
   'c   (-): Sum of Supplier amounts Paid/Refunded ( allocated /purchase payment
'    If boolConsolidatedStatement = 1 Then
            whereProperty = "(P.UnitID in (" & ListOfProperties & ") OR isnull(P.UnitID) OR P.UnitID='' ) AND "
'    Else
'            whereProperty = "P.UnitID in (" & ListOfProperties & ") AND "
'    End If


'szSQL = "Select SUM(SWITCH(P.TYPE=24,PS.PaymentAmount,P.TYPE=8,-PS.PaymentAmount,P.TYPE=9,-PS.PaymentAmount)) as AMT from tlbPayment P,tlbPaymentSplit S," & _
'            "PaytransactionsSplit PS,Supplier SP where  PS.TransactionID=S.PayTransactionIDSplit AND SP.SupplierID=P.SageAccountNumber AND P.TransactionID=S.PayHeader " & _
'            "AND PS.Deleteflag=False AND S.ClientStatementID=" & trtoinclude & " AND Sp.Type='Supplier' AND P.ClientID ='" & szSelectedClient & "' " & _
'            "AND P.PDate <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "#"
'Calculating supplier payment amount here
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
    getAvailablefundsModified = getAvailablefundsModified + dblAmt
'Take the vat amount from the allocation table
'    If bolisAgentToSubmit = True Then
        szSQL = "Select  SUM(AL.VatAmount)  as AMT from tlbPayment P,PayTransactionsSplit AL,tlbPaymentSplit S,Fund F,Supplier SP,Property PR,GLobalData G " & _
                "where  G.PropertyID=PR.PropertyID AND isAgentToSubmit=true  AND " & _
                "AL.TransactionID=S.PayTransactionIDSplit and SP.SupplierID=P.SageAccountNumber AND AL.Deleteflag=false and PR.propertyID=P.UNITID and " & _
                "Sp.Type='Supplier' AND P.TransactionID=S.PayHeader AND S.ClientStatementID=" & trtoinclude & "" & _
                "AND S.FundID=F.FundID AND AL.FromTran=P.transactionID AND P.ClientID ='" & szSelectedClient & "' " & _
                "AND P.PDate <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "#"
        rsPayment.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
        If Not rsPayment.EOF Then
                dblAmt = IIf(IsNull(rsPayment.Fields.Item("AMT").Value), 0, rsPayment.Fields.Item("AMT").Value)
                'result is -15
        End If
        rsPayment.Close
        Set rsPayment = Nothing
         getAvailablefundsModified = getAvailablefundsModified - dblAmt
'    End If
    'd)  Add (+): Sum of Bank payments and receipts
'    If boolConsolidatedStatement = 1 Then
            whereProperty = "(B.PropertyID in (" & ListOfProperties & ") OR isnull(B.PropertyID) OR B.PropertyID ='' ) AND "
'    Else
'            whereProperty = "B.PropertyID in (" & ListOfProperties & ") AND "
'    End If
    
     szSQL = "Select  SUM(SWITCH(TransactionType=11 ,-B.NET_AMOUNT,TransactionType=12 ,B.NET_AMOUNT)) as AMT from tlbBankPayment B, Fund F  where B.DEPT_ID=F.FundID " & _
            "AND B.RentSumStatement='" & trtoinclude & "' and clientID='" & szSelectedClient & "' " & _
            "AND  B.TRAN_DATE <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "#"
    rsBankPaymentAndRcpt.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
    If Not rsBankPaymentAndRcpt.EOF Then
        dblAmt = IIf(IsNull(rsBankPaymentAndRcpt.Fields.Item("AMT").Value), 0, rsBankPaymentAndRcpt.Fields.Item("AMT").Value)
           'result is 0
    End If
    rsBankPaymentAndRcpt.Close
    Set rsBankPaymentAndRcpt = Nothing
    getAvailablefundsModified = getAvailablefundsModified + dblAmt
    'f)  Less (-): Supplier OS Account balances for the client selected
    
        dblAmt = GetSupplierOSAmount
    'If negative then ignore this
    
        getAvailablefundsModified = getAvailablefundsModified - IIf(dblAmt < 0, 0, dblAmt)
    
    'it should be -40

    Dim rsNLposting As New ADODB.Recordset

'g)  Less (-): Client /Landlord OS balances for the client selected  and property selected amounts due to Client/Landlord not paid
         dblAmt = GetClientACBalance ' GetClientACBalanceModPreview(trToinclude)  ' -GetClientACBalance
          getAvailablefundsModified = getAvailablefundsModified + dblAmt
         'client payment
                whereProperty = "(P.UnitID in (" & ListOfProperties & ") OR isnull(P.UnitID) OR P.UnitID='' ) AND "
'                szSQL = "Select   SUM(SWITCH(P.TYPE=24,PS.PaymentAmount,P.TYPE=8,-PS.PaymentAmount,P.TYPE=9,-PS.PaymentAmount)) as AMT from tlbPayment P,Paytransactions PS,Fund F,Supplier SP where  " & _
'                "PS.FROMTran=P.TransactionID AND SP.SupplierID=P.SageAccountNumber  AND PS.Deleteflag=False AND P.RentSumStatement='" & trtoinclude & "' AND " & _
'                "PS.FundID=F.FundID and Sp.Type='Client' AND P.ClientID ='" & szSelectedClient & "' " & _
'                 "AND P.PDate <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "#"
                 
'                szSQL = "Select   SUM(SWITCH(P.TYPE=24,PS.Amount,P.TYPE=8,-PS.Amount,P.TYPE=9,-PS.Amount)) as AMT from tlbPayment P,tlbPaymentSplit PS,Fund F,Supplier SP where  " & _
'                "PS.PayHeader=P.TransactionID AND SP.SupplierID=P.SageAccountNumber  AND PS.Amount>PS.OSAmount AND PS.ClientStatementID=" & trtoinclude & " AND " & _
'                "PS.FundID=F.FundID and Sp.Type='Client' AND P.ClientID ='" & szSelectedClient & "' " & _
'                "AND P.PDate <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "#"
                szSQL = "Select   SUM(SWITCH(P.TYPE=24,PS.Amount,P.TYPE=8,-PS.Amount,P.TYPE=9,-PS.Amount)) as AMT from tlbPayment P,tlbPaymentSplit PS,Fund F,Supplier SP," & _
                "tblPurInv V,tlbPayment PI,PayTransactions PT where PT.fromtran=PI.transactionID and P.transactionID=PT.Totran and " & _
                "PS.Payheader=P.TransactionID AND SP.SupplierID=P.SageAccountNumber AND PS.Amount>PS.OSAmount AND PS.ClientStatementID=" & trtoinclude & " AND " & _
                "PS.FundID=F.FundID and Sp.Type='Client' AND P.ClientID ='" & szSelectedClient & "' AND V.MY_ID=PI.PI  " & _
                 "AND P.PDate <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "#"

                rsPayment.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
                If Not rsPayment.EOF Then
                        dblAmt = IIf(IsNull(rsPayment.Fields.Item("AMT").Value), 0, rsPayment.Fields.Item("AMT").Value)
                    'result is -837
                End If
                rsPayment.Close
                Set rsPayment = Nothing
    
    
         'COMING -35  dblAmt is negative then ignore
    'getAvailablefundsModified = getAvailablefundsModified + GetClientACBalance + GetLandLordACBalance
        getAvailablefundsModified = getAvailablefundsModified + dblAmt

          
        'landlord payment
          whereProperty = "(P.UnitID in (" & ListOfProperties & ") OR isnull(P.UnitID) OR P.UnitID='') AND "
'                szSQL = "Select   SUM(SWITCH(P.TYPE=24,PS.PaymentAmount,P.TYPE=8,-PS.PaymentAmount,P.TYPE=9,-PS.PaymentAmount)) as AMT from tlbPayment P,Paytransactions PS,Fund F,Supplier SP where  " & _
'                "PS.FROMTran=P.TransactionID AND SP.SupplierID=P.SageAccountNumber  AND PS.Deleteflag=False AND  P.RentSumStatement='" & trtoinclude & "' AND " & _
'                "PS.FundID=F.FundID and Sp.Type='LLORD' and  P.ClientID ='" & szSelectedClient & "' " & _
'                "AND P.PDate <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "#"
'                szSQL = "Select   SUM(SWITCH(P.TYPE=24,PS.Amount,P.TYPE=8,-PS.Amount,P.TYPE=9,-PS.Amount)) as AMT from tlbPayment P,tlbPaymentSplit PS,Fund F,Supplier SP where  " & _
'                "PS.Payheader=P.TransactionID AND SP.SupplierID=P.SageAccountNumber AND PS.Amount>PS.OSAmount AND PS.ClientStatementID=" & trtoinclude & " AND " & _
'                "PS.FundID=F.FundID and Sp.Type='LLORD' and  P.ClientID ='" & szSelectedClient & "' " & _
'                "AND P.PDate <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "#"
                szSQL = "Select   SUM(SWITCH(P.TYPE=24,PS.Amount,P.TYPE=8,-PS.Amount,P.TYPE=9,-PS.Amount)) as AMT from tlbPayment P,tlbPaymentSplit PS,Fund F,Supplier SP," & _
                "tblPurInv V,tlbPayment PI,PayTransactions PT where PT.fromtran=PI.transactionID and P.transactionID=PT.Totran and " & _
                "PS.Payheader=P.TransactionID AND SP.SupplierID=P.SageAccountNumber AND PS.Amount>PS.OSAmount AND PS.ClientStatementID=" & trtoinclude & " AND " & _
                "PS.FundID=F.FundID and Sp.Type='LLORD' AND P.ClientID ='" & szSelectedClient & "' AND V.MY_ID=PI.PI  " & _
                 "AND P.PDate <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "#"


                rsPayment.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
                If Not rsPayment.EOF Then
                        dblAmt = IIf(IsNull(rsPayment.Fields.Item("AMT").Value), 0, rsPayment.Fields.Item("AMT").Value)
                    'result is -837
                End If
                rsPayment.Close
                Set rsPayment = Nothing
    
    
         'COMING -35  dblAmt is negative then ignore
    'getAvailablefundsModified = getAvailablefundsModified + GetClientACBalance + GetLandLordACBalance
        getAvailablefundsModified = getAvailablefundsModified + dblAmt
        
        
        
        dblAmt = -GetLandLordACBalance '-GetLandLordACBalanceMODpreview(trToinclude) '-GetLandLordACBalance
        'GetLandLordACBalanceMODpreview
          'if dblAmt is negative then ignore
        getAvailablefundsModified = getAvailablefundsModified + dblAmt
     
     
    
    'getAvailablefundsPreview = getAvailablefundsPreview + GetClientACBalance + GetLandLordACBalance
       ' getAvailablefundsModified = getAvailablefundsModified + dblAmt
        
    'h)  Less (-): Managing Agent OS Balances for the client selected Management Fees due but not paid
    'GetAgentBalanceModPreview
    'dblAmt = -GetAgentBalance ' -GetAgentBalanceModPreview(trToinclude) 'GetAgentBalance '
    If chkShowDue.Value = 1 Then
        dblAmt = -GetAgentBalance ' -GetAgentBalanceModPreview(trToinclude) 'GetAgentBalance '
    End If
    getAvailablefundsModified = getAvailablefundsModified + dblAmt
    '
    dblAmt = GetAGENTPaymentsModified(trtoinclude)
    getAvailablefundsModified = getAvailablefundsModified + dblAmt
    getAvailablefundsModified = Round(getAvailablefundsModified, 2)
    
    Debug.Print getAvailablefundsModified
    
    
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
    getAvailablefundsModified = getAvailablefundsModified + dblAmt
    'j)  Less (-): Tenant Deposits received for the client selected
    'REMOVE THIS AS PER SPEC
    'getAvailablefundsModified = getAvailablefundsModified - GetRentDeposit
    rsNLposting.Open "Select sum(AMOUNT) as AMT from RetentionDetails where  isDeleted=false and BankCode='" & szSelectedBankAccount & "' AND " & _
                    "ClientID='" & szSelectedClient & "' AND statementID=" & _
                    szCurrentStatementID & " ", adoConn, adOpenStatic, adLockReadOnly
    If Not rsNLposting.EOF Then
        txtRetention.text = IIf(IsNull(rsNLposting.Fields.Item("AMT").Value), 0, rsNLposting.Fields.Item("AMT").Value)
    End If
    rsNLposting.Close


    getAvailablefundsModified = getAvailablefundsModified - Val(txtRetention.text)
    MsgBox "Available fund is: " & Round(getAvailablefundsModified, 2)
    txtAvailableFunds.text = Round(getAvailablefundsModified, 2)
    If bEditMode = False Then
        txtRentPayable.text = txtAvailableFunds.text
    End If
     
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
    "AND S.FundID=F.FundID AND R.BankCODE='" & szSelectedBankAccount & "' and F.FundCode<>'TENANTDEPOSIT'  AND (R.RentSumStatement='' or isnull(R.RentSumStatement)) AND ClientID ='" & szSelectedClient & "' " & _
    "AND R.RDate >#" & Format(txtLastStatementDate1.text, "dd/mmm/yyyy") & "# AND R.RDate <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "#"
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
            "SP.SupplierID=P.SageAccountNumber AND P.TransactionID=S.PayHeader AND P.TYPE IN(8,9,24) AND (P.RentSumStatement='' or isnull(P.RentSumStatement)) AND " & _
            "P.PDate >#" & Format(txtLastStatementDate1.text, "dd/mmm/yyyy") & "# AND P.PDate <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "# AND " & _
            "S.FundID=F.FundID and  P.BankCODE='" & szSelectedBankAccount & "' and  F.FundCode in (" & ListOfFunds & ") AND ClientID ='" & szSelectedClient & "' "
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
    rsBankPaymentAndRcpt.Close
    Set rsBankPaymentAndRcpt = Nothing
 End Function

Private Sub GenerateSummaryStatementModify(szStatmentID As String, ByVal szReportGenID As Long)

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
             szCurrentStatementID = szStatmentID
            !statementNo = GetLastStatementNoByClient + 1
            !ClientIDLandlordID = szSelectedClient
            !BankCode = szSelectedBankAccount
            !PreviousStatementDate = IIf(GetLastStatementDateByClient(szStatmentID) = "", "01/01/2000", GetLastStatementDateByClient(szStatmentID)) 'This is Fromdate
            !StatementDate = Format(txtStatementDate1.text, "dd/mmmm/yyyy") 'This is todate
            !StatementOpBal = dblLasClosingBalance
           
            !Clearretentions = False 'Will need to come again

            !AccrualsAcBalance = GetAccrualsControlBalance
            !SupplierAcBalance = GetSupplierOSAmount 'GetBalance("Supplier") 'GetBalanceSupplier'wrong
            !ManagingAgentAcBalance = GetAgentBalance 'GetBalance("Agent") 'GetBalanceAgent'wrong
            !ClientACBalance = GetClientACBalance
            !LandlordACBalance = GetLandLordACBalance
            !ListOffundID = ListOfFundsForDBSave ' szSelectedFund
'            !ListOfPayableTypeID = ListOfPayableTypesForDBSave ' ListOfPayableTypes
            !TenantDepositsReceived = GetRentDeposit(szStatmentID)
            '!Availablefunds = getAvailablefunds(dblLasClosingBalance)
            !Availablefunds = getAvailablefundsModified(dblLasClosingBalance, szStatmentID)
            !Retentions = Val(txtRetention.text) 'we need to further analyse detail/add/deduct retension
            !ListOfinputProperties = ListOfProperties
            !PaymentsonAccount = -GetPaymentsonAccount(szStatmentID)
            'New fields added 2021-01-24
            !TenantReceipts = GetTenantReceiptsFinalized(szStatmentID)
            !SupplierPayments = GetSupplierPayment(szStatmentID)
            '!BankPaymentReceipts = GetBankPaymentReceipts
            !ClientLandlordBalance = GetClientACBalance + GetLandLordACBalance
            !ClientPayments = GetClientPayments(szStatmentID)
            !LandlordPayments = GetLandLordPayments(szStatmentID)
            !ManagingAgentPayments = GetAGENTPayments(szStatmentID)
            !PayableAmount = 0 'txtRentPayable.text
            '!StatementClosingBal = getClosingBalance(dblLasClosingBalance)
            !StatementClosingBal = !Availablefunds
            !BankPayment = GetBankPaymentPreview(CLng(szStatmentID))
            !BankReceipts = GetBankreceiptsPreview(CLng(szStatmentID))

            !PINumber = ""
            !Generated_Date = Format(Now, "dd/mmmm/yyyy")
            !BankACBalance = BankAccBalance(adoConn, szSelectedBankAccount, szSelectedClient)
            !Printed = False
            !Emailed = False
            !Invoiced = False
            !PostToHistory = False
            !ReportGenID = szReportGenID
'            !InclSupplierOS = bolInclSupplierOS
'            !InclMngtFeesDue = InclMngtFeesDue
            !InclSupplierOS = IIf(chkExcludeSupOS.Value = 1, True, False)
            !InclMngtFeesDue = IIf(chkShowDue.Value = 1, True, False)
            .Update
    End With
    rsRentSummaryStatement.Close
    Set rsRentSummaryStatement = Nothing
    Call WorkOnMgtfeedueSupplierOS(adoConn, szStatmentID)  'added by anol 15/08/2023
    Call frmRentPayable.loadflxPayFees("")
    adoConn.Close
    Set adoConn = Nothing
End Sub


'Private Sub GenerateSummaryStatementModified(szStatmentID As String, ByVal szReportGenID As Long)
'
'    Dim adoConn As New ADODB.Connection
'    Dim rsRentSummaryStatement As New ADODB.Recordset
'    adoConn.Open getConnectionString
'    Dim dblLasClosingBalance As Double
'    Dim szSQL As String
''    szSQL = "Select StatementClosingBal from RentSummaryStatement where StatementID=" & szStatmentID & " AND ClientIDLandlordID='" & _
''    szSelectedClient & "'"
''    rsRentSummaryStatement.Open szSQL, adoconn, adOpenKeyset, adLockReadOnly
''    If Not rsRentSummaryStatement.EOF Then
''        dblLasClosingBalance = rsRentSummaryStatement!StatementClosingBal
''    End If
''
''    rsRentSummaryStatement.Close
''    Set rsRentSummaryStatement = Nothing
'
'    'szStatementID
'    'Exit Sub
'    rsRentSummaryStatement.Open "Select * from RentSummaryStatement where StatementID=" & szStatmentID & "", adoConn, adOpenDynamic, adLockOptimistic
'    With rsRentSummaryStatement
'            '.AddNew
''            !StatementID = szStatmentID 'we are setting this column atutomatically
''             szCurrentStatementID = szStatmentID
''            !StatementNo = GetLastStatementNoByClient + 1
''            !ClientIDLandlordID = szSelectedClient
'            !BankCode = szSelectedBankAccount
'            !PreviousStatementDate = IIf(GetLastStatementDateByClient(szStatmentID) = "", "01/01/2000", GetLastStatementDateByClient(szStatmentID)) 'This is Fromdate
'            !StatementDate = Format(txtStatementDate1.text, "dd/mmmm/yyyy") 'This is todate
''            !StatementOpBal = dblLasClosingBalance
''            !Retentions = Val(txtRetention.text) 'we need to further analyse detail/add/deduct retension
''            !Clearretentions = False 'Will need to come again
'
'            !AccrualsAcBalance = GetAccrualsControlBalance
'            !SupplierAcBalance = GetSupplierOSAmount 'GetBalance("Supplier") 'GetBalanceSupplier'wrong
'            !ManagingAgentAcBalance = GetAgentBalance 'GetBalance("Agent") 'GetBalanceAgent'wrong
'            !ClientAcBalance = GetClientACBalance
'            !LandlordACBalance = GetLandLordACBalance
'            !ListOffundID = ListOfFundsForDBSave ' szSelectedFund
''            !ListOfPayableTypeID = ListOfPayableTypesForDBSave ' ListOfPayableTypes
'            !TenantDepositsReceived = GetRentDeposit(szStatmentID)
'            !Availablefunds = getAvailablefundsModified(dblLasClosingBalance, szStatmentID)
'            !ListOfinputProperties = ListOfProperties
'            !PaymentsonAccount = -GetPaymentsonAccount
'            'New fields added 2021-01-24
'            !TenantReceipts = GetTenantReceiptsFinalized(szStatmentID)
'            !SupplierPayments = GetSupplierPayment(szStatmentID)
'            !BankPaymentReceipts = GetBankPaymentReceipts
'            !ClientLandlordBalance = GetClientACBalance + GetLandLordACBalance
'            !ClientPayments = GetClientPayments(szStatmentID)
'            !LandlordPayments = GetLandLordPayments(szStatmentID)
'            !ManagingAgentPayments = GetAGENTPayments(szStatmentID)
'            '!PayableAmount = 0 'txtRentPayable.text
'            !StatementClosingBal = getClosingBalanceFinalized(dblLasClosingBalance, szStatmentID)
'            !PINumber = ""
'            !Generated_Date = Format(Now, "dd/mmmm/yyyy")
'            !BankACBalance = BankAccBalance(adoConn, szSelectedBankAccount, szSelectedClient)
'            !Printed = False
'            !Emailed = False
'            !Invoiced = False
'            !PostTohistory = False
'            !ReportGenID = szReportGenID
'            .Update
'    End With
'    rsRentSummaryStatement.Close
'    Set rsRentSummaryStatement = Nothing
'    Call SaveRetentionDetails(adoConn)
'    Call frmRentPayable.loadflxPayFees("")
'    adoConn.Close
'    Set adoConn = Nothing
'End Sub

Private Sub GenerateSummaryStatementFinalized(szStatmentID As String, ByVal szReportGenID As Long)

    Dim adoConn As New ADODB.Connection
    Dim rsRentSummaryStatement As New ADODB.Recordset
    adoConn.Open getConnectionString
    Dim dblLasClosingBalance As Double
    Dim szSQL As String
'    szSQL = "Select StatementClosingBal from RentSummaryStatement where StatementID=" & szStatmentID & " AND ClientIDLandlordID='" & _
'    szSelectedClient & "'"
'    rsRentSummaryStatement.Open szSQL, adoconn, adOpenKeyset, adLockReadOnly
'    If Not rsRentSummaryStatement.EOF Then
'        dblLasClosingBalance = rsRentSummaryStatement!StatementClosingBal
'    End If
'
'    rsRentSummaryStatement.Close
'    Set rsRentSummaryStatement = Nothing
   
    'szStatementID
    'Exit Sub
    rsRentSummaryStatement.Open "Select * from RentSummaryStatement where StatementID=" & szStatmentID & "", adoConn, adOpenDynamic, adLockOptimistic
    With rsRentSummaryStatement
            '.AddNew
'            !StatementID = szStatmentID 'we are setting this column atutomatically
'             szCurrentStatementID = szStatmentID
'            !StatementNo = GetLastStatementNoByClient + 1
'            !ClientIDLandlordID = szSelectedClient
            !BankCode = szSelectedBankAccount
            !PreviousStatementDate = IIf(GetLastStatementDateByClient(szStatmentID) = "", "01/01/2000", GetLastStatementDateByClient(szStatmentID)) 'This is Fromdate
            !StatementDate = Format(txtStatementDate1.text, "dd/mmmm/yyyy") 'This is todate
'            !StatementOpBal = dblLasClosingBalance
           
'            !Clearretentions = False 'Will need to come again
            
            !AccrualsAcBalance = GetAccrualsControlBalance
            !SupplierAcBalance = GetSupplierOSAmount 'GetBalance("Supplier") 'GetBalanceSupplier'wrong
            !ManagingAgentAcBalance = GetAgentBalance 'GetBalance("Agent") 'GetBalanceAgent'wrong
            !ClientACBalance = GetClientACBalance
            !LandlordACBalance = GetLandLordACBalance
            !ListOffundID = ListOfFundsForDBSave ' szSelectedFund
'            !ListOfPayableTypeID = ListOfPayableTypesForDBSave ' ListOfPayableTypes
            !TenantDepositsReceived = GetRentDeposit(szStatmentID)
           ' !Availablefunds = getAvailablefundsModified(dblLasClosingBalance, szStatmentID)
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
            !ManagingAgentPayments = GetAGENTPaymentFinalized(szStatmentID) 'GetAGENTPayments
            !BankPayment = GetBankPaymentPreview(CLng(szStatmentID))
            !BankReceipts = GetBankreceiptsPreview(CLng(szStatmentID))
            '!PayableAmount = 0 'txtRentPayable.text
            '!StatementClosingBal = !Availablefunds ' getClosingBalanceFinalized(dblLasClosingBalance, szStatmentID)
            '!PINumber = ""
            !Generated_Date = Format(Now, "dd/mmmm/yyyy")
            !BankACBalance = BankAccBalance(adoConn, szSelectedBankAccount, szSelectedClient)
            '!Printed = False
            '!Emailed = False
            '!Invoiced = False
            '!PostTohistory = False
            !ReportGenID = szReportGenID
            .Update
    End With
    rsRentSummaryStatement.Close
    Set rsRentSummaryStatement = Nothing
'    Call SaveRetentionDetails(adoConn)
'    Call frmRentPayable.loadflxPayFees("")
    adoConn.Close
    Set adoConn = Nothing
End Sub
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
    'B )***********************  Sum of Rent received Paid/Refunded ***********************************8148.24


    szSQL = "Select  SUM(switch(type=3,S.Amount,type=4,S.Amount,type=23,-S.Amount)) as DR from tlbReceipt R,tlbReceiptSplit S,Fund F where R.TransactionID= S.RptHeader AND TYPE IN(3,4,23) " & _
    "AND S.FundID=F.FundID AND R.BankCODE='" & szSelectedBankAccount & "' and F.FundCode<>'TENANTDEPOSIT'  AND  S.ClientStatementID=" & trxToinclude & "  AND ClientID ='" & szSelectedClient & "' " & _
    "AND S.Amount>S.OSAmount AND R.RDate <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "#"
    rsReceipt.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
    If Not rsReceipt.EOF Then
            dblAmt = IIf(IsNull(rsReceipt.Fields.Item("Dr").Value), 0, rsReceipt.Fields.Item("Dr").Value)
             'result is 255
    End If
    rsReceipt.Close
    Set rsReceipt = Nothing
    getClosingBalanceFinalized = dblLasClosingBalance + dblAmt
 
   'c   (-): Sum of Supplier amounts Paid/Refunded (Both allocated and unallocated)-2207.55
 
    szSQL = "Select  SUM(SWITCH(P.TYPE=24,S.Amount,P.TYPE=8,-S.Amount,P.TYPE=9,-S.Amount)) as AMT from tlbPayment P,tlbPaymentSplit S,Fund F,Supplier SP where  " & _
            "SP.SupplierID=P.SageAccountNumber AND P.TransactionID=S.PayHeader AND P.TYPE IN(8,9,24) AND S.ClientStatementID=" & trxToinclude & " AND " & _
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
'Private Sub GenerateSummaryStatementNonConsolidated(szStatmentID As String, ByVal PropertyID As String, ByVal szReportGenID As Long)
'
'    Dim adoconn As New ADODB.Connection
'    Dim rsRentSummaryStatement As New ADODB.Recordset
'    adoconn.Open getConnectionString
'    Dim dblLasClosingBalance As Double
'    Dim szSQL As String
'    szSQL = "Select StatementClosingBal from RentSummaryStatement where StatementNo=" & GetLastStatementNoByClient & " AND ClientIDLandlordID='" & _
'    szSelectedClient & "'"
'    rsRentSummaryStatement.Open szSQL, adoconn, adOpenKeyset, adLockReadOnly
'    If Not rsRentSummaryStatement.EOF Then
'        dblLasClosingBalance = rsRentSummaryStatement!StatementClosingBal
'    End If
'
'    rsRentSummaryStatement.Close
'    Set rsRentSummaryStatement = Nothing
'
'    'szStatementID
'    rsRentSummaryStatement.Open "Select * from RentSummaryStatement where 1=2", adoconn, adOpenDynamic, adLockOptimistic
'    With rsRentSummaryStatement
'            .AddNew
'            !statementID = szStatmentID 'we are setting this column automatically
'             szCurrentStatementID = szStatmentID
'            !StatementNo = GetLastStatementNoByClient + 1
'            !ClientIDLandlordID = szSelectedClient
'            !BankCode = szSelectedBankAccount
'            !PreviousStatementDate = IIf(GetLastStatementDateByClient(szStatmentID) = "", "01/01/2000", GetLastStatementDateByClient(szStatmentID)) 'This is Fromdate
'            !StatementDate = Format(txtStatementDate1.text, "dd/mmmm/yyyy") 'This is todate
'            !StatementOpBal = dblLasClosingBalance
'            !Retentions = Val(txtRetention.text) 'we need to further analyse detail/add/deduct retension
'            !Clearretentions = False 'Will need to come again
'
'            !AccrualsAcBalance = GetAccrualsControlBalanceNonConsolidated(PropertyID)
'            !SupplierAcBalance = GetSupplierOSAmountNonConsolidated(PropertyID) 'GetBalance("Supplier") 'GetBalanceSupplier'wrong
'            !ManagingAgentAcBalance = GetAgentBalanceNonConsolidated(PropertyID)  'GetBalance("Agent") 'GetBalanceAgent'wrong
'            !ClientAcBalance = GetClientACBalance
'            !LandlordACBalance = GetLandLordACBalance
'            !ListOffundID = ListOfFundsForDBSave ' szSelectedFund
''            !ListOfPayableTypeID = ListOfPayableTypesForDBSave ' ListOfPayableTypes
'            !TenantDepositsReceived = GetRentDeposit(szStatmentID)
'            !Availablefunds = getAvailablefunds(dblLasClosingBalance)
'            !ListOfinputProperties = ListOfProperties
'            !PaymentsonAccount = -GetPaymentsonAccount
'            'New fields added 2021-01-24
'            !TenantReceipts = GetTenantReceiptsNonConsolidated(PropertyID)
'            !SupplierPayments = GetSupplierPayment
'            !BankPaymentReceipts = GetBankPaymentReceipts
'            !ClientLandlordBalance = GetClientACBalance + GetLandLordACBalance
'            !ClientPayments = GetClientPayments
'            !LandlordPayments = GetLandLordPayments
'            '!ManagingAgentPayments = GetAGENTPaymentsNonConsolidated(PropertyID) need to write this function later
'            !PayableAmount = txtRentPayable.text
'            !StatementClosingBal = getClosingBalance(dblLasClosingBalance)
'            !PINumber = ""
'            !Generated_Date = Format(Now, "dd/mmmm/yyyy")
'            !BankACBalance = BankAccBalance(adoconn, szSelectedBankAccount, szSelectedClient)
'            !Printed = False
'            !Emailed = False
'            !Invoiced = False
'            !PostTohistory = False
'            !ReportGenID = szReportGenID
'            !StatementNobyProperty = GetStatementNobyProperty(PropertyID)
'            !ConsolidatedPropID = PropertyID
'            '!isConsolidated = True
'            .Update
'    End With
'    rsRentSummaryStatement.Close
'    Set rsRentSummaryStatement = Nothing
'    Call SaveRetentionDetails(adoconn)
'    Call frmRentPayable.loadflxPayFees("")
'    adoconn.Close
'    Set adoconn = Nothing
'End Sub
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

Private Function GetBankPaymentPreview(ByRef idToinclude As Long) As Double
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
            "B.DEPT_ID=F.FundID   AND  B.RentSumStatement='" & idToinclude & "' AND B.ClientID ='" & _
            szSelectedClient & "' AND " & whereProperty
    
    rsPayment.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
    If Not rsPayment.EOF Then
        GetBankPaymentPreview = IIf(IsNull(rsPayment.Fields.Item("DR").Value), 0, rsPayment.Fields.Item("DR").Value)
    End If
    rsPayment.Close
    adoConn.Close
    Set adoConn = Nothing
End Function
Private Function GetBankreceiptsPreview(ByRef idToinclude As Long) As Double
    Dim rsPayment As New ADODB.Recordset
    Dim szSQL As String
    Dim adoConn As New ADODB.Connection
    'F.CategoryCode = 1 Fund category 1 Means rent
    'Implement switch here in SQL
    adoConn.Open getConnectionString
    Dim whereProperty As String
'    If boolConsolidatedStatement = 1 Then
'            whereProperty = "(B.PropertyID in (" & ListOfProperties & ") OR isnull(B.PropertyID) OR B.PropertyID='' ) "
'    Else
'            whereProperty = "B.PropertyID in (" & ListOfProperties & ") "
'    End If
    szSQL = "Select  SUM(SWITCH(TransactionType=12,B.NET_AMOUNT,TransactionType=11,-B.NET_AMOUNT)) as DR from tlbBankPayment B,Fund F where TransactionType IN(12) AND " & _
            "B.TRAN_DATE <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "# AND " & _
            "B.DEPT_ID=F.FundID  AND B.RentSumStatement='" & idToinclude & "'  AND B.ClientID ='" & szSelectedClient & "' " & whereProperty
    
    rsPayment.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
    If Not rsPayment.EOF Then
        GetBankreceiptsPreview = IIf(IsNull(rsPayment.Fields.Item("DR").Value), 0, rsPayment.Fields.Item("DR").Value)
    End If
    rsPayment.Close
    adoConn.Close
    Set adoConn = Nothing
End Function

Private Function GetBankPaymentReceiptsPreview(ByRef idToinclude As Long) As Double
    Dim rsPayment As New ADODB.Recordset
    Dim szSQL As String
    Dim adoConn As New ADODB.Connection
    'F.CategoryCode = 1 Fund category 1 Means rent
    'Implement switch here in SQL
    adoConn.Open getConnectionString

    szSQL = "Select  SUM(SWITCH(TransactionType=12,B.NET_AMOUNT,TransactionType=11,-B.NET_AMOUNT)) as DR from tlbBankPayment B,Fund F where TransactionType IN(11,12) AND " & _
            "B.TRAN_DATE >#" & Format(txtLastStatementDate1.text, "dd/mmm/yyyy") & "# AND B.TRAN_DATE <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "# AND " & _
            "B.DEPT_ID=F.FundID  AND (B.RentSumStatement=''  OR B.RentSumStatement='" & idToinclude & "' OR isnull(B.RentSumStatement))  AND B.ClientID ='" & szSelectedClient & "'"
    
    rsPayment.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
    If Not rsPayment.EOF Then
        GetBankPaymentReceiptsPreview = IIf(IsNull(rsPayment.Fields.Item("DR").Value), 0, rsPayment.Fields.Item("DR").Value)
    End If
    rsPayment.Close
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
        GetAgentBalance = Round(GetAgentBalance, 2) + IIf(IsNull(rsPayment.Fields.Item("AMT").Value), 0, rsPayment.Fields.Item("AMT").Value)
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
        GetAgentBalanceNonConsolidated = Round(GetAgentBalanceNonConsolidated, 2) + IIf(IsNull(rsPayment.Fields.Item("AMT").Value), 0, rsPayment.Fields.Item("AMT").Value)
    End If
    rsPayment.Close
    
    
    adoConn.Close
    Set adoConn = Nothing
End Function
Private Function GetSupplierOSAmount() As Double   'This function return result as minus'This is getting supplier balance
'Temporarily remming it 2023-08-11 by anol ' we are not using boolean supplier os amount

    Dim rsPayment As New ADODB.Recordset
    Dim szSQL As String
    Dim adoConn As New ADODB.Connection
    'F.CategoryCode = 1 Fund category 1 Means rent
    'Implement switch here in SQL
    'Bank code does not exits in PI,so do not put it in where clause
    If chkExcludeSupOS.Value = 1 Then
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
'            " SP.SupplierID=P.SageaccountNumber AND " & whereProperty & "  P.TransactionID=S.PayHeader AND P.ClientID ='" & szSelectedClient & "'  and SP.TYPE in ('Client') " & _
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
Private Function GetClientPaymentsPreview(ByRef idToinclude As Long) As Double   'This function return result as minus'This is getting supplier balance
    Dim rsPayment As New ADODB.Recordset
    Dim szSQL As String
    Dim adoConn As New ADODB.Connection
    'F.CategoryCode = 1 Fund category 1 Means rent
    'Implement switch here in SQL
    'Bank code does not exits in PI,so do not put it in where clause
    Dim whereProperty As String
    adoConn.Open getConnectionString
    If boolConsolidatedStatement = 1 Then
            whereProperty = "(P.UNITID IN (" & ListOfProperties & ") OR isnull(P.UNITID)  OR P.UNITID='' ) AND "
    Else
            whereProperty = "P.UNITID in (" & ListOfProperties & ") AND "
    End If
    szSQL = "Select  SUM(SWITCH(P.TYPE =24,S.Amount,P.TYPE =6,S.Amount,P.TYPE =7,-S.Amount,P.TYPE =8,-S.Amount,P.TYPE =9,-S.Amount)) as AMT " & _
            "from tlbPayment P,tlbPaymentSplit S,Fund F,Supplier SP where " & _
            " SP.SupplierID=P.SageaccountNumber AND " & whereProperty & "  P.TransactionID=S.PayHeader AND P.TYPE " & _
            "IN(7,8,9) AND S.FundID=F.FundID AND   F.FundCode in (" & _
             ListOfFunds & ") AND P.ClientID ='" & szSelectedClient & "'  and SP.TYPE in ('Client') " & _
             "and (P.RentSumStatement='' OR P.RentSumStatement='" & idToinclude & "' OR isnull(P.RentSumStatement)) and " & _
               " P.PDate >#" & Format(txtLastStatementDate1.text, "dd/mmm/yyyy") & "# AND P.PDate <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "#"
    
    rsPayment.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
    If Not rsPayment.EOF Then
        GetClientPaymentsPreview = IIf(IsNull(rsPayment.Fields.Item("AMT").Value), 0, rsPayment.Fields.Item("AMT").Value)
    End If
    rsPayment.Close
    adoConn.Close
    Set adoConn = Nothing
End Function
Private Function GetLandLordPaymentsPreview(ByRef idToinclude As Long) As Double   'This function return result as minus'This is getting supplier balance
    Dim rsPayment As New ADODB.Recordset
    Dim szSQL As String
    Dim adoConn As New ADODB.Connection
    'F.CategoryCode = 1 Fund category 1 Means rent
    'Implement switch here in SQL
    'Bank code does not exits in PI,so do not put it in where clause
    Dim whereProperty As String
    adoConn.Open getConnectionString
      If boolConsolidatedStatement = 1 Then
            whereProperty = "(P.UNITID IN (" & ListOfProperties & ") OR isnull(P.UNITID)  OR P.UNITID='' ) AND "
    Else
            whereProperty = "P.UNITID in (" & ListOfProperties & ") AND "
    End If
    szSQL = "Select  SUM(SWITCH(P.TYPE =24,S.Amount,P.TYPE =6,S.Amount,P.TYPE =7,-S.Amount,P.TYPE =8,-S.Amount,P.TYPE =9,-S.Amount)) as AMT from " & _
            "tlbPayment P,tlbPaymentSplit S,Fund F,Supplier SP where " & _
            " SP.SupplierID=P.SageaccountNumber AND " & whereProperty & "   P.TransactionID=S.PayHeader AND P.TYPE " & _
            "IN(7,8,9) AND S.FundID=F.FundID AND  (P.RentSumStatement='' OR P.RentSumStatement='" & idToinclude & "' OR isnull(P.RentSumStatement)) and F.FundCode in (" & _
             ListOfFunds & ") AND P.ClientID ='" & szSelectedClient & "'  and SP.TYPE in ('LLORD')" & _
            "AND P.PDate >#" & Format(txtLastStatementDate1.text, "dd/mmm/yyyy") & "# AND P.PDate <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "#"
    
    rsPayment.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
    If Not rsPayment.EOF Then
        GetLandLordPaymentsPreview = IIf(IsNull(rsPayment.Fields.Item("AMT").Value), 0, rsPayment.Fields.Item("AMT").Value)
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
            " SP.SupplierID=P.SageaccountNumber AND " & whereProperty & "   P.TransactionID=S.PayHeader AND S.ClientStatementID=" & trxToinclude & "  AND P.TYPE " & _
            "IN(8,9,24)  AND P.ClientID ='" & szSelectedClient & "'  and SP.TYPE in ('LLORD')" & _
            "AND P.PDate <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "#"

    rsPayment.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
    If Not rsPayment.EOF Then
        GetLandLordPayments = IIf(IsNull(rsPayment.Fields.Item("AMT").Value), 0, rsPayment.Fields.Item("AMT").Value)
    End If
    rsPayment.Close
    adoConn.Close
    Set adoConn = Nothing
End Function
Private Function GetAGENTPayments(ByVal trxToinclude As Long) As Double    'This function return result as minus'This is getting supplier balance
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

             
    szSQL = "Select   SUM(PS.PaymentAmount) as AMT from tlbPayment P,tlbPaymentSplit S,Paytransactions PS,Fund F,Supplier SP where  " & _
            "PS.FROMTran=P.TransactionID AND SP.SupplierID=P.SageAccountNumber AND P.TransactionID=S.PayHeader AND PS.Deleteflag=False AND  S.ClientStatementID=" & trxToinclude & "   AND " & _
            "PS.FundID=F.FundID and P.type=24 and " & whereProperty & " P.BankCODE='" & szSelectedBankAccount & "' and Sp.Type='AGENT' AND P.ClientID ='" & szSelectedClient & "' " & _
            "AND P.PDate <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "#"

rsPayment.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
    If Not rsPayment.EOF Then
        GetAGENTPayments = IIf(IsNull(rsPayment.Fields.Item("AMT").Value), 0, rsPayment.Fields.Item("AMT").Value)
    End If
    rsPayment.Close
'    whereProperty = "(S.PropertyID IN (" & ListOfProperties & ") OR isnull(S.PropertyID)  OR S.PropertyID='' ) AND " P.RentSumStatement='" & trxToinclude & "'

         szSQL = "Select   SUM(PS.PaymentAmount) as AMT from tlbPayment P,tlbPaymentSplit S,PaytransactionsSplit PS,Fund F,Supplier SP where  " & _
            "PS.TransactionID=S.PayTransactionIDSplit and PS.FromTran=P.TransactionID AND SP.SupplierID=P.SageAccountNumber  AND PS.Deleteflag=False " & _
            "AND S.ClientStatementID=" & trxToinclude & "  AND " & _
            " P.TransactionID=S.PayHeader AND PS.FundID=F.FundID and P.type IN(8,9) and " & whereProperty & " P.BankCODE='" & szSelectedBankAccount & "' " & _
            "and Sp.Type='AGENT' AND P.ClientID ='" & szSelectedClient & "'  AND P.PDate <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "#"

    
    rsPayment.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
    If Not rsPayment.EOF Then
        GetAGENTPayments = GetAGENTPayments - IIf(IsNull(rsPayment.Fields.Item("AMT").Value), 0, rsPayment.Fields.Item("AMT").Value)
    End If
    rsPayment.Close
    adoConn.Close
    Set adoConn = Nothing
End Function
Private Function GetAGENTPaymentFinalized(ByVal trtoinclude As Long) As Double    'This function return result as minus'This is getting supplier balance
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
            "S.Payheader=P.TransactionID AND SP.SupplierID=P.SageAccountNumber  AND S.Amount>S.OSAmount  AND S.ClientStatementID=" & trtoinclude & " AND " & _
            "S.FundID=F.FundID and P.type=24 and " & whereProperty & " P.BankCODE='" & szSelectedBankAccount & "' and Sp.Type='AGENT' and P.ClientID ='" & szSelectedClient & "' " & _
            "AND P.PDate <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "#"

rsPayment.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
    If Not rsPayment.EOF Then
        GetAGENTPaymentFinalized = IIf(IsNull(rsPayment.Fields.Item("AMT").Value), 0, rsPayment.Fields.Item("AMT").Value)
    End If
    rsPayment.Close
    whereProperty = "(S.PropertyID IN (" & ListOfProperties & ") OR isnull(S.PropertyID)  OR S.PropertyID='' ) AND "

'Otherwise you shall get duplicated value
         szSQL = "Select   SUM(S.Amount) as AMT from tlbPayment P,tlbPaymentSplit S,Fund F,Supplier SP where  " & _
            "S.Payheader=P.TransactionID  AND SP.SupplierID=P.SageAccountNumber  AND  S.Amount>S.OSAmount AND S.ClientStatementID=" & trtoinclude & " AND " & _
            " P.TransactionID=S.PayHeader AND S.FundID=F.FundID and P.type IN(8,9) and " & whereProperty & " P.BankCODE='" & szSelectedBankAccount & "' " & _
            "and Sp.Type='AGENT' AND P.ClientID ='" & szSelectedClient & "'  AND  P.PDate <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "#"

    
    rsPayment.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
    If Not rsPayment.EOF Then
        GetAGENTPaymentFinalized = GetAGENTPaymentFinalized - IIf(IsNull(rsPayment.Fields.Item("AMT").Value), 0, rsPayment.Fields.Item("AMT").Value)
    End If
    rsPayment.Close
    adoConn.Close
    Set adoConn = Nothing
End Function
Private Function GetAGENTPaymentsModified(ByVal idToinclude As Long) As Double   'This function return result as minus'This is getting supplier balance
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
        GetAGENTPaymentsModified = IIf(IsNull(rsPayment.Fields.Item("AMT").Value), 0, rsPayment.Fields.Item("AMT").Value)
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
        GetAGENTPaymentsModified = GetAGENTPaymentsModified - IIf(IsNull(rsPayment.Fields.Item("AMT").Value), 0, rsPayment.Fields.Item("AMT").Value)
    End If
    rsPayment.Close
    adoConn.Close
    Set adoConn = Nothing
    
End Function
'Private Function GetAGENTPaymentsModified(ByRef idToinclude As Long) As Double   'This function return result as minus'This is getting supplier balance
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
'
'
'    szSQL = "Select   SUM(PS.PaymentAmount) as AMT from tlbPayment P,Paytransactions PS,Fund F,Supplier SP where  " & _
'            "PS.FROMTran=P.TransactionID AND SP.SupplierID=P.SageAccountNumber  AND PS.Deleteflag=False AND  " & _
'               " P.RentSumStatement='" & idToinclude & "'  AND " & _
'            "PS.FundID=F.FundID and P.type=24 and " & whereProperty & " P.BankCODE='" & szSelectedBankAccount & "' and Sp.Type='AGENT' and " & _
'            " F.FundCode in (" & ListOfFunds & ") AND P.ClientID ='" & szSelectedClient & "' " & _
'            "AND P.PDate <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "#"
'
'
'
'    rsPayment.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'    If Not rsPayment.EOF Then
'        GetAGENTPaymentsModified = IIf(IsNull(rsPayment.Fields.Item("AMT").Value), 0, rsPayment.Fields.Item("AMT").Value)
'    End If
'    rsPayment.Close
'
'    whereProperty = "(S.PropertyID IN (" & ListOfProperties & ") OR isnull(S.PropertyID)  OR S.PropertyID='' ) AND "
'
''if you want to join with PaytransactionsSplit to get the correct Property on the payment you need to join PaytransactionsSplit with tlbPaymentSplit via PayTransactionIDSplit
''Otherwise you shall get duplicated value
'         szSQL = "Select   SUM(PS.PaymentAmount) as AMT from tlbPayment P,tlbPaymentSplit S,PaytransactionsSplit PS,Fund F,Supplier SP where  " & _
'            "PS.TransactionID=S.PayTransactionIDSplit and PS.FromTran=P.TransactionID AND SP.SupplierID=P.SageAccountNumber  AND " & _
'             "PS.Deleteflag=False AND  P.RentSumStatement='" & idToinclude & "' AND " & _
'            " P.TransactionID=S.PayHeader AND PS.FundID=F.FundID and P.type IN(8,9) and " & whereProperty & " P.BankCODE='" & szSelectedBankAccount & "' " & _
'            "and Sp.Type='AGENT' AND P.ClientID ='" & szSelectedClient & "'  AND  P.PDate <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "#"
'
'
'    rsPayment.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'    If Not rsPayment.EOF Then
'        GetAGENTPaymentsModified = GetAGENTPaymentsModified - IIf(IsNull(rsPayment.Fields.Item("AMT").Value), 0, rsPayment.Fields.Item("AMT").Value)
'    End If
'    rsPayment.Close
'    adoConn.Close
'    Set adoConn = Nothing
'End Function
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

    adoConn.Open getConnectionString
    szSQL = "Select  SUM(SWITCH(P.TYPE =24,S.Amount,P.TYPE =8,-S.Amount,P.TYPE =9,-S.Amount)) as AMT from tlbPayment P,tlbPaymentSplit S,Fund F,Supplier SP where " & _
            "SP.SupplierID=P.SageaccountNumber AND  P.TransactionID=S.PayHeader AND P.TYPE IN(8,9,24) AND " & _
            "P.PDate >#" & Format(txtLastStatementDate1.text, "dd/mmm/yyyy") & "# AND P.PDate <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "# AND " & _
            "S.FundID=F.FundID AND  P.BankCODE='" & szSelectedBankAccount & "'  AND " & _
            "(P.RentSumStatement='' OR P.RentSumStatement='" & idToinclude & "' OR isnull(P.RentSumStatement)) and  F.FundCode in (" & _
             ListOfFunds & ") AND P.ClientID ='" & szSelectedClient & "' AND (SP.Type='Supplier' ) " & whereProperty
    
    rsPayment.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
    If Not rsPayment.EOF Then
        GetSupplierPaymentPreview = IIf(IsNull(rsPayment.Fields.Item("AMT").Value), 0, rsPayment.Fields.Item("AMT").Value)
    End If
    rsPayment.Close
            whereProperty = "AND (G.PropertyID IN (" & ListOfProperties & ") OR isnull(G.PropertyID) OR G.PropertyID='') "

             szSQL = "Select  SUM(AL.VatAmount)  as AMT from tlbPayment P,PayTransactionsSplit AL,tlbPaymentSplit S,Fund F,Supplier SP,Property PR,GLobalData G   " & _
                "where  G.PropertyID=PR.PropertyID AND isAgentToSubmit=true  AND " & _
                "AL.TransactionID=S.PayTransactionIDSplit and SP.SupplierID=P.SageAccountNumber AND AL.Deleteflag=false and PR.propertyID=p.UNITID AND  " & _
                "P.TransactionID=S.PayHeader AND P.TYPE IN(8,9,24) AND (P.RentSumStatement='' OR P.RentSumStatement='" & idToinclude & "' or isnull(P.RentSumStatement)) AND " & _
                "S.FundID=F.FundID AND AL.FromTran=P.transactionID " & whereProperty & " AND P.BankCODE='" & szSelectedBankAccount & "' and   " & _
                "F.FundCode in (" & ListOfFunds & ") AND P.ClientID ='" & szSelectedClient & "' " & _
                "AND P.PDate >#" & Format(txtLastStatementDate1.text, "dd/mmm/yyyy") & "# AND P.PDate <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "#"
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
Private Function GetTenantReceipts(ByVal trxToinclude As Long) As Double
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
Private Function GetTenantReceiptsPreview(ByRef idToinclude As Long) As Double
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
            "AND R.amount>R.OSamount AND RS.FundID=F.FundID AND  R.BankCODE='" & szSelectedBankAccount & "'  AND (R.RentSumStatement='' OR R.RentSumStatement='" & _
             idToinclude & "' OR isnull(R.RentSumStatement)) and  F.FundCode in (" & _
             ListOfFunds & ") AND R.UnitID=U.UnitNumber AND " & whereProperty & " R.ClientID ='" & szSelectedClient & "' "
    
    rsPayment.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
    If Not rsPayment.EOF Then
        GetTenantReceiptsPreview = IIf(IsNull(rsPayment.Fields.Item("AMT").Value), 0, rsPayment.Fields.Item("AMT").Value)
    End If
    rsPayment.Close
            szSQL = "Select  SUM(AL.VatAmount) as DR from tlbReceipt R,rptTransactions AL,tlbReceiptSplit S,Fund F, Units U,GLobalData G where G.PropertyID=U.PropertyID " & _
            " AND G.isAgentToSubmit=true AND R.TransactionID= S.RptHeader AND TYPE IN(3,4,23) AND R.amount>R.OSamount AND S.FundID=F.FundID " & _
            "AND R.BankCODE='" & szSelectedBankAccount & "' and F.FundCode<>'TENANTDEPOSIT'  AND (R.RentSumStatement='' OR R.RentSumStatement='" & idToinclude & "' " & _
            " or isnull(R.RentSumStatement)) AND ClientID ='" & szSelectedClient & "' AND R.UnitID=U.UnitNumber and AL.Deleteflag=false and AL.FromTran=R.TransactionID  " & _
            "AND  F.FundCode in (" & ListOfFunds & ") AND " & whereProperty & " R.RDate >#" & Format(txtLastStatementDate1.text, "dd/mmm/yyyy") & "# AND R.RDate <=#" & _
            Format(txtStatementDate1.text, "dd/mmm/yyyy") & "#"
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
    Dim rCount As Integer
    Dim szStatmentID As String
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
    Dim rsRentSummaryStatement As New ADODB.Recordset
    adoConn.Open getConnectionString
    Dim dblLasClosingBalance As Double
    Dim szSQL As String
     szStatmentID = Replace(szCurrentStatementID, "CS", "")
    'Before writing this table you need to delete this table
    adoConn.Execute "Delete from  RentSummaryStatementPreview"
    szSQL = "Select StatementClosingBal from RentSummaryStatement where StatementNo=" & GetLastStatementNoByClient & " AND ClientIDLandlordID='" & szSelectedClient & "'"
    rsRentSummaryStatement.Open szSQL, adoConn, adOpenKeyset, adLockReadOnly
    If Not rsRentSummaryStatement.EOF Then
        dblLasClosingBalance = rsRentSummaryStatement!StatementClosingBal
    End If
    rsRentSummaryStatement.Close
    Set rsRentSummaryStatement = Nothing
    adoConn.Close
    Set adoConn = Nothing
    txtAvailableFunds.text = Format(getAvailablefundsModified(dblLasClosingBalance, szStatmentID), "0.00")
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

Private Function GetControlAccountForPayableString(adoConn As ADODB.Connection, szSelectedPayableTypeID As String) As Boolean
    Dim rsPayableTypes As New ADODB.Recordset
    rsPayableTypes.Open "Select * from  PayableTypes where ID=" & szSelectedPayableTypeID & "", adoConn, adOpenStatic, adLockReadOnly
    If Not rsPayableTypes.EOF Then
            GetControlAccountForPayableString = rsPayableTypes!PayNCAmt 'PayNCAmt
    End If
    rsPayableTypes.Close
    Set rsPayableTypes = Nothing

End Function


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
'        Dim dblDemandTypeId As Integer
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
            szSQLManagingAgent = "SELECT DISTINCT agr.ManagingAgentID " & _
              "FROM tlbAgreement agr, ClientProAgr CPA,  ChargeTypes C,Fund  F,SECONDARYCODE SC " & _
              "WHERE agr.CPA_ID = CPA.CPA_ID And (F.FundID)=agr.fund AND  SC.CODE=agr.CHARGE_METHOD AND " & _
              "CPA.ClientID = '" & szSelectedClient & "'  And C.ID = agr.CHARGE_TYPE And " & _
              "CPA.PropertyID = '" & szPropertySelection1 & "'"
              'to fix current problem add
              'AND CPA.PropertyID=D.PropertyID
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
              "WHERE P.CPA_ID = CPA.CPA_ID AND agr.CPA_ID = CPA.CPA_ID AND  P.ClientID=CPA.ClientID AND P.PAY_FUND=agr.fund And F.FundID=agr.fund AND  SC.CODE=agr.CHARGE_METHOD AND " & _
              "CPA.ClientID = '" & szSelectedClient & "'  And C.ID = agr.CHARGE_TYPE And " & _
              "CPA.PropertyID = '" & szPropertySelection1 & "' AND agr.ManagingAgentID='" & Trim(szManagingAgent(iManagingAgentCount)) & "'"
               'to fix current problem add
              'AND CPA.PropertyID=D.PropertyID
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
'                  dblDemandTypeId = rsCharge.Fields.Item("DEMAND_TYPE").Value
                  
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
'                                            szPropertySelection1 & "' and R.ISMGTFEE=false AND Rs.FundID=" & dblFundId & ""
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
'
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
                                                 
                                                        'need to consider the selected property in where clause
                                                        'modified on 20211103
'                                                        szSQLFrom = "Select  min(DS.DateFrom) as  DateFromMin ,max(DS.DateTO) as DateTOMAX   from tlbReceipt R,tlbReceiptsplit RS,tlbReceipt R1, " & _
'                                                        "rptTransactionsSPlit AL, DemandSplitRecords DS, Units U where R1.DemandRef=DS.DemandID and AL.TOTRAN=R1.TransactionID AND RS.SPLITID=DS.SPLITID AND AL.deleteflag=false AND " & _
'                                                        "AL.TransactionID= RS.RptTransactionsIDSplit AND  R.TransactionID=RS.RptHeader AND R.RDate<=#" & Format(txtStatementDate1.text, "dd MMM yyyy") & "# " & _
'                                                        "AND R.RDate>#" & Format(strLastChargeDate, "dd MMM yyyy") & "# and R.Type in (3,4,23)  AND U.UnitNumber=R.UnitID AND U.PropertyID='" & _
'                                                        szPropertySelection1 & "' and R.ISMGTFEE=false AND Rs.FundID=" & dblFundId & " AND DS.TypeOfDemand=" & rsCharge("DEMAND_TYPE").Value & " "

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
Private Function RentPayablePayment(szSelectedPI As String) As Boolean
'    Dim strTemp As String
'    Dim temp
'    temp = Split(szSelectedPI, " ")
'    Dim connn As New ADODB.Connection
'    Dim rsRentPayable As New ADODB.Recordset
'    Dim i As Integer
'    connn.Open getConnectionString
'    For i = 1 To UBound(temp)
'        strTemp = temp(i)
'        rsRentPayable.Open "Select * from tlbPayment where slnumber=" & StrDigitVal(strTemp) & " and type=6", connn, adOpenStatic, adLockReadOnly
'        If Not rsRentPayable.EOF Then
'            If rsRentPayable("OSamount") > 0 Then
'                     If MsgBox("You have rent payable invoices to pay. Do you wish to pay these invoices " & strTemp & " before finalising your client statement?", vbYesNo, "Please confirm") = vbYes Then
'                        connn.Close
'                        RentPayablePayment = True
'                        frmPurchaseExpense.tabPurExp.Tab = 1
'                        LoadForm frmPurchaseExpense
'                        frmPurchaseExpense.txtClientIDPurPay.text = szSelectedClient
'                        frmPurchaseExpense.txtBankCode.text = ""
'                        frmPurchaseExpense.txtBankAc.text = ""
'                        Exit Function
'                     Else
'                        'proceed do nothing
'                     End If
'            End If
'        End If
'        rsRentPayable.Close
'    Next i
'    connn.Close
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
Private Sub ShowRenPayable()
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
      MsgBox "Please select a statement.", vbInformation + vbOKOnly, "Statement Selection"
'      chkSelectAllDemands.Value = 0
      'ClearGridSelection
      Exit Sub
   End If
   
    iIncDec = 0
    
    
    Dim isitPlus As Boolean
    For rCount = 1 To frmRentPayable.flxPayFees.Rows - 1
         If frmRentPayable.flxPayFees.TextMatrix(rCount, 0) = "X" Then
             If frmRentPayable.flxPayFees.TextMatrix(rCount, 1) = "+" Or frmRentPayable.flxPayFees.TextMatrix(rCount, 1) = ">" Then
                isitPlus = True
             Else
                isitPlus = False
             End If
             iIncDec = iIncDec + 1
             selRow = rCount
         End If
    Next
    If iIncDec < 1 Or isitPlus = False Then
       MsgBox "Please select a statement at header level.", vbInformation + vbOKOnly, "Statement Selection"
       Exit Sub
    End If
    If szCurrentStatementID = "" Then
         Exit Sub
    End If
    'frmRentPayable.flxPayFees.TextMatrix(i, 3) is the statement ID by client
    '66) It should only be possible to modify a statement provided a rent payable
    ' PI has not been generated against the statement and a subsequent statement has not been produced.
    If isitPlus = True And frmRentPayable.flxPayFees.TextMatrix(selRow, 29) <> "" And frmRentPayable.flxPayFees.TextMatrix(selRow, 3) <= GetLastStatementNoByClient + 1 Then
        MsgBox "A Rent Payable invoice  " & frmRentPayable.flxPayFees.TextMatrix(selRow, 29) & " has already been generated against this statement.", vbInformation + vbOKOnly, "statement Selection"
        Exit Sub
    ElseIf isitPlus = False Then 'when you selected"-" it wont let you modify
        MsgBox "Please select a statement to modify.", vbInformation + vbOKOnly, "Statement Selection"
        Exit Sub
    End If
    
   If szCurrentStatementID <> "" Then
'            Frame1(6).Visible = False
'            Frame4.Caption = "Create PI from Statement: SS" & szCurrentStatementID
            frmGenaratePayable.strRef = "CS" & szCurrentStatementID & "/" & frmRentPayable.flxPayFees.TextMatrix(selRow, 3)
            frmGenaratePayable.szCurrentStatementID = szCurrentStatementID
            frmGenaratePayable.szClientID = frmRentPayable.flxPayFees.TextMatrix(selRow, 4)
            'frmGenaratePayable.txtClientAccount.text = XX
'            frmGenaratePayable.Show
'            frmGenaratePayable.ZOrder 0
            LoadForm frmGenaratePayable
'            Frame4.Left = 2070
'            Frame4.Top = 180
'            Frame4.Visible = True
'            txtAvailableFund1.text = szAvailableFund1
'            FocusControl txtRentPayable1
'            txtRentPayable1.SelStart = 0
'            txtRentPayable1.SelLength = Len(txtRentPayable1.text)
    End If
    
End Sub
Private Sub cmdFinalizeStatement_Click()
        Dim connn As New ADODB.Connection
        Dim szSelectedPI As String
         'Backup code start
   If MsgBox("Do you wish to take a backup?", vbYesNo + vbQuestion, "Data Backup") = vbYes Then
      If BackupDB Then
         MsgBox "Backup successful.", vbInformation, "The Backup has been successful"
      Else
         MsgBox "Backup failed. Please try again", vbInformation, "Warning"
      End If
   End If
   'Backup code End
        If validateIsFinalised = False Then
           Exit Sub
        End If
        Dim szSelectedClient As String
        connn.Open getConnectionString
        Dim szCurrentStatementIDID As String
        Dim rCount, iIncDec, selRow As Integer
               For rCount = 1 To frmRentPayable.flxPayFees.Rows - 1
                    If frmRentPayable.flxPayFees.TextMatrix(rCount, 0) = "X" Then
                        iIncDec = iIncDec + 1
                        selRow = rCount
                    End If
               Next
               szCurrentStatementIDID = frmRentPayable.flxPayFees.TextMatrix(selRow, 2)
               szCurrentStatementID = Replace(szCurrentStatementIDID, "CS", "")
        bEditMode = True
        whichFieldToCheck = "RentSumStatement"
        If bEditMode = True Then
            If PIvalidation = False Then
                MsgBox "Please select a statement at header level", vbInformation, "Warning"
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
             
'        If Val(txtRentPayable.text) > Val(txtAvailableFunds.text) Then
'                MsgBox "Rent Payable amount cannot be greater than the Available funds", vbInformation, "Warning!"
'                Exit Sub
'        End If
        If Trim(txtStatementDate1.text) = "" Then
              MsgBox "Please enter statment ", vbInformation, "Statement Date!!!"
              FocusControl txtStatementDate1
                Exit Sub
        End If
         For rCount = 1 To flxClients.Rows - 1
            If flxClients.TextMatrix(rCount, 0) = "X" Then
              szSelectedClient = flxClients.TextMatrix(rCount, 1)
            End If
         Next rCount
    
    
XX:
    
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
    Dim szSelectedPropertyID As String
    If szSelectedBankAccount = "" Then
        MsgBox "Please select a Bank account", vbInformation, "Warning "
        Exit Sub
    End If
    
   
'    If isAnyTransactionAvailable = False Then
'                MsgBox "There are no transactions for the statement period selected", vbInformation, "Warning!"
'                Exit Sub
'    End If
'    If GetSupplierOSAmount > 0 Then
'        If MsgBox("You have outstanding supplier balances to pay. Do you wish to pay them before finalising your client statement?", vbYesNo, "Supplier Os Balance") = vbYes Then
'                LoadForm frmPurchaseExpense
'                frmPurchaseExpense.tabPurExp.Tab = 1
'                frmPurchaseExpense.txtClientIDPurPay.text = szSelectedClient
'                frmPurchaseExpense.txtBankCode.text = ""
'                frmPurchaseExpense.txtBankAc.text = ""
'                Exit Sub
'        Else
'            'proceed
'        End If
'    End If
'The Finalise procedure needs to check if the rent payable has been paid
    'connn.Open getConnectionString
    Dim rsRentPayable As New ADODB.Recordset
    Dim rsRentPayableDetails As New ADODB.Recordset
    Dim rsPayment As New ADODB.Recordset
    rsRentPayable.Open "Select * from RentSummaryStatement where StatementID=" & szCurrentStatementID & "", connn, adOpenStatic, adLockReadOnly
    If Not rsRentPayable.EOF Then
        'PINumber
            If IsNull(rsRentPayable("PINumber").Value) Then
                If MsgBox("Rent payable has not been generated for this statement. Do you wish to generate Rent payable?", vbYesNo, "Warning") = vbYes Then
                    ' Show rent Payable
                    rsRentPayable.Close
                    connn.Close
                    Call ShowRenPayable
                    Exit Sub
                 Else
                    'proceed to generate
                 End If
            End If
            If (rsRentPayable("PINumber").Value) = "" Then
                If MsgBox("Rent payable has not been generated for this statement. Do you wish to generate Rent payable?", vbYesNo, "Warning") = vbYes Then
                    ' Show rent Payable
                    Call ShowRenPayable
                        rsRentPayable.Close
                        connn.Close
                        Exit Sub
                 Else
                     'proceed to generate
                 End If
            End If
        
    End If
    rsRentPayableDetails.Open "Select * from RentSummaryStatementDetails where StatementID=" & szCurrentStatementID & "", connn, adOpenStatic, adLockReadOnly
    While Not rsRentPayableDetails.EOF
         rsPayment.Open "Select * from tlbPayment P where P.slnumber=" & StrDigitVal(rsRentPayableDetails("PINumber").Value) & " ", connn, adOpenStatic, adLockReadOnly
         If Not rsPayment.EOF Then
                If rsPayment("OSAmount").Value > 0 Then
                         MsgBox "Rent payable amount hasn't been paid", vbInformation, "Warning " & rsRentPayableDetails("PINumber").Value
                         rsPayment.Close
                          connn.Close
                          Exit Sub
                End If
                
         End If
         rsPayment.Close
         rsRentPayableDetails.MoveNext
    Wend
    connn.Close
'' 15/08/2023 rem by anol
''    If Feestogenerate Then
''            If MsgBox("You have management fees to generate. Do you wish to generate  " & _
''                "them before finalise your client statement?", vbYesNo, "Please confirm") = vbYes Then
''                 frmManagementFeeSelection.Caption = "Management Fee Preview"
''                 frmManagementFeeSelection.szCallingFrom = "ManagementFee Preview"
''                 LoadForm frmManagementFeeSelection
''                 Exit Sub
''            Else
''                'proceed
''            End If
''    End If
    'check if management fee has been paid or not
    
    
''    If FeeshasBeenPaid Then
''        If MsgBox("You have management fees to pay. Do you wish to pay  " & _
''                "them before finalising your client statement?", vbYesNo, "Please confirm") = vbYes Then
''                LoadForm frmPurchaseExpense
''                frmPurchaseExpense.tabPurExp.Tab = 1
''                frmPurchaseExpense.txtClientIDPurPay.text = szSelectedClient
''                frmPurchaseExpense.txtBankCode.text = ""
''                frmPurchaseExpense.txtBankAc.text = ""
''                 Exit Sub
''            Else
''                'Do not proceed
''                Exit Sub
''            End If
''    End If
    
    Dim szStatmentID As String
    Dim szReportGenID As String
    Dim temp
    szReportGenID = ReportGenID
    For rCount = 1 To frmRentPayable.flxPayFees.Rows - 1
         If frmRentPayable.flxPayFees.TextMatrix(rCount, 0) = "X" Then
            szSelectedStatement = frmRentPayable.flxPayFees.TextMatrix(rCount, 2)
            szSelectedPI = frmRentPayable.flxPayFees.TextMatrix(rCount, 29)
            Exit For
         End If
    Next
    'here is the warning validation for non payment of rent payable
'    If RentPayablePayment(szSelectedPI) = False Then
'        Exit Sub
'    End If
  Dim strTemp As String
    'Dim temp
    temp = Split(szSelectedPI, " ")
    'Dim connn As New ADODB.Connection
    'Dim rsRentPayable As New ADODB.Recordset
    Dim i As Integer
    If connn.State = 0 Then
        connn.Open getConnectionString
    End If
    For i = 1 To UBound(temp)
        strTemp = temp(i)
        rsRentPayable.Open "Select * from tlbPayment where slnumber=" & StrDigitVal(strTemp) & " and type=6", connn, adOpenStatic, adLockReadOnly
        If Not rsRentPayable.EOF Then
            If rsRentPayable("OSamount") > 0 Then
                     If MsgBox("You have rent payable invoices to pay. Do you wish to pay these invoices " & strTemp & " before finalising your client statement?", vbYesNo, "Please confirm") = vbYes Then
                        connn.Close
'                        RentPayablePayment = True
                        frmPurchaseExpense.tabPurExp.Tab = 1
                        LoadForm frmPurchaseExpense
                        frmPurchaseExpense.txtClientIDPurPay.text = szSelectedClient
                        'Bank code and name was not coming automatically
                        ' Fixed by anol 2023-04-21
                        frmPurchaseExpense.txtBankCode.text = szSelectedBankAC1
                        frmPurchaseExpense.txtBankAc.text = szSelectedBankACName1
                        Exit Sub
                     Else
                        'proceed do nothing
                         Exit Sub
                     End If
            End If
        End If
        rsRentPayable.Close
    Next i
    'connn.Close
    
    szStatmentID = Replace(szSelectedStatement, "CS", "")
    If szStatmentID = "" Then Exit Sub
    If MsgBox("Are you sure, you wish to finalise this statement?", vbYesNo, "Please confirm") = vbYes Then
        
        whichFieldToCheck = "RentSumStatement"
        If boolConsolidatedStatement = 1 Then
            Call MarkAllTransactionsWithSSFinalized(szStatmentID)
            Call GenerateSummaryStatementFinalized(szStatmentID, szReportGenID)  'Write into SummaryStatement table in this function
            
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
          'If bEditDone Then
       connn.Execute "Update rentSummarystatement set isfinalized=1,DateFinalized=#" & Format(Date, "dd/MMM/yyyy") & "# where statementID=" & szCurrentStatementID & ""
       Call Sleep(100)
       'Call frmRentPayable.loadflxPayFees("")
       'msgbox""
    'End If
       connn.Close
       MsgBox "This statement has been finalised."
       Call frmRentPayable.loadflxPayFees("")
    End If
    
End Sub
Private Function validateIsFinalised() As Boolean
     Dim iIncDec As Long
    iIncDec = 0
    Dim rCount As Integer
    Dim selRow As Integer
    Dim isitPlus As Boolean
    Dim adoConn As New ADODB.Connection
    Dim rsFinalized As New ADODB.Recordset
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
    If isitPlus = True Then
        adoConn.Open getConnectionString
        rsFinalized.Open "Select * from RentSummaryStatement where StatementID=" & Replace(frmRentPayable.flxPayFees.TextMatrix(selRow, 2), "CS", "") & "", adoConn, adOpenStatic, adLockReadOnly
        If Not rsFinalized.EOF Then
            If rsFinalized("isfinalized").Value = "1" Then
                    rsFinalized.Close
                    adoConn.Close
                    MsgBox "This statement has already been finalised.", vbInformation + vbOKOnly, "Statement Selection"
                    validateIsFinalised = True
                    Exit Function
            End If
        End If
        adoConn.Close
      '  MsgBox "This statement has already been finalised.", vbInformation + vbOKOnly, "Statement Selection"
       ' Exit Function
    ElseIf isitPlus = True And frmRentPayable.flxPayFees.TextMatrix(selRow, 29) = "" Then
          adoConn.Open getConnectionString
        rsFinalized.Open "Select * from RentSummaryStatement where StatementID=" & frmRentPayable.flxPayFees.TextMatrix(selRow, 2) & "", adoConn, adOpenStatic, adLockReadOnly
        If Not rsFinalized.EOF Then
            If rsFinalized("PINumber").Value <> "" Then
                    rsFinalized.Close
                    adoConn.Close
                    'MsgBox "This statement has already been finalised.", vbInformation + vbOKOnly, "Statement Selection"
                    MsgBox "This statement cannot be finalised, because a Rent Payable invoice " & frmRentPayable.flxPayFees.TextMatrix(selRow, 29) & " has not been generated against it.", vbInformation + vbOKOnly, "statement Selection"
                    Exit Function
            End If
        End If
        adoConn.Close
      '  MsgBox "This statement has already been finalised.", vbInformation + vbOKOnly, "Statement Selection"
       ' Exit Function
        
        'MsgBox "This statement cannot be finalised, because a Rent Payable invoice " & frmRentPayable.flxPayFees.TextMatrix(selRow, 29) & " has not been generated against it.", vbInformation + vbOKOnly, "statement Selection"
       ' Exit Function
    ElseIf isitPlus = False Then 'when you selected"-" it wont let you modify
        MsgBox "Please select one statement to modify Rent Summary Statement", vbInformation + vbOKOnly, "Statement Selection"
        Exit Function
    End If
    validateIsFinalised = True
End Function
Private Sub cmdOKInouts_Click()
'    'validaton
'    'at least one bank , one fund, one property ,one payable type is selected.
'    'this procedure shall make visible only that
''    If Val(txtRentPayable.text) > Val(txtAvailableFunds.text) Then
''        MsgBox "Rent Payable amount cannot be greater than the Available funds", vbInformation, "Warning!"
''        Exit Sub
''    End If
'    If isAnyTransactionAfterClientStatementDate = True Then
'        MsgBox "There are transactions found after the statement Date.", vbInformation, "Warning!"
'    End If
'
'    If txtLastStatementDate1.text = "" And szStatementNo = 1 Then
'        txtLastStatementDate1.Locked = False
'        txtLastStatementDate1.text = "01/01/2000"
'        MsgBox "Please enter last statement date", vbInformation, "Warning!"
'        FocusControl txtLastStatementDate1
'        Exit Sub
'    End If
'    'for the first time they can modify the last statement date
'    'if szStatementNo is 1 that means it is first statement for the client
''    If bEditMode = True Then
'            If txtLastStatementDate1.text = "" Then
'                    txtLastStatementDate1.Locked = False
'                    If szStatementNo = 1 Then GoTo XX ' you can keep empty for the first statement date else you need to must enter date
'                    MsgBox "Please enter last statement date", vbInformation, "Warning!"
'
'                    Exit Sub
'             Else
'                    If DateDiff("d", txtStatementDate1.text, txtLastStatementDate1.text) >= 0 And bEditMode = False Then
'                        MsgBox "A statement already exists for this date. Please enter a date after the 'Last Statement Date'", vbInformation, "Statement Date!"
'                        Exit Sub
'                    End If
'
'             End If
''    Else
''
''    End If
'XX:
'    'Validation for Client
'    Dim rCount As Integer
'    Dim selRow As Integer
'    Dim iIncDec As Long
'    For rCount = 1 To flxClients.Rows - 1
'         If flxClients.TextMatrix(rCount, 0) = "X" Then
'             iIncDec = iIncDec + 1
'             selRow = rCount
'         End If
'    Next
'    If iIncDec <> 1 Then
'       MsgBox "Please select a client.", vbInformation + vbOKOnly, "Client Selection"
'       Exit Sub
'    End If
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
'    iIncDec = 0
'    For rCount = 1 To flxProperties.Rows - 1
'         If flxProperties.TextMatrix(rCount, 0) = "X" Then
'             iIncDec = iIncDec + 1
'             selRow = rCount
'         End If
'    Next
''    If iIncDec <> 1 And boolConsolidatedStatement = 0 Then
''       MsgBox "Please select only one property.", vbInformation + vbOKOnly, "Property Selection"
''       Exit Sub
''    End If
'
'
'
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
'    If szSelectedBankAccount = "" Then
'
'        MsgBox "Please select a Bank account", vbInformation, "Warning!"
'        Exit Sub
'    End If
''    If ListOfPayableTypes = "" Then
''        MsgBox "Please select a Payable Type", vbInformation, "Warning!"
''        Exit Sub
''    End If
''        If isAnyTransactionAvailable = False Then
''                MsgBox "There are no transactions for the statement period selected", vbInformation, "Warning!"
''                Exit Sub
''        End If
'    Dim szStatmentID As String
'   iIncDec = 0
'   iIncDec = 0
'   Dim adoconn As New ADODB.Connection
'   For rCount = 1 To frmRentPayable.flxPayFees.Rows - 1
'        If frmRentPayable.flxPayFees.TextMatrix(rCount, 0) = "X" Then
'            iIncDec = iIncDec + 1
'            selRow = rCount
'        End If
'   Next
'
'
'           whichFieldToCheck = "RentSumStatementPreview"
'
'
'                szStatmentID = frmRentPayable.flxPayFees.TextMatrix(selRow, 2)
'                szStatmentID = Replace(szStatmentID, "CS", "")
'
'           Call GeneratePreview(szStatmentID)
'           Call MarkAllTransactionsWithSS(szStatmentID)
'
'
'            Dim reportApp As New CRAXDRT.Application
'            Dim Report As CRAXDRT.Report
'            Set Report = reportApp.OpenReport(App.Path & szReportPath & "\RentSummaryStatementPreview.rpt")
'            Report.EnableParameterPrompting = False
'            Report.DiscardSavedData
'            Report.ParameterFields(1).AddCurrentValue CInt(szStatmentID)
'            Load frmReport
'            frmReport.LoadReportViewer Report

End Sub
Private Sub printClientStatement(szCurrentStatementID As String, selRow As Integer)
   Dim iIncDec As Long
   iIncDec = 0
   Dim rCount As Integer
'   Dim selRow As Integer
   Dim adoConn As New ADODB.Connection
'   For rCount = 1 To frmRentPayable.flxPayFees.Rows - 1
'        If frmRentPayable.flxPayFees.TextMatrix(rCount, 0) = "X" Then
'            iIncDec = iIncDec + 1
'            selRow = rCount
'        End If
'   Next
'   If iIncDec < 1 Then
'      MsgBox "Please select one statement only.", vbInformation + vbOKOnly, "statement Selection"
'      Exit Sub
'   End If
'            whichFieldToCheck = "RentSumStatement"
'            Call GeneratePreview(szCurrentStatementID)
            'Call MarkAllTransactionsWithSS(szCurrentStatementID)
   ' szCurrentStatementID = frmRentPayable.flxPayFees.TextMatrix(selRow, 2)
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
'            whichFieldToCheck = "RentSumStatement"
'            Call GeneratePreview(szCurrentStatementID)
            'Call MarkAllTransactionsWithSS(szCurrentStatementID)
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

Private Function isAnyTransactionAfterClientStatementDate() As Boolean
        Dim adoConn As New ADODB.Connection
        Dim whereProperty As String
        Dim szSQL As String
        Dim rsReceipt As New ADODB.Recordset
        Dim dblAmt As Double
        If ListOfProperties = "" Then
            MsgBox "Please select a Property", vbInformation, "Warning"
            Exit Function
        End If
        adoConn.Open getConnectionString
        'We cannot filter by fund here because we are not counting tlbPaymentSplit in this SQL
        whereProperty = "(P.UNITID IN (" & ListOfProperties & ") OR isnull(P.UNITID) or P.UNITID='' ) AND "
        szSQL = "Select  SUM(S.Amount) as AMT from tlbPayment P,tlbPaymentSplit S,Supplier SP where " & _
                " S.Payheader=P.TransactionID AND SP.SupplierID=P.SageaccountNumber AND P.type in(7,8,9) AND " & whereProperty & " " & _
                " P.PDate >#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "# AND " & _
                " P.ClientID ='" & szSelectedClient & "'"
                
                '" R.RDate >=#" & Format(txtLastStatementDate1.text, "dd/mmm/yyyy") & "# AND R.RDate <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "#" & _
        rsReceipt.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
        rsReceipt.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
        If Not rsReceipt.EOF Then
             dblAmt = IIf(IsNull(rsReceipt.Fields.Item("amt").Value), 0, rsReceipt.Fields.Item("amt").Value)
        End If
        rsReceipt.Close
        whereProperty = "(R.UNITID IN (" & ListOfProperties & ") OR isnull(R.UNITID) or R.UNITID='' ) AND "
        szSQL = "Select  SUM(R.Amount) as DR from tlbReceipt R,Fund F, Units U " & _
                "where R.UnitID=U.UnitNumber AND R.type in(3,4,23) AND " & whereProperty & " F.FundCode IN(" & ListOfFunds & ") AND  R.FundID=F.FundID AND ClientID ='" & _
                " R.RDate > #" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "#" & _
                 szSelectedClient & "'"
        rsReceipt.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
        If Not rsReceipt.EOF Then
             dblAmt = dblAmt + IIf(IsNull(rsReceipt.Fields.Item("Dr").Value), 0, rsReceipt.Fields.Item("Dr").Value)
        End If
        rsReceipt.Close
        
        whereProperty = "(B.PROPERTYID IN (" & ListOfProperties & ") OR isnull(B.PROPERTYID) or B.PROPERTYID='' ) AND "
        szSQL = "Select  SUM(B.NET_AMOUNT) as DR from tlbBankPayment B,Fund F where TransactionType " & _
                "IN(11,12) AND " & whereProperty & " B.DEPT_ID=F.FundID and " & _
                "  B.TRAN_DATE >#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "# " & _
                "AND B.ClientID ='" & szSelectedClient & "'"

        rsReceipt.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
        If Not rsReceipt.EOF Then
             dblAmt = dblAmt + IIf(IsNull(rsReceipt.Fields.Item("Dr").Value), 0, rsReceipt.Fields.Item("Dr").Value)
        End If
        rsReceipt.Close
        If dblAmt > 0 Then
            isAnyTransactionAfterClientStatementDate = True
        End If
        adoConn.Close
        Set adoConn = Nothing
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
   
    szSQL = "Select  SUM(SWITCH(TYPE=1,S.Amount,TYPE=2,S.Amount,TYPE=3,-S.Amount,TYPE=4,-S.Amount,TYPE=23,-S.Amount)) as DR from tlbReceipt R,tlbReceiptSplit S,Fund F, Units U " & _
            "where R.transactionID =S.rptHeader and  R.UnitID=U.UnitNumber AND S.ClientStatementID=" & trxToinclude & " AND " & whereProperty & "  TYPE IN(3,4,23) AND R.FundID=F.FundID and F.FundCode='RENTDEPOSIT' AND ClientID ='" & _
             szSelectedClient & "' AND R.RDate <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "# "

'    If boolConsolidatedStatement = 1 Then
'            whereProperty = "(B.PROPERTYID IN (" & ListOfProperties & ") OR isnull(B.PROPERTYID) or B.PROPERTYID='' ) AND "
'    Else
'            whereProperty = "B.PROPERTYID  in (" & ListOfProperties & ") AND "
'    End If
    
    
    szSQL2 = "Select  SUM(SWITCH(TransactionType=11,B.NET_AMOUNT,TransactionType=12,-B.NET_AMOUNT)) as DR from tlbBankPayment B,Fund F where TransactionType " & _
            " IN(11,12) AND " & whereProperty & " B.DEPT_ID=F.FundID and F.FundCode='RENTDEPOSIT' AND B.RentSumStatement='" & trxToinclude & "'  AND B.ClientID ='" & szSelectedClient & "'" & _
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
    
    szSQL = "Select Sum(AMOUNT) as SumAmount from NLPOSTING where NOMINAL_CODE='" & AccrualsCode & "' AND ClientID='" & _
    szSelectedClient & "' AND PROPERTY_ID in (" & ListOfProperties & ") AND DeleteFlag=false AND TRANSACTION_DATE <=#" & Format(txtStatementDate1.text, "dd/mmm/yyyy") & "#"
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





'Private Sub printClientStatement()
'   Dim iIncDec As Long
'   iIncDec = 0
'   Dim rCount As Integer
'   Dim selRow As Integer
'   For rCount = 1 To frmRentPayable.frmRentPayable.flxPayFees.Rows - 1
'        If frmRentPayable.frmRentPayable.flxPayFees.TextMatrix(rCount, 0) = "X" Then
'            iIncDec = iIncDec + 1
'            selRow = rCount
'        End If
'   Next
'   If iIncDec < 1 Then
'      MsgBox "Please select one statement only.", vbInformation + vbOKOnly, "statement Selection"
'      Exit Sub
'   End If
''            whichFieldToCheck = "RentSumStatement"
''            Call GeneratePreview(szCurrentStatementID)
'            'Call MarkAllTransactionsWithSS(szCurrentStatementID)
'            szCurrentStatementID = frmRentPayable.frmRentPayable.flxPayFees.TextMatrix(selRow, 1)
'            'run TestReportForRentSummary.rpt
'            Dim reportApp As New CRAXDRT.Application
'            Dim Report As CRAXDRT.Report
'
'           ' Set Report = reportApp.OpenReport(App.Path & szReportPath & "\TestReportForRentSummary.rpt")
'             Set Report = reportApp.OpenReport(App.Path & szReportPath & "\TestReportForRentSummary.rpt")
'            Report.EnableParameterPrompting = False
'            Report.DiscardSavedData
'            Report.ParameterFields(1).AddCurrentValue CInt(Right(szCurrentStatementID, Len(szCurrentStatementID) - 2))
'
'            '               Report.ParameterFields(1).AddCurrentValue CStr(txtLLID.text)
'            '               Report.ParameterFields(2).AddCurrentValue CDate(txtFromDate.text)
'            '               Report.ParameterFields(3).AddCurrentValue CDate(txtToDate.text)
'            '                Report.ParameterFields(4).AddCurrentValue cboCategory.text
'            Load frmReport
'            frmReport.LoadReportViewer Report
'End Sub


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
Private Function PIvalidationModification() As Boolean
    Dim iIncDec As Long
    iIncDec = 0
    Dim rCount As Integer
    Dim selRow As Integer
    Dim isitPlus As Boolean
    Dim rsPIExists As New ADODB.Recordset
    Dim adoConn As New ADODB.Connection
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
    
        adoConn.Open getConnectionString
        rsPIExists.Open "Select * from RentSummaryStatement where StatementID=" & Replace(frmRentPayable.flxPayFees.TextMatrix(selRow, 2), "CS", "") & "", adoConn, adOpenStatic, adLockReadOnly
        If Not rsPIExists.EOF Then
            If IsNull(rsPIExists("PINumber").Value) Or rsPIExists("PINumber").Value = "" Then
                    rsPIExists.Close
                    adoConn.Close
                     PIvalidationModification = True
                     Exit Function
            End If
            If rsPIExists("PINumber").Value <> "" Then
                    rsPIExists.Close
                    adoConn.Close
                    MsgBox "You cannot modify this statement because a Rent Payable Invoice has been generated against it.", vbInformation, "Warning!"
                    PIvalidationModification = False
                    Exit Function
            End If
        End If
        adoConn.Close
    
    
'         MsgBox "You cannot modify this statement because a Rent Payable Invoice has been generated against it.", vbInformation, "Warning!"
'        PIvalidationModification = False
    ElseIf isitPlus = True And frmRentPayable.flxPayFees.TextMatrix(selRow, 30) <> "" Then
       ' MsgBox "You cannot modify this statement because a Rent Payable Invoice has been generated against it.", vbInformation, "Warning!"
        MsgBox "You cannot modify this statement because this statement is finalized.", vbInformation, "Warning!"
        PIvalidationModification = False
    Else
        PIvalidationModification = True
    End If
    
    'PIvalidation = True
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
Private Sub cmdSave_Click()
    Dim rCount As Integer
    Dim selRow As Integer
    Dim iIncDec As Long

    
    If PIvalidationModification = False Then
         
        Exit Sub
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
             
             
'    If Val(txtRentPayable.text) > Val(txtAvailableFunds.text) Then
'        MsgBox "Rent Payable amount cannot be greater than the Available funds", vbInformation, "Warning!"
'        Exit Sub
'    End If
    If Trim(txtStatementDate1.text) = "" Then
              MsgBox "Please enter statment ", vbInformation, "Statement Date!!!"
              FocusControl txtStatementDate1
        Exit Sub
    End If
    
    
XX:
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
    Dim szSelectedPropertyID As String
    If szSelectedBankAccount = "" Then
        MsgBox "Please select a Bank account", vbInformation, "Warning "
        Exit Sub
    End If
'    If isAnyTransactionAvailable = False Then
'                MsgBox "There are no transactions for the statement period selected", vbInformation, "Warning!"
'                Exit Sub
'    End If

    Dim intmaxStatementNo As Integer
    Dim adoConn As New ADODB.Connection
    Dim rsRentSummaryStatement As New ADODB.Recordset
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
    
    
    
    adoConn.Open getConnectionString
    Dim szSQL As String
    szSQL = "Select max(StatementID) as IDbyCL from RentSummaryStatement where ClientIDLandlordID='" & szSelectedClient & "'"
    rsRentSummaryStatement.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
    If Not rsRentSummaryStatement.EOF Then
        intmaxStatementNo = IIf(IsNull(rsRentSummaryStatement!IDbyCL), 0, rsRentSummaryStatement!IDbyCL)
    End If
    rsRentSummaryStatement.Close
    Set rsRentSummaryStatement = Nothing
    adoConn.Close
    Set adoConn = Nothing
    If intmaxStatementNo = szCurrentStatementID Then
    Else
        MsgBox "You can only modify last statement", vbInformation, "Warning"
        Exit Sub
    End If
    
    If GetSupplierOSAmount > 0 Then
        If MsgBox("You have outstanding supplier balances to pay. Do you wish to pay them before modifying your client statement?", vbYesNo, "Supplier Os Balance") = vbYes Then
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
        If MsgBox("You have outstanding Client balances to pay. Do you wish to pay them before modifying your client statement?", vbYesNo, "Client Os Balance") = vbYes Then
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
'    If GetAgentOSAmount > 0 Then
'        If MsgBox("You have outstanding Managing Agent balances to pay. Do you wish to pay them before modifying your client statement?", vbYesNo, "Managing Agent Os Balance") = vbYes Then
'                LoadForm frmPurchaseExpense
'                frmPurchaseExpense.tabPurExp.Tab = 1
'                frmPurchaseExpense.txtClientIDPurPay.text = szSelectedClient
'                frmPurchaseExpense.txtBankCode.text = ""
'                frmPurchaseExpense.txtBankAc.text = ""
'                Exit Sub
'        Else
'            'proceed
'        End If
'    End If
    
    If Feestogenerate Then
            If MsgBox("You have Management Fees to generate. Do you wish to generate  " & _
                "them before modifying your client statement", vbYesNo, "Please confirm") = vbYes Then
                 LoadForm frmManagementFeeSelection
                 frmManagementFeeSelection.szCallingFrom = "ManagementFee Preview"
                 Exit Sub
            Else
                'proceed
            End If
    End If
    'check if management fee has been paid or not
    If FeeshasBeenPaid Then
        If MsgBox("You have Management Fees to pay. Do you wish to pay  " & _
                "them before modifying your client statement", vbYesNo, "Please confirm") = vbYes Then
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
    'If bPreviewMode = False Then
    Dim szStatmentID As String
    Dim szReportGenID As String
    szReportGenID = ReportGenID
    
    If MsgBox("Are you sure, you wish to modify this statement?", vbYesNo, "Please confirm") = vbYes Then
'        If bEditMode = False Then
'            szStatmentID = GetLastStatementID + 1
'            bEditDone = False
'        Else
                szStatmentID = Replace(szCurrentStatementID, "CS", "")
'              'When you are modifying a statement then it is using the selected statment ID from the parant form(rentPayable  form) )
                'Dim adoConn As New ADODB.Connection
                adoConn.Open getConnectionString
''                Dim rsCheckFlag As New ADODB.Recordset
''                rsCheckFlag.Open "Select InclSupplierOS,InclMngtFeesDue from RentSummaryStatement where  StatementID=" & szStatmentID & "", adoconn, adOpenStatic, adLockReadOnly
''                If Not rsCheckFlag.EOF Then
''                        bolInclSupplierOS = IIf(IsNull(rsCheckFlag("InclSupplierOS").Value) = True, 0, rsCheckFlag("InclSupplierOS").Value)
''                        InclMngtFeesDue = IIf(IsNull(rsCheckFlag("InclMngtFeesDue").Value) = True, 0, rsCheckFlag("InclMngtFeesDue").Value)
''                End If
''                rsCheckFlag.Close
                adoConn.Execute "Delete from RentSummaryStatement where  StatementID=" & szStatmentID & ""
                adoConn.Execute "Delete from RentSummaryStatementPreview where  StatementID=" & szStatmentID & ""
                adoConn.Execute "Update tlbBankPayment Set RentSumStatement=null where  RentSumStatement='" & szStatmentID & "'"
                adoConn.Execute "Update tlbPaymentSplit Set ClientStatementID=null where  ClientStatementID=" & szStatmentID & ""
                adoConn.Execute "Update tlbReceiptSplit Set ClientStatementID=null where  ClientStatementID=" & szStatmentID & ""
                adoConn.Close
                bEditDone = True
'        End If
        whichFieldToCheck = "ClientStatementID"
        If boolConsolidatedStatement = 1 Then
            Call MarkAllTransactionsWithSS(szStatmentID)
            Call GenerateSummaryStatementModify(szStatmentID, szReportGenID)  'Write into SummaryStatement table in this function
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
           
    End If
     Call frmRentPayable.loadflxPayFees("")
'    Frame1(6)  .Visible = False
    Unload Me
End Sub

Private Sub cmdTestReport_Click()
    Dim rCount As Integer
    Dim szSelectedStatement As String
    Dim szStatmentID As String
    Dim adoConn As New ADODB.Connection
    Dim rsRentSummaryStatement As New ADODB.Recordset
    For rCount = 1 To frmRentPayable.flxPayFees.Rows - 1
         If frmRentPayable.flxPayFees.TextMatrix(rCount, 0) = "X" Then
            szSelectedStatement = frmRentPayable.flxPayFees.TextMatrix(rCount, 2)
            Exit For
         End If
    Next

    szStatmentID = Replace(szSelectedStatement, "CS", "")
    If szStatmentID = "" Then Exit Sub
    If MsgBox("Are you sure, you wish to fix available fund?", vbYesNo, "Please confirm") = vbYes Then

 adoConn.Open getConnectionString

        rsRentSummaryStatement.Open "Select * from RentSummaryStatement where StatementID=" & szStatmentID & "", adoConn, adOpenDynamic, adLockOptimistic
            With rsRentSummaryStatement
                 !Availablefunds = getAvailablefundsModified(0, szStatmentID)
        End With
        rsRentSummaryStatement.Update
      rsRentSummaryStatement.Close
      Sleep (100)
      'Call frmRentPayable.loadflxPayFees("")
      adoConn.Close
      Call frmRentPayable.loadflxPayFees("")
      MsgBox "Av Fund has been fixed"
   End If
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
    chkInFunds.Value = 0
     
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



Private Sub flxClients_Click()
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
           ' flxProperties.RowHeight(iRow) = 280
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
'this one selects multiple lines
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

Private Sub Form_Activate()
    Dim adoConn As New ADODB.Connection
    Dim iRow As Integer
    Dim szSelectedFunds1  As String
    Dim szSelectedProperties1 As String
   
    If bEditMode = True Then
        If szCurrentStatementID = "" Then Exit Sub
        adoConn.Open getConnectionString
        Dim szSQL As String
        Dim rsRentSummaryStatement As New ADODB.Recordset
'        Dim adoconn As New ADODB.Connection
'        adoconn.Open getConnectionString
        szSQL = "Select StatementNo,ClientIDLandlordID,ListOfFundId,ListOfinputProperties,BankCode,Retentions,PayableAmount,AvailableFunds,PreviousStatementDate," & _
                "StatementDate,InclSupplierOS,InclMngtFeesDue  from RentSummaryStatement where StatementID=" & Replace(szCurrentStatementID, "CS", "") & ""
        rsRentSummaryStatement.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
        If Not rsRentSummaryStatement.EOF Then
            For iRow = 1 To flxClients.Rows - 1
                If flxClients.TextMatrix(iRow, 1) = rsRentSummaryStatement("ClientIDLandlordID").Value Then
                    flxClients.TextMatrix(iRow, 0) = "X"
                    szSelectedClient = flxClients.TextMatrix(iRow, 1)
                    szSelectedFunds1 = rsRentSummaryStatement("ListOfFundId").Value
                    szSelectedProperties1 = rsRentSummaryStatement("ListOfinputProperties").Value
                    szSelectedBankAC1 = rsRentSummaryStatement("BankCode").Value
                    chkExcludeSupOS.Value = IIf((IIf(IsNull(rsRentSummaryStatement("InclSupplierOS").Value), 0, rsRentSummaryStatement("InclSupplierOS").Value)), 1, 0)
                    chkShowDue.Value = IIf((IIf(IsNull(rsRentSummaryStatement("InclMngtFeesDue").Value), 0, rsRentSummaryStatement("InclMngtFeesDue").Value)), 1, 0)
               
                    Dim rstClient   As New ADODB.Recordset
                    szSQL = "SELECT MY_ID, NominalCode, Bank_AC_Name, CLIENT_ID " & _
                               "FROM tlbClientBanks where CLIENT_ID='" & szSelectedClient & "' and NominalCode='" & szSelectedBankAC1 & "'" & _
                               "ORDER BY NominalCode;"

                     rstClient.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
                    If Not rstClient.EOF Then
                            szSelectedBankACName1 = rstClient("Bank_AC_Name").Value
                    End If
                    rstClient.Close
                    Set rstClient = Nothing
                       
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
                 flxBankAccounts.TextMatrix(iRow, 0) = ""
        Next
        For iRow = 1 To flxBankAccounts.Rows - 1
             If InStr(1, szSelectedBankAC1, flxBankAccounts.TextMatrix(iRow, 2)) > 0 Then
                    flxBankAccounts.TextMatrix(iRow, 0) = "X"
             End If
        Next
        For iRow = 1 To flxInFunds.Rows - 1
              flxInFunds.TextMatrix(iRow, 0) = ""
        Next
        For iRow = 1 To flxInFunds.Rows - 1
             If InStr(1, szSelectedFunds1, flxInFunds.TextMatrix(iRow, 1)) > 0 Then
                    flxInFunds.TextMatrix(iRow, 0) = "X"
             End If
        Next
        For iRow = 1 To flxProperties.Rows - 1
              flxProperties.TextMatrix(iRow, 0) = ""
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
    Frame1(6).BackColor = MODULEBACKCOLOR
    Frame1(12).BackColor = MODULEBACKCOLOR
    Frame1(8).BackColor = MODULEBACKCOLOR
    Frame1(9).BackColor = MODULEBACKCOLOR
    Frame1(10).BackColor = MODULEBACKCOLOR
    chkAllProperties.BackColor = MODULEBACKCOLOR
    chkExcludeSupOS.Value = 0
    chkShowDue.Value = 0
    chkExcludeSupOS.BackColor = MODULEBACKCOLOR
    chkExcludeSupOS.Enabled = False
    chkShowDue.BackColor = MODULEBACKCOLOR
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
    szSQL = "SELECT Distinct FundID, FundName, FundCode,CategoryCode FROM Fund LEFT JOIN tlbPayable PB ON PB.Pay_fund=(Fund.FUNDID) where PB.clientID='" & _
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
    szSQL = "SELECT Distinct FundID, FundName, FundCode,CategoryCode FROM Fund LEFT JOIN tlbPayable PB ON PB.Pay_fund=(Fund.FUNDID) where PB.clientID='" & _
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
      .ColWidth(2) = 3800 'Label2(2).Left - Label2(1).Left 'Property Name
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




Private Sub txtRentPayable_KeyPress(KeyAscii As Integer)
    DigitTextKeyPress txtRentPayable, KeyAscii
End Sub



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

