VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmBkRecHistReport 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Bank Reconciliation History Report"
   ClientHeight    =   5775
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13560
   BeginProperty Font 
      Name            =   "Myriad Web"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmBkRecHistReport.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5775
   ScaleWidth      =   13560
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraEnterDates 
      Caption         =   "Transaction Date:"
      Height          =   2055
      Left            =   1320
      TabIndex        =   25
      Top             =   2640
      Width           =   3645
      Begin VB.TextBox txtStClosingBal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1560
         TabIndex        =   31
         Top             =   1440
         Width           =   1935
      End
      Begin VB.TextBox txtFromDate 
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   1560
         MaxLength       =   10
         TabIndex        =   27
         Text            =   "01/01/2000"
         Top             =   360
         Width           =   1935
      End
      Begin VB.TextBox txtToDate 
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   1560
         MaxLength       =   10
         TabIndex        =   26
         Top             =   900
         Width           =   1935
      End
      Begin VB.Label lblStClosingBal 
         BackStyle       =   0  'Transparent
         Caption         =   "Statement Balance:"
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label lblSpecifyDateRange 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "From Date:"
         Height          =   195
         Left            =   120
         TabIndex        =   29
         Top             =   360
         Width           =   765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To Date:"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   28
         Top             =   900
         Width           =   585
      End
   End
   Begin VB.Frame fraOptions 
      Caption         =   "Reconciliation Type:"
      Height          =   2055
      Left            =   120
      TabIndex        =   24
      Top             =   120
      Width           =   1845
      Begin MSForms.OptionButton optUnReconBoth 
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   1440
         Width           =   735
         VariousPropertyBits=   746588179
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DisplayStyle    =   5
         Size            =   "1296;450"
         Value           =   "0"
         Caption         =   "Both"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.OptionButton optReconciled 
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   960
         Width           =   1215
         VariousPropertyBits=   746588179
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DisplayStyle    =   5
         Size            =   "2143;450"
         Value           =   "0"
         Caption         =   "Reconciled"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.OptionButton optUnreconciled 
         Height          =   255
         Left            =   240
         TabIndex        =   0
         Top             =   480
         Width           =   1335
         VariousPropertyBits=   746588179
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DisplayStyle    =   5
         Size            =   "2355;450"
         Value           =   "1"
         Caption         =   "Unreconciled"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
   Begin VB.CheckBox chkDetails 
      Caption         =   "Detailed Transaction"
      Height          =   255
      Left            =   1560
      TabIndex        =   23
      Top             =   5280
      Width           =   2295
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
      Height          =   2460
      Left            =   6000
      ScaleHeight     =   2430
      ScaleWidth      =   5520
      TabIndex        =   20
      Top             =   240
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
         TabIndex        =   21
         Top             =   10
         Width           =   295
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxClientList 
         Height          =   2295
         Left            =   45
         TabIndex        =   22
         Top             =   75
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   4048
         _Version        =   393216
         Appearance      =   0
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
      End
   End
   Begin VB.TextBox txtLlName 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   8640
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   4320
      Width           =   3375
   End
   Begin VB.Frame fraFunds 
      Caption         =   "Funds:"
      Height          =   1935
      Left            =   120
      TabIndex        =   19
      Top             =   3240
      Width           =   5565
      Begin VB.CheckBox chkFunds 
         Caption         =   "All Funds"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   1095
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxFunds 
         Height          =   1335
         Left            =   120
         TabIndex        =   11
         Top             =   525
         Width           =   5320
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
      Left            =   6000
      TabIndex        =   18
      Top             =   2880
      Width           =   5565
      Begin VB.CheckBox chkBankAccounts 
         Caption         =   "All Accounts"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   1215
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxBankAccounts 
         Height          =   1335
         Left            =   120
         TabIndex        =   9
         Top             =   525
         Width           =   5320
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
   Begin VB.Frame fraSelectDates 
      Caption         =   "Statement Date:"
      Height          =   2055
      Left            =   2040
      TabIndex        =   15
      Top             =   120
      Width           =   3645
      Begin MSForms.ComboBox cboLtStDt 
         Height          =   315
         Left            =   1680
         TabIndex        =   4
         Top             =   1140
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
      Begin VB.Label lblLast 
         BackStyle       =   0  'Transparent
         Caption         =   "Last Statement Date: "
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   1140
         Width           =   1455
      End
      Begin MSForms.ComboBox cboCurStDt 
         Height          =   315
         Left            =   1680
         TabIndex        =   3
         Top             =   600
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
         Caption         =   "Current Date:"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   16
         Top             =   600
         Width           =   1095
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
      Left            =   12000
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4320
      Width           =   285
   End
   Begin VB.TextBox txtClientID 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   7560
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   4320
      Width           =   1090
   End
   Begin VB.CommandButton cmdGenReport 
      Caption         =   "&Generate Report"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   2280
      Width           =   1575
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   375
      Left            =   4080
      TabIndex        =   6
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFDFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Client:"
      Height          =   195
      Index           =   3
      Left            =   6720
      TabIndex        =   14
      Top             =   4320
      Width           =   465
   End
End
Attribute VB_Name = "frmBkRecHistReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Samrat Rahman: On 07/02/2012
'First this form was created for Landlord Summary Statement.
'We are changing this form for Bank Reconciliation History Report.
'This refer to the modification document on page xxx: Bank Reconciliation SAVING the closing balance.

Option Explicit

Private szDemandTypes      As String
Private bCallingFromGrid   As Boolean
'Private szBanks            As String
Private szFunds            As String
'Private szFundList         As String
Private Const INI_HEIGHT   As Integer = 3315
Private cBBF               As Currency
Private ReConDates()       As String
Private dtLastStDate       As Date
Private bX                 As Boolean

Public szClientID          As String
Public szBankID            As String

Private Sub cboCurStDt_Click()
   Dim i As Integer

   If cboCurStDt.text = "" Then Exit Sub
   cboLtStDt.Clear
   cboLtStDt.Column() = ReConDates()

   If optUnReconBoth.Value Then txtToDate.text = cboCurStDt.text

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

Private Sub chkBankAccounts_Click()
   If bCallingFromGrid Then
      bCallingFromGrid = False
      Exit Sub
   End If
   
   Dim iRow As Integer

   For iRow = 1 To flxBankAccounts.Rows - 1
      If flxBankAccounts.RowHeight(iRow) > 0 And flxBankAccounts.TextMatrix(iRow, 0) = "X" Then
         SelectFlxGridRow 0, flxBankAccounts, iRow
      End If
   Next iRow
   
   For iRow = 1 To flxBankAccounts.Rows - 1
      If flxBankAccounts.RowHeight(iRow) > 0 And chkBankAccounts.Value Then
         SelectFlxGridRow 0, flxBankAccounts, iRow
      End If
   Next iRow
End Sub

Private Sub chkFunds_Click()
   If bCallingFromGrid Then
      bCallingFromGrid = False
      Exit Sub
   End If

   Dim iRow As Integer

   For iRow = 1 To flxFunds.Rows - 1
      If flxFunds.RowHeight(iRow) > 0 And flxFunds.TextMatrix(iRow, 0) = "X" Then
         SelectFlxGridRow 0, flxFunds, iRow
      End If
   Next iRow

   For iRow = 1 To flxFunds.Rows - 1
      If flxFunds.RowHeight(iRow) > 0 And chkFunds.Value Then
         SelectFlxGridRow 0, flxFunds, iRow
      End If
   Next iRow
End Sub

Private Sub cmdClient_Click()
   picClientList.Top = txtClientID.Top + txtClientID.Height + 5
   picClientList.Left = Label1(3).Left + 5
   picClientList.Visible = True
   picClientList.ZOrder 0
   Me.Height = 3360
End Sub

Private Sub cmdClose_Click()
   Unload Me
   frmCashbook.Enabled = True
End Sub

Private Sub cmdGenReport_Click()
  Dim fReport As New frmReport
   If (optUnreconciled.Value Or optUnReconBoth.Value) And txtToDate.text = "" Then
      MsgBox "Please enter the To Date", vbCritical + vbOKOnly, "To Date"
      txtToDate.SetFocus
      Exit Sub
   End If
   If (optUnreconciled.Value Or optUnReconBoth.Value) And txtStClosingBal.text = "" Then
      MsgBox "Please enter the statement closing balance", vbCritical + vbOKOnly, "Closing Balance"
      txtStClosingBal.SetFocus
      Exit Sub
   End If

   If optReconciled.Value And cboCurStDt.text = "" Then
      MsgBox "Please select current statement date.", vbInformation + vbOKOnly, "From Date"
      cboCurStDt.SetFocus
      Exit Sub
   End If

   If optReconciled.Value And cboLtStDt.text = "" Then
      If MsgBox("Do you wish to leave blank last statement date?", vbInformation + vbYesNo, "To Date") = vbYes Then
         dtLastStDate = Format(#1/1/2000#, "dd/mm/yyyy")
      Else
         cboLtStDt.SetFocus
         Exit Sub
      End If
   End If
'
'   szBanks = ""
'   If Not IsBankSelected Then
'      MsgBox "Please select a bank account.", vbInformation + vbOKOnly, "Bank Account"
'      chkBankAccounts.SetFocus
'      Exit Sub
'   End If
   
'   szFunds = ""
'   szFundList = ""
'   If Not IsFundSelected Then
'      MsgBox "Please select a fund.", vbInformation + vbOKOnly, "Funds"
'      chkFunds.SetFocus
'      Exit Sub
'   End If

   Dim reportApp As New CRAXDRT.Application
   Dim Report As CRAXDRT.Report
   Dim rep As frmReport
   Dim adoConn As New ADODB.Connection
   Dim zVal As Double

   adoConn.Open getConnectionString
   If frmCashbook.txtClientList.text = "Consolidated" Then
        If optUnreconciled.Value Then              'Show Consolidated unreconciled transactions onl
              
              'Resolved By BOSL. Added By Asif. Issue: 0000523. Date: 21-02-2015
              'If the balance is empty, prompt the user whether the program should load last statement balance
              'or let the user enter the project balance
              If txtStClosingBal.text = "" Then
                 txtToDate_LostFocus
              End If
              
              Set Report = reportApp.OpenReport(App.Path & szReportPath & "\CB_Hist_UnRecon_Cons.rpt")
        
              Report.Database.Tables(1).ConnectionProperties.Item("Database Password") = accessDBPws
        
              Report.EnableParameterPrompting = False
              Report.DiscardSavedData
        
              Report.ParameterFields(1).AddCurrentValue frmCashbook.SelectedConBankID
              ' modified by anol 20230608
              
             Dim dtReqDate As Date
             dtReqDate = txtToDate.text
             
              'Report.ParameterFields(2).AddCurrentValue AccountBalanceDatedConsolidated(adoConn, txtToDate.text, frmCashbook.SelectedConBankID)
              Dim dblABC As Double
              dblABC = BankAccBalanceConsolidated(adoConn, dtReqDate, CInt(frmCashbook.SelectedConBankID))
              Report.ParameterFields(2).AddCurrentValue dblABC
              Report.ParameterFields(3).AddCurrentValue CDate(txtToDate.text)
              Report.ParameterFields(4).AddCurrentValue CDbl(txtStClosingBal.text)
              Report.ParameterFields(5).AddCurrentValue frmCashbook.txtClientList.Tag
              
              '*********
                      
'              Report.ParameterFields(1).AddCurrentValue frmCashbook.txtBC.Tag 'Passing bank code
'              zVal = AccountBalanceDatedConsolidated(adoConn, txtToDate.text, frmCashbook.SelectedConBankID)
'              Report.ParameterFields(2).AddCurrentValue zVal
'              Report.ParameterFields(3).AddCurrentValue 0 ' no need opening bal Format(frmCashbook.txtStOpenBal.text, "0.00") 'Opening statement Balance
'              Report.ParameterFields(4).AddCurrentValue CDate(txtToDate.text) 'statement date
'              Report.ParameterFields(5).AddCurrentValue CDbl(txtStClosingBal.text)
'              Report.ParameterFields(6).AddCurrentValue CDbl(txtStClosingBal.text) 'This is unrecon bal
'              Report.ParameterFields(7).AddCurrentValue frmCashbook.SelectedConBankID
               
               
        
              Set rep = New frmReport
              Load rep
              rep.LoadReportViewer Report
              GoTo CloseConnection          'Exit sub
           End If

           If optReconciled.Value Then              'Show RECONCILED transactions only
              If cboLtStDt.text = "" Then
                 If MsgBox("Do you wish to run this report with a blank last reconciled statement date?", vbQuestion + vbYesNo, "To Date") = vbYes Then
                    dtLastStDate = Format(#1/1/2000#, "dd/mm/yyyy")
                 Else
                    cboLtStDt.SetFocus
                    GoTo CloseConnection          'Exit sub
                 End If
              Else
                 dtLastStDate = CDate(cboLtStDt.text)
              End If

              Set Report = reportApp.OpenReport(App.Path & szReportPath & "\CB_HistRecon_Cons.rpt")

              Report.EnableParameterPrompting = False
              Report.DiscardSavedData

              Report.ParameterFields(1).AddCurrentValue CDate(dtLastStDate)
              Report.ParameterFields(2).AddCurrentValue CDate(cboCurStDt.text)
              Report.ParameterFields(3).AddCurrentValue CInt(frmCashbook.SelectedConBankID) 'pass consolidated Bank Id here
              'we are not using ClosingBalance parameter int he report. which is parameter no 4
              
        '      Report.ParameterFields(4).AddCurrentValue StatementClosingBalance(adoConn, cboCurStDt.text, frmCashbook.txtBC.Tag)
           'Modified by anol 20 Jan 2015
            'issue 523
              'Report.ParameterFields(4).AddCurrentValue CStr(frmCashbook.cboClientID.Value)
               'modified by anol 20230608
             'Dim dtReqDate As Date
             dtReqDate = cboCurStDt.text
             'Dim dblABC As Double
             dblABC = BankAccBalanceConsolidated(adoConn, dtReqDate, frmCashbook.SelectedConBankID)
              Report.ParameterFields(5).AddCurrentValue dblABC
              If cboLtStDt.text = "" Then
                 Report.ParameterFields(6).AddCurrentValue 0
              Else
                    'If I send a date range 16/05/2023 -15/05/2023. It sends date  5/05/2023( cboLtStDt.text)
                 Report.ParameterFields(6).AddCurrentValue StatementOpeningBalance(adoConn, cboLtStDt.text, frmCashbook.txtAccountName.text, frmCashbook.txtClientList.Tag)
              End If
              'issue 523
              'Modified by anol 20 Jan 2015
              Report.ParameterFields(7).AddCurrentValue frmCashbook.txtClientList.Tag

              Load fReport
              fReport.LoadReportViewer Report
           End If

           If optUnReconBoth.Value Then              'Show unreconciled & reconciled transactions BOTH
              If txtToDate.text = "" Then Exit Sub

              'Resolved By BOSL. Added By Asif. Issue: 0000523. Date: 21-02-2015
              'If the balance is empty, prompt the user whether the program should load last statement balance
              'or let the user enter the project balance

              If txtStClosingBal.text = "" Then
                 txtToDate_LostFocus
              End If

              Set Report = reportApp.OpenReport(App.Path & szReportPath & "\CB_Hist_Both_Cons.rpt")

              Report.Database.Tables(1).ConnectionProperties.Item("Database Password") = accessDBPws

              Report.EnableParameterPrompting = False
              Report.DiscardSavedData

              Report.ParameterFields(1).AddCurrentValue frmCashbook.SelectedConBankID
              'issue 523
              'modified by anol 18 Feb 2015
              'Report.ParameterFields(2).AddCurrentValue AccountBalanceDated(adoConn, txtToDate.text, frmCashbook.txtBC.Tag, frmCashbook.txtClientList.Tag)
               'modified by anol 20230608
             ' Report.ParameterFields(2).AddCurrentValue AccountBalanceDatedConsolidated(adoConn, txtToDate.text, frmCashbook.SelectedConBankID)

             'Dim dtReqDate As Date
             dtReqDate = txtToDate.text
             'Dim dblABC As Double
             dblABC = BankAccBalanceConsolidated(adoConn, dtReqDate, frmCashbook.SelectedConBankID)
             Report.ParameterFields(2).AddCurrentValue dblABC
              Report.ParameterFields(3).AddCurrentValue Format(frmCashbook.txtStOpenBal.text, "0.00")
             ' Report.ParameterFields(4).AddCurrentValue CDate(cboCurStDt.text)
              Report.ParameterFields(5).AddCurrentValue CDbl(frmCashbook.Label1(27).Caption)
              'issue 523
              'modified by anol 20 Jan 2015
              'Val(frmCashbook.txtAcBal.text)  '
              'Val(txtStClosingBal.text) '

              'Resolved By BOSL. Modified By Asif. Issue: 0000523. Date: 21-02-2015
              'Should pass the value from the closing balance text box as it could be the last closing balance before
              'the date entered or the projected balance entered by the user.

              'Report.ParameterFields(6).AddCurrentValue StatementClosingBalance(adoConn, txtToDate.text, frmCashbook.txtBC.Tag, frmCashbook.cboClientID.Value)
              Report.ParameterFields(6).AddCurrentValue CDbl(txtStClosingBal.text)


              If txtFromDate.text = "" Then
                 Report.ParameterFields(7).AddCurrentValue CDate("01/01/2000")
              Else
                 Report.ParameterFields(7).AddCurrentValue CDate(txtFromDate.text)
              End If
              Report.ParameterFields(8).AddCurrentValue CDate(txtToDate.text)
              Report.ParameterFields(9).AddCurrentValue frmCashbook.txtClientList.Tag
              Set rep = New frmReport
              Load rep
              rep.LoadReportViewer Report
              GoTo CloseConnection          'Exit sub
           End If
   Else 'Non consolidated and normal report part
           If optUnreconciled.Value Then              'Show unreconciled transactions onl
              
              'Resolved By BOSL. Added By Asif. Issue: 0000523. Date: 21-02-2015
              'If the balance is empty, prompt the user whether the program should load last statement balance
              'or let the user enter the project balance
              If txtStClosingBal.text = "" Then
                 txtToDate_LostFocus
              End If
              
              Set Report = reportApp.OpenReport(App.Path & szReportPath & "\CB_Hist_UnRecon.rpt")
        
              Report.Database.Tables(1).ConnectionProperties.Item("Database Password") = accessDBPws
        
              Report.EnableParameterPrompting = False
              Report.DiscardSavedData
        
              Report.ParameterFields(1).AddCurrentValue frmCashbook.txtBC.Tag
              'Below line is modified by anol 19 Feb 2015 issue 530
              Report.ParameterFields(2).AddCurrentValue AccountBalanceDated(adoConn, txtToDate.text, frmCashbook.txtBC.Tag, frmCashbook.txtClientList.Tag)
              Report.ParameterFields(3).AddCurrentValue CDate(txtToDate.text)
              
              'Resolved By BOSL. Modified By Asif. Issue: 0000523. Date: 21-02-2015
              'Should pass the value from the closing balance text box as it could be the last closing balance before
              'the date entered or the projected balance entered by the user.
              
              'Report.ParameterFields(4).AddCurrentValue StatementClosingBalance(adoConn, txtToDate.text, frmCashbook.txtBC.Tag, frmCashbook.cboClientID.Value)
              Report.ParameterFields(4).AddCurrentValue CDbl(txtStClosingBal.text)
              
              'issue 523
              'Val(frmCashbook.txtAcBal.text) '
              'CDbl(txtStClosingBal.text) '
              'Modified by anol 20 Jan 2015
              Report.ParameterFields(5).AddCurrentValue frmCashbook.txtClientList.Tag
        
              Set rep = New frmReport
              Load rep
              rep.LoadReportViewer Report
              GoTo CloseConnection          'Exit sub
           End If
        
           If optReconciled.Value Then              'Show RECONCILED transactions only
              If cboLtStDt.text = "" Then
                 If MsgBox("Do you wish to run this report with a blank last reconciled statement date?", vbQuestion + vbYesNo, "To Date") = vbYes Then
                    dtLastStDate = Format(#1/1/2000#, "dd/mm/yyyy")
                 Else
                    cboLtStDt.SetFocus
                    GoTo CloseConnection          'Exit sub
                 End If
              Else
                 dtLastStDate = CDate(cboLtStDt.text)
              End If
        
              Set Report = reportApp.OpenReport(App.Path & szReportPath & "\CB_HistRecon.rpt")
        
              Report.EnableParameterPrompting = False
              Report.DiscardSavedData
        
              Report.ParameterFields(1).AddCurrentValue CDate(dtLastStDate)
              Report.ParameterFields(2).AddCurrentValue CDate(cboCurStDt.text)
              Report.ParameterFields(3).AddCurrentValue CInt(frmCashbook.cmdBC.Tag)
        '      Report.ParameterFields(4).AddCurrentValue StatementClosingBalance(adoConn, cboCurStDt.text, frmCashbook.txtBC.Tag)
           'Modified by anol 20 Jan 2015
            'issue 523
              'Report.ParameterFields(4).AddCurrentValue CStr(frmCashbook.cboClientID.Value)
              Report.ParameterFields(5).AddCurrentValue AccountBalanceDated(adoConn, cboCurStDt.text, frmCashbook.txtBC.Tag, frmCashbook.txtClientList.Tag)
              If cboLtStDt.text = "" Then
                 Report.ParameterFields(6).AddCurrentValue 0
              Else
                 Report.ParameterFields(6).AddCurrentValue StatementOpeningBalance(adoConn, cboLtStDt.text, frmCashbook.txtBC.Tag, frmCashbook.txtClientList.Tag)
              End If
              'issue 523
              'Modified by anol 20 Jan 2015
              Report.ParameterFields(7).AddCurrentValue frmCashbook.txtClientList.Tag
             
              Load fReport
              fReport.LoadReportViewer Report
           End If
        
           If optUnReconBoth.Value Then              'Show unreconciled & reconciled transactions BOTH
              If txtToDate.text = "" Then Exit Sub
        
              'Resolved By BOSL. Added By Asif. Issue: 0000523. Date: 21-02-2015
              'If the balance is empty, prompt the user whether the program should load last statement balance
              'or let the user enter the project balance
              
              If txtStClosingBal.text = "" Then
                 txtToDate_LostFocus
              End If
              
              Set Report = reportApp.OpenReport(App.Path & szReportPath & "\CB_Hist_Both.rpt")
        
              Report.Database.Tables(1).ConnectionProperties.Item("Database Password") = accessDBPws
        
              Report.EnableParameterPrompting = False
              Report.DiscardSavedData
        
              Report.ParameterFields(1).AddCurrentValue frmCashbook.txtBC.Tag
              'issue 523
              'modified by anol 18 Feb 2015
              Report.ParameterFields(2).AddCurrentValue AccountBalanceDated(adoConn, txtToDate.text, frmCashbook.txtBC.Tag, frmCashbook.txtClientList.Tag)
              Report.ParameterFields(3).AddCurrentValue Format(frmCashbook.txtStOpenBal.text, "0.00")
             ' Report.ParameterFields(4).AddCurrentValue CDate(cboCurStDt.text)
              Report.ParameterFields(5).AddCurrentValue CDbl(frmCashbook.Label1(27).Caption)
              'issue 523
              'modified by anol 20 Jan 2015
              'Val(frmCashbook.txtAcBal.text)  '
              'Val(txtStClosingBal.text) '
              
              'Resolved By BOSL. Modified By Asif. Issue: 0000523. Date: 21-02-2015
              'Should pass the value from the closing balance text box as it could be the last closing balance before
              'the date entered or the projected balance entered by the user.
              
              'Report.ParameterFields(6).AddCurrentValue StatementClosingBalance(adoConn, txtToDate.text, frmCashbook.txtBC.Tag, frmCashbook.cboClientID.Value)
              Report.ParameterFields(6).AddCurrentValue CDbl(txtStClosingBal.text)
              
              
              If txtFromDate.text = "" Then
                 Report.ParameterFields(7).AddCurrentValue CDate("01/01/2000")
              Else
                 Report.ParameterFields(7).AddCurrentValue CDate(txtFromDate.text)
              End If
              Report.ParameterFields(8).AddCurrentValue CDate(txtToDate.text)
              Report.ParameterFields(9).AddCurrentValue frmCashbook.txtClientList.Tag
              Set rep = New frmReport
              Load rep
              rep.LoadReportViewer Report
              GoTo CloseConnection          'Exit sub
           End If
End If
CloseConnection:
   adoConn.Close
   Set adoConn = Nothing
End Sub
Private Function StatementOpeningBalanceConsolidated(adoConn As ADODB.Connection, dtLstStDate As Date, szBankAc As String, ClientID As String) As Double
'Resolved by BOSL
'added a parameter client ID 20 Jan 2015
'issue 523 opeing balance is not showing for specific client in the report
   Dim szSQL   As String
   Dim adoRst  As New ADODB.Recordset

   szSQL = "SELECT ProjClBal " & _
           "FROM   tlbBankReconClosingBal AS C " & _
           "WHERE  StatementDate = #" & Format(dtLstStDate, "dd mmmm yyyy") & "# AND " & _
                  "BankCode = '" & szBankAc & "' And clientID='" & ClientID & "' ;"
   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If Not adoRst.EOF Then
      StatementOpeningBalanceConsolidated = CDbl(adoRst.Fields.Item("ProjClBal").Value)
   Else
      StatementOpeningBalanceConsolidated = 0
   End If

   adoRst.Close
   Set adoRst = Nothing
End Function
Private Function BankAccBalanceConsolidated(adoConn As ADODB.Connection, ByVal dtReqDate As Date, ByVal SelectedConBankID As Integer) As Currency
  
   'dtReqDate = Date
   Dim szSQL As String
   Dim adoRst As New ADODB.Recordset

   szSQL = "SELECT SUM(R.Amount) AS AMT, Type AS T " & _
           "FROM tlbReceipt AS R, " & _
                "Units AS U, Property AS P, tlbClientBanks AS B " & _
           "WHERE R.BankCode =B.NominalCode  AND B.ConsolidatedBankID =  " & SelectedConBankID & " AND " & _
                 "U.UnitNumber = R.UnitID AND U.PropertyID = P.PropertyID AND " & _
                 "B.NominalCode = R.BankCode AND R.RDate <= #" & Format(dtReqDate, "dd mmmm yyyy") & "#   AND " & _
                 "B.CLIENT_ID = P.ClientID AND " & _
                 "R.Amount > 0 " & _
           "GROUP BY Type " & _
           "UNION "
   szSQL = szSQL & _
           "SELECT SUM(P.Amount) AS AMT, Type AS T " & _
           "FROM tlbPayment AS P, tlbClientBanks AS B " & _
           "WHERE P.BankCode = B.NominalCode AND B.Client_ID=P.clientID AND B.ConsolidatedBankID =" & SelectedConBankID & " AND " & _
                 "B.NominalCode = P.BankCode AND  P.PDate <= #" & Format(dtReqDate, "dd mmmm yyyy") & "#  AND " & _
                 "P.Amount > 0 " & _
           "GROUP BY TYPE " & _
           "UNION "
   szSQL = szSQL & _
           "SELECT SUM (BP.NET_AMOUNT + BP.VAT) AS AMT, TRANS AS T " & _
           "FROM tlbBankPayment AS BP, tlbClientBanks AS CB " & _
           "WHERE BP.BANK_AC = CB.NominalCode AND CB.Client_ID=BP.clientID AND CB.ConsolidatedBankID =" & SelectedConBankID & "  AND " & _
                  "CB.NominalCode = BP.BANK_AC  AND  BP.TRAN_DATE <= #" & Format(dtReqDate, "dd mmmm yyyy") & "#  AND " & _
               "(BP.NET_AMOUNT + BP.VAT) > 0 " & _
           "GROUP BY TRANS " & _
           "ORDER BY T;"
'Debug.Print szSQL
   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   While Not adoRst.EOF
      If adoRst.Fields.Item("T").Value = "3" Then _
         BankAccBalanceConsolidated = BankAccBalanceConsolidated + adoRst.Fields.Item("AMT").Value
      If adoRst.Fields.Item("T").Value = "4" Then _
         BankAccBalanceConsolidated = BankAccBalanceConsolidated + adoRst.Fields.Item("AMT").Value
      If adoRst.Fields.Item("T").Value = "8" Then _
         BankAccBalanceConsolidated = BankAccBalanceConsolidated - adoRst.Fields.Item("AMT").Value
      If adoRst.Fields.Item("T").Value = "9" Then _
         BankAccBalanceConsolidated = BankAccBalanceConsolidated - adoRst.Fields.Item("AMT").Value
      If adoRst.Fields.Item("T").Value = "BP" Then _
         BankAccBalanceConsolidated = BankAccBalanceConsolidated - adoRst.Fields.Item("AMT").Value
      If adoRst.Fields.Item("T").Value = "BR" Then _
         BankAccBalanceConsolidated = BankAccBalanceConsolidated + adoRst.Fields.Item("AMT").Value
      If adoRst.Fields.Item("T").Value = "23" Then _
         BankAccBalanceConsolidated = BankAccBalanceConsolidated - adoRst.Fields.Item("AMT").Value
      If adoRst.Fields.Item("T").Value = "24" Then _
         BankAccBalanceConsolidated = BankAccBalanceConsolidated + adoRst.Fields.Item("AMT").Value

      adoRst.MoveNext
   Wend

   adoRst.Close
   Set adoRst = Nothing
End Function
'Private Function AccountBalanceDatedConsolidated(ByVal adoConn As adodb.Connection, ByVal dtReqDate As Date, _
'                                    ByVal szConBankID As Long) As Double
'   Dim szSQL   As String
'   Dim cR      As Currency
'   Dim cP      As Currency
'   Dim cB      As Currency
'   Dim adoRst  As New adodb.Recordset
'
''----------------------------------------------  SR & RoA
'   szSQL = "SELECT SUM(S.Amount) AS T " & _
'           "FROM   tlbReceipt AS R, tlbReceiptSplit AS S, " & _
'                  "Units AS U, Property AS P, tlbClientBanks AS B " & _
'           "WHERE  R.TransactionID = S.RptHeader AND " & _
'                  "R.RDate <= #" & Format(dtReqDate, "dd mmmm yyyy") & "# AND " & _
'                  "R.BankCode = B.NominalCode AND " & _
'                  "U.UnitNumber = R.UnitID AND U.PropertyID = P.PropertyID AND " & _
'                  "B.NominalCode = R.BankCode AND " & _
'                  "B.CLIENT_ID = P.ClientID AND B.ConsolidatedBankID=" & szConBankID & " AND " & _
'                  "R.TYPE IN (3, 4);"
'   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'
'   If Not adoRst.EOF Then
'      cR = IIf(IsNull(adoRst.Fields.Item(0).Value), 0, adoRst.Fields.Item(0).Value)
'   Else
'      cR = 0
'   End If
'   adoRst.Close
''----------------------------------------------  SRR
'   szSQL = "SELECT SUM(S.Amount) AS T " & _
'           "FROM   tlbReceipt AS R, tlbReceiptSplit AS S, " & _
'                  "Units AS U, Property AS P, tlbClientBanks AS B " & _
'           "WHERE  R.TransactionID = S.RptHeader AND " & _
'                  "R.RDate <= #" & Format(dtReqDate, "dd mmmm yyyy") & "# AND " & _
'                  "R.BankCode =  B.NominalCode  AND " & _
'                  "U.UnitNumber = R.UnitID AND U.PropertyID = P.PropertyID AND " & _
'                  "B.ConsolidatedBankID = " & szConBankID & " AND B.NominalCode = R.BankCode AND " & _
'                  "B.CLIENT_ID = P.ClientID AND " & _
'                  "R.TYPE = 23;"
''Debug.Print szSQL
'   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'
'   If Not adoRst.EOF Then cR = cR - IIf(IsNull(adoRst.Fields.Item(0).Value), 0, adoRst.Fields.Item(0).Value)
'   adoRst.Close
''----------------------------------------------  PP & PoA
'   szSQL = "SELECT SUM(S.Amount) AS T " & _
'           "FROM   tlbPayment AS P, tlbPaymentSplit AS S, tlbClientBanks AS B " & _
'           "WHERE  P.TransactionID = S.PayHeader AND " & _
'                  "P.PDate <= #" & Format(dtReqDate, "dd mmmm yyyy") & "# AND " & _
'                  "P.BankCode = B.NominalCode AND " & _
'                  "B.ConsolidatedBankID = " & szConBankID & " AND " & _
'                  "B.NominalCode = P.BankCode AND B.CLIENT_ID = P.ClientID AND " & _
'                  "P.TYPE IN (8, 9);"
''Debug.Print szSQL
'   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'
'   If Not adoRst.EOF Then
'      cP = IIf(IsNull(adoRst.Fields.Item(0).Value), 0, adoRst.Fields.Item(0).Value)
'   Else
'      cP = 0
'   End If
'   adoRst.Close
''----------------------------------------------  PPR
'   szSQL = "SELECT SUM(S.Amount) AS T " & _
'           "FROM   tlbPayment AS P, tlbPaymentSplit AS S, tlbClientBanks AS B " & _
'           "WHERE  P.TransactionID = S.PayHeader AND " & _
'                  "P.PDate <= #" & Format(dtReqDate, "dd mmmm yyyy") & "# AND " & _
'                  "P.BankCode = B.NominalCode  AND " & _
'                  "B.ConsolidatedBankID = " & szConBankID & " AND " & _
'                  "B.NominalCode = P.BankCode AND B.CLIENT_ID = P.ClientID AND " & _
'                  "P.TYPE = 24;"
'   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'
'   If Not adoRst.EOF Then cP = cP - IIf(IsNull(adoRst.Fields.Item(0).Value), 0, adoRst.Fields.Item(0).Value)
'   adoRst.Close
''----------------------------------------------  BR
'   szSQL = "SELECT SUM(B.NET_AMOUNT + B.VAT) AS T " & _
'           "FROM   tlbBankPayment AS B, tlbClientBanks AS CB " & _
'           "WHERE  B.TRAN_DATE <= #" & Format(dtReqDate, "dd mmmm yyyy") & "# AND " & _
'                  "B.BANK_AC = CB.NominalCode AND " & _
'                  "CB.ConsolidatedBankID = " & szConBankID & " AND " & _
'                  "CB.NominalCode = B.BANK_AC AND CB.CLIENT_ID = B.ClientID AND " & _
'                  "B.TransactionType = 12;"
'   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'
'   If Not adoRst.EOF Then
'      cB = IIf(IsNull(adoRst.Fields.Item(0).Value), 0, adoRst.Fields.Item(0).Value)
'   Else
'      cB = 0
'   End If
'   adoRst.Close
''----------------------------------------------  BP
'   szSQL = "SELECT SUM(B.NET_AMOUNT + B.VAT) AS T " & _
'           "FROM   tlbBankPayment AS B, tlbClientBanks AS CB " & _
'           "WHERE  B.TRAN_DATE <= #" & Format(dtReqDate, "dd mmmm yyyy") & "# AND " & _
'                  "B.BANK_AC =  CB.NominalCode AND " & _
'                  "CB.ConsolidatedBankID = " & szConBankID & " AND " & _
'                  "CB.NominalCode = B.BANK_AC AND CB.CLIENT_ID = B.ClientID AND " & _
'                  "B.TransactionType = 11;"
'   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'
'   If Not adoRst.EOF Then cB = cB - IIf(IsNull(adoRst.Fields.Item(0).Value), 0, adoRst.Fields.Item(0).Value)
'   adoRst.Close
''######################################################################
'
'   AccountBalanceDatedConsolidated = cR - cP + cB
'
'   Set adoRst = Nothing
'End Function
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
           "WHERE  RC.RefID = P.szTransactionID AND " & _
                  "P.TransactionID = S.PayHeader AND " & _
                  "S.FundID = F.FundID AND " & _
                  "P.BankCode = CB.NominalCode AND " & _
                  "RC.ReconDate < #" & Format(dtLastStDate, "dd mmmm yyyy") & "# AND " & _
                  "F.SelFund = 'Y' AND CB.SelBanks = 'Y' AND " & _
                  "P.TYPE IN (8, 9);"
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
           "WHERE  RC.RefID = P.szTransactionID AND " & _
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
   Set adoRst = Nothing
End Sub

Private Sub flxBankAccounts_Click()
   If flxBankAccounts.row = 0 Then Exit Sub

   SelectFlxGridRow 0, flxBankAccounts, flxBankAccounts.row

   Dim iRow As Integer

   For iRow = 1 To flxBankAccounts.Rows - 1
      If flxBankAccounts.TextMatrix(iRow, 0) <> "X" And chkBankAccounts.Value Then
         bCallingFromGrid = True
         chkBankAccounts.Value = 0
         Exit For
      End If
   Next iRow
End Sub

Private Sub flxClientList_Click()
   If flxClientList.TextMatrix(flxClientList.row, 1) = "" Then Exit Sub

   txtClientID.text = flxClientList.TextMatrix(flxClientList.row, 1)
   txtLlName.text = flxClientList.TextMatrix(flxClientList.row, 2)

   picClientList.Visible = False
   Me.Height = INI_HEIGHT

   fraBankAccounts.Enabled = True
   fraFunds.Enabled = True

   FilteringBanks
End Sub

Private Sub FilteringDates()
   Dim iCmb As Integer
   Dim iRow As Integer

   cboCurStDt.Clear
   cboCurStDt.Column() = ReConDates()
   cboLtStDt.Clear
   cboLtStDt.Column() = ReConDates()
   For iRow = 1 To flxBankAccounts.Rows - 1
      If flxBankAccounts.RowHeight(iRow) = 0 Then
         iCmb = 0
         Do While (iCmb <= cboCurStDt.ListCount - 1)
            If cboCurStDt.Column(1, iCmb) = flxBankAccounts.TextMatrix(iRow, 2) Then
               cboCurStDt.RemoveItem (iCmb)
               iCmb = iCmb - 1
            End If
            iCmb = iCmb + 1
         Loop

         iCmb = 0
         Do While (iCmb <= cboLtStDt.ListCount - 1)
            If cboLtStDt.Column(1, iCmb) = flxBankAccounts.TextMatrix(iRow, 2) Then
               cboLtStDt.RemoveItem (iCmb)
               iCmb = iCmb - 1
            End If
            iCmb = iCmb + 1
         Loop
      End If
   Next iRow
End Sub

Private Sub FilteringBanks()
   Dim iRow As Integer

   For iRow = 1 To flxBankAccounts.Rows - 1
      flxBankAccounts.RowHeight(iRow) = 240
   Next iRow

   For iRow = 1 To flxBankAccounts.Rows - 1
      If flxBankAccounts.TextMatrix(iRow, 4) <> txtClientID.text Then
         flxBankAccounts.RowHeight(iRow) = 0
      End If
   Next iRow

   FilteringDates
End Sub

Private Sub flxFunds_Click()
   If flxFunds.row = 0 Then Exit Sub

   SelectFlxGridRow 0, flxFunds, flxFunds.row

   Dim iRow As Integer

   For iRow = 1 To flxFunds.Rows - 1
      If flxFunds.TextMatrix(iRow, 0) <> "X" And chkFunds.Value Then
         bCallingFromGrid = True
         chkFunds.Value = 0
         Exit For
      End If
   Next iRow
End Sub

Private Sub Form_Activate()
   optUnreconciled_Click
   fraEnterDates.Left = fraSelectDates.Left
   txtToDate.text = Format(Now, "dd/mm/yyyy")
End Sub

Private Sub Form_Load()
   Dim conClient As New ADODB.Connection
'   frmMMain.Arrange vbCascade
'   Me.ZOrder 0
   Me.Height = INI_HEIGHT
   Me.Width = 5865
   Me.BackColor = MODULEBACKCOLOR
   fraOptions.BackColor = MODULEBACKCOLOR
   fraSelectDates.BackColor = MODULEBACKCOLOR
   fraFunds.BackColor = MODULEBACKCOLOR
   fraEnterDates.BackColor = MODULEBACKCOLOR
   fraBankAccounts.BackColor = MODULEBACKCOLOR
   chkFunds.BackColor = MODULEBACKCOLOR
   chkDetails.BackColor = MODULEBACKCOLOR
   chkBankAccounts.BackColor = MODULEBACKCOLOR

   bCallingFromGrid = False

   conClient.Open getConnectionString

   LoadTransactionDates conClient

   conClient.Close
   Set conClient = Nothing

'   chkBankAccounts.Value = 1
   chkFunds.Value = 1
   Call WheelHook(Me.hWnd)
End Sub

Private Sub PrepareList(conClient As ADODB.Connection)
'   ConfigFlxClientList
'   ConfigFlxBankAccounts
'   ConfigFlxFunds

'   LoadFlxClientList conClient
'   LoadFlxBankAccounts conClient
'   LoadFlxFunds conClient
   LoadTransactionDates conClient
End Sub

Private Sub LoadFlxFunds(conClient As ADODB.Connection)
   Dim rstClient   As New ADODB.Recordset
   Dim szSQL       As String
   Dim iRow As Integer

   On Error GoTo ErrorHandler

   szSQL = "SELECT F.FundID, F.FundName, S.Value " & _
           "FROM Fund AS F, SecondaryCode AS S " & _
           "WHERE F.CategoryCode = CBYTE(S.Code) AND S.PrimaryCode = 'DCTG' " & _
           "ORDER BY FundID;"

   rstClient.Open szSQL, conClient, adOpenStatic, adLockReadOnly

   iRow = 1

   While Not rstClient.EOF
      flxFunds.TextMatrix(iRow, 1) = rstClient!fundID
      flxFunds.TextMatrix(iRow, 2) = rstClient!FundName
      flxFunds.TextMatrix(iRow, 3) = rstClient!Value
      rstClient.MoveNext
      If Not rstClient.EOF Then flxFunds.AddItem ""
      iRow = iRow + 1
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
   Dim iRow As Integer

   On Error GoTo ErrorHandler

   szSQL = "SELECT MY_ID, NominalCode, Bank_AC_Name, CLIENT_ID " & _
           "FROM tlbClientBanks " & _
           "ORDER BY CLIENT_ID;"

   rstClient.Open szSQL, conClient, adOpenStatic, adLockReadOnly

   iRow = 1

   While Not rstClient.EOF
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
   Exit Sub

ErrorHandler:
   MsgBox Err.description & "::" & Err.Number

   Set rstClient = Nothing
End Sub

Private Sub LoadflxClientList(conClient As ADODB.Connection)
   Dim rstClient   As New ADODB.Recordset
   Dim szSQL       As String
   Dim iRow As Integer

   On Error GoTo ErrorHandler
   szSQL = "SELECT ClientID, ClientName " & _
           "FROM Client " & _
           "ORDER BY ClientName;"

   rstClient.Open szSQL, conClient, adOpenStatic, adLockReadOnly

   iRow = 1

   While Not rstClient.EOF
      flxClientList.TextMatrix(iRow, 1) = rstClient!ClientID
      flxClientList.TextMatrix(iRow, 2) = rstClient!ClientName
      flxClientList.TextMatrix(iRow, 3) = "Client"
      rstClient.MoveNext
      If Not rstClient.EOF Then flxClientList.AddItem ""
      iRow = iRow + 1
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
'issue 523 Combo is not loading with proper recon dated
'Resolved by BOSL
'added by anol 20 Jan 2015
        If frmCashbook.txtClientList.text = "Consolidated" Then
            szSQL = "SELECT StatementDate, BankCode " & _
                    "FROM tlbBankReconClosingBal " & _
                    "WHERE BankCode = '" & frmCashbook.txtAccountName.text & "' " & _
                    "AND ClientID = '" & frmCashbook.txtBC.Tag & "' GROUP BY StatementDate, BankCode " & _
                    "ORDER BY StatementDate DESC;"
        Else
            szSQL = "SELECT StatementDate, BankCode " & _
                    "FROM tlbBankReconClosingBal " & _
                    "WHERE BankCode = '" & frmCashbook.txtBC.Tag & "' " & _
                    "AND ClientID = '" & frmCashbook.txtClientList.Tag & "' GROUP BY StatementDate, BankCode " & _
                    "ORDER BY StatementDate DESC;"
           End If
   
'Debug.Print szSQL
   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   Dim TotalRow As Integer, TotalCol As Integer
   Dim i As Integer, j As Integer

   TotalRow = adoRst.RecordCount
   TotalCol = adoRst.Fields.Count - 1

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
'
'Private Function IsBankSelected() As Boolean
'   Dim iRow          As Integer
'
'   For iRow = 1 To flxBankAccounts.Rows - 1
'      If flxBankAccounts.TextMatrix(iRow, 0) = "X" And flxBankAccounts.RowHeight(iRow) > 0 Then
'         szBanks = flxBankAccounts.TextMatrix(iRow, 1) & ", " & szBanks
'      End If
'   Next iRow
'   If Len(szBanks) > 2 Then
'      szBanks = Left(szBanks, Len(szBanks) - 2)
'      IsBankSelected = True
'   Else
'      IsBankSelected = False
'   End If
'End Function
'
'Private Function IsFundSelected() As Boolean
'   Dim iRow                As Integer
'
'   For iRow = 1 To flxFunds.Rows - 1
'      If flxFunds.TextMatrix(iRow, 0) = "X" And flxFunds.RowHeight(iRow) > 0 Then
'         szFunds = flxFunds.TextMatrix(iRow, 1) & ", " & szFunds
'         szFundList = flxFunds.TextMatrix(iRow, 2) & ", " & szFundList
'      End If
'   Next iRow
'   If Len(szFunds) > 2 Then
'      szFunds = Left(szFunds, Len(szFunds) - 2)
'      szFundList = Left(szFundList, Len(szFundList) - 2)
'      IsFundSelected = True
'   Else
'      IsFundSelected = False
'      Exit Function
'   End If
'End Function

Private Sub MarkBankFund()
   Dim adoConn    As New ADODB.Connection

   adoConn.Open getConnectionString

'Clear any existing marking
   adoConn.Execute "UPDATE tlbClientBanks " & _
                   "SET    SelBanks = '';"
'Clear any existing marking
   adoConn.Execute "UPDATE Fund " & _
                   "SET    SelFund = '', FundList = '';"

   adoConn.Execute "UPDATE tlbClientBanks " & _
                   "SET    SelBanks = 'Y' " & _
                   "WHERE  MY_ID = " & szBankID & ";"

   adoConn.Execute "UPDATE Fund " & _
                   "SET    SelFund = 'Y';"

   CalculateBBF adoConn

   adoConn.Close
   Set adoConn = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
   frmCashbook.Enabled = True
   Call WheelUnHook(Me.hWnd)
End Sub

Private Sub optReconciled_Click()
   If optReconciled.Value Then
      fraSelectDates.Top = fraOptions.Top
      fraSelectDates.Height = fraOptions.Height
      lblLast.Visible = True
      cboLtStDt.Visible = True
      cboCurStDt.SetFocus

      fraEnterDates.Top = Me.Height + 1000
      bX = False
   End If
End Sub

Private Sub optReconciled_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   bX = True
End Sub

Private Sub optUnReconBoth_Click()
   If optUnReconBoth.Value Then
      txtToDate.Top = 900
      Label1(2).Top = txtToDate.Top
      fraEnterDates.Top = fraOptions.Top
      fraEnterDates.Height = fraOptions.Height
      fraSelectDates.Top = Me.Height + 1000
      lblSpecifyDateRange.Visible = True
      txtFromDate.Visible = True
      txtFromDate.SetFocus
   End If
End Sub

Private Sub optUnreconciled_Click()
   If optUnreconciled.Value Then
      fraSelectDates.Top = Me.Height + 1000
      fraEnterDates.Top = fraOptions.Top
      fraEnterDates.Height = fraOptions.Height
      lblSpecifyDateRange.Visible = False
      txtFromDate.Visible = False
      txtToDate.Top = 660
      Label1(2).Top = txtToDate.Top
      txtToDate.SetFocus
   End If
End Sub

Private Sub txtFromDate_Change()
   TextBoxChangeDate txtFromDate
End Sub

Private Sub txtFromDate_GotFocus()
   SelTxtInCtrl txtFromDate
End Sub

Private Sub txtFromDate_KeyPress(KeyAscii As Integer)
   TextBoxKeyPrsDate txtFromDate, KeyAscii
End Sub

Private Sub txtFromDate_LostFocus()
   TextBoxFormatDate txtFromDate

   If txtFromDate.text = "" Or txtToDate.text = "" Then Exit Sub

   If CDate(txtFromDate.text) > CDate(txtToDate.text) Then
      txtFromDate.text = ""
      ShowMsgInTaskBar "From date cannot be after the To date", "Y", "N"
      txtFromDate.SetFocus
   End If
End Sub

Private Sub txtStClosingBal_KeyPress(KeyAscii As Integer)
   DigitTextKeyPress txtStClosingBal, KeyAscii
End Sub

Private Sub txtStClosingBal_LostFocus()
'   cmdGenReport.SetFocus
End Sub

Private Sub txtToDate_Change()
   TextBoxChangeDate txtToDate
End Sub

Private Sub txtToDate_GotFocus()
   SelTxtInCtrl txtToDate
End Sub

Private Sub txtToDate_KeyPress(KeyAscii As Integer)
   TextBoxKeyPrsDate txtToDate, KeyAscii
End Sub

Private Sub txtToDate_LostFocus()
   TextBoxFormatDate txtToDate

   If txtFromDate.text = "" And txtToDate.text = "" Then Exit Sub
   
   If bX Then Exit Sub

'   If optUnReconBoth.Value And cboCurStDt.text <> "" Then
'      If CDate(txtToDate.text) > CDate(cboCurStDt.text) Then
'         txtToDate.text = ""
'         ShowMsgInTaskBar "To date cannot be after the statement date", "Y", "N"
'         txtToDate.SetFocus
'         Exit Sub
'      End If
'   End If
'
'   If txtToDate.text <> "" And txtFromDate.text <> "" Then
'      If CDate(txtFromDate.text) > CDate(txtToDate.text) Then
'         txtToDate.text = ""
'         ShowMsgInTaskBar "To date cannot be before the From date", "Y", "N"
'         txtToDate.SetFocus
'         Exit Sub
'      End If
'   End If

   If optUnreconciled.Value Or optUnReconBoth.Value Then
      Dim cStBal As Currency
      Dim adoConn As New ADODB.Connection

      adoConn.Open getConnectionString
'issue 523
'Modified by anol 20 Jan 2015
      If IsDate(txtToDate.text) = False Then Exit Sub
      If frmCashbook.txtClientList.text = "Consolidated" Then
            cStBal = StatementClosingBalance(adoConn, txtToDate.text, frmCashbook.txtAccountName.text, frmCashbook.txtClientList.Tag) 'txtClientList.Tag is BankCode and txtAccountName.text is consolidated account number
      Else
             cStBal = StatementClosingBalance(adoConn, txtToDate.text, frmCashbook.txtBC.Tag, frmCashbook.txtClientList.Tag)
      End If
      
      If cStBal = 0 Then
         
         'Resolved By BOSL. Modified By Asif. Issue: 0000523. Date: 21-02-2015
         'If the statement balance is 0 which means that the user did not enter the statement date or a date
         'which has no statement balance, then prompt the user whether the program should load the last
         'statement balance before the date entered or the user wants to enter the projected balance instead.
      
         Dim message As String
         message = "The statement closing balance on " & txtToDate.text & " is not found. Do you want the program to load the last closing balance instead?"
         
         If MsgBox(message, vbYesNo, "Closing Balance Confirmation") = vbYes Then
             If frmCashbook.txtClientList.text = "Consolidated" Then
                cStBal = LastStatementBalanceBefore(adoConn, txtToDate.text, frmCashbook.txtAccountName.text, frmCashbook.txtClientList.Tag) 'txtClientList.Tag is BankCode and txtAccountName.text is consolidated account number
             Else
                cStBal = LastStatementBalanceBefore(adoConn, txtToDate.text, frmCashbook.txtBC.Tag, frmCashbook.txtClientList.Tag)
             End If
            
            txtStClosingBal.text = Format(cStBal, "0.00")
'            cmdGenReport.SetFocus
            txtStClosingBal.Locked = False
         Else
            txtStClosingBal.text = "0.00"
            txtStClosingBal.SetFocus
            txtStClosingBal.Locked = False
         End If
         'END OF MODIFICATION
      Else
         txtStClosingBal.text = Format(cStBal, "0.00")
         cmdGenReport.SetFocus
         txtStClosingBal.Locked = True
      End If

      adoConn.Close
      Set adoConn = Nothing
   End If
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
          PictureBoxZoom ctl, MouseKeys, Rotation, Xpos, Ypos

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


