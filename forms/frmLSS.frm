VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmLSS 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Client Summary Statement"
   ClientHeight    =   8880
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10545
   BeginProperty Font 
      Name            =   "Myriad Web"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLSS.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8880
   ScaleWidth      =   10545
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraProperties 
      Caption         =   "Properties:"
      Height          =   1710
      Left            =   225
      TabIndex        =   29
      Top             =   1215
      Width           =   5565
      Begin VB.CheckBox chkAllProperties 
         Caption         =   "All Properties"
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   240
         Width           =   2025
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxProperties 
         Height          =   1110
         Left            =   120
         TabIndex        =   31
         Top             =   540
         Width           =   5325
         _ExtentX        =   9393
         _ExtentY        =   1958
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
      Height          =   555
      Left            =   2160
      TabIndex        =   23
      Top             =   8235
      Width           =   1725
      Begin VB.TextBox txtRetention 
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
         Height          =   285
         Left            =   90
         MaxLength       =   10
         TabIndex        =   10
         Text            =   "0.00"
         Top             =   180
         Width           =   1485
      End
   End
   Begin VB.Frame fraInFunds 
      Caption         =   "Income Funds:"
      Height          =   1845
      Left            =   240
      TabIndex        =   21
      Top             =   4560
      Width           =   5565
      Begin VB.CheckBox chkInFunds 
         Caption         =   "All Funds"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   1095
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxInFunds 
         Height          =   1245
         Left            =   120
         TabIndex        =   7
         Top             =   525
         Width           =   5325
         _ExtentX        =   9393
         _ExtentY        =   2196
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
   Begin VB.Frame fraFunds 
      Caption         =   "Expenditure Funds:"
      Height          =   1845
      Left            =   240
      TabIndex        =   20
      Top             =   6435
      Width           =   5565
      Begin VB.CheckBox chkFunds 
         Caption         =   "All Funds"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   1095
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxFunds 
         Height          =   1245
         Left            =   120
         TabIndex        =   8
         Top             =   525
         Width           =   5325
         _ExtentX        =   9393
         _ExtentY        =   2196
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
      Height          =   1665
      Left            =   240
      TabIndex        =   19
      Top             =   2910
      Width           =   5565
      Begin VB.CheckBox chkBankAccounts 
         Caption         =   "All Accounts"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   1215
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxBankAccounts 
         Height          =   1020
         Left            =   120
         TabIndex        =   6
         Top             =   525
         Width           =   5325
         _ExtentX        =   9393
         _ExtentY        =   1799
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
      Height          =   3495
      Left            =   6120
      ScaleHeight     =   3465
      ScaleWidth      =   5520
      TabIndex        =   18
      Top             =   7635
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
         TabIndex        =   13
         Top             =   10
         Width           =   295
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxClientList 
         Height          =   2805
         Left            =   45
         TabIndex        =   24
         Top             =   585
         Width           =   5445
         _ExtentX        =   9604
         _ExtentY        =   4948
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
      Begin MSForms.Label lblClientID 
         Height          =   195
         Left            =   75
         TabIndex        =   28
         Top             =   45
         Width           =   735
         VariousPropertyBits=   8388627
         Caption         =   "Client ID"
         Size            =   "1296;353"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label lblClientName 
         Height          =   195
         Left            =   1785
         TabIndex        =   27
         Top             =   15
         Width           =   1185
         VariousPropertyBits=   8388627
         Caption         =   "Client Name"
         Size            =   "2090;344"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtSearchClientID 
         Height          =   255
         Left            =   45
         TabIndex        =   26
         Top             =   270
         Width           =   1530
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         Size            =   "2699;450"
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtSearchClientName 
         Height          =   255
         Left            =   1620
         TabIndex        =   25
         Top             =   270
         Width           =   3825
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         Size            =   "6747;450"
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Shape Shape4 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C0C0C0&
         BorderStyle     =   6  'Inside Solid
         FillColor       =   &H00E0FFFF&
         FillStyle       =   0  'Solid
         Height          =   240
         Index           =   15
         Left            =   0
         Top             =   0
         Width           =   5760
      End
   End
   Begin VB.TextBox txtLlName 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   120
      Width           =   3375
   End
   Begin VB.TextBox txtClientID 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   1090
   End
   Begin VB.CommandButton cmdClient 
      Caption         =   ".."
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
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   120
      Width           =   285
   End
   Begin VB.Frame Frame2 
      Caption         =   "Transaction Date:"
      Height          =   735
      Left            =   240
      TabIndex        =   14
      Top             =   480
      Width           =   5565
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
         Height          =   285
         Left            =   720
         MaxLength       =   10
         TabIndex        =   3
         Text            =   "01/01/2000"
         Top             =   300
         Width           =   1575
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
         Height          =   285
         Left            =   3825
         MaxLength       =   10
         TabIndex        =   4
         Top             =   270
         Width           =   1575
      End
      Begin VB.Label lblSpecifyDateRange 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "From:"
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   300
         Width           =   390
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To:"
         Height          =   195
         Index           =   1
         Left            =   3450
         TabIndex        =   15
         Top             =   315
         Width           =   210
      End
   End
   Begin VB.CommandButton cmdGenReport 
      Caption         =   "&Generate Report"
      Height          =   375
      Left            =   240
      TabIndex        =   9
      Top             =   8385
      Width           =   1575
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   375
      Left            =   4200
      TabIndex        =   11
      Top             =   8385
      Width           =   1575
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFDFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Client:"
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   17
      Top             =   120
      Width           =   465
   End
End
Attribute VB_Name = "frmLSS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Samrat Rahman: On 07/02/2012
'First this form was created for Landlord Summary Statement.
'We are changing this form for Client Summary Statement.
'This refer to the modification document on page 126: Client Summary Statement.

Option Explicit

Private szDemandTypes      As String
Private szBanks            As String
Private szFunds            As String
Private szInFunds          As String
Private bCallingFromGrid   As Boolean
Private Const INI_HEIGHT   As Integer = 9360
Dim szPropertyList As String
Private Sub chkAllProperties_Click()
   If bCallingFromGrid Then
      bCallingFromGrid = False
      Exit Sub
   End If

   Dim iRow As Integer

   For iRow = 1 To flxProperties.Rows - 1
      If flxProperties.RowHeight(iRow) > 0 And flxProperties.TextMatrix(iRow, 0) = "X" Then
         SelectFlxGridRow 0, flxProperties, iRow
      End If
      'MsgBox flxProperties.RowHeight(iRow)
   Next iRow

   For iRow = 1 To flxProperties.Rows - 1
      If flxProperties.RowHeight(iRow) > 0 And chkAllProperties.Value Then
         SelectFlxGridRow 0, flxProperties, iRow
      End If
   Next iRow
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

Private Sub chkInFunds_Click()
   If bCallingFromGrid Then
      bCallingFromGrid = False
      Exit Sub
   End If

   Dim iRow As Integer

   For iRow = 1 To flxInFunds.Rows - 1
      If flxInFunds.RowHeight(iRow) > 0 And flxInFunds.TextMatrix(iRow, 0) = "X" Then
         SelectFlxGridRow 0, flxInFunds, iRow
      End If
   Next iRow

   For iRow = 1 To flxInFunds.Rows - 1
      If flxInFunds.RowHeight(iRow) > 0 And chkInFunds.Value Then
         SelectFlxGridRow 0, flxInFunds, iRow
      End If
   Next iRow
End Sub

Private Sub cmdClose_Click()
   Unload Me
End Sub

Private Sub cmdGenReport_Click()
   If txtClientID.text = "" Then
      MsgBox "Please a Client.", vbInformation + vbOKOnly, "Client"
      cmdClient.SetFocus
      Exit Sub
   End If
   If txtRetention.text = "" Then
        MsgBox "Please enter retention amount", vbInformation, "Invalid amount"
        Exit Sub
   End If
   If txtFromDate.text = "" Then
      MsgBox "Please enter from date.", vbInformation + vbOKOnly, "From Date"
      txtFromDate.SetFocus
      Exit Sub
   End If
   If txtToDate.text = "" Then
      MsgBox "Please enter to date.", vbInformation + vbOKOnly, "To Date"
      txtToDate.SetFocus
      Exit Sub
   End If
   'fixed by anol 03 Nov 2015
   If IsDate(txtFromDate.text) = False Then
        MsgBox "Please enter a valid date"
        txtFromDate.SetFocus
        Exit Sub
   End If
    If IsDate(txtToDate.text) = False Then
        MsgBox "Please enter a valid date"
        txtToDate.SetFocus
        Exit Sub
   End If
   szPropertyList = ""
   If CreateListOfProp = 0 Then
      ShowMsgInTaskBar "Please select Property from the grid."
      flxProperties.SetFocus
      Exit Sub
   End If
   
   If Not IsBankSelected Then
      MsgBox "Please select a bank account.", vbInformation + vbOKOnly, "Bank Account"
      chkBankAccounts.SetFocus
      Exit Sub
   End If
   If Not IsInFundSelected Then
      MsgBox "Please select an income fund.", vbInformation + vbOKOnly, "Income Funds"
      chkInFunds.SetFocus
      Exit Sub
   End If
   If Not IsFundSelected Then
      MsgBox "Please select an expenditure fund.", vbInformation + vbOKOnly, "Expenditure Funds"
      chkFunds.SetFocus
      Exit Sub
   End If
   cmdGenReport.Enabled = False
   MarkBankFund

   Dim reportApp As New CRAXDRT.Application
   Dim Report As CRAXDRT.Report

   Set Report = reportApp.OpenReport(App.Path & szReportPath & "\CSS.rpt")

   Report.EnableParameterPrompting = False
   Report.DiscardSavedData

   Report.ParameterFields(1).AddCurrentValue CStr(txtClientID.text)
   Report.ParameterFields(2).AddCurrentValue CDate(txtFromDate.text)
   Report.ParameterFields(3).AddCurrentValue CDate(txtToDate.text)
   Report.ParameterFields(4).AddCurrentValue BBF2 '0 ' BBF()
   Report.ParameterFields(5).AddCurrentValue -Val(txtRetention.text)
   Load frmReport
   frmReport.LoadReportViewer Report
   cmdGenReport.Enabled = True
End Sub
Private Function CreateListOfProp() As Integer
   Dim i As Integer

   szPropertyList = ""
   
   For i = 0 To flxProperties.Rows - 1
      If flxProperties.TextMatrix(i, 0) = "X" And flxProperties.RowHeight(i) > 0 Then
         szPropertyList = "'" & flxProperties.TextMatrix(i, 1) & "'" & ", " & szPropertyList
      End If
   Next i
   If Len(szPropertyList) > 2 Then
      szPropertyList = Left(szPropertyList, Len(szPropertyList) - 2)
      CreateListOfProp = Len(szPropertyList)
      Exit Function
   End If
   CreateListOfProp = 0
End Function
Private Sub MarkPropOfSelection(adoconn As ADODB.Connection)
   Dim szSQL As String

   szSQL = "UPDATE PROPERTY " & _
           "SET    RAS = '';"
   adoconn.Execute szSQL

   szSQL = "UPDATE PROPERTY " & _
           "SET    RAS = 'X' " & _
           "WHERE  PropertyID IN (" & szPropertyList & ");"
'Debug.Print szSQL
   adoconn.Execute szSQL
End Sub
Private Function BBF2() As Double
   Dim i As Double
   Dim adoconn    As New ADODB.Connection
   adoconn.Open getConnectionString
   Dim iRow          As Integer
   szBanks = vbNullString
   For iRow = 1 To flxBankAccounts.Rows - 1
      If flxBankAccounts.TextMatrix(iRow, 0) = "X" And flxBankAccounts.RowHeight(iRow) > 0 Then
         szBanks = flxBankAccounts.TextMatrix(iRow, 2)
         'Fixed by anol 27 Oct 2015  duducting 1 day when taking BBF
         'i = i + BankAccBalanceDate(adoConn, szBanks, txtClientID.text, txtFromDate.text)
         i = i + BankAccBalanceDate(adoconn, szBanks, txtClientID.text, DateAdd("d", -1, txtFromDate.text))
      End If
   Next iRow
   adoconn.Close
   Set adoconn = Nothing
    BBF2 = i
End Function
Private Sub MarkBankFund()
   Dim adoconn    As New ADODB.Connection
   
   adoconn.Open getConnectionString

   adoconn.Execute "UPDATE tlbClientBanks " & _
                   "SET    SelBanks = NULL;"
   adoconn.Execute "UPDATE Fund " & _
                   "SET    SelFund = NULL, SelInFund = NULL;"

   adoconn.Execute "UPDATE tlbClientBanks " & _
                   "SET    SelBanks = 'Y' " & _
                   "WHERE  MY_ID IN (" & szBanks & ")"

   adoconn.Execute "UPDATE Fund " & _
                   "SET    SelFund = 'Y' " & _
                   "WHERE  FundID IN (" & szFunds & ")"
   
   adoconn.Execute "UPDATE Fund " & _
                   "SET    SelInFund = 'Y' " & _
                   "WHERE  FundID IN (" & szInFunds & ")"
'Debug.Print szFunds + " S "
   Call MarkPropOfSelection(adoconn)
   adoconn.Close
   Set adoconn = Nothing
End Sub

Private Function IsBankSelected() As Boolean
   Dim iRow          As Integer
   Dim adoconn As New ADODB.Connection
   adoconn.Open getConnectionString
   adoconn.Execute "Update tlbClientBanks Set SelBanks=''"
   szBanks = vbNullString
   For iRow = 1 To flxBankAccounts.Rows - 1
      If flxBankAccounts.TextMatrix(iRow, 0) = "X" And flxBankAccounts.RowHeight(iRow) > 0 Then
         szBanks = flxBankAccounts.TextMatrix(iRow, 1) & ", " & szBanks
         adoconn.Execute "Update tlbClientBanks Set SelBanks='Y' where NominalCode='" & flxBankAccounts.TextMatrix(iRow, 2) & "'"
      End If
   Next iRow
   If Len(szBanks) > 2 Then
      szBanks = Left(szBanks, Len(szBanks) - 2)
      IsBankSelected = True
   Else
      IsBankSelected = False
   End If
   adoconn.Close
   Set adoconn = Nothing
End Function

Private Function IsInFundSelected() As Boolean
   Dim iRow                As Integer

   szInFunds = vbNullString
   For iRow = 1 To flxInFunds.Rows - 1
      If flxInFunds.TextMatrix(iRow, 0) = "X" And flxInFunds.RowHeight(iRow) > 0 Then
         szInFunds = flxInFunds.TextMatrix(iRow, 1) & ", " & szInFunds
      End If
   Next iRow
   If Len(szInFunds) > 2 Then
      szInFunds = Left(szInFunds, Len(szInFunds) - 2)
      IsInFundSelected = True
   Else
      IsInFundSelected = False
      Exit Function
   End If
End Function

Private Function IsFundSelected() As Boolean
   Dim iRow                As Integer

   szFunds = vbNullString
   For iRow = 1 To flxFunds.Rows - 1
      If flxFunds.TextMatrix(iRow, 0) = "X" And flxFunds.RowHeight(iRow) > 0 Then
         szFunds = flxFunds.TextMatrix(iRow, 1) & ", " & szFunds
      End If
   Next iRow
   If Len(szFunds) > 2 Then
      szFunds = Left(szFunds, Len(szFunds) - 2)
      IsFundSelected = True
   Else
      IsFundSelected = False
      Exit Function
   End If
End Function

Private Function BBF() As Double
   Dim adoconn    As New ADODB.Connection
   Dim adoRst     As New ADODB.Recordset
   Dim szSQL      As String

   Dim cNR        As Currency
   Dim cLess      As Currency
   Dim cRentColl  As Currency
   Dim cAmtPaid   As Currency
   Dim cUnAlloc   As Currency

   adoconn.Open getConnectionString

'  Rent Collection
'  tblPurInv.Supp_AC = LL_ID, Sum of tblPurInv.TOTAL_AMOUNT
'  Transaction date < From Date
'  Fund category 1:RENT
   szSQL = "SELECT SUM(I.TOTAL_AMOUNT) AS RentColl " & _
           "FROM tblPurInv AS I, tblPurInvSRec AS S, " & _
               "Fund AS F, Property AS P " & _
           "WHERE I.TRAN_DATE < #" & Format(txtFromDate.text, "dd mmmm yyyy") & "# AND " & _
               "P.ClientID = '" & txtClientID.text & "' AND " & _
               "I.MY_ID = S.ParentID AND " & _
               "S.DEPT_ID = F.FundID AND " & _
               "F.CategoryCode = 1 AND " & _
               "P.PropertyID = I.PropertyID;"

'Debug.Print szSQL
   adoRst.Open szSQL, adoconn, adOpenStatic, adLockReadOnly

   If adoRst.EOF Then
      cRentColl = 0
   Else
      cRentColl = IIf(IsNull(adoRst.Fields.Item(0).Value), 0, adoRst.Fields.Item(0).Value)
   End If
   adoRst.Close

'  cAmtPaid: PoA
'  tlbPayment.Type = 9
'  Transaction date < From Date
'  tlbPayment.SageAccountNumber = LL_ID
'  Sum of tlbPayment.Amount
'  Fund category 1:RENT
   szSQL = "SELECT SUM(P.Amount) AS AmtPaid " & _
           "FROM tlbPayment AS P, Fund AS F " & _
           "WHERE P.Type = 9 AND " & _
               "P.PDate < #" & Format(txtFromDate.text, "dd mmmm yyyy") & "# AND " & _
               "P.SageAccountNumber = '" & txtClientID.text & "' AND " & _
               "P.FundID = F.FundID AND " & _
               "F.CategoryCode = 1;"

'Debug.Print szSQL
   adoRst.Open szSQL, adoconn, adOpenStatic, adLockReadOnly

   If adoRst.EOF Then
      cAmtPaid = 0
   Else
      cAmtPaid = IIf(IsNull(adoRst.Fields.Item(0).Value), 0, adoRst.Fields.Item(0).Value)
   End If
   adoRst.Close

'  cAmtPaid: PP
'  tlbPayment.Type = 8
'  Transaction date < From Date
'  tlbPayment.SageAccountNumber = LL_ID
'  Sum of tlbPayment.Amount
'  Fund category 1:RENT
   szSQL = "SELECT SUM(S.Amount) AS AmtPaid " & _
           "FROM tlbPayment AS P, Fund AS F, tlbPaymentSplit AS S " & _
           "WHERE P.Type = 8 AND " & _
               "P.PDate < #" & Format(txtFromDate.text, "dd mmmm yyyy") & "# AND " & _
               "P.SageAccountNumber = '" & txtClientID.text & "' AND " & _
               "S.FundID = F.FundID AND " & _
               "F.CategoryCode = 1 AND S.PayHeader = P.TransactionID;"

'Debug.Print szSQL
   adoRst.Open szSQL, adoconn, adOpenStatic, adLockReadOnly

   If Not adoRst.EOF Then
      cAmtPaid = cAmtPaid + IIf(IsNull(adoRst.Fields.Item(0).Value), 0, adoRst.Fields.Item(0).Value)
   End If
   adoRst.Close

'  cLess:
'  Less Outgoing expenses, both Non-Recoverable & Recoverable
'  Sum of PayTransactions.PaymentAmount
'  Transaction date < From Date
   szSQL = "SELECT SUM(S.Amount) AS Less " & _
           "FROM tlbPayment         AS P1, tlbPayment AS P2, " & _
                 "PayTransactions   AS PT, tblPurInv  AS I, " & _
                 "tlbPaymentSplit   AS S, Fund        AS F " & _
           "WHERE (P1.Type = 8 OR P1.Type = 9) AND " & _
               "P1.PDate < #" & Format(txtFromDate.text, "dd mmmm yyyy") & "# AND " & _
               "P1.TransactionID = PT.FromTran AND " & _
               "P2.TransactionID = PT.ToTran AND " & _
               "I.CL_ID = '" & txtClientID.text & "' AND " & _
               "P2.PI = I.MY_ID AND " & _
               "S.FundID = F.FundID AND " & _
               "F.CategoryCode = 1 AND S.PayHeader = P1.TransactionID;"
'Debug.Print szSQL
   adoRst.Open szSQL, adoconn, adOpenStatic, adLockReadOnly

   If adoRst.EOF Then
      cLess = 0
   Else
      cLess = IIf(IsNull(adoRst.Fields.Item(0).Value), 0, adoRst.Fields.Item(0).Value)
   End If
   adoRst.Close

   cUnAlloc = 0
''  cUnAlloc:
''  tlbPayment.Type = 8
''  Transaction date < From Date
''  tlbPayment.OSAmount = LL_ID
''  Sum of tlbPayment.Amount
''  Fund category 1:RENT
'   szSQL = "SELECT SUM(S.Amount) AS UnAlloc " & _
'           "FROM tlbPayment AS P, tlbPaymentSplit AS S, Fund AS F " & _
'           "WHERE (P.Type = 8 OR P.Type = 9) AND " & _
'               "P.PDate < #" & txtFromDate.text & "# AND " & _
'               "I.CL_ID = '" & txtClientID.text & "' AND " & _
'               "S.FundID = F.FundID AND " & _
'               "F.CategoryCode = 1 AND S.PayHeader = P.TransactionID;"
'
''Debug.Print szSQL
'   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'
'   If adoRst.EOF Then
'      cUnAlloc = 0
'   Else
'      cUnAlloc = IIf(IsNull(adoRst.Fields.Item(0).Value), 0, adoRst.Fields.Item(0).Value)
'   End If
'   adoRst.Close

   adoconn.Close
   Set adoRst = Nothing
   Set adoconn = Nothing

   BBF = cRentColl - cAmtPaid - cLess - cUnAlloc
End Function

Private Sub cmdGridUnitLookup_Click()
   picClientList.Visible = False
   Me.Height = INI_HEIGHT
End Sub

Private Sub cmdClient_Click()
   picClientList.Top = txtClientID.Top + txtClientID.Height + 5
   picClientList.Left = Label1(0).Left + 5
   picClientList.Visible = True
   picClientList.ZOrder 0
   Me.Height = 3360
   txtSearchClientID.SetFocus
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
   Call loadflxProperties
   chkAllProperties.Value = 1
   FilteringBanks
End Sub
Private Sub loadflxProperties()
    Dim szSQL   As String
    Dim r       As Integer
    Dim adoconn As New ADODB.Connection
    Dim adoRst As New ADODB.Recordset
    adoconn.Open getConnectionString
     szSQL = "SELECT   PROPERTYID, PROPERTYNAME, ClientID " & _
           "FROM     PROPERTY where ClientID='" & txtClientID.text & "'" & _
           "ORDER BY PROPERTYID;"
   adoRst.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
   ConfigFlxProperties
   r = 1
   
'      flxProperties.TextMatrix(r, 1) = "ZZZZ"
'      flxProperties.TextMatrix(r, 2) = "Common Properties"
'      flxProperties.TextMatrix(r, 3) = txtClientList.Tag
'      flxProperties.AddItem ""
'   r = 2

   While Not adoRst.EOF
      flxProperties.TextMatrix(r, 1) = adoRst.Fields.Item("PROPERTYID").Value
      flxProperties.TextMatrix(r, 2) = adoRst.Fields.Item("PROPERTYNAME").Value
      flxProperties.TextMatrix(r, 3) = adoRst.Fields.Item("ClientID").Value
      flxProperties.RowHeight(r) = 240
      r = r + 1
      
      adoRst.MoveNext
      If Not adoRst.EOF Then flxProperties.AddItem ""
   Wend
    Debug.Print r
   adoRst.Close
   Set adoRst = Nothing
   adoconn.Close
   Set adoconn = Nothing
   flxProperties.row = 0
End Sub
Private Sub ConfigFlxProperties()
   Dim szHeader As String
   flxProperties.Clear
   flxProperties.Rows = 2
   szHeader$ = "<|<|<|<"
   With flxProperties
      .FormatString = szHeader
      .Cols = 4
      .RowHeight(0) = 0
      .ColWidth(0) = 200 'Label2(0).Left - .Left '200                 '"X"
      .ColWidth(1) = 2000 'Label2(1).Left - Label2(0).Left 'Property ID
      .ColWidth(2) = 2500 'Label2(2).Left - Label2(1).Left 'Property Name
      .ColWidth(3) = 0 '.Width + .Left - Label2(2).Left - 300 'Client ID
   End With
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
End Sub

Private Sub flxFunds_Click()
   If flxFunds.row = 0 Then Exit Sub

   SelectFlxGridRow 0, flxFunds, flxFunds.row

   Dim iRow As Integer

   For iRow = 1 To flxFunds.Rows - 1
      If flxFunds.TextMatrix(iRow, 0) <> "X" And chkFunds.Value = 1 Then
         bCallingFromGrid = True
         chkFunds.Value = 0
         Exit For
      End If
   Next iRow
End Sub

Private Sub flxInFunds_Click()
   If flxInFunds.row = 0 Then Exit Sub

   SelectFlxGridRow 0, flxInFunds, flxInFunds.row

   Dim iRow As Integer

   For iRow = 1 To flxInFunds.Rows - 1
      If flxInFunds.TextMatrix(iRow, 0) <> "X" And chkInFunds.Value = 1 Then
         bCallingFromGrid = True
         chkInFunds.Value = 0
         Exit For
      End If
   Next iRow
End Sub

Private Sub flxProperties_Click()
    If flxProperties.row = 0 Then Exit Sub

   SelectFlxGridRow 0, flxProperties, flxProperties.row

   Dim iRow As Integer

   For iRow = 1 To flxProperties.Rows - 1
      If flxProperties.TextMatrix(iRow, 0) <> "X" And chkAllProperties.Value Then
         bCallingFromGrid = True
         chkAllProperties.Value = 0
         Exit For
      End If
   Next iRow
End Sub

Private Sub Form_Load()
   Dim conClient As New ADODB.Connection
    
   Me.Height = INI_HEIGHT
   Me.Width = 6075
   frmMMain.Arrange vbCascade
   Me.ZOrder 0
   Me.BackColor = MODULEBACKCOLOR
   Frame2.BackColor = MODULEBACKCOLOR
   fraBankAccounts.BackColor = MODULEBACKCOLOR
   fraProperties.BackColor = MODULEBACKCOLOR
   chkAllProperties.BackColor = MODULEBACKCOLOR
   fraFunds.BackColor = MODULEBACKCOLOR
   fraInFunds.BackColor = MODULEBACKCOLOR
   chkBankAccounts.BackColor = fraBankAccounts.BackColor
   chkFunds.BackColor = fraFunds.BackColor
   chkInFunds.BackColor = fraInFunds.BackColor
   txtToDate.text = Format(Now, "dd/mm/yyyy")
   bCallingFromGrid = False

   conClient.Open getConnectionString

   Call PrepareList(conClient)

   conClient.Close
   Set conClient = Nothing

   chkBankAccounts.Value = 1
   chkFunds.Value = 1
   chkInFunds.Value = 1
   Call WheelHook(Me.hWnd)
End Sub

Private Sub PrepareList(conClient As ADODB.Connection)
   ConfigflxClientList
   ConfigFlxBankAccounts
   ConfigFlxFunds
   ConfigFlxInFunds
' szSQL = "SELECT ClientID, ClientName " & _
'           "FROM Client " & _
'           "ORDER BY ClientName;"
   LoadflxClientList conClient, "SELECT ClientID, ClientName " & _
           "FROM Client " & _
           "ORDER BY ClientName;"
   LoadFlxBankAccounts conClient
   LoadFlxFunds conClient
   LoadflxInFunds conClient
End Sub

Private Sub ConfigflxClientList()
   Dim szHeader As String

   flxClientList.Cols = 4
   flxClientList.Clear
   szHeader$ = "|<ID|<Name|<Type"
   flxClientList.FormatString = szHeader$
   flxClientList.ColWidth(0) = 0        'Solid column
   flxClientList.ColWidth(1) = 1600        'Client ID
   flxClientList.ColWidth(2) = 3000       'Client Name
   flxClientList.ColWidth(3) = 0        'Post Code
   flxClientList.Rows = 2

   flxClientList.RowHeight(0) = 0
End Sub
Private Sub txtSearchClientID_Change()
    'Updated by anol 10 Dec 2015
   Dim i As Integer

   If Len(txtSearchClientID.text) > 0 Then
        txtSearchClientName.text = ""
   End If

   For i = flxClientList.Rows - 1 To 1 Step -1
      flxClientList.RowHeight(i) = 240

      If InStr(1, UCase(flxClientList.TextMatrix(i, 1)), UCase(txtSearchClientID.text), vbTextCompare) = 0 Then
            flxClientList.RowHeight(i) = 0
      End If
      If flxClientList.RowHeight(i) = 240 Then
            flxClientList.row = i
      End If
   Next i
'    Dim conClient As New ADODB.Connection
'    conClient.Open getConnectionString
'    ConfigflxClientList
'    LoadflxClientList conClient, "SELECT ClientID, ClientName " & _
'    "FROM Client where ClientID like '%" & txtSearchClientID.text & "' " & _
'    "ORDER BY ClientName;"
'
'    conClient.Close
'    Set conClient = Nothing
End Sub

Private Sub txtSearchClientID_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 13 Then
        flxClientList.SetFocus
    End If
End Sub

Private Sub txtSearchClientID_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 13 Then
        flxClientList.SetFocus
    End If
End Sub

Private Sub txtSearchClientName_Change()
   'Updated by anol 10 Dec 2015
   Dim i As Integer

   If Len(txtSearchClientName.text) > 0 Then
        txtSearchClientID.text = ""
   End If

   For i = flxClientList.Rows - 1 To 1 Step -1
      flxClientList.RowHeight(i) = 240
      If InStr(1, UCase(flxClientList.TextMatrix(i, 2)), UCase(txtSearchClientName.text), vbTextCompare) = 0 Then
            flxClientList.RowHeight(i) = 0
      End If
      If flxClientList.RowHeight(i) = 240 Then
            flxClientList.row = i
      End If
   Next i
'    ConfigflxClientList
'    Dim conClient As New ADODB.Connection
'    conClient.Open getConnectionString
'    LoadflxClientList conClient, "SELECT ClientID, ClientName " & _
'         "FROM Client where ClientName like '%" & txtSearchClientName.text & "' " & _
'         "ORDER BY ClientName;"
'    conClient.Close
'    Set conClient = Nothing
End Sub
Private Sub LoadflxClientList(conClient As ADODB.Connection, szSQL As String)
   Dim rstClient   As New ADODB.Recordset
  ' Dim szSQL       As String
   Dim iRow As Integer

   On Error GoTo ErrorHandler
  

   rstClient.Open szSQL, conClient, adOpenStatic, adLockReadOnly

   iRow = 1

   While Not rstClient.EOF
      flxClientList.TextMatrix(iRow, 1) = rstClient!clientID
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

Private Sub ConfigFlxInFunds()
   Dim szHeader As String

   flxInFunds.Cols = 4
   flxInFunds.Clear
   szHeader$ = "|<ID|<Name|<Category"
   flxInFunds.FormatString = szHeader$
   flxInFunds.ColWidth(0) = 280        'Solid column
   flxInFunds.ColWidth(1) = 400        'ID
   flxInFunds.ColWidth(2) = 2800       'Name
   flxInFunds.ColWidth(3) = 1500       'Post Code
   flxInFunds.Rows = 2

   flxInFunds.RowHeightMin = 255
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

Private Sub LoadflxInFunds(conClient As ADODB.Connection)
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
      flxInFunds.TextMatrix(iRow, 1) = rstClient!fundID
      flxInFunds.TextMatrix(iRow, 2) = rstClient!FundName
      flxInFunds.TextMatrix(iRow, 3) = rstClient!Value
      rstClient.MoveNext
      If Not rstClient.EOF Then flxInFunds.AddItem ""
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

Private Sub Form_Unload(Cancel As Integer)
   Call WheelUnHook(Me.hWnd)
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
   If IsDate(txtFromDate.text) = False Then
        MsgBox "Please enter a valid date"
        txtFromDate.SetFocus
   End If
   If txtFromDate.text <> "" Then
      TextBoxFormatDate txtFromDate
   End If
End Sub

Private Sub txtRetention_KeyPress(KeyAscii As Integer)
     DigitTextKeyPress txtRetention, KeyAscii
End Sub

Private Sub txtRetention_LostFocus()
    If IsNumeric(txtRetention.text) = False Then
        txtRetention.text = "0.00"
    End If
    txtRetention.text = Format(txtRetention.text, "0.00")
End Sub

Private Sub txtSearchClientName_KeyPress(KeyAscii As MSForms.ReturnInteger)
     If KeyAscii = 13 Then
        flxClientList.SetFocus
    End If
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
   'TextBoxFormatDate txtToDate
   If IsDate(txtToDate.text) = False Then
        MsgBox "Please enter a valid date"
        txtToDate.SetFocus
   End If
    If txtToDate.text <> "" Then
      TextBoxFormatDate txtToDate
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
