VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Prestige - Import Data from Excel to Access"
   ClientHeight    =   8625
   ClientLeft      =   6600
   ClientTop       =   3270
   ClientWidth     =   7020
   BeginProperty Font 
      Name            =   "Myriad Web"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8625
   ScaleWidth      =   7020
   Begin VB.CommandButton cmdBrowse 
      Height          =   615
      Index           =   0
      Left            =   6000
      Picture         =   "frmMain.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      Width           =   735
   End
   Begin VB.PictureBox cdlgBrowser 
      Height          =   480
      Left            =   6120
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   35
      Top             =   120
      Width           =   1200
   End
   Begin VB.TextBox txtSourceFile 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   22
      Top             =   240
      Width           =   4215
   End
   Begin MSForms.CheckBox chkChargeTypes 
      Height          =   375
      Left            =   360
      TabIndex        =   34
      Top             =   4995
      Width           =   1695
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      DisplayStyle    =   4
      Size            =   "2990;661"
      Value           =   "1"
      Caption         =   "Charge Types"
      FontName        =   "Myriad Web"
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.CheckBox chkPayableTypes 
      Height          =   375
      Left            =   360
      TabIndex        =   33
      Top             =   4590
      Width           =   1695
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      DisplayStyle    =   4
      Size            =   "2990;661"
      Value           =   "1"
      Caption         =   "Payable Types"
      FontName        =   "Myriad Web"
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.CheckBox chkAddPayDateSet 
      Height          =   375
      Left            =   360
      TabIndex        =   32
      Top             =   3765
      Width           =   3015
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      DisplayStyle    =   4
      Size            =   "5318;661"
      Value           =   "1"
      Caption         =   "Additional Payment Date Sets"
      FontName        =   "Myriad Web"
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Select Excel file:"
      Height          =   375
      Left            =   240
      TabIndex        =   23
      Top             =   240
      Width           =   1455
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0000C000&
      Height          =   5130
      Index           =   0
      Left            =   240
      Top             =   750
      Width           =   6615
   End
   Begin MSForms.CheckBox chkServiceBudget 
      Height          =   375
      Left            =   4320
      TabIndex        =   14
      Top             =   2190
      Width           =   2295
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      DisplayStyle    =   4
      Size            =   "4048;661"
      Value           =   "1"
      Caption         =   "Service Charge Budget"
      FontName        =   "Myriad Web"
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.CheckBox chkInsuranceBudget 
      Height          =   375
      Left            =   4320
      TabIndex        =   15
      Top             =   2580
      Width           =   1935
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      DisplayStyle    =   4
      Size            =   "3413;661"
      Value           =   "1"
      Caption         =   "Insurance Budget"
      FontName        =   "Myriad Web"
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.CheckBox chkRentBudget 
      Height          =   375
      Left            =   4320
      TabIndex        =   13
      Top             =   1800
      Width           =   1455
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      DisplayStyle    =   4
      Size            =   "2566;661"
      Value           =   "1"
      Caption         =   "Rent Budget"
      FontName        =   "Myriad Web"
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.CheckBox chkRC_SC_IC 
      Height          =   735
      Left            =   4320
      TabIndex        =   17
      Top             =   3360
      Width           =   1935
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      DisplayStyle    =   4
      Size            =   "3413;1296"
      Value           =   "1"
      Caption         =   "Rent Charges Service Charges Insurance Charge"
      FontName        =   "Myriad Web"
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.CheckBox chkSchedule 
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   1455
      Width           =   1215
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      DisplayStyle    =   4
      Size            =   "2143;661"
      Value           =   "1"
      Caption         =   "Schedule"
      FontName        =   "Myriad Web"
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.CheckBox chkReportCategory 
      Height          =   375
      Left            =   360
      TabIndex        =   7
      Top             =   3000
      Width           =   2055
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      DisplayStyle    =   4
      Size            =   "3625;661"
      Value           =   "1"
      Caption         =   "Nominal Categories"
      FontName        =   "Myriad Web"
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00404080&
      BorderWidth     =   2
      Height          =   5085
      Index           =   3
      Left            =   240
      Top             =   750
      Width           =   6615
   End
   Begin MSForms.CheckBox chkNC 
      Height          =   375
      Left            =   360
      TabIndex        =   8
      Top             =   3375
      Width           =   1575
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      DisplayStyle    =   4
      Size            =   "2778;661"
      Value           =   "1"
      Caption         =   "Nominal Code"
      FontName        =   "Myriad Web"
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.CommandButton cmdImportData 
      Default         =   -1  'True
      Height          =   975
      Left            =   3600
      TabIndex        =   20
      Top             =   7455
      Width           =   1455
      Caption         =   "Import Data"
      Size            =   "2566;1720"
      FontName        =   "Myriad Web"
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdExitProgram 
      Height          =   975
      Left            =   5235
      TabIndex        =   21
      Top             =   7455
      Width           =   1455
      Caption         =   "EXIT"
      Size            =   "2566;1720"
      FontName        =   "Myriad Web"
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label lblPlsWait 
      Height          =   375
      Index           =   0
      Left            =   1200
      TabIndex        =   31
      Top             =   7380
      Visible         =   0   'False
      Width           =   1575
      ForeColor       =   32768
      Caption         =   "Please wait"
      Size            =   "2778;661"
      FontName        =   "Myriad Web"
      FontEffects     =   1073741825
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin VB.Label lblMessage 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   495
      Index           =   2
      Left            =   240
      TabIndex        =   30
      Top             =   6780
      Width           =   6495
   End
   Begin MSForms.CheckBox chkSuppliers 
      Height          =   375
      Left            =   4320
      TabIndex        =   18
      Top             =   4110
      Width           =   1455
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      DisplayStyle    =   4
      Size            =   "2566;661"
      Value           =   "1"
      Caption         =   "Suppliers"
      FontName        =   "Myriad Web"
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.CheckBox chkFund 
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   1065
      Width           =   855
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      DisplayStyle    =   4
      Size            =   "1508;661"
      Value           =   "1"
      Caption         =   "Fund"
      FontName        =   "Myriad Web"
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.CheckBox chkLessee 
      Height          =   375
      Left            =   4320
      TabIndex        =   12
      Top             =   1410
      Width           =   1695
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      DisplayStyle    =   4
      Size            =   "2990;661"
      Value           =   "1"
      Caption         =   "Lessee/Tenant"
      FontName        =   "Myriad Web"
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label lblMessage 
      BackColor       =   &H00E0E0E0&
      Caption         =   "1.   Lessees' details should be in the database before running this process."
      ForeColor       =   &H00008000&
      Height          =   495
      Index           =   1
      Left            =   240
      TabIndex        =   29
      Top             =   6480
      Width           =   6495
   End
   Begin MSForms.CheckBox chkAll 
      Height          =   375
      Left            =   2160
      TabIndex        =   1
      Top             =   840
      Width           =   2055
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      DisplayStyle    =   4
      Size            =   "3625;661"
      Value           =   "1"
      Caption         =   "Select All Options"
      FontName        =   "Myriad Web"
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.CheckBox chkBankDetails 
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   1830
      Width           =   1455
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      DisplayStyle    =   4
      Size            =   "2566;661"
      Value           =   "1"
      Caption         =   "Bank Branch"
      FontName        =   "Myriad Web"
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.CheckBox chkLease 
      Height          =   375
      Left            =   4320
      TabIndex        =   16
      Top             =   2970
      Width           =   1575
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      DisplayStyle    =   4
      Size            =   "2778;661"
      Value           =   "1"
      Caption         =   "Lease Details"
      FontName        =   "Myriad Web"
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.CheckBox chkAPD 
      Height          =   375
      Left            =   4320
      TabIndex        =   19
      Top             =   4500
      Width           =   2535
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      DisplayStyle    =   4
      Size            =   "4471;661"
      Value           =   "1"
      Caption         =   "Additional Payment Dates"
      FontName        =   "Myriad Web"
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.CheckBox chkUnit 
      Height          =   375
      Left            =   4320
      TabIndex        =   11
      Top             =   1020
      Width           =   855
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      DisplayStyle    =   4
      Size            =   "1508;661"
      Value           =   "1"
      Caption         =   "Units"
      FontName        =   "Myriad Web"
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.CheckBox chkProperty 
      Height          =   375
      Left            =   360
      TabIndex        =   10
      Top             =   5355
      Width           =   1335
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      DisplayStyle    =   4
      Size            =   "2355;661"
      Value           =   "1"
      Caption         =   "Properties"
      FontName        =   "Myriad Web"
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.CheckBox chkDemandType 
      Height          =   375
      Left            =   360
      TabIndex        =   9
      Top             =   4155
      Width           =   1695
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      DisplayStyle    =   4
      Size            =   "2990;661"
      Value           =   "1"
      Caption         =   "Demand Types"
      FontName        =   "Myriad Web"
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.CheckBox chkClientBank 
      Height          =   375
      Left            =   360
      TabIndex        =   6
      Top             =   2610
      Width           =   1575
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      DisplayStyle    =   4
      Size            =   "2778;661"
      Value           =   "1"
      Caption         =   "Client's Banks"
      FontName        =   "Myriad Web"
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.CheckBox chkClients 
      Height          =   375
      Left            =   360
      TabIndex        =   5
      Top             =   2220
      Width           =   975
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      DisplayStyle    =   4
      Size            =   "1720;661"
      Value           =   "1"
      Caption         =   "Clients"
      FontName        =   "Myriad Web"
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label lblMessage 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Conditions for Multiple Selection:"
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   28
      Top             =   6030
      Width           =   6495
   End
   Begin MSForms.Label lblPlsWait 
      Height          =   375
      Index           =   4
      Left            =   1200
      TabIndex        =   27
      Top             =   7335
      Visible         =   0   'False
      Width           =   1695
      ForeColor       =   32768
      Caption         =   "Please wait . . ."
      Size            =   "2990;661"
      FontName        =   "Myriad Web"
      FontEffects     =   1073741825
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.Label lblPlsWait 
      Height          =   375
      Index           =   2
      Left            =   1200
      TabIndex        =   26
      Top             =   7335
      Visible         =   0   'False
      Width           =   1575
      ForeColor       =   32768
      Caption         =   "Please wait ."
      Size            =   "2778;661"
      FontName        =   "Myriad Web"
      FontEffects     =   1073741825
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.Label lblPlsWait 
      Height          =   495
      Index           =   1
      Left            =   720
      TabIndex        =   25
      Top             =   7860
      Visible         =   0   'False
      Width           =   2535
      Caption         =   "While data is uploading"
      Size            =   "4471;873"
      FontName        =   "Myriad Web"
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label lblPlsWait 
      Height          =   375
      Index           =   3
      Left            =   1200
      TabIndex        =   24
      Top             =   7335
      Visible         =   0   'False
      Width           =   1695
      ForeColor       =   32768
      Caption         =   "Please wait . ."
      Size            =   "2990;661"
      FontName        =   "Myriad Web"
      FontEffects     =   1073741825
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Enum PayingTime
   Pay_Yearly = 1
   Pay_Half_Yearly = 2
   Pay_Quarterly = 3
   Pay_Monthly = 4
   Pay_Fortnightly = 5
   Pay_Weekly = 6
End Enum

Dim fpFile As String
Dim iMsgNum As Byte
Dim baMsg(10) As Boolean

Const DELAY_MAX As Long = 99999

Public szGDYearly As String
Public szGDHalfYearly1 As String
Public szGDHalfYearly2 As String
Public szGDQuarterly1 As String
Public szGDQuarterly2 As String
Public szGDQuarterly3 As String
Public szGDQuarterly4 As String
Private szaMonthlyGD(12) As String
Dim isSupplier As Boolean


Private Sub cmdImportData_Click()
   
   If Trim(txtSourceFile.text) = "" Then
      MsgBox "Please SELECT an excel file", vbCritical + vbOKOnly, "Update Data"
      cmdBrowse(0).SetFocus
      Exit Sub
   End If

'   On Error GoTo ErrReport

   Dim oConnEx          As New ADODB.Connection    'MS Excel
   Dim oConnAcc         As New ADODB.Connection    'MS Access

'  Recordset for Excel
   Dim oRstXls          As New ADODB.Recordset
   Dim oNCsXls          As New ADODB.Recordset
   Dim oRC_Xls          As New ADODB.Recordset
   Dim oSC_Xls          As New ADODB.Recordset
   Dim oIC_Xls          As New ADODB.Recordset

'  Recordset for Access
   Dim adoRstMdb        As New ADODB.Recordset
   Dim adoRst           As New ADODB.Recordset
   Dim adoClient        As New ADODB.Recordset
   Dim adoClientAll        As New ADODB.Recordset
   Dim adoRST_          As New ADODB.Recordset

   Dim sTableNameEx     As String
   Dim sTableNameAcc    As String
   Dim szSQL            As String
'   Dim iLoop            As Integer
   Dim szTemp           As String
   Dim szTemp2           As String
   Dim iKount           As Integer
   Dim szClients        As String
   Dim bFreq            As Boolean
   Dim dtStDt           As Date
   Dim dtEndDt          As Date
'   Dim szPropertyID() As String
   Dim i As Integer
   'ADDED BY ANOL 26 MAR 2015
   Call UpdateDatabase
   
   Open App.Path & "\Logs\UpdateLog_" & Format(Now, "yymmddhhmmss") & ".dat" For Output As #1

'  this assumes that the first row contains headers
   oConnEx.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & txtSourceFile.text & ";" & "Extended Properties=""Excel 8.0;HDR=YES;IMEX=1;"""
'  Create Connection with Access Database
   oConnAcc.Open "DSN=" & frmStartUp.cboShopCentre.Value & ";UID=;PWD=RDSWKDPP"
  
   lblPlsWait(1).Visible = True
   lblPlsWait(2).Visible = True
   DoEvents

   Print #1, "####### Log file has been opened ######### Date:" & Format(Now, "dd/mm/yyyy") & " " & Format(Now, "hh:mm:ss") & "" & Chr$(13) + Chr$(10)
   Print #1, "Excel File has been opened Successfull." & Chr$(13) + Chr$(10)
   Print #1, "Starting.... Writing into the Database." & Chr$(13) + Chr$(10)
'   Dim SQL As String
'   Call Imp_Lessee_phone(oConnEx, oConnAcc)
'GoTo END_MILESTONE:
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ Fund ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
FUND_MILESTONE:
If Not chkFund.Value Then GoTo SCHEDULE_MILESTONE


   Print #1, "" & Chr$(13) + Chr$(10)
   Print #1, "" & Chr$(13) + Chr$(10)
   Print #1, "Data writing START: ********** Fund **********" & Chr$(13) + Chr$(10)
   lblPlsWait(1).Caption = "While data is uploading: Fund"
   'WaitFewSec
   DoEvents
   If chkFund.Value Then
        Call Imp_Funds(oConnEx, oConnAcc)
   End If

   Print #1, "Data writing FINISH: ********** Fund **********" & Chr$(13) + Chr$(10)

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ Schedule ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
SCHEDULE_MILESTONE:
'If Not chkFund.Value Then GoTo BANKDETAILS_MILESTONE
  If Not chkSchedule.Value Then GoTo BANKDETAILS_MILESTONE

   Print #1, "" & Chr$(13) + Chr$(10)
   Print #1, "" & Chr$(13) + Chr$(10)
   Print #1, "Data writing START: ********** Schedule **********" & Chr$(13) + Chr$(10)
   lblPlsWait(1).Caption = "While data is uploading: Schedule"
   
   DoEvents
   If Not chkSchedule.Value Then
        Call Imp_Schedule(oConnEx, oConnAcc)
   End If

   Print #1, "Data writing FINISH: ********** Schedule **********" & Chr$(13) + Chr$(10)

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ Bank Details ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BANKDETAILS_MILESTONE:
   If Not chkBankDetails.Value Then GoTo CLIENTS_MILESTONE:

   Print #1, "Data writing START: ********** BankDetails **********" & Chr$(13) + Chr$(10)
   lblPlsWait(1).Caption = "While data is uploading: Bank Details"
   'WaitFewSec
   DoEvents
   If chkBankDetails.Value Then
        Call Imp_BankDetails(oConnEx, oConnAcc)
   End If
   
   Print #1, "Data writing FINISH: ********** BankDetails **********" & Chr$(13) + Chr$(10)
CLIENTS_MILESTONE:
'Current procedure was long. Complier could not allow it here.I have seperated the code in new sub procedure anol 2019-11-11
'Import Clients
 If chkClients.Value Then
    lblPlsWait(1).Caption = "While data is uploading: Clients"
    WaitFewSec
    DoEvents
    If Imp_Clients(oConnEx, oConnAcc) = False Then
        GoTo ErrReport
    End If
 End If
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ Nominal CATEGORY Code ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'NOMINALCATEGORY_MILESTONE:

If chkReportCategory.Value Then
    Call import_nominalCategory(oConnEx, oConnAcc)
End If
   
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ Nominal Code ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
NOMINALCODE_MILESTONE:
   If Not chkNC.Value Then GoTo CLIENT_BANK_MILESTONE
   If chkNC.Value Then
        If import_NominalCode(oConnEx, oConnAcc) = False Then
            MsgBox "Import fails"
            Exit Sub
        End If
   End If

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ Client's Banks ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
CLIENT_BANK_MILESTONE:
If Not chkClientBank.Value Then GoTo PROPERTY_MILESTONE
    If chkClientBank.Value Then
        Call imp_ClientBank(oConnEx, oConnAcc)
        Call imp_ConsolidatedBank(oConnEx, oConnAcc)
        
   End If
'

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ Properties ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
PROPERTY_MILESTONE:
If Not chkProperty.Value Then GoTo ADD_PAY_DATE_SET_MILESTONE
   Print #1, "" & Chr$(13) + Chr$(10)
   Print #1, "" & Chr$(13) + Chr$(10)
   Print #1, "Data writing START: ********** Properties **********" & Chr$(13) + Chr$(10)
   lblPlsWait(1).Caption = "While data is uploading: Properties"
   'WaitFewSec
   DoEvents

   If Imp_Properties(oConnEx, oConnAcc) = False Then 'import properties
        GoTo ErrReport
   End If

   Print #1, "Data writing FINISH: ********** Properties **********" & Chr$(13) + Chr$(10)

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ Additional Payment DateSet ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
ADD_PAY_DATE_SET_MILESTONE:
If Not chkAddPayDateSet.Value Then GoTo DEMAND_TYPE_MILESTONE

   Print #1, "" & Chr$(13) + Chr$(10)
   Print #1, "" & Chr$(13) + Chr$(10)
   Print #1, "Data writing START: ********** Addistional Due Date **********" & Chr$(13) + Chr$(10)
   lblPlsWait(1).Caption = "While data is uploading: Properties"
   'WaitFewSec
   DoEvents

   Imp_AddPayDates oConnEx, oConnAcc

   Print #1, "Data writing FINISH: ********** Additional Due Date **********" & Chr$(13) + Chr$(10)

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ Demand Types ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
DEMAND_TYPE_MILESTONE:
If chkDemandType.Value Then
    'function has been written seperately by anol 2019-11-11
   'Importing DemandTypes
   Call imp_demandTypes(oConnEx, oConnAcc)
End If
   

    If chkChargeTypes.Value Then
        'function has been written seperately by anol 2021-08-31
        Call imp_ChargeTypes(oConnEx, oConnAcc)
    End If
    
      If chkPayableTypes.Value Then
        'function has been written seperately by anol 2021-08-31
        Call imp_PayableTypes(oConnEx, oConnAcc)
    End If
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ Unit ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
UNIT_MILESTONE:
If Not chkUnit.Value Then GoTo LESSEE_MILESTONE

   Print #1, "" & Chr$(13) + Chr$(10)
   Print #1, "" & Chr$(13) + Chr$(10)
   Print #1, "Data writing START: ********** Unit **********" & Chr$(13) + Chr$(10)
   lblPlsWait(1).Caption = "While data is uploading: Units"
   'WaitFewSec
   DoEvents

   Imp_Units oConnEx, oConnAcc

   Print #1, "Data writing FINISH: ********** Unit **********" & Chr$(13) + Chr$(10)

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ Lessee/Tenant ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
LESSEE_MILESTONE:
   If Not chkLessee.Value Then GoTo RENT_BUDGET_MILESTONE

   Print #1, "" & Chr$(13) + Chr$(10)
   Print #1, "" & Chr$(13) + Chr$(10)
   Print #1, "Data writing START: ********** Lessee **********" & Chr$(13) + Chr$(10)
   lblPlsWait(1).Caption = "While data is uploading: Lessee"
   'WaitFewSec
   DoEvents
   'added by anol 25 Jan 2015
   If (CheckLessee(oConnEx)) = True Then
      MsgBox "Import failed. Please check import log for details.", vbCritical + vbOKOnly, "Data Import"
      'Print #1, 0 / 0
      Exit Sub
   End If
   
   If Imp_Lessee(oConnEx, oConnAcc) = False Then
        GoTo ErrReport
   End If
   
   Print #1, "Data writing FINISH: ********** Lessee **********" & Chr$(13) + Chr$(10)

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ Rent Budget ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
RENT_BUDGET_MILESTONE:
If Not chkRentBudget.Value Then GoTo INSURANCE_BUDGET_MILESTONE

   Print #1, "" & Chr$(13) + Chr$(10)
   Print #1, "" & Chr$(13) + Chr$(10)
   Print #1, "Data writing START: ********** Rent Budget **********" & Chr$(13) + Chr$(10)
   lblPlsWait(1).Caption = "While data is uploading: Rent Bduget"
   'WaitFewSec
   DoEvents

   Imp_RentBduget oConnEx, oConnAcc

   Print #1, "Data writing FINISH: ********** Rent Budget **********" & Chr$(13) + Chr$(10)

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ Insurance Budget ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
INSURANCE_BUDGET_MILESTONE:
If Not chkInsuranceBudget.Value Then GoTo SC_BUDGET_MILESTONE

   Print #1, "" & Chr$(13) + Chr$(10)
   Print #1, "" & Chr$(13) + Chr$(10)
   Print #1, "Data writing START: ********** Insurance Budget **********" & Chr$(13) + Chr$(10)
   lblPlsWait(1).Caption = "While data is uploading: Insurance Bduget"
   'WaitFewSec
   DoEvents

   Imp_InsuranceBduget oConnEx, oConnAcc

   Print #1, "Data writing FINISH: ********** Insurance Budget **********" & Chr$(13) + Chr$(10)

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ Service Charge Budget ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
SC_BUDGET_MILESTONE:
If Not chkServiceBudget.Value Then GoTo LEASE_DETAILS_MILESTONE

   Print #1, "" & Chr$(13) + Chr$(10)
   Print #1, "" & Chr$(13) + Chr$(10)
   Print #1, "Data writing START: ********** Service Charge Budget **********" & Chr$(13) + Chr$(10)
   lblPlsWait(1).Caption = "While data is uploading: Service Charge Budget"
   'WaitFewSec
   DoEvents

   If Imp_ServiceChargeBduget(oConnEx, oConnAcc) = False Then
        GoTo ErrReport
   End If

   Print #1, "Data writing FINISH: ********** Service Charge Budget **********" & Chr$(13) + Chr$(10)

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ Lease Details ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
LEASE_DETAILS_MILESTONE:
If Not chkLease.Value Then GoTo SUPPLIERS_MILESTONE

   Print #1, "" & Chr$(13) + Chr$(10)
   Print #1, "" & Chr$(13) + Chr$(10)
   Print #1, "Data writing START: ********** Lease **********" & Chr$(13) + Chr$(10)
   lblPlsWait(1).Caption = "While data is uploading: Leases"
   'WaitFewSec
   DoEvents
   'issue 521
   'solved by anol 22 Jan 2015
'anol implement "need to check before writing on Rentcharges, Servicecharges and Insurancecharges if
'"UnitID. exists in [unit and leases] - UnitID." here

   If (checkRent_Service_insurence_charges(oConnEx)) = True Then
      MsgBox "Import failed. Please check import log for details.", vbCritical + vbOKOnly, "Data Import"
      'Print #1, 0 / 0
      Exit Sub
   End If

   szSQL = Imp_Lease(oConnEx, oConnAcc) ' inside this procedure you have rent review, rent charge, service charge and insurance charge
   If szSQL = "" Then
      Print #1, "Data writing FINISH: ********** Lease **********" & Chr$(13) + Chr$(10)
   Else
      Print #1, "Data writing STOPPED: ********** Lease **********" & Chr$(13) + Chr$(10)
      Print #1, "There are some duplicate Lessee ID found in the Lease import" & Chr$(13) + Chr$(10)
      Print #1, szSQL
      Print #1, 0 / 0
   End If

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ Supplier Details ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
SUPPLIERS_MILESTONE:
   If Not chkSuppliers.Value Then GoTo END_MILESTONE

   Call imp_suppler(oConnEx, oConnAcc)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

END_MILESTONE:
      If adoRstMdb.State = 1 Then
         adoRstMdb.Close
      End If
   Print #1, "####### Log file close ######### Date: " & Format(Now, "hh:mm:ss") & "" & Chr$(13) + Chr$(10)
   Call duplicationofID(oConnAcc)
   Close 1
  
   Set oRstXls = Nothing
   Set adoRstMdb = Nothing
   Set adoRst = Nothing
   oConnEx.Close
   Set oConnEx = Nothing
   oConnAcc.Close
   Set oConnAcc = Nothing

   lblPlsWait(0).Visible = False
   lblPlsWait(1).Visible = False
   lblPlsWait(2).Visible = False
   txtSourceFile.text = ""

   MsgBox "Data has been imported successfully. Thanks.", vbInformation + vbOKOnly, "Update Data"

   Exit Sub

ErrReport:
   If isSupplier Then
      isSupplier = False
      lblMessage(2).Caption = lblMessage(2).Caption & vbNewLine & " No Supplier sheet found in excel file!! Moving to next section.." & vbNewLine
      Print #1, "Data writing aborted : No Supplier sheet found in excel file!! Moving to next section.." & Chr$(13) + Chr$(10)
      GoTo CLIENTS_MILESTONE
   End If
   MsgBox "Import failed. Please check import log for details.", vbCritical + vbOKOnly, "Data Import"

   Print #1, "**********************************************" & Format(Now, "hh:mm:ss") & "" & Chr$(13) + Chr$(10)
   Print #1, "There is an error occured. All data has not been updated successfully." & Format(Now, "hh:mm:ss") & "" & Chr$(13) + Chr$(10)
   Print #1, "Error Description: " & Err.description & Chr$(13) + Chr$(10)
   Print #1, "Error Number: " & Err.Number & Chr$(13) + Chr$(10)
   Close 1
   lblPlsWait(0).Visible = False
   lblPlsWait(1).Visible = False
   lblPlsWait(2).Visible = False
'==========================================================================================================================================
'//////////////////////////////// update data model ////////////// DON'T DELETE ///////////////////////////////////////////////////////////
'------------------------------------------------------------------------------------------------------------------------------------------
''  the table name is the worksheet name
'   sTableNameEx = "SELECT * FROM [Clients$]"
''  Get the recordset
'   oRstXls.Open sTableNameEx, oConnEx, adOpenStatic, adLockOptimistic
''  Get the Access Database table
'   adoRstMdb.Open "SELECT * FROM Client", oConnAcc, adOpenDynamic, adLockOptimistic
'
'   Print #1, "" & Chr$(13) + Chr$(10)
'   Print #1, "" & Chr$(13) + Chr$(10)
'   Print #1, "Data writing START: ********** Clients **********" & Chr$(13) + Chr$(10)
'
'   With adoRstMdb
'      While Not oRstXls.EOF
'         .AddNew
'         .Fields.Item("").Value = oRstXls.Fields.Item().Value
'         .Update
'
'         Print #1, "Data - " & oRstXls.Fields.Item(0).Value & "" & Chr$(13) + Chr$(10)
'
'         oRstXls.MoveNext
'      Wend
'   End With
'   Print #1, "Data writing FINISH: ********** BankDetails **********" & Chr$(13) + Chr$(10)
'
'   oRstXls.Close
'   adoRstMdb.Close
'------------------------------------------------------------------------------------------------------------------------------------------
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'===========================================================================================================================================
End Sub
Private Function duplicationofID(Conn1 As ADODB.Connection)
    Dim Rst1 As New ADODB.Recordset
   ' Dim Conn1 As New ADODB.connection
    Dim szSQL As String
    'Conn1.Open getacc
DUPLICATE_ID_LCLSA:
   szSQL = "SELECT ID "
   szSQL = szSQL & "FROM (SELECT ID, COUNT(ID) AS C "
   szSQL = szSQL & "FROM ("
   szSQL = szSQL & "SELECT SupplierID AS ID "
   szSQL = szSQL & "FROM Supplier WHERE TYPE = 'SUPPLIER' UNION ALL "
   szSQL = szSQL & "SELECT ClientID AS ID "
   szSQL = szSQL & "FROM Client UNION ALL "
   szSQL = szSQL & "SELECT SageAccountNumber AS ID "
   szSQL = szSQL & "FROM Tenants UNION ALL "
   szSQL = szSQL & "SELECT AgentID AS ID "
   szSQL = szSQL & "FROM Agent UNION ALL "
   szSQL = szSQL & "SELECT LandlordID AS ID "
   szSQL = szSQL & "From Landlord "
   szSQL = szSQL & ") "
   szSQL = szSQL & " GROUP BY ID "
   szSQL = szSQL & ") "
   szSQL = szSQL & "WHERE C > 1;"

''Debug.Print szSQL
   'Debug.Print "DUPLICATE_ID_LCLSA" & time
   Rst1.Open szSQL, Conn1, adOpenStatic, adLockReadOnly
    'Debug.Print "DUPLICATE_ID_LCLSA" & time

   If Not Rst1.EOF Then                                  'Duplicate ID found
'      Conn1.Execute "DELETE Tenants.* " & _
'                    "FROM Tenants, Landlord " & _
'                    "WHERE Tenants.SageAccountNumber = Landlord.LandlordID;"
      Rst1.Close
      Rst1.Open szSQL, Conn1, adOpenStatic, adLockReadOnly
      If Not Rst1.EOF Then                                  'Duplicate ID found
         szSQL = SQL2String(Rst1, 0)
         Rst1.Close
        'keep a log in the error log table
        'Conn1.Execute "Insert into SpareTable5(ClientID,Code,CC) values('Login','" & Date & "' ,'he following ID(s) are duplicating(Lessee, Client, Landlord, Supplier):" & szSQL & "' )"
         MsgBox "The following ID(s) are duplicating: " & szSQL & ". Please contact with PCM Consulting.", vbCritical & vbOKOnly, "Data need to fix"
          Print #1, "The following ID(s) are duplicating: " & szSQL & ". Please contact with PCM Consulting." & Chr$(13) + Chr$(10)
      Else
         Rst1.Close
      End If
   Else
      Rst1.Close
   End If
End Function

Private Function imp_suppler(oConnEx As ADODB.Connection, oConnAcc As ADODB.Connection) As Boolean
    Dim oRstAccPL As New ADODB.Recordset
    Dim isEmptySheet As Boolean
    Dim sTableNameEx As String
    Dim oRstXls As New ADODB.Recordset
    Dim adoRst As New ADODB.Recordset
    Dim oNCsXls As New ADODB.Recordset
    Dim adoClient As New ADODB.Recordset
    Dim adoRstMdb As New ADODB.Recordset
    Dim adoRST_ As New ADODB.Recordset
   
   isSupplier = True
'  the table name is the worksheet name
   sTableNameEx = "SELECT * FROM [Suppliers$]"
'  Get the recordset
   oRstXls.Open sTableNameEx, oConnEx, adOpenStatic, adLockOptimistic
'  Get the Access Database table

   Print #1, "Data writing START: ********** Suppliers Details **********" & Chr$(13) + Chr$(10)
   lblPlsWait(1).Caption = "While data is uploading: Suppliers Details"
   'WaitFewSec
   DoEvents

   isEmptySheet = False
   Do While Not oRstXls.EOF
      If Not IsEmpty(oRstXls.Fields.Item("SUPPLIER ACCOUNT").Value) And Not IsNull(oRstXls.Fields.Item("SUPPLIER ACCOUNT").Value) Then
          With adoRstMdb
            .Open "SELECT * FROM Supplier where SupplierID='" & Trim(oRstXls.Fields.Item(0).Value) & "'", oConnAcc, adOpenDynamic, adLockOptimistic
            If .EOF = True Then .AddNew

            .Fields.Item("SupplierID").Value = Left(oRstXls.Fields.Item("SUPPLIER ACCOUNT").Value, 10)
            .Fields.Item("SupplierName").Value = oRstXls.Fields.Item("SUPPLIER NAME").Value
            .Fields.Item("SupplierAddressLine1").Value = oRstXls.Fields.Item("SUPPLIERADDRESSLINE1").Value
            .Fields.Item("SupplierAddressLine2").Value = oRstXls.Fields.Item("SUPPLIERADDRESSLINE2").Value
            .Fields.Item("SupplierAddressLine3").Value = oRstXls.Fields.Item("SUPPLIERADDRESSLINE3").Value
            .Fields.Item("SupplierAddressLine4").Value = oRstXls.Fields.Item("SUPPLIERADDRESSLINE4").Value
            .Fields.Item("SupplierPostCode").Value = oRstXls.Fields.Item("SUPPLIERPOSTCODE").Value
            .Fields.Item("SupplierOfficeEmail").Value = oRstXls.Fields.Item("SUPPLIERSTATEMENTEMAIL").Value
            .Fields.Item("SupplierPersonalEmail").Value = oRstXls.Fields.Item("SUPPLIEREMAIL").Value
            .Fields.Item("SupplierHomeTel").Value = oRstXls.Fields.Item("SUPPLIERHOMTEL").Value
            .Fields.Item("SupplierMobile").Value = oRstXls.Fields.Item("SUPPLIERMOBILE").Value
            .Fields.Item("SupplierOfficeAddressLine1").Value = oRstXls.Fields.Item("SUPPLIERSTATEMENTADDRESSLINE1").Value
            .Fields.Item("SupplierOfficeAddressLine2").Value = oRstXls.Fields.Item("SUPPLIERSTATEMENTADDRESSLINE2").Value
            .Fields.Item("SupplierOfficeAddressLine3").Value = oRstXls.Fields.Item("SUPPLIERSTATEMENTADDRESSLINE3").Value
            .Fields.Item("SupplierOfficeAddressLine4").Value = oRstXls.Fields.Item("SUPPLIERSTATEMENTADDRESSLINE4").Value
            .Fields.Item("SupplierOfficePostCode").Value = oRstXls.Fields.Item("SUPPLIERSTATEMENTPOSTCODE").Value
            .Fields.Item("SupplierOfficeTel").Value = oRstXls.Fields.Item("SUPPLIERSTATEMENTTEL").Value
            .Fields.Item("VATReg").Value = oRstXls.Fields.Item("VAT REG").Value
            .Fields.Item("SupplierType").Value = UpdateSupplierType(IIf(IsNull(oRstXls.Fields.Item("SUPPLIER TYPE").Value), "", oRstXls.Fields.Item("SUPPLIER TYPE").Value), oConnAcc)
            .Fields.Item("CreditLimit").Value = IIf(IsNull(oRstXls.Fields.Item("CREDIT LIMIT").Value), 0, oRstXls.Fields.Item("CREDIT LIMIT").Value)
            'Modified by anol
            '31 Dec 2014
            '.Fields.Item("VATCode").Value = GetVatCode(oRstXls.Fields.Item("VAT CODE").Value, oConnAcc)
            .Fields.Item("VATCode").Value = GetVatCode(IIf(IsNull(oRstXls.Fields.Item("TAX CODE").Value), "", oRstXls.Fields.Item("TAX CODE").Value), oConnAcc) 'oRstXls.Fields.Item("TAX CODE").Value
            .Fields.Item("AccountType").Value = Left(oRstXls.Fields.Item("ACCOUNTTYPE").Value, 10)

            .Fields.Item("PaymentType").Value = "CHQ" ' IIf(IsNull(oRstXls.Fields.Item("PAYMENTTYPE").Value) Or _
                                                    oRstXls.Fields.Item("PAYMENTTYPE").Value = "", "CHQ", _
                                                    oRstXls.Fields.Item("PAYMENTTYPE").Value)
            If IsNull(oRstXls.Fields.Item("PAYMENTTERMS").Value) Then
               .Fields.Item("PaymentTerms").Value = 0
            Else
               .Fields.Item("PaymentTerms").Value = Val(oRstXls.Fields.Item("PAYMENTTERMS").Value)
            End If
            .Fields.Item("SortCode").Value = Left(oRstXls.Fields.Item("BANK SORT CODE").Value, 8)
            .Fields.Item("AcNo").Value = Left(oRstXls.Fields.Item("BANK ACCOUNT NUMBER").Value, 8)
            .Fields.Item("AcName").Value = oRstXls.Fields.Item("BANK ACCOUNT NAME").Value
            .Fields.Item("BPR").Value = oRstXls.Fields.Item("BANK PAYMENT REFERENCE").Value
            If IsNull(oRstXls.Fields.Item("PURCHASE LEDGER CONTROL").Value) Then
               .Fields.Item("PLControl").Value = 0
            Else
               .Fields.Item("PLControl").Value = Val(oRstXls.Fields.Item("PURCHASE LEDGER CONTROL").Value)
            End If
            .Fields.Item("TYPE").Value = "SUPPLIER"
            If Not IsNull(oRstXls.Fields.Item("PURCHASE LEDGER CONTROL").Value) Then
               oRstAccPL.Open "SELECT Name FROM NominalLedger " & _
                              "WHERE code = '" & Trim(oRstXls.Fields.Item("PURCHASE LEDGER CONTROL").Value) & "';", oConnAcc, adOpenDynamic, adLockOptimistic
               If Not oRstAccPL.EOF Then
                  .Fields.Item("PLControlName").Value = oRstAccPL.Fields("Name")
               End If
               oRstAccPL.Close
            End If
            .Update
         End With
         Print #1, "Data - " & oRstXls.Fields.Item(1).Value & "" & Chr$(13) + Chr$(10)
         adoRstMdb.Close
      Else
         Print #1, "Supplier Import - No data found in sheet" & Chr$(13) + Chr$(10)
         GoTo ENDSUPPLIER
      End If

      oRstXls.MoveNext
      If oRstXls.EOF Then Exit Do
      If oRstXls.Fields.Item(0).Value = "END" Then Exit Do
   Loop
   Set oRstAccPL = Nothing

ENDSUPPLIER:
   Print #1, "Data writing FINISH: ********** Suppliers Details **********" & Chr$(13) + Chr$(10)
   oRstXls.Close
   isSupplier = False
End Function

Private Function NextID(adoConn As ADODB.Connection) As Long
   Dim szSQL As String
   Dim adoRst As New ADODB.Recordset
   szSQL = "SELECT MAX(Cint(conBankID))+1 AS Ref FROM ConsolidatedBankList;"
   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
        NextID = IIf(adoRst.EOF, 1, IIf(IsNull(adoRst!ref), 1, adoRst!ref))
   adoRst.Close
   Set adoRst = Nothing
End Function
Private Function imp_ConsolidatedBank(oConnEx As ADODB.Connection, oConnAcc As ADODB.Connection) As Boolean
 'Get the recordset from Excel
'    On Error GoTo ERR
    Dim sTableNameEx As String
    Dim oRstXls As New ADODB.Recordset
    Dim adoRst As New ADODB.Recordset
    Dim oNCsXls As New ADODB.Recordset
    Dim adoClient As New ADODB.Recordset
    Dim adoRstMdb As New ADODB.Recordset
    Dim adoRST_ As New ADODB.Recordset
    
   oRstXls.Open "SELECT * FROM [ConsolBankDetails$]", oConnEx, adOpenStatic, adLockOptimistic
   
'  Get the Access Database table
   adoRst.Open "SELECT * FROM ConsolidatedBankList", oConnAcc, adOpenDynamic, adLockOptimistic

   Print #1, "" & Chr$(13) + Chr$(10)
   Print #1, "" & Chr$(13) + Chr$(10)
   Print #1, "Data writing START: ********** Client's Banks **********" & Chr$(13) + Chr$(10)
   lblPlsWait(1).Caption = "While data is uploading: Client's Bank"
   'WaitFewSec
   DoEvents

   With adoRstMdb
      Do While Not oRstXls.EOF
         If Not IsEmpty(oRstXls.Fields.Item("BANK CODE").Value) And _
                  Not IsNull(oRstXls.Fields.Item("BANK CODE").Value) Then
            
                  .Open "SELECT * FROM ConsolidatedBankList " & _
                  "WHERE BankCode = '" & oRstXls.Fields.Item("BANK CODE").Value & "' AND " & _
                        "BankACNumber = '" & oRstXls.Fields.Item("BANK ACCOUNT NUMBER").Value & "' AND " & _
                        "SortCode = '" & Replace(oRstXls.Fields.Item("BANK SORT CODE").Value, "'", "''") & "';", oConnAcc, adOpenDynamic, adLockOptimistic
                  
                     If .EOF Then
                            .AddNew
                            .Fields.Item("conBankID").Value = NextID(oConnAcc)
                     End If
                            .Fields.Item("BankCode").Value = oRstXls.Fields.Item("BANK CODE").Value
'
'                    If adoRst.EOF Then
'                            MsgBox "Import failed. Bank ID was not found for the following sortCode in tlbBank :  " & oRstXls.Fields.Item("BANK CODE").Value & vbCrLf & " on excel sheet ConsolBankDetails$", vbOKOnly, "Importing BankAccountDetails!!"
'                            Exit Function
'                    End If
                    .Fields.Item("BankName").Value = oRstXls.Fields.Item("BANK NAME").Value
                    .Fields.Item("BankACNumber").Value = oRstXls.Fields.Item("BANK ACCOUNT NUMBER").Value
                    .Fields.Item("SortCode").Value = oRstXls.Fields.Item("BANK SORT CODE").Value
'                    If "YES" = UCase(oRstXls.Fields.Item("CONSOLIDATED").Value) Then
'                             .Fields.Item("ConsBankACNumber").Value = oRstXls.Fields.Item("BANK ACCOUNT NUMBER").Value
'                             .Fields.Item("ConsSortCode").Value = oRstXls.Fields.Item("BANK SORT CODE").Value
'                             .Fields.Item("conBankCode").Value = oRstXls.Fields.Item("BANK CODE").Value
'                             .Fields.Item("conBankReadOnly").Value = 1
'                    End If
                    .Update
                    .Close
        
                    Print #1, "Data - " & oRstXls.Fields.Item("BANK NAME").Value & ", A/C Num:" & oRstXls.Fields("BANK ACCOUNT NUMBER").Value & "" & Chr$(13) + Chr$(10)
                    
         End If

         oRstXls.MoveNext
         If oRstXls.EOF Then Exit Do
         If oRstXls.Fields.Item(0).Value = "END" Then Exit Do
      Loop
   End With
   oRstXls.Close
   adoRst.Close
   'Now update consolidated bank extra fields
   'join query from two excel sheet and then with the  result you update tlbclientBanks
   Dim my_ID As String
   Dim ConsBankACNumber As String
   oRstXls.Open "SELECT *,C.[BANK ACCOUNT NUMBER] as BAN,C.[BANK SORT CODE] AS BSC FROM [ConsolBankDetails$] as C INNER JOIN  [BankAccountDetails$] AS B ON B.[BANK ACCOUNT NUMBER] =C.[BANK ACCOUNT NUMBER] AND " & _
                        "B.[BANK SORT CODE] =C.[BANK SORT CODE] where B.[CONSOLIDATED]='YES'", oConnEx, adOpenStatic, adLockOptimistic
    While Not oRstXls.EOF
        adoRst.Open "Select conBankID,BankACNumber from ConsolidatedBankList where BankCode ='" & oRstXls.Fields.Item("BANK CODE").Value & "' AND  BankACNumber='" & _
                    oRstXls.Fields.Item("BAN").Value & "' AND SortCode='" & oRstXls.Fields.Item("BSC").Value & "' ", oConnAcc, adOpenDynamic, adLockOptimistic
        If Not adoRst.EOF Then
            my_ID = adoRst("conBankID").Value
            ConsBankACNumber = adoRst("BankACNumber").Value
            oConnAcc.Execute "Update tlbClientBanks set ConsolidatedBankID='" & my_ID & "' , ConsBankACNumber='" & ConsBankACNumber & "', ConsSortCode='" & oRstXls.Fields.Item("BANK SORT CODE").Value & "' " & _
                ",conBankCode='" & oRstXls.Fields.Item("BANK CODE").Value & "' where BANK_AC_NUM='" & _
                oRstXls.Fields.Item("BANK ACCOUNT NUMBER").Value & "' AND BANK_SC ='" & oRstXls.Fields.Item("BANK SORT CODE").Value & "'"
        End If
        adoRst.Close
        oRstXls.MoveNext
    Wend
oRstXls.Close
   oConnAcc.Execute "UPDATE tlbClientBanks SET Spare1 = MY_ID;"
   imp_ConsolidatedBank = True
   Print #1, "Data writing FINISH: ********** Consolidate Banks **********" & Chr$(13) + Chr$(10)
    Exit Function
Err:
    Print #1, Err.description & Chr$(13) + Chr$(10)
    MsgBox "err.Description"
End Function

Private Function imp_ClientBank(oConnEx As ADODB.Connection, oConnAcc As ADODB.Connection) As Boolean
      'Get the recordset from Excel
    Dim sTableNameEx As String
    Dim oRstXls As New ADODB.Recordset
    Dim adoRst As New ADODB.Recordset
    Dim oNCsXls As New ADODB.Recordset
    Dim adoClient As New ADODB.Recordset
    Dim adoRstMdb As New ADODB.Recordset
    Dim adoRST_ As New ADODB.Recordset
'    On Error GoTo ERR
   oRstXls.Open "SELECT * FROM [BankAccountDetails$]", oConnEx, adOpenStatic, adLockOptimistic
   
'  Get the Access Database table
   adoRst.Open "SELECT * FROM tlbBank", oConnAcc, adOpenDynamic, adLockOptimistic

   Print #1, "" & Chr$(13) + Chr$(10)
   Print #1, "" & Chr$(13) + Chr$(10)
   Print #1, "Data writing START: ********** Client's Banks **********" & Chr$(13) + Chr$(10)
   lblPlsWait(1).Caption = "While data is uploading: Client's Bank"
   'WaitFewSec
   DoEvents
   '******************* check deafaultbank is twice true for a client
    adoRST_.Open "SELECT COUNT([DEFAULT BANK]) as C,[CLIENT ID]  FROM  [BankAccountDetails$] " & _
                                 " Where [DEFAULT BANK]='YES' Group BY [CLIENT ID] ", oConnEx, adOpenStatic, adLockReadOnly
                    While Not adoRST_.EOF
                       If adoRST_.Fields.Item("C").Value > 1 Then
                            MsgBox "Client : " & adoRST_.Fields.Item("CLIENT ID").Value & " have more than one default bank"
                           Print #1, "Client : " & adoRST_.Fields.Item("CLIENT ID").Value & " have more than one default bank" & Chr$(13) + Chr$(10)
                       End If
                       adoRST_.MoveNext
                    Wend
                    adoRST_.Close
                    
                    
   With adoRstMdb
      Do While Not oRstXls.EOF
         If Not IsEmpty(oRstXls.Fields.Item("CLIENT ID").Value) And _
                  Not IsNull(oRstXls.Fields.Item("CLIENT ID").Value) Then
                  'Below appostrophe problem has been fixed by anol 04 Oct 2015
                  If "JIANGCAI" = oRstXls.Fields.Item("CLIENT ID").Value Then
                    Debug.Print ""
                  End If
            
            .Open "SELECT * FROM tlbClientBanks " & _
                  "WHERE CLIENT_ID = '" & oRstXls.Fields.Item("CLIENT ID").Value & "' AND " & _
                        "BANK_AC_NUM = '" & oRstXls.Fields.Item("BANK ACCOUNT NUMBER").Value & "' AND " & _
                        "Bank_AC_Name = '" & Replace(oRstXls.Fields.Item("BANK ACCOUNT NAME").Value, "'", "''") & "' AND NominalCode ='" & Replace(oRstXls.Fields.Item("NOMINAL CODE").Value, "'", "''") & "';", oConnAcc, adOpenDynamic, adLockOptimistic

           ' .Close
'            .Open "SELECT * FROM tlbClientBanks " & _
'                  "WHERE CLIENT_ID = '" & oRstXls.Fields.Item("CLIENT ID").Value & "' AND " & _
'                        "BANK_AC_NUM = '" & oRstXls.Fields.Item("BANK ACCOUNT NUMBER").Value & "' AND " & _
'                        "Bank_AC_Name = '" & Replace(oRstXls.Fields.Item("BANK ACCOUNT NAME").Value, "'", "''") & "' AND  ", _
'                        "NominalCode ='" & Replace(oRstXls.Fields.Item("NOMINAL CODE").Value, "'", "''") & "';", oConnAcc, adOpenDynamic, adLockOptimistic


' .Open "SELECT * FROM tlbClientBanks " & _
'                  "WHERE Bank_AC_Name = '" & Replace(oRstXls.Fields.Item("BANK ACCOUNT NAME").Value, "'", "''") & "'" & _
'                        " " & _
'                        "", _
'                        "", oConnAcc, adOpenDynamic, adLockOptimistic
                 

                  'Replace(oRstXls.Fields.Item("BANK ACCOUNT NAME").Value, "''", "'")
                     If .EOF Then .AddNew
                    .Fields.Item("CLIENT_ID").Value = oRstXls.Fields.Item("CLIENT ID").Value
                    adoRst.Find "SORT_CODE = '" & oRstXls.Fields.Item("BANK SORT CODE").Value & "' ", , , 1
                    'validation warning added by anol 27 May 2016
                    If adoRst.EOF Then
                        MsgBox "Import failed. Bank ID was not found for the following sortCode in tlbBank :  " & oRstXls.Fields.Item("BANK SORT CODE").Value & vbCrLf & " on excel sheet BankAccountDetails$", vbOKOnly, "Importing BankAccountDetails!!"
                        Exit Function
                    End If
                    If Not adoRst.EOF Then .Fields.Item("BANK_ID").Value = adoRst.Fields.Item("BANK_ID").Value
                    
                    .Fields.Item("Bank_AC_Name").Value = oRstXls.Fields.Item("BANK ACCOUNT NAME").Value
        'If oRstXls.Fields.Item("CLIENT ID").Value = "FRMC" Then
        '   Print #1, "Data - " & oRstXls.Fields.Item("BANK ACCOUNT NAME").Value & Chr$(13) + Chr$(10) & _
        '             ", A/C Num: " & oRstXls.Fields.Item(2).Value & "" & Chr$(13) + Chr$(10) & _
        '             .Fields.Item("BANK_AC_NUM").Value
        'End If
                    .Fields.Item("BANK_AC_NUM").Value = CStr(oRstXls.Fields.Item("BANK ACCOUNT NUMBER").Value)
                    .Fields.Item("BANK_SC").Value = oRstXls.Fields.Item("BANK SORT CODE").Value
                    'added by anol 2021-06-15
                    .Fields.Item("CONSOLIDATED").Value = IIf(UCase(oRstXls.Fields.Item("CONSOLIDATED").Value) = "YES", True, False)
                    adoRST_.Open "SELECT COUNT(CLIENT_ID) AS C FROM tlbClientBanks " & _
                                 "WHERE CLIENT_ID = '" & oRstXls.Fields.Item("CLIENT ID").Value & "' " & _
                                 "GROUP BY CLIENT_ID;", oConnAcc, adOpenStatic, adLockReadOnly
                    If adoRST_.EOF Then
                       .Fields.Item("DEFAULT_AC").Value = True
                    Else
                       .Fields.Item("DEFAULT_AC").Value = IIf(UCase(oRstXls.Fields.Item("DEFAULT BANK").Value) = "YES", True, False)
                    End If
                    adoRST_.Close
                    If IsNull(oRstXls.Fields.Item("NOMINAL CODE").Value) Then
                        MsgBox "Import failed.Nominal Code is empty on excel sheet BankAccountDetails$", vbOKOnly, "Warning!!"
                        Exit Function
                    End If
                    .Fields.Item("NominalCode").Value = oRstXls.Fields.Item("NOMINAL CODE").Value
                    .Fields.Item("PaymentMethod").Value = oRstXls.Fields.Item("PAYMENT METHOD").Value
                    .Fields.Item("BacsRef").Value = oRstXls.Fields.Item("BACS REFERENCE").Value
                    .Update
                    .Close
        
                    Print #1, "Data - " & oRstXls.Fields.Item("BANK ACCOUNT NAME").Value & ", A/C Num:" & oRstXls.Fields.Item(2).Value & "" & Chr$(13) + Chr$(10)
         End If

         oRstXls.MoveNext
         If oRstXls.EOF Then Exit Do
         If oRstXls.Fields.Item(0).Value = "END" Then Exit Do
      Loop
   End With
   oRstXls.Close
   adoRst.Close

   oConnAcc.Execute "UPDATE tlbClientBanks SET Spare1 = MY_ID;"
   imp_ClientBank = True
   Print #1, "Data writing FINISH: ********** Client's Banks **********" & Chr$(13) + Chr$(10)
   Exit Function
Err:
   Print #1, Err.description & Chr$(13) + Chr$(10)
   MsgBox Err.description
End Function
Private Function import_NominalCode(oConnEx As ADODB.Connection, oConnAcc As ADODB.Connection) As Boolean
    Dim sTableNameEx As String
    Dim oRstXls As New ADODB.Recordset
    Dim adoRst As New ADODB.Recordset
    Dim oNCsXls As New ADODB.Recordset
    Dim adoClient As New ADODB.Recordset
    Dim adoRstMdb As New ADODB.Recordset
    Dim adoRST_ As New ADODB.Recordset
    Dim adoRsDuplicateCodeinClient As New ADODB.Recordset
    Dim rsControlAccountbyClient As New ADODB.Recordset
    Print #1, "Data writing START: ******** NLSubTypes ******************" & Chr$(13) + Chr$(10)
   lblPlsWait(1).Caption = "While data is uploading: Nominal Code"
   'WaitFewSec
   DoEvents
'  IMPORT/UPDATE ~~~~~~~~~~~~~~~~~~~~~~~~ NLSubTypes table ~~~~~~~~~~~~~~~~~~~~~~~~

   sTableNameEx = "SELECT SUBTYPE FROM [NominalCodes$] GROUP BY [SUBTYPE];"
'  get the recordset in the excel file
   oRstXls.Open sTableNameEx, oConnEx, adOpenStatic, adLockOptimistic

   With adoRst
      .Open "SELECT * FROM NLSubTypes", oConnAcc, adOpenDynamic, adLockOptimistic

      Do While Not oRstXls.EOF
         If Not IsNull(oRstXls.Fields.Item("SUBTYPE").Value) Then
            .Find "STName = '" & oRstXls.Fields.Item("SUBTYPE").Value & "' ", , , 1
            If .EOF And Not IsNull(oRstXls.Fields.Item("SUBTYPE").Value) Then
               .AddNew
               .Fields.Item("STCode").Value = CreateSubTypeID(oRstXls.Fields.Item("SUBTYPE").Value)
               .Fields.Item("STName").Value = oRstXls.Fields.Item("SUBTYPE").Value
               .Fields.Item("STDescription").Value = oRstXls.Fields.Item("SUBTYPE").Value
               .Update
            End If
         End If
         oRstXls.MoveNext
         If oRstXls.EOF Then Exit Do
         If oRstXls.Fields.Item(0).Value = "END" Then Exit Do
      Loop
'     .close      'Dont close this recordset
   End With
   oRstXls.Close

   Print #1, "Data writing START: ******** NominalCode ******************" & Chr$(13) + Chr$(10)
   lblPlsWait(1).Caption = "While data is uploading: Nominal Code"
   'WaitFewSec
   DoEvents
   'We dont need this we are handling defaulat when no client specified
'''                                Now Nominal Code is loading      ~~~~~
''   oNCsXls.Open "SELECT * FROM [NominalCodes$] WHERE [CLIENT ID] = 'DEFAULT';", oConnEx, adOpenStatic, adLockOptimistic
''   If Not oNCsXls.EOF Then
''           oConnAcc.Execute "Delete FROM NominalLedger WHERE ClientID='NONE'"
''           Dim rsCheck As New ADODB.Recordset
''           rsCheck.Open "SELECT * FROM NominalLedger WHERE 1=2", oConnAcc, adOpenDynamic, adLockOptimistic
''           With rsCheck
''           Do While Not oNCsXls.EOF
''                   .AddNew
''                   .Fields.Item("Code").Value = oNCsXls.Fields.Item("CODE").Value
''                   .Fields.Item("Name").Value = IIf(IsNull(oNCsXls.Fields.Item("NOMINAL NAME").Value), "Not defined", oNCsXls.Fields.Item("NOMINAL NAME").Value)
''                   .Fields.Item("Type").Value = IIf(IsNull(oNCsXls.Fields.Item("TYPE").Value), 0, _
''                      IIf(UCase(Left(oNCsXls.Fields.Item("TYPE").Value, 1)) = "B", 1, IIf(UCase(Left(oNCsXls.Fields.Item("TYPE").Value, 1)) = "P", 2, 0)))
''                   .Fields.Item("DrCr").Value = IIf(IsNull(oNCsXls.Fields.Item("CREDIT OR DEBIT").Value), "", _
''                      IIf(UCase(oNCsXls.Fields.Item("CREDIT OR DEBIT").Value) = "DEBIT", "Dr", "Cr"))
''                   .Fields.Item("ClientID").Value = "NONE" 'adoClient.Fields.Item("ClientID").Value
''                   'added by anol 2021-06-15
''                   .Fields.Item("Posting").Value = IIf(UCase(oNCsXls.Fields.Item("POSTING ALLOWED").Value = "YES"), True, False)
''                   adoRST.Find "STName = '" & oNCsXls.Fields.Item("SUBTYPE").Value & "' ", , , 1
''                   .Fields.Item("SubType").Value = IIf(IsNull(adoRST.Fields.Item("STCode").Value), "", adoRST.Fields.Item("STCode").Value)
''                   'issue 505: Import module needs to be modified
''                    'added by anol 10 Feb 2015
''
''                   If IsNull(oNCsXls.Fields.Item("CONTROL ACCOUNT").Value) = True Then
''                      .Fields.Item("CAFixed").Value = "0"
''                      '.Fields.Item("CAName").Value = ""
''                      '.Fields.Item("CADisorder").Value = ""
''                      'added by anol issue 521 note 1
''                      'Date 25 Jan 2015
''                       .Fields.Item("CAType").Value = ""
''                   Else
''                      .Fields.Item("CAType").Value = ""
''                      .Fields.Item("CAFixed").Value = "-1"
''                      If UCase(oNCsXls.Fields.Item("CONTROL ACCOUNT").Value) = "SALES LEDGER CONTROL" Then
''                           .Fields.Item("CAName").Value = "Sales Ledger Control"
''                           .Fields.Item("CAType").Value = "S"
''                           .Fields.Item("CADisorder").Value = "1"
''                      ElseIf UCase(oNCsXls.Fields.Item("CONTROL ACCOUNT").Value) = "PURCHASE LEDGER CONTROL" Then
''                           .Fields.Item("CAName").Value = "Purchase Ledger Control"
''                           .Fields.Item("CAType").Value = "P"
''                           .Fields.Item("CADisorder").Value = "2"
''                      ElseIf UCase(oNCsXls.Fields.Item("CONTROL ACCOUNT").Value) = "INPUT VAT" Then
''                          .Fields.Item("CAName").Value = "Input VAT"
''                          .Fields.Item("CAType").Value = "I"
''                          .Fields.Item("CADisorder").Value = "3"
''                      ElseIf UCase(oNCsXls.Fields.Item("CONTROL ACCOUNT").Value) = "OUTPUT VAT" Then
''                           .Fields.Item("CAName").Value = "Output VAT"
''                           .Fields.Item("CAType").Value = "O"
''                           .Fields.Item("CADisorder").Value = "4"
''                      ElseIf UCase(oNCsXls.Fields.Item("CONTROL ACCOUNT").Value) = "RETAINED EARNINGS" Then
''                         .Fields.Item("CAName").Value = "Retained Earnings"
''                         .Fields.Item("CAType").Value = "R"
''                         .Fields.Item("CADisorder").Value = "5"
''                      End If
''                    End If
''                   'End of modification
''
''                   .Update
''
''                   'NJ_CC table needs to updated here before go to next row
''        '           adoRST_.Open "SELECT * FROM NJ_CC WHERE ClientID = '" & adoClient.Fields.Item("ClientID").Value & "' AND " & _
''        '                                                  "Code     = '" & oNCsXls.Fields.Item("CODE").Value & "' AND " & _
''        '                                                  "CC       = '" & oNCsXls.Fields.Item("CATEGORY CODE").Value & "' ", _
''        '                   oConnAcc, adOpenDynamic, adLockOptimistic
''        '           If adoRST_.EOF Then adoRST_.AddNew
''        '           adoRST_.Fields.Item("ClientID").Value = "NONE"
''        '           adoRST_.Fields.Item("Code").Value = oNCsXls.Fields.Item("CODE").Value
''        '           adoRST_.Fields.Item("CC").Value = oNCsXls.Fields.Item("CATEGORY CODE").Value
''        '           adoRST_.Update
''        '           adoRST_.Close
''                   oNCsXls.MoveNext
''                   If oNCsXls.EOF Then Exit Do
''                   If oNCsXls.Fields.Item(0).Value = "END" Then Exit Do
''                Loop
''            End With
''        End If
''  oNCsXls.Close
''         'End of Modification 16 Feb 2015
  
    
'Below code was modified by anol 24 Sep 2015
'PPM import amendment
    'Now for ashwood it will not modify the existing nominal code (PPM data) it will just amend for the new cleint
    '1/ The nominal ledger import should update the client ID None from the upload sheet. It should then only update those existing client
    'IDs contained in the upload sheet. Any new clients should have a nominal codes based on the upload sheet and appended to the existing nominal table.
'   oNCsXls.Open "SELECT [CLIENT ID] FROM [NominalCodes$] GROUP BY [CLIENT ID];", oConnEx, adOpenDynamic, adLockReadOnly
'   szTemp = SQL2String(oNCsXls, 0)
'   oNCsXls.Close

'was not here originally wriiten by anol
'   oNCsXls.Open "SELECT [CLIENT ID] FROM [Clients$] GROUP BY [CLIENT ID];", oConnEx, adOpenDynamic, adLockReadOnly
'   szTemp = SQL2String(oNCsXls, 0)
'   oNCsXls.Close
'comment out by anol for ashwood it will not modify the existing nominal code (PPM data)
  ' adoClient.Open "SELECT * FROM Client WHERE ClientID NOT IN (" & szTemp & ");", oConnAcc, adOpenDynamic, adLockOptimistic
  Dim strmsg As String
  Dim strmsg1 As String
'   adoClient.Open "SELECT [CLIENT ID] as ClientID FROM [Clients$] where [CLIENT ID] is not null;", oConnEx, adOpenStatic, adLockReadOnly
'     If Not adoClient.EOF Then
'                rsControlAccountbyClient.Open "SELECT CODE,[CLIENT ID] from [NominalCodes$] where CODE is not null and " & _
'                   "UCASE([CONTROL ACCOUNT]) in ('SALES LEDGER CONTROL') AND [CLIENT ID]='" & adoClient.Fields.Item("ClientID").Value & "' ", oConnEx, adOpenDynamic, adLockOptimistic
'                If rsControlAccountbyClient.EOF Then
'                     Print #1, "The main Control Accounts needs to be specified for client and control is : " & adoClient.Fields.Item("ClientID").Value & ", 'SALES LEDGER CONTROL' " & Chr(13) + Chr$(10)
'                     strmsg1 = strmsg1 & "The main Control Accounts needs to be specified for client and control is : " & adoClient.Fields.Item("ClientID").Value & ", 'SALES LEDGER CONTROL' " & Chr(13) + Chr$(10)
'                End If
'                rsControlAccountbyClient.Close
'                 rsControlAccountbyClient.Open "SELECT CODE,[CLIENT ID] from [NominalCodes$] where CODE is not null and " & _
'                   "UCASE([CONTROL ACCOUNT]) in ('PURCHASE LEDGER CONTROL') AND [CLIENT ID]='" & adoClient.Fields.Item("ClientID").Value & "' ", oConnEx, adOpenDynamic, adLockOptimistic
'                If rsControlAccountbyClient.EOF Then
'                     Print #1, "The main Control Accounts needs to be specified for client and control is : " & adoClient.Fields.Item("ClientID").Value & ",  'PURCHASE LEDGER CONTROL' " & Chr(13) + Chr$(10)
'                     strmsg1 = strmsg1 & "The main Control Accounts needs to be specified for client and control is : " & adoClient.Fields.Item("ClientID").Value & ",  'PURCHASE LEDGER CONTROL' " & Chr(13) + Chr$(10)
'                End If
'                rsControlAccountbyClient.Close
'                 rsControlAccountbyClient.Open "SELECT CODE,[CLIENT ID] from [NominalCodes$] where CODE is not null and " & _
'                   "UCASE([CONTROL ACCOUNT]) in ('INPUT VAT' ) AND [CLIENT ID]='" & adoClient.Fields.Item("ClientID").Value & "' ", oConnEx, adOpenDynamic, adLockOptimistic
'                   If rsControlAccountbyClient.EOF Then
'                         Print #1, "The main Control Accounts needs to be specified for client and control is : " & adoClient.Fields.Item("ClientID").Value & ",  'INPUT VAT' " & Chr(13) + Chr$(10)
'                         strmsg1 = strmsg1 & "The main Control Accounts needs to be specified for client and control is : " & adoClient.Fields.Item("ClientID").Value & ",  'INPUT VAT' " & Chr(13) + Chr$(10)
'                   End If
'                rsControlAccountbyClient.Close
'                 rsControlAccountbyClient.Open "SELECT CODE,[CLIENT ID] from [NominalCodes$] where CODE is not null and " & _
'                   "UCASE([CONTROL ACCOUNT]) in (" & _
'                   "'OUTPUT VAT' ) AND [CLIENT ID]='" & adoClient.Fields.Item("ClientID").Value & "' ", oConnEx, adOpenDynamic, adLockOptimistic
'                   If rsControlAccountbyClient.EOF Then
'                         Print #1, "The main Control Accounts needs to be specified for client and control is : " & adoClient.Fields.Item("ClientID").Value & ",  'OUTPUT VAT' " & Chr(13) + Chr$(10)
'                         strmsg1 = strmsg1 & "The main Control Accounts needs to be specified for client and control is : " & adoClient.Fields.Item("ClientID").Value & ", 'OUTPUT VAT' " & Chr(13) + Chr$(10)
'                End If
'                rsControlAccountbyClient.Close
'                 rsControlAccountbyClient.Open "SELECT CODE,[CLIENT ID] from [NominalCodes$] where CODE is not null and " & _
'                   "UCASE([CONTROL ACCOUNT]) in ( " & _
'                   " 'RETAINED EARNINGS' ) AND [CLIENT ID]='" & adoClient.Fields.Item("ClientID").Value & "' ", oConnEx, adOpenDynamic, adLockOptimistic
'                   If rsControlAccountbyClient.EOF Then
'                         Print #1, "The main Control Accounts needs to be specified for client and control is : " & adoClient.Fields.Item("ClientID").Value & ", 'RETAINED EARNINGS' " & Chr(13) + Chr$(10)
'                         strmsg1 = strmsg1 & "The main Control Accounts needs to be specified for client and control is : " & adoClient.Fields.Item("ClientID").Value & ",  'RETAINED EARNINGS' " & Chr(13) + Chr$(10)
'                End If
'                rsControlAccountbyClient.Close
                
             adoRsDuplicateCodeinClient.Open "SELECT CODE,[CLIENT ID], count(*) as cnt from [NominalCodes$] where CODE is not null  group by " & _
                "CODE,[CLIENT ID] having count(*)>1 ORDER BY [CLIENT ID],CODE", oConnEx, adOpenDynamic, adLockOptimistic
'                If adoClient.Fields.Item("ClientID").Value = "KRCAPITAL" Then
'                    Debug.Print ""
'                End If
             While Not adoRsDuplicateCodeinClient.EOF
                    If adoRsDuplicateCodeinClient("cnt").Value > 1 Then
                        strmsg = strmsg & adoRsDuplicateCodeinClient.Fields.Item("CLIENT ID").Value & ", " & adoRsDuplicateCodeinClient("CODE").Value & "  "
                        'adoRsDuplicateCodeinClient.Close
'                        Exit Do
                    End If
                    adoRsDuplicateCodeinClient.MoveNext
            Wend
            adoRsDuplicateCodeinClient.Close
            If Len(strmsg) > 0 Then
                    Print #1, "Duplicate nominal code for this client and code is : " & Chr(13) + Chr$(10) & strmsg & Chr$(13) + Chr$(10)
                    MsgBox "Duplicate nominal code for this client and code is : " & Chr(13) + Chr$(10) & strmsg, vbInformation, "Warning"
            End If
            If Len(strmsg1) > 0 Then
                MsgBox strmsg1
                Exit Function
            End If
            ' + Chr$(10)
            strmsg1 = ""
'     End If
     If Len(strmsg) > 0 Then
        Exit Function
     End If
  adoClient.Open "SELECT [CLIENT ID] as ClientID FROM [Clients$] where [CLIENT ID] is not null;", oConnEx, adOpenStatic, adLockReadOnly
  Do While Not adoClient.EOF
'      adoClient.MoveFirst
                  'Below part is added by anol 20161209
'                  adoClient.MovePrevious
        ' one error is coming for duplicate  in the
                oNCsXls.Open "SELECT * FROM [NominalCodes$] WHERE [CLIENT ID] ='" & adoClient.Fields.Item("ClientID").Value & "';", oConnEx, adOpenStatic, adLockOptimistic
                If adoClient.Fields.Item("ClientID").Value = "SEGALL" Then
                                 Debug.Print ""
                            End If
                If oNCsXls.EOF Then
                    oNCsXls.Close
                    'This is the validation part for control account has been set for default
                    oNCsXls.Open "SELECT * FROM [NominalCodes$] WHERE [CLIENT ID] = 'DEFAULT';", oConnEx, adOpenStatic, adLockOptimistic
                                rsControlAccountbyClient.Open "SELECT CODE,[CLIENT ID] from [NominalCodes$] where CODE is not null and " & _
                               "UCASE([CONTROL ACCOUNT]) in ('SALES LEDGER CONTROL') AND [CLIENT ID]='DEFAULT' ", oConnEx, adOpenDynamic, adLockOptimistic
                            If rsControlAccountbyClient.EOF Then
                                 Print #1, "The main Control Accounts needs to be specified for client and control is : " & adoClient.Fields.Item("ClientID").Value & ", 'SALES LEDGER CONTROL' " & Chr(13) + Chr$(10)
                                 strmsg1 = strmsg1 & "The main Control Accounts needs to be specified for client and control is : " & adoClient.Fields.Item("ClientID").Value & ", 'SALES LEDGER CONTROL' " & Chr(13) + Chr$(10)
                            End If
                            rsControlAccountbyClient.Close
                             rsControlAccountbyClient.Open "SELECT CODE,[CLIENT ID] from [NominalCodes$] where CODE is not null and " & _
                               "UCASE([CONTROL ACCOUNT]) in ('PURCHASE LEDGER CONTROL') AND [CLIENT ID]='DEFAULT' ", oConnEx, adOpenDynamic, adLockOptimistic
                            If rsControlAccountbyClient.EOF Then
                                 Print #1, "The main Control Accounts needs to be specified for client and control is : " & adoClient.Fields.Item("ClientID").Value & ",  'PURCHASE LEDGER CONTROL' " & Chr(13) + Chr$(10)
                                 strmsg1 = strmsg1 & "The main Control Accounts needs to be specified for client and control is : " & adoClient.Fields.Item("ClientID").Value & ",  'PURCHASE LEDGER CONTROL' " & Chr(13) + Chr$(10)
                            End If
                            rsControlAccountbyClient.Close
                             rsControlAccountbyClient.Open "SELECT CODE,[CLIENT ID] from [NominalCodes$] where CODE is not null and " & _
                               "UCASE([CONTROL ACCOUNT]) in ('INPUT VAT' ) AND [CLIENT ID]='DEFAULT' ", oConnEx, adOpenDynamic, adLockOptimistic
                               If rsControlAccountbyClient.EOF Then
                                     Print #1, "The main Control Accounts needs to be specified for client and control is : " & adoClient.Fields.Item("ClientID").Value & ",  'INPUT VAT' " & Chr(13) + Chr$(10)
                                     strmsg1 = strmsg1 & "The main Control Accounts needs to be specified for client and control is : " & adoClient.Fields.Item("ClientID").Value & ",  'INPUT VAT' " & Chr(13) + Chr$(10)
                               End If
                            rsControlAccountbyClient.Close
                             rsControlAccountbyClient.Open "SELECT CODE,[CLIENT ID] from [NominalCodes$] where CODE is not null and " & _
                               "UCASE([CONTROL ACCOUNT]) in (" & _
                               "'OUTPUT VAT' ) AND [CLIENT ID]='DEFAULT' ", oConnEx, adOpenDynamic, adLockOptimistic
                               If rsControlAccountbyClient.EOF Then
                                     Print #1, "The main Control Accounts needs to be specified for client and control is : " & adoClient.Fields.Item("ClientID").Value & ",  'OUTPUT VAT' " & Chr(13) + Chr$(10)
                                     strmsg1 = strmsg1 & "The main Control Accounts needs to be specified for client and control is : " & adoClient.Fields.Item("ClientID").Value & ", 'OUTPUT VAT' " & Chr(13) + Chr$(10)
                            End If
                            rsControlAccountbyClient.Close
                             rsControlAccountbyClient.Open "SELECT CODE,[CLIENT ID] from [NominalCodes$] where CODE is not null and " & _
                               "UCASE([CONTROL ACCOUNT]) in ( " & _
                               " 'RETAINED EARNINGS' ) AND [CLIENT ID]='DEFAULT' ", oConnEx, adOpenDynamic, adLockOptimistic
                               If rsControlAccountbyClient.EOF Then
                                     Print #1, "The main Control Accounts needs to be specified for client and control is : " & adoClient.Fields.Item("ClientID").Value & ", 'RETAINED EARNINGS' " & Chr(13) + Chr$(10)
                                     strmsg1 = strmsg1 & "The main Control Accounts needs to be specified for client and control is : " & adoClient.Fields.Item("ClientID").Value & ",  'RETAINED EARNINGS' " & Chr(13) + Chr$(10)
                            End If
                            rsControlAccountbyClient.Close
                Else
                    'This is the validation part for control account has been set for specific clients
                        rsControlAccountbyClient.Open "SELECT CODE,[CLIENT ID] from [NominalCodes$] where CODE is not null and " & _
                           "UCASE([CONTROL ACCOUNT]) in ('SALES LEDGER CONTROL') AND [CLIENT ID]='" & adoClient.Fields.Item("ClientID").Value & "' ", oConnEx, adOpenDynamic, adLockOptimistic
                        If rsControlAccountbyClient.EOF Then
                             Print #1, "The main Control Accounts needs to be specified for client and control is : " & adoClient.Fields.Item("ClientID").Value & ", 'SALES LEDGER CONTROL' " & Chr(13) + Chr$(10)
                             strmsg1 = strmsg1 & "The main Control Accounts needs to be specified for client and control is : " & adoClient.Fields.Item("ClientID").Value & ", 'SALES LEDGER CONTROL' " & Chr(13) + Chr$(10)
                        End If
                        rsControlAccountbyClient.Close
                         rsControlAccountbyClient.Open "SELECT CODE,[CLIENT ID] from [NominalCodes$] where CODE is not null and " & _
                           "UCASE([CONTROL ACCOUNT]) in ('PURCHASE LEDGER CONTROL') AND [CLIENT ID]='" & adoClient.Fields.Item("ClientID").Value & "' ", oConnEx, adOpenDynamic, adLockOptimistic
                        If rsControlAccountbyClient.EOF Then
                             Print #1, "The main Control Accounts needs to be specified for client and control is : " & adoClient.Fields.Item("ClientID").Value & ",  'PURCHASE LEDGER CONTROL' " & Chr(13) + Chr$(10)
                             strmsg1 = strmsg1 & "The main Control Accounts needs to be specified for client and control is : " & adoClient.Fields.Item("ClientID").Value & ",  'PURCHASE LEDGER CONTROL' " & Chr(13) + Chr$(10)
                        End If
                        rsControlAccountbyClient.Close
                         rsControlAccountbyClient.Open "SELECT CODE,[CLIENT ID] from [NominalCodes$] where CODE is not null and " & _
                           "UCASE([CONTROL ACCOUNT]) in ('INPUT VAT' ) AND [CLIENT ID]='" & adoClient.Fields.Item("ClientID").Value & "' ", oConnEx, adOpenDynamic, adLockOptimistic
                           If rsControlAccountbyClient.EOF Then
                                 Print #1, "The main Control Accounts needs to be specified for client and control is : " & adoClient.Fields.Item("ClientID").Value & ",  'INPUT VAT' " & Chr(13) + Chr$(10)
                                 strmsg1 = strmsg1 & "The main Control Accounts needs to be specified for client and control is : " & adoClient.Fields.Item("ClientID").Value & ",  'INPUT VAT' " & Chr(13) + Chr$(10)
                           End If
                        rsControlAccountbyClient.Close
                         rsControlAccountbyClient.Open "SELECT CODE,[CLIENT ID] from [NominalCodes$] where CODE is not null and " & _
                           "UCASE([CONTROL ACCOUNT]) in (" & _
                           "'OUTPUT VAT' ) AND [CLIENT ID]='" & adoClient.Fields.Item("ClientID").Value & "' ", oConnEx, adOpenDynamic, adLockOptimistic
                           If rsControlAccountbyClient.EOF Then
                                 Print #1, "The main Control Accounts needs to be specified for client and control is : " & adoClient.Fields.Item("ClientID").Value & ",  'OUTPUT VAT' " & Chr(13) + Chr$(10)
                                 strmsg1 = strmsg1 & "The main Control Accounts needs to be specified for client and control is : " & adoClient.Fields.Item("ClientID").Value & ", 'OUTPUT VAT' " & Chr(13) + Chr$(10)
                        End If
                        rsControlAccountbyClient.Close
                         rsControlAccountbyClient.Open "SELECT CODE,[CLIENT ID] from [NominalCodes$] where CODE is not null and " & _
                           "UCASE([CONTROL ACCOUNT]) in ( " & _
                           " 'RETAINED EARNINGS' ) AND [CLIENT ID]='" & adoClient.Fields.Item("ClientID").Value & "' ", oConnEx, adOpenDynamic, adLockOptimistic
                           If rsControlAccountbyClient.EOF Then
                                 Print #1, "The main Control Accounts needs to be specified for client and control is : " & adoClient.Fields.Item("ClientID").Value & ", 'RETAINED EARNINGS' " & Chr(13) + Chr$(10)
                                 strmsg1 = strmsg1 & "The main Control Accounts needs to be specified for client and control is : " & adoClient.Fields.Item("ClientID").Value & ",  'RETAINED EARNINGS' " & Chr(13) + Chr$(10)
                        End If
                        rsControlAccountbyClient.Close
                End If

                  With adoRstMdb
'                    Do While Not adoClient.EOF             'Create the nominal code for each CLIENT
                    If oNCsXls.EOF Then
                         oNCsXls.Close
                         Exit Do
                    End If
                    'now test for duplicate nominal code in excel for each client
                        adoRsDuplicateCodeinClient.Open "SELECT CODE,[CLIENT ID], count(*) as cnt from [NominalCodes$] where [CLIENT ID]  ='" & adoClient.Fields.Item("ClientID").Value & "' AND CODE is not null  group by " & _
                            "CODE,[CLIENT ID] having count(*)>1", oConnEx, adOpenDynamic, adLockOptimistic
                            If adoClient.Fields.Item("ClientID").Value = "KRCAPITAL" Then
                                Debug.Print ""
                            End If
                         If Not adoRsDuplicateCodeinClient.EOF Then
                                If adoRsDuplicateCodeinClient("cnt").Value > 1 Then
                                    MsgBox "Duplicate nominal code for this client: " & adoClient.Fields.Item("ClientID").Value & ", Code is : " & adoRsDuplicateCodeinClient("CODE").Value, vbInformation, "Warning"
                                    adoRsDuplicateCodeinClient.Close
                                    Exit Do
                                End If
                         End If
                        adoRsDuplicateCodeinClient.Close
                            
                        .Open "SELECT * FROM NominalLedger WHERE ClientID = '" & adoClient.Fields.Item("ClientID").Value & "';", oConnAcc, adOpenDynamic, adLockOptimistic
                        
                        If .EOF Then        'THE CLIENT DOES NOT HAVE NC INTO DATABASE
                           'NOW THE CLIENT WILL HAVE A SET OF NOMINAL CODE BY THE DEFAULT SET OF THE INPUT SHEET
                           Do While Not oNCsXls.EOF
                              .AddNew
                              .Fields.Item("Code").Value = oNCsXls.Fields.Item("CODE").Value
                              .Fields.Item("Name").Value = IIf(IsNull(oNCsXls.Fields.Item("NOMINAL NAME").Value), "Not defined", oNCsXls.Fields.Item("NOMINAL NAME").Value)
                              .Fields.Item("Type").Value = IIf(IsNull(oNCsXls.Fields.Item("TYPE").Value), 0, _
                                 IIf(UCase(Left(oNCsXls.Fields.Item("TYPE").Value, 1)) = "B", 1, IIf(UCase(Left(oNCsXls.Fields.Item("TYPE").Value, 1)) = "P", 2, 0)))
                              .Fields.Item("DrCr").Value = IIf(IsNull(oNCsXls.Fields.Item("CREDIT OR DEBIT").Value), "", _
                                 IIf(UCase(oNCsXls.Fields.Item("CREDIT OR DEBIT").Value) = "DEBIT", "Dr", "Cr"))
                              .Fields.Item("ClientID").Value = adoClient.Fields.Item("ClientID").Value
                             
                            If adoClient.Fields.Item("ClientID").Value = "SEGALL" Then
                                 Debug.Print ""
                            End If
                              .Fields.Item("Posting").Value = IIf(UCase(oNCsXls.Fields.Item("POSTING ALLOWED").Value = "YES"), True, False)
                              adoRst.Find "STName = '" & oNCsXls.Fields.Item("SUBTYPE").Value & "' ", , , 1
                              .Fields.Item("SubType").Value = IIf(IsNull(adoRst.Fields.Item("STCode").Value), "", adoRst.Fields.Item("STCode").Value)
                              'issue 505: Import module needs to be modified
                               'added by anol 10 Feb 2015
                           
                              If IsNull(oNCsXls.Fields.Item("CONTROL ACCOUNT").Value) = True Then
                                 .Fields.Item("CAFixed").Value = "0"
                                 '.Fields.Item("CAName").Value = ""
                                 '.Fields.Item("CADisorder").Value = ""
                                 'added by anol issue 521 note 1
                                 'Date 25 Jan 2015
                                  .Fields.Item("CAType").Value = ""
                              Else
                                 .Fields.Item("CAType").Value = ""
                                 .Fields.Item("CAFixed").Value = "-1"
                                 If UCase(oNCsXls.Fields.Item("CONTROL ACCOUNT").Value) = "SALES LEDGER CONTROL" Then
                                      .Fields.Item("CAName").Value = "Sales Ledger Control"
                                      .Fields.Item("CAType").Value = "S"
                                      .Fields.Item("CADisorder").Value = "1"
                                 ElseIf UCase(oNCsXls.Fields.Item("CONTROL ACCOUNT").Value) = "PURCHASE LEDGER CONTROL" Then
                                      .Fields.Item("CAName").Value = "Purchase Ledger Control"
                                      .Fields.Item("CAType").Value = "P"
                                      .Fields.Item("CADisorder").Value = "2"
                                 ElseIf UCase(oNCsXls.Fields.Item("CONTROL ACCOUNT").Value) = "INPUT VAT" Then
                                     .Fields.Item("CAName").Value = "Input VAT"
                                     .Fields.Item("CAType").Value = "I"
                                     .Fields.Item("CADisorder").Value = "3"
                                 ElseIf UCase(oNCsXls.Fields.Item("CONTROL ACCOUNT").Value) = "OUTPUT VAT" Then
                                      .Fields.Item("CAName").Value = "Output VAT"
                                      .Fields.Item("CAType").Value = "O"
                                      .Fields.Item("CADisorder").Value = "4"
                                 ElseIf UCase(oNCsXls.Fields.Item("CONTROL ACCOUNT").Value) = "RETAINED EARNINGS" Then
                                    .Fields.Item("CAName").Value = "Retained Earnings"
                                    .Fields.Item("CAType").Value = "R"
                                    .Fields.Item("CADisorder").Value = "5"
                                 End If
                               End If
                              'End of modification
                                 
                              .Update
            
                              'NJ_CC table needs to updated here before go to next row
                              adoRST_.Open "SELECT * FROM NJ_CC WHERE ClientID = '" & adoClient.Fields.Item("ClientID").Value & "' AND " & _
                                                                     "Code     = '" & oNCsXls.Fields.Item("CODE").Value & "' AND " & _
                                                                     "CC       = '" & oNCsXls.Fields.Item("CATEGORY CODE").Value & "' ", _
                                      oConnAcc, adOpenDynamic, adLockOptimistic
                              If adoRST_.EOF Then adoRST_.AddNew
                              adoRST_.Fields.Item("ClientID").Value = adoClient.Fields.Item("ClientID").Value
                              adoRST_.Fields.Item("Code").Value = oNCsXls.Fields.Item("CODE").Value
                              adoRST_.Fields.Item("CC").Value = oNCsXls.Fields.Item("CATEGORY CODE").Value
                              adoRST_.Update
                              adoRST_.Close
                              oNCsXls.MoveNext
                              If oNCsXls.EOF Then
'                                    oNCsXls.Close
                                    Exit Do
                              End If
                              If Trim(oNCsXls.Fields.Item(0).Value) = "END" Then
'                                    oNCsXls.Close
                                    Exit Do
                              End If
                              If oNCsXls.Fields.Item(0).Value = "" Then
'                                    oNCsXls.Close
                                    Exit Do
                              End If
                           Loop
                          
                        Else           'The client has some NC already in the system
                           Do While Not oNCsXls.EOF
                              .Find "Code = '" & oNCsXls.Fields.Item("CODE").Value & "' ", , , 1
                              
                              If .EOF Then
                                 .AddNew
                                 .Fields.Item("Code").Value = oNCsXls.Fields.Item("CODE").Value
                                 .Fields.Item("ClientID").Value = adoClient.Fields.Item("ClientID").Value
                              End If
                              If adoClient.Fields.Item("ClientID").Value = "SEGALL" Then
                                Debug.Print ""
                              End If
                              .Fields.Item("Name").Value = IIf(IsNull(oNCsXls.Fields.Item("NOMINAL NAME").Value), "Not defined", oNCsXls.Fields.Item("NOMINAL NAME").Value)
                              .Fields.Item("Type").Value = IIf(IsNull(oNCsXls.Fields.Item("TYPE").Value), 0, _
                                 IIf(UCase(Left(oNCsXls.Fields.Item("TYPE").Value, 1)) = "B", 1, IIf(UCase(Left(oNCsXls.Fields.Item("TYPE").Value, 1)) = "P", 2, 0)))
                              .Fields.Item("DrCr").Value = IIf(IsNull(oNCsXls.Fields.Item("CREDIT OR DEBIT").Value), "", _
                                 IIf(UCase(oNCsXls.Fields.Item("CREDIT OR DEBIT").Value) = "DEBIT", "Dr", "Cr"))
                              .Fields.Item("Posting").Value = IIf(UCase(oNCsXls.Fields.Item("POSTING ALLOWED").Value = "YES"), True, False)
                              If oNCsXls.Fields.Item("CODE").Value = 3490 Then
                                Debug.Print ""
                              End If
                              adoRst.Find "STName = '" & oNCsXls.Fields.Item("SUBTYPE").Value & "' ", , , 1
                              .Fields.Item("SubType").Value = IIf(IsNull(adoRst.Fields.Item("STCode").Value), "", adoRst.Fields.Item("STCode").Value)
                               'issue 505: Import module needs to be modified
                               'added by anol 10 Feb 2015
                              If IsNull(oNCsXls.Fields.Item("CONTROL ACCOUNT").Value) = True Then
                                 .Fields.Item("CAFixed").Value = "0"
                                 '.Fields.Item("CAName").Value = ""
                                 '.Fields.Item("CADisorder").Value = ""
                                 'added by anol issue 521 note 1
                                 'Date 25 Jan 2015
                                 .Fields.Item("CAType").Value = ""
                              Else
                                 .Fields.Item("CAFixed").Value = "-1"
                                 If UCase(oNCsXls.Fields.Item("CONTROL ACCOUNT").Value) = "SALES LEDGER CONTROL" Then
                                      .Fields.Item("CAName").Value = "Sales Ledger Control"
                                      .Fields.Item("CAType").Value = "S"
                                      .Fields.Item("CADisorder").Value = "1"
                                 ElseIf UCase(oNCsXls.Fields.Item("CONTROL ACCOUNT").Value) = "PURCHASE LEDGER CONTROL" Then
                                      .Fields.Item("CAName").Value = "Purchase Ledger Control"
                                      .Fields.Item("CAType").Value = "P"
                                      .Fields.Item("CADisorder").Value = "2"
                                 ElseIf UCase(oNCsXls.Fields.Item("CONTROL ACCOUNT").Value) = "INPUT VAT" Then
                                     .Fields.Item("CAName").Value = "Input VAT"
                                     .Fields.Item("CAType").Value = "I"
                                     .Fields.Item("CADisorder").Value = "3"
                                 ElseIf UCase(oNCsXls.Fields.Item("CONTROL ACCOUNT").Value) = "OUTPUT VAT" Then
                                      .Fields.Item("CAName").Value = "Output VAT"
                                      .Fields.Item("CAType").Value = "O"
                                      .Fields.Item("CADisorder").Value = "4"
                                 ElseIf UCase(oNCsXls.Fields.Item("CONTROL ACCOUNT").Value) = "RETAINED EARNINGS" Then
                                    .Fields.Item("CAName").Value = "Retained Earnings"
                                    .Fields.Item("CAType").Value = "R"
                                    .Fields.Item("CADisorder").Value = "5"
                                 End If
                                End If
                              'End of modification
                              
                              .Update
            
                              adoRST_.Open "SELECT * FROM NJ_CC WHERE ClientID = '" & adoClient.Fields.Item("ClientID").Value & "' AND " & _
                                                                     "Code     = '" & oNCsXls.Fields.Item("CODE").Value & "' AND " & _
                                                                     "CC       = '" & oNCsXls.Fields.Item("CATEGORY CODE").Value & "' ", _
                                      oConnAcc, adOpenDynamic, adLockOptimistic
                              If adoRST_.EOF Then adoRST_.AddNew
                              adoRST_.Fields.Item("ClientID").Value = adoClient.Fields.Item("ClientID").Value
                              adoRST_.Fields.Item("Code").Value = oNCsXls.Fields.Item("CODE").Value
                              adoRST_.Fields.Item("CC").Value = oNCsXls.Fields.Item("CATEGORY CODE").Value
                              adoRST_.Update
                              adoRST_.Close
                              oNCsXls.MoveNext
                              If oNCsXls.EOF Then
                                    Exit Do
                              End If
                              If Trim(oNCsXls.Fields.Item(0).Value) = "END" Then
                                    Exit Do
                              End If
                              If oNCsXls.Fields.Item(0).Value = "" Then
                                    Exit Do
                              End If
                           Loop
                           
                        End If
                        .Close
            
                        adoClient.MoveNext
                        'added by anol 24 Sep 2015
                        If Trim(adoClient.Fields.Item(0).Value) = "END" Then
                            Exit Do
                        End If
                        If adoClient.Fields.Item(0).Value = "" Then
                            Exit Do
                        End If
                        If adoClient.Fields.Item("ClientID").Value = "END" Then
                            Exit Do
                        End If

                         If .State = 1 Then .Close
                  End With

                  oNCsXls.Close
'                  adoClient.MoveNext
            Loop
            adoRst.Close
            If oNCsXls.State = 1 Then
                oNCsXls.Close
            End If
            adoClient.Close
            Print #1, "Data writing FINISH: ********** NominalCode **********" & Chr$(13) + Chr$(10)
            import_NominalCode = True
End Function

Private Function GetVatCode(szVAT As String, ByVal oConnAcc As ADODB.Connection) As Integer
    Dim adoRst As New ADODB.Recordset

   'adoRST.Open "SELECT VAT_ID FROM tlbVatCode WHERE VAT_RATE = " & Val(szVAT) & ";", oConnAcc, adOpenStatic, adLockReadOnly
   'modified by anol 30 Dec 2014
adoRst.Open "SELECT VAT_ID FROM tlbVatCode WHERE VAT_Code = '" & szVAT & "';", oConnAcc, adOpenStatic, adLockReadOnly
   If adoRst.EOF Then
      GetVatCode = 0
   Else
      GetVatCode = adoRst.Fields.Item(0).Value
   End If

   adoRst.Close
   Set adoRst = Nothing
   
End Function

Private Function leaseId(szLinkID As String, ByVal oConnAcc As ADODB.Connection) As String
   Dim adoRst As New ADODB.Recordset
   Dim szSQL As String

   szSQL = "SELECT LeaseID FROM LeaseDetails WHERE Text3 = '" & szLinkID & "';"
   adoRst.Open szSQL, oConnAcc, adOpenStatic, adLockReadOnly

   leaseId = IIf(adoRst.EOF, Null, adoRst!leaseId)

   adoRst.Close
   Set adoRst = Nothing
End Function

Private Function PropertyNumber_Acc(szProperty As String, ByVal oConn As ADODB.Connection) As String
   Dim adoRst As New ADODB.Recordset
   Dim szSQL As String

   szSQL = "SELECT * FROM Property WHERE PropertyName = '" & replaceApostrophe(szProperty) & "';"
'Debug.Print szSQL
   adoRst.Open szSQL, oConn, adOpenStatic, adLockReadOnly

   If adoRst.EOF Then
      PropertyNumber_Acc = ""
   Else
      PropertyNumber_Acc = adoRst.Fields.Item("PropertyID").Value
   End If

   adoRst.Close
   Set adoRst = Nothing
End Function

Private Function UnitNumber(szProperty As String, szUnit As String, ByVal oConnAcc As ADODB.Connection) As String
   Dim adoRst As New ADODB.Recordset
   Dim szSQL As String, szPropertyID As String

   szSQL = "SELECT PropertyID FROM Property WHERE PropertyName = '" & replaceApostrophe(Trim(szProperty)) & "';"
'Debug.Print szSQL
   adoRst.Open szSQL, oConnAcc, adOpenStatic, adLockReadOnly

   If adoRst.EOF Then
      Set adoRst = Nothing
      UnitNumber = Null
      Exit Function
   Else
      szPropertyID = adoRst!propertyID
      adoRst.Close
   End If

   szSQL = "SELECT UnitNumber FROM Units " & _
           "WHERE PropertyID = '" & szPropertyID & "' AND " & _
               "UnitName = '" & replaceApostrophe(szUnit) & "';"
   adoRst.Open szSQL, oConnAcc, adOpenStatic, adLockReadOnly

   UnitNumber = IIf(adoRst.EOF, Null, adoRst!UnitNumber)

   adoRst.Close
   Set adoRst = Nothing
End Function

Private Function LesseeID(szLesseeName As String, ByVal oConnAcc As ADODB.Connection) As String
   Dim adoRst As New ADODB.Recordset
   Dim szSQL As String

   szSQL = "SELECT SageAccountNumber FROM Tenants WHERE Name = '" & replaceApostrophe(szLesseeName) & "';"
   adoRst.Open szSQL, oConnAcc, adOpenStatic, adLockReadOnly

   LesseeID = IIf(adoRst.EOF, "", adoRst!SageAccountNumber)

   adoRst.Close
   Set adoRst = Nothing
End Function

Private Function LesseeNAME(szLesseeID As String, ByVal oConnAcc As ADODB.Connection) As String
   Dim adoRst As New ADODB.Recordset
   Dim szSQL As String

   szSQL = "SELECT Name FROM Tenants WHERE SageAccountNumber = '" & szLesseeID & "';"
   adoRst.Open szSQL, oConnAcc, adOpenStatic, adLockReadOnly

   LesseeNAME = IIf(adoRst.EOF, "", adoRst!Name)

   adoRst.Close
   Set adoRst = Nothing
End Function

Public Function GenerateUnitNumber(szPropertyID As String, ByVal oConnAcc As ADODB.Connection) As String
   Dim rstUnitNumber As New ADODB.Recordset
   Dim sSQLQuery_ As String
   Dim MAX_UNIT_ As String
   Dim UNIT_NUMBER_ As String

   'CLREATE SQL QUERY ON OPTION BUTTON SELECTION
   sSQLQuery_ = "SELECT MAX(RIGHT(Units.UnitNumber,3)) + 1 AS  MAX_UNIT " & _
                "From Units " & _
                "WHERE Units.PropertyID = '" & szPropertyID & "';"

   rstUnitNumber.Open sSQLQuery_, oConnAcc, adOpenStatic, adLockReadOnly

   If rstUnitNumber.EOF Or rstUnitNumber.BOF Then
      MAX_UNIT_ = "1"
   End If

   While Not rstUnitNumber.EOF
      MAX_UNIT_ = IIf(IsNull(rstUnitNumber!MAX_UNIT), "1", rstUnitNumber!MAX_UNIT)
      rstUnitNumber.MoveNext
   Wend

   GenerateUnitNumber = szPropertyID & "-" & Lpad(MAX_UNIT_, "0", 3)

   rstUnitNumber.Close
   Set rstUnitNumber = Nothing
End Function

Function Lpad(MyValue As String, MyPadCharacter As String, MyPaddedLength As Integer)
    Lpad = String(MyPaddedLength - Len(MyValue), MyPadCharacter) & MyValue
End Function

Private Function ClientID_Acc(ByVal szClientName, oConn As ADODB.Connection) As String
   Dim oRst As New ADODB.Recordset
   Dim szSQL As String

'  the table name is the worksheet name
   szSQL = "select * from Client where ClientName = '" & replaceApostrophe(szClientName) & "';"
'  Get the recordset
   oRst.Open szSQL, oConn, adOpenStatic, adLockOptimistic

   If Not oRst.EOF Then
      ClientID_Acc = oRst.Fields.Item("ClientID").Value
   End If

   oRst.Close
   Set oRst = Nothing
End Function

Private Function BankSortCode(ByVal szBankID, oConnEx As ADODB.Connection) As String
   Dim oRstEx As New ADODB.Recordset
   Dim sTableNameEx As String

'  the table name is the worksheet name
   sTableNameEx = "select * from [BankDetails$] where [Bank ID] = " & szBankID & ";"
'  Get the recordset
   oRstEx.Open sTableNameEx, oConnEx, adOpenStatic, adLockReadOnly

   If Not oRstEx.EOF Then
      BankSortCode = oRstEx.Fields.Item(1).Value
   End If

   oRstEx.Close
   Set oRstEx = Nothing
End Function

Private Function BankSortCode_Acc(ByVal szBankID, oConn As ADODB.Connection) As String
   Dim oRst As New ADODB.Recordset
   Dim szSQL As String

'  the table name is the worksheet name
   szSQL = "select * from tlbBank where BANK_ID = '" & szBankID & "';"
'  Get the recordset
   oRst.Open szSQL, oConn, adOpenStatic, adLockReadOnly

   If Not oRst.EOF Then
      BankSortCode_Acc = oRst.Fields.Item("SORT_CODE").Value
   End If

   oRst.Close
   Set oRst = Nothing
End Function

Private Function GetBankID(ByVal szPropID, oConn As ADODB.Connection) As String
   Dim oRst As New ADODB.Recordset
   Dim szSQL As String

'  the table name is the worksheet name
   szSQL = "SELECT MY_ID FROM tlbClientBanks, Property WHERE CLIENT_ID = ClientID and PropertyID = '" & szPropID & "' AND DEFAULT_AC;"
'  Get the recordset
   oRst.Open szSQL, oConn, adOpenStatic, adLockReadOnly

   If Not oRst.EOF Then
      GetBankID = oRst.Fields.Item("MY_ID").Value
   End If

   oRst.Close
   Set oRst = Nothing
End Function

Private Sub cmdBrowse_Click(Index As Integer)
   cdlgBrowser.Filename = CurDir & "\*.xls"  'set default location
   cdlgBrowser.ShowOpen

   txtSourceFile.text = cdlgBrowser.Filename
End Sub

Private Sub Form_Load()
   'Me.Move (Screen.Width - Width) / 2, (Screen.Height - Height) / 2
   'add all the shopping centre names from control database to cboshopcentre
   modLists.loadCompanies
   iMsgNum = 1

   If UCase(SystemUser) = "SAMRAT" And UCase(WS_Name) = "WS1" Then
      txtSourceFile.text = "salia"
      txtSourceFile.text = "C:\Samrat\PropertyManagementProgram\Non_Client_Server\Non_SAGE\BlockMng\ImportModule\WorkingFolderNew\Internal Import Template.xls"
      txtSourceFile.text = "C:\Samrat\PropertyManagementProgram\Non_Client_Server\Non_SAGE\BlockMng\ImportModule\WorkingFolderNew\Itsyourplace - Internal Import Template.xls"

'      cmdImportData_Click
   End If
End Sub

Private Function FinePropertyID_LeaseLink(szLinkID As String, ByVal oConnAcc As ADODB.Connection, ByVal oConnEx As ADODB.Connection) As String
   Dim oRstEx As New ADODB.Recordset
   Dim szSQL As String, szProperty As String

   szSQL = "select * from [Lease$] WHERE LeaseLinkID = " & szLinkID & ";"
   oRstEx.Open szSQL, oConnEx, adOpenStatic, adLockOptimistic

   szProperty = oRstEx.Fields.Item(1).Value
   oRstEx.Close
   Set oRstEx = Nothing

   FinePropertyID_LeaseLink = PropertyNumber_Acc(szProperty, oConnAcc)
End Function

Private Sub Form_Unload(Cancel As Integer)
   txtSourceFile.text = ""
   frmStartUp.Enabled = True
   'FocusControl frmStartUp
   lblMessage(2).Caption = ""

End Sub

Public Sub WaitFewSec()
   If UCase(SystemUser) = "SAMRAT" And UCase(WS_Name) = "WS1" Then
      Exit Sub
   End If
   
   Dim i As Long

   For i = 0 To DELAY_MAX
   Next i
End Sub

Private Sub MessageDetails()
   lblMessage(2).Caption = ""
   iMsgNum = 1

   If chkAll.Value Then Exit Sub

   If Not chkBankDetails.Value And chkClientBank.Value Then
      iMsgNum = iMsgNum + 1
      lblMessage(2).Caption = lblMessage(2).Caption & _
                              iMsgNum & ".  " & "Please make sure already bank details have been entered into the system." & Chr(13)
   End If
   If Not chkClients.Value And (chkClientBank.Value Or chkProperty.Value) Then
      iMsgNum = iMsgNum + 1
      lblMessage(2).Caption = lblMessage(2).Caption & _
                              iMsgNum & ".  " & "Please make sure already client details have been entered into the system." & Chr(13)
   End If
'   If Not chkProperty.Value And (chkUnit.Value Or chkGlobalData.Value) Then
'      iMsgNum = iMsgNum + 1
'      lblMessage(2).Caption = lblMessage(2).Caption & _
'                              iMsgNum & ".  " & "Please make sure already property details have been entered into the system." & Chr(13)
'   End If
   If Not chkUnit.Value And chkLease.Value Then
      iMsgNum = iMsgNum + 1
      lblMessage(2).Caption = lblMessage(2).Caption & _
                              iMsgNum & ".  " & "Please make sure already unit details have been entered into the system." & Chr(13)
   End If
End Sub

Private Function CreateClientId(szName As String, adoConn As ADODB.Connection) As String
   Dim szSQL As String, i As Integer, szChar As String, j As Integer
   Dim adoRSTClient As New ADODB.Recordset, adoRSTLandlord As New ADODB.Recordset

   For i = 1 To Len(szName) - 1
      szChar = UCase(Mid(szName, i, 1))
      If (szChar >= "A" And szChar <= "Z") Then
         CreateClientId = CreateClientId & szChar
         j = j + 1
      End If
      If j = 8 Then Exit For
   Next i

   If j < 8 Then CreateClientId = Left(CreateClientId & "01234567", 8)

   szSQL = "SELECT ClientID " & _
           "FROM Client " & _
           "WHERE Client.ClientID = '" & CreateClientId & "';"
   adoRSTClient.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   szSQL = "SELECT LandlordID " & _
           "FROM Landlord " & _
           "WHERE Landlord.LandlordID = '" & CreateClientId & "';"
   adoRSTLandlord.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   j = 1
   Do
      If adoRSTClient.EOF And adoRSTLandlord.EOF Then Exit Do

      adoRSTClient.Close
      adoRSTLandlord.Close

      CreateClientId = Left(CreateClientId & "01234567", 6) & Format(j, "00")
      szSQL = "SELECT ClientID " & _
              "FROM Client " & _
              "WHERE Client.ClientID = '" & CreateClientId & "';"
      adoRSTClient.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

      szSQL = "SELECT LandlordID " & _
              "FROM Landlord " & _
              "WHERE Landlord.LandlordID = '" & CreateClientId & "';"
      adoRSTLandlord.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

      j = j + 1
   Loop

   adoRSTClient.Close
   adoRSTLandlord.Close
   Set adoRSTClient = Nothing
   Set adoRSTLandlord = Nothing
End Function

Private Function replaceApostrophe(ByVal str As String) As String
 replaceApostrophe = Replace(str, "'", "''")
End Function

Private Function CreateSubTypeID(szSubType As String) As String
   Dim szaTemp()  As String
   Dim i          As Integer

   szaTemp = Split(szSubType)
   
   For i = 0 To UBound(szaTemp)
      If i >= 5 Then Exit Function
      If IsOnlyAlpha(Left(szaTemp(i), 1)) Then _
         CreateSubTypeID = CreateSubTypeID & UCase(Left(szaTemp(i), 1))
   Next i
End Function

Private Function IsOnlyAlpha(szChar As String) As Boolean
   If (Asc(szChar) > 64 And Asc(szChar) < 91) Or (Asc(szChar) > 96 And Asc(szChar) < 123) Then
      IsOnlyAlpha = True
   Else
      IsOnlyAlpha = False
   End If
End Function
Private Sub cboShopCentre_KeyPress(KeyAscii As MSForms.ReturnInteger)
   If KeyAscii = 27 Then Unload Me
End Sub

Private Sub chkAll_Click()
    chkBankDetails.Value = chkAll.Value
    chkClients.Value = chkAll.Value
    chkClientBank.Value = chkAll.Value
    chkDemandType.Value = chkAll.Value
    chkChargeTypes.Value = chkAll.Value
    chkPayableTypes.Value = chkAll.Value
    chkProperty.Value = chkAll.Value
    chkUnit.Value = chkAll.Value
    chkAPD.Value = chkAll.Value
    chkLease.Value = chkAll.Value
    chkFund.Value = chkAll.Value
    chkSchedule.Value = chkAll.Value
    chkReportCategory.Value = chkAll.Value
    chkNC.Value = chkAll.Value
    chkAddPayDateSet.Value = chkAll.Value
    chkLessee.Value = chkAll.Value
    chkRC_SC_IC.Value = chkAll.Value
    chkSuppliers.Value = chkAll.Value
    chkRentBudget.Value = chkAll.Value
    chkServiceBudget.Value = chkAll.Value
    chkInsuranceBudget.Value = chkAll.Value
   'chkControlAccount.Value = chkAll.Value
    chkRentBudget.Value = chkAll.Value
    chkServiceBudget.Value = chkAll.Value
    chkInsuranceBudget.Value = chkAll.Value
   
   MessageDetails
End Sub

Private Sub chkAPD_Click()
   If Not chkAPD.Value Then chkAll.Value = False

   MessageDetails
End Sub

Private Sub chkBankDetails_Click()
   If Not chkBankDetails.Value Then chkAll.Value = False

   MessageDetails
End Sub

Private Sub chkClientBank_Click()
   If Not chkClientBank.Value Then chkAll.Value = False

   MessageDetails
End Sub

Private Sub chkClients_Click()
   If Not chkClients.Value Then chkAll.Value = False

   MessageDetails
End Sub

Private Sub chkDemandType_Click()
   If Not chkDemandType.Value Then chkAll.Value = False

   MessageDetails
End Sub

Private Sub chkGlobalData_Click()
   MessageDetails
End Sub

Private Sub chkLease_Click()
   If Not chkLease.Value Then chkAll.Value = False

   MessageDetails
End Sub

Private Sub chkProperty_Click()
   If Not chkProperty.Value Then chkAll.Value = False

   MessageDetails
End Sub

Private Sub chkUnit_Click()
   If Not chkUnit.Value Then chkAll.Value = False

   MessageDetails
End Sub

Private Sub cmdExitProgram_Click()
   Unload Me
End Sub
Private Sub UpdateDatabase()
'added by anol 26 mAR 2015
   On Error GoTo Err
   Dim oConnAcc As New ADODB.Connection    'MS Access
   oConnAcc.Open "DSN=" & frmStartUp.cboShopCentre.Value & ";UID=;PWD=RDSWKDPP"
   Dim temp As Integer
    Do
         temp = UpdateDatabase1(oConnAcc)
         If temp = 0 Then
            oConnAcc.Close
            Exit Sub
         End If
      Loop While temp <> 0
      Exit Sub
Err:
     If Err.Number = -2147467259 Then
        MsgBox "Could not find the database file", vbCritical, "Warning!"
        Exit Sub
     End If
End Sub
Private Function UpdateDatabase1(oConnAcc As ADODB.Connection) As Integer
'added by anol 26 mAR 2015
   Dim rsCheck As New ADODB.Recordset
On Error GoTo add4
   UpdateDatabase1 = 0
   rsCheck.Open "Select PropertyID from tblPurInv", oConnAcc, adOpenKeyset, adLockReadOnly
   If rsCheck.Fields.Item("PropertyID").DefinedSize < 10 Then
      rsCheck.Close
      Set rsCheck = Nothing
      oConnAcc.Execute "ALTER TABLE tblPurInv ALTER COLUMN PropertyID text(10)"
      UpdateDatabase1 = 1
      Exit Function
   Else
      rsCheck.Close
      Set rsCheck = Nothing
      GoTo add4
   End If
add4:
   On Error GoTo Modadd4
   rsCheck.Open "Select ClientAddressLine4 from client", oConnAcc, adOpenKeyset, adLockReadOnly
   rsCheck.Close
   Set rsCheck = Nothing
   GoTo DataTypeDirectLine2
Modadd4:
   oConnAcc.Execute "ALTER TABLE client add COLUMN ClientAddressLine4 text(250)"
   UpdateDatabase1 = 1
   Exit Function
DataTypeDirectLine2:
   On Error GoTo Err
   rsCheck.Open "Select DirectLine2 from tenants", oConnAcc, adOpenKeyset, adLockReadOnly
   If rsCheck.Fields.Item("DirectLine2").DefinedSize = 20 Then
        rsCheck.Close
        oConnAcc.Execute "ALTER TABLE tenants ALTER COLUMN DirectLine2 text(40)"
        oConnAcc.Execute "ALTER TABLE tenants ALTER COLUMN HOTelephone text(40)"
        oConnAcc.Execute "ALTER TABLE tenants ALTER COLUMN BillFax text(40)"
        oConnAcc.Execute "ALTER TABLE tenants ALTER COLUMN HOFax text(40)"
        UpdateDatabase1 = 1
   Else
        rsCheck.Close
   End If
   
   rsCheck.Open "Select SageAccountNumber from tenants", oConnAcc, adOpenKeyset, adLockReadOnly
   If rsCheck.Fields.Item("SageAccountNumber").DefinedSize = 8 Then
   rsCheck.Close
        oConnAcc.Execute "ALTER TABLE tenants ALTER COLUMN SageAccountNumber text(30)"
        oConnAcc.Execute "ALTER TABLE   LeaseDetails    ALTER COLUMN SageAccountNumber text(30)"
        oConnAcc.Execute "ALTER TABLE   DemandRecords    ALTER COLUMN SageAccountNumber text(30)"
        oConnAcc.Execute "ALTER TABLE   DemandRecPreview     ALTER COLUMN SageAccountNumber text(30)"
        oConnAcc.Execute "ALTER TABLE   tlbChildDemandRecord     ALTER COLUMN SageAccountNumber text(30)"
        oConnAcc.Execute "ALTER TABLE   tlbDRCurrentPrint    ALTER COLUMN SageAccountNumber text(30)"
        oConnAcc.Execute "ALTER TABLE   tlbLetterReports     ALTER COLUMN SageAccountNumber text(30)"
        oConnAcc.Execute "ALTER TABLE   tlbReceipt   ALTER COLUMN SageAccountNumber text(30)"
        oConnAcc.Execute "ALTER TABLE   Units    ALTER COLUMN SageAccountNumber text(30)"
        oConnAcc.Execute "ALTER TABLE   PropertyMaintHistory     ALTER COLUMN ReportedBy text(30)"
        oConnAcc.Execute "ALTER TABLE   NLPosting    ALTER COLUMN ACCOUNT_NUMBER text(30)"
        UpdateDatabase1 = 1
   Else
        rsCheck.Close
   End If
   
   rsCheck.Open "Select PropertyID from Property", oConnAcc, adOpenKeyset, adLockReadOnly
   If rsCheck.Fields.Item("PropertyID").DefinedSize < 10 Then
        rsCheck.Close
        Set rsCheck = Nothing
        oConnAcc.Execute "ALTER TABLE Property ALTER COLUMN PropertyID text(10)"
        oConnAcc.Execute "ALTER TABLE ClientGlobalData ALTER COLUMN PropertyID text(10)"
        oConnAcc.Execute "ALTER TABLE ClientProAgr ALTER COLUMN PropertyID text(10)"
        oConnAcc.Execute "ALTER TABLE DemandTypes ALTER COLUMN PropertyID text(10)"
        oConnAcc.Execute "ALTER TABLE GlobalData ALTER COLUMN PropertyID text(10)"
        oConnAcc.Execute "ALTER TABLE GlobalInsurance ALTER COLUMN PropertyID text(10)"
        oConnAcc.Execute "ALTER TABLE GlobalRC  ALTER COLUMN PropertyID text(10)"
        oConnAcc.Execute "ALTER TABLE GlobalSC ALTER COLUMN PropertyID text(10)"
        oConnAcc.Execute "ALTER TABLE InterestRates  ALTER COLUMN PropertyID text(10)"
        oConnAcc.Execute "ALTER TABLE NLPosting ALTER COLUMN PROPERTY_ID text(10)"
        oConnAcc.Execute "ALTER TABLE PropertyAnalysis  ALTER COLUMN PropertyID text(10)"
        oConnAcc.Execute "ALTER TABLE PropertyInsurance  ALTER COLUMN PropertyID text(10)"
        oConnAcc.Execute "ALTER TABLE PropertyLandlord ALTER COLUMN PropertyID text(10)"
        oConnAcc.Execute "ALTER TABLE PropertyMaintHistory  ALTER COLUMN PropertyID text(10)"
        oConnAcc.Execute "ALTER TABLE PropertySafety  ALTER COLUMN PropertyID text(10)"
        oConnAcc.Execute "ALTER TABLE PropertyUtilities  ALTER COLUMN PropertyID text(10)"
        oConnAcc.Execute "ALTER TABLE tblBatchPayment  ALTER COLUMN PropertyID text(10)"
        oConnAcc.Execute "ALTER TABLE tblBatchReceipt  ALTER COLUMN PropertyID text(10)"
        oConnAcc.Execute "ALTER TABLE tblBatchTransaction  ALTER COLUMN PropertyID text(10)"
        oConnAcc.Execute "ALTER TABLE tblPurInv  ALTER COLUMN PropertyID text(10)"
        oConnAcc.Execute "ALTER TABLE tlbBankPayment ALTER COLUMN PropertyID text(10)"
        oConnAcc.Execute "ALTER TABLE tlbPayment  ALTER COLUMN UnitID text(10)"
        oConnAcc.Execute "ALTER TABLE tlbPaymentSplit ALTER COLUMN TRANS text(10)"
        oConnAcc.Execute "ALTER TABLE FundMatrix ALTER COLUMN PropertyID text(10)"
      UpdateDatabase1 = 1
      Exit Function
   Else
      rsCheck.Close
      Set rsCheck = Nothing
   End If
   'Increase property field size on table ChargeTypes to 10
   rsCheck.Open "Select PropertyID from ChargeTypes", oConnAcc, adOpenKeyset, adLockReadOnly
   If rsCheck.Fields.Item("PropertyID").DefinedSize < 10 Then
        rsCheck.Close
        Set rsCheck = Nothing
        oConnAcc.Execute "ALTER TABLE ChargeTypes ALTER COLUMN PropertyID text(10)"
        oConnAcc.Execute "ALTER TABLE PayableTypes ALTER COLUMN PropertyID text(10)"
        UpdateDatabase1 = 1
        Exit Function
   Else
      rsCheck.Close
      Set rsCheck = Nothing
   End If
ADD_Client_ClientAddressLine5:
        On Error GoTo MOD_Client_ClientAddressLine5
        rsCheck.Open "Select ClientAddressLine5 from Client;", oConnAcc, adOpenKeyset, adLockReadOnly
        rsCheck.Close
        GoTo Modify_DemandTypes_ID
        Exit Function
        
MOD_Client_ClientAddressLine5:
        oConnAcc.Execute "ALTER TABLE Client ADD COLUMN ClientAddressLine5 TEXT(250);"
        oConnAcc.Execute "ALTER TABLE Client ADD COLUMN ClientOfficeAddressLine5 TEXT(250);"
        oConnAcc.Execute "ALTER TABLE Client ADD COLUMN RegAdd5 TEXT(250);"
        UpdateDatabase1 = 1
        Exit Function
        
Modify_DemandTypes_ID:
         
        rsCheck.Open "SELECT ID FROM DemandTypes;", oConnAcc, adOpenStatic, adLockReadOnly
        If rsCheck.Fields(0).DefinedSize = 1 Then '3 means type long integer and 17 means Byte
            If rsCheck.State = 1 Then
               rsCheck.Close
            End If
            GoTo Mod_DemandTypes_ID
        Else
            If rsCheck.State = 1 Then
               rsCheck.Close
            End If
            GoTo Modify_LServiceCharges_SCDemandType
        End If
        Exit Function
Mod_DemandTypes_ID:
        oConnAcc.Execute "ALTER TABLE DemandTypes ALTER COLUMN ID Int;"
        UpdateDatabase1 = 1
        
Modify_LServiceCharges_SCDemandType:
         
        rsCheck.Open "SELECT SCDemandType FROM LServiceCharges;", oConnAcc, adOpenStatic, adLockReadOnly
        If rsCheck.Fields(0).DefinedSize = 1 Then '1 means type BIT integer and 4 means INT
            If rsCheck.State = 1 Then
               rsCheck.Close
            End If
            GoTo Mod_LServiceCharges_SCDemandType
        Else
            If rsCheck.State = 1 Then
               rsCheck.Close
            End If
            GoTo Modify_LServiceCharges_BRDemandType
        End If
        Exit Function
Mod_LServiceCharges_SCDemandType:
        oConnAcc.Execute "ALTER TABLE LServiceCharges ALTER COLUMN SCDemandType Int;"
        UpdateDatabase1 = 1
        
Modify_LServiceCharges_BRDemandType:
        rsCheck.Open "SELECT SCDemandType FROM LServiceCharges;", oConnAcc, adOpenStatic, adLockReadOnly
        If rsCheck.Fields(0).DefinedSize = 1 Then '1 means type BIT integer and 4 means INT
            If rsCheck.State = 1 Then
               rsCheck.Close
            End If
            GoTo Mod_LServiceCharges_BRDemandType
        Else
            If rsCheck.State = 1 Then
               rsCheck.Close
            End If
            GoTo Modify_LInsuranceCharges_InsuranceDemandType
        End If
        Exit Function
Mod_LServiceCharges_BRDemandType:
        oConnAcc.Execute "ALTER TABLE LServiceCharges ALTER COLUMN BRDemandType Int;"
        UpdateDatabase1 = 1
   
Modify_LInsuranceCharges_InsuranceDemandType:
        rsCheck.Open "SELECT InsuranceDemandType FROM LInsuranceCharges;", oConnAcc, adOpenStatic, adLockReadOnly
        If rsCheck.Fields(0).DefinedSize = 1 Then '1 means type BIT integer and 4 means INT
            If rsCheck.State = 1 Then
               rsCheck.Close
            End If
            GoTo Mod_LInsuranceCharges_InsuranceDemandType
        Else
            If rsCheck.State = 1 Then
               rsCheck.Close
            End If
            GoTo Modify_LRentCharges_BRDemandType
        End If
        Exit Function
Mod_LInsuranceCharges_InsuranceDemandType:
        oConnAcc.Execute "ALTER TABLE LInsuranceCharges ALTER COLUMN InsuranceDemandType Int;"
        UpdateDatabase1 = 1
        
Modify_LRentCharges_BRDemandType:
        rsCheck.Open "SELECT BRDemandType FROM LRentCharges;", oConnAcc, adOpenStatic, adLockReadOnly
        If rsCheck.Fields(0).DefinedSize = 1 Then '1 means type BIT integer and 4 means INT
            If rsCheck.State = 1 Then
               rsCheck.Close
            End If
            GoTo Mod_LRentCharges_BRDemandType
        Else
            If rsCheck.State = 1 Then
               rsCheck.Close
            End If
            GoTo ADD_PayableTypes_PrintTemplate
        End If
        Exit Function
Mod_LRentCharges_BRDemandType:
        oConnAcc.Execute "ALTER TABLE LRentCharges ALTER COLUMN BRDemandType Int;"
        UpdateDatabase1 = 1
        
ADD_PayableTypes_PrintTemplate:
        On Error GoTo MOD_PayableTypes_PrintTemplate
        rsCheck.Open "Select PrintTemplate from PayableTypes;", oConnAcc, adOpenStatic, adLockReadOnly
        rsCheck.Close
'        GoTo ADD_ClientProAgr_agreementstartdate
        Exit Function
MOD_PayableTypes_PrintTemplate:
        oConnAcc.Execute "ALTER TABLE PayableTypes ADD COLUMN PrintTemplate Text(255);"
        oConnAcc.Execute "ALTER TABLE PayableTypes ADD COLUMN EmailTemplate Text(255);"
         UpdateDatabase1 = 1
Err:
   
End Function
Private Function imp_PayableTypes(oConnEx As ADODB.Connection, oConnAcc As ADODB.Connection) As Boolean
'First we have to import the Payable types which property ID's are not Default
 'Secondly we have to import the default
    Print #1, "" & Chr$(13) + Chr$(10)
    Print #1, "" & Chr$(13) + Chr$(10)
    Print #1, "Data writing START: ********** Payable Types **********" & Chr$(13) + Chr$(10)
    Dim oNCsXls As New ADODB.Recordset
    Dim adoRstMdb As New ADODB.Recordset
    Dim oRstXls As New ADODB.Recordset
    Dim adoRST_ As New ADODB.Recordset
    Dim adoRst As New ADODB.Recordset
    Dim iLoop            As Integer
    
    Dim szTemp2           As String
    Dim iKount           As Integer
    Dim szClients        As String
    Dim bFreq            As Boolean
    Dim dtStDt           As Date
    Dim dtEndDt          As Date
   
    Dim szTemp As String
    Dim allPropertyID As String
    Dim szPropertyID() As String
    Dim i As Integer
    lblPlsWait(1).Caption = "While data is uploading: Payable Types"
'  Get the list of the Property in demand type worksheet
   oNCsXls.Open "SELECT [Property] FROM [PayableTypes$] where [Property] <> 'DEFAULT' AND LEFT([Property],7) <> 'Default' GROUP BY [Property]", oConnEx, adOpenStatic, adLockOptimistic
   szTemp = SQL2String(oNCsXls, 0)
   oNCsXls.Close
   If Len(szTemp) > 0 Then
        szPropertyID = Split(szTemp, ",")
        '  check for the property Column has the valid column values added by anol 20161124
        '   if you use property Name instead of property ID in the [PayableTypes$] then check the property sheet and return validation result.
            For i = 0 To UBound(szPropertyID)
                oNCsXls.Open "SELECT [PROPERTY ID] FROM [Properties$] WHERE [PROPERTY ID] =" & szPropertyID(i) & ";", oConnEx, adOpenStatic, adLockOptimistic
                If oNCsXls.EOF Then
                        oNCsXls.Close
                        Print #1, "PROPERTY ID is not valid in the PayableTypes Sheet PROPERTY column. " & szPropertyID(i) & Chr$(13) + Chr$(10)
                        MsgBox "PROPERTY ID is not valid in the PayableTypes Sheet PROPERTY column. : " & szPropertyID(i) & Chr$(13) + Chr$(10)
                        Print #1, 0 / 0
                        Exit Function
                Else
                    oNCsXls.Close
                End If
            Next i
      
   DoEvents
'first import Payable types for the property that is  specified and one next section import default property
   With adoRstMdb
       For i = 0 To UBound(szPropertyID) 'Loops for the new propertyID that is going to be inserted for new demand type
              .Open "SELECT * FROM PayableTypes WHERE PropertyID = " & Trim(szPropertyID(i)) & ";", _
                    oConnAcc, adOpenDynamic, adLockOptimistic
              oRstXls.Open "SELECT * FROM [PayableTypes$] WHERE [Property] = " & Trim(szPropertyID(i)) & ";", oConnEx, adOpenStatic, adLockOptimistic
              Do While Not oRstXls.EOF
                        If Not IsEmpty(oRstXls.Fields.Item("Payable TYPE NAME").Value) And Not IsNull(oRstXls.Fields.Item("Payable TYPE NAME").Value) Then
                                   .Find "PayType = '" & oRstXls.Fields.Item("Payable TYPE NAME").Value & "' ", , , 1
                                   If .EOF Then
                                        adoRST_.Open "SELECT MAX(ID) AS A FROM PayableTypes", oConnAcc, adOpenStatic, adLockReadOnly
                                        iLoop = IIf(adoRST_.EOF, 1, IIf(IsNull(adoRST_.Fields.Item("A").Value), 0, adoRST_.Fields.Item("A").Value) + 1)
                                        adoRST_.Close
                                        .AddNew
                                        .Fields.Item("ID").Value = iLoop
                                        .Fields.Item("PayType").Value = oRstXls.Fields.Item("Payable TYPE NAME").Value
                                        .Fields.Item("PayIC").Value = "X"
                                   End If
                    
                                   .Fields.Item("PaySagePrefix").Value = oRstXls.Fields.Item("PREFIX").Value
                    
                                   If UCase(oRstXls.Fields.Item("PAYMENT DATES").Value) = "DEFAULT" Then
                                      .Fields.Item("PaymentDates").Value = 0
                                   Else
                                      adoRST_.Open "SELECT DateSetID FROM PaymentDates WHERE NameOfSet = '" & oRstXls.Fields.Item("PAYMENT DATES").Value & "';", oConnAcc, adOpenStatic, adLockReadOnly
                                      If adoRST_.EOF Then
                                         .Fields.Item("PaymentDates").Value = 0
                                      Else
                                         .Fields.Item("PaymentDates").Value = adoRST_.Fields.Item("DateSetID").Value
                                      End If
                                      adoRST_.Close
                                   End If
                                   .Fields.Item("PropertyID").Value = Trim(Replace(szPropertyID(i), "'", ""))
                                   .Fields.Item("ClientID").Value = FindClient(.Fields.Item("PropertyID").Value, oConnAcc)
                                    If .Fields.Item("ClientID").Value = "" Then
                                            Print #1, "Client Id was not found for the property in 'Payable Types' excel sheet : " & .Fields.Item("PropertyID").Value & Chr$(13) + Chr$(10)
                                            MsgBox "Client Id was not found for the property in 'Payable Types' excel sheet  : " & .Fields.Item("PropertyID").Value
                                            End
                                     End If
                                    .Fields.Item("PrintTemplate").Value = "ClientStatement.rpt"
                                   .Fields.Item("EmailTemplate").Value = "ClientStatement.rpt"
                                   .Update
                    
                                   Print #1, "Data - P: " & Trim(Replace(szPropertyID(i), "'", "")) & " DT: " & oRstXls.Fields.Item(1).Value & "" & Chr$(13) + Chr$(10)
                                
                    
                                    oRstXls.MoveNext
                                    If oRstXls.EOF Then Exit Do
                                    If oRstXls.Fields.Item(0).Value = "END" Then Exit Do
                        End If
            Loop ' Lopp ends for excel sheet Payable TYPE NAME loop
'         adoRST.MoveNext
         .Close
         oRstXls.Close
        Next 'End of propertyID loop

'    End If
        End With
    ' 2nd part start
    End If
'  Get the recordset
   ' 2nd part start.................this section import default property
        With adoRstMdb
' oNCsXls.Close
        oNCsXls.Open "SELECT [Property] FROM [PayableTypes$] where [Property] = 'DEFAULT' OR LEFT([Property],7) = 'Default' GROUP BY [Property]", oConnEx, adOpenStatic, adLockOptimistic
        szTemp2 = SQL2String(oNCsXls, 0)
        If Len(szTemp2) > 0 Then
      'problem was in below line  fixed by anol 25 Sep 2015
        If Len(szTemp) > 0 Then
            adoRst.Open "SELECT [PROPERTY ID] as PropertyID FROM [Properties$] where [PROPERTY ID] NOT IN (" & szTemp & ");", oConnEx, adOpenStatic, adLockReadOnly
        Else
            adoRst.Open "SELECT [PROPERTY ID] as PropertyID FROM [Properties$] ;", oConnEx, adOpenStatic, adLockReadOnly
        End If
        szTemp = SQL2String(adoRst, 0)
        szPropertyID = Split(szTemp, ",")

       For i = 0 To UBound(szPropertyID)
        
        oRstXls.Open "SELECT * FROM [PayableTypes$] WHERE [Property] = 'DEFAULT' OR LEFT([Property],7) = 'Default' ;", oConnEx, adOpenStatic, adLockOptimistic
         Do While Not oRstXls.EOF
            If Not IsEmpty(oRstXls.Fields.Item("Payable TYPE NAME").Value) And Not IsNull(oRstXls.Fields.Item("Payable TYPE NAME").Value) Then
                     
                    .Open "SELECT * FROM PayableTypes WHERE PropertyID = " & Trim(szPropertyID(i)) & ";", _
                          oConnAcc, adOpenDynamic, adLockOptimistic
                           .Find "Type = '" & oRstXls.Fields.Item("Payable TYPE NAME").Value & "' ", , , 1
        
                        If .EOF Then 'if Payable TYPE NAME that exists in excel not found in the database then append it to the database
                           adoRST_.Open "SELECT MAX(ID) AS A FROM PayableTypes", oConnAcc, adOpenStatic, adLockReadOnly
                           iLoop = IIf(adoRST_.EOF, 1, IIf(IsNull(adoRST_.Fields.Item("A").Value), 0, adoRST_.Fields.Item("A").Value) + 1)
                           adoRST_.Close
                           .AddNew
                           .Fields.Item("ID").Value = iLoop
                           .Fields.Item("PayType").Value = oRstXls.Fields.Item("Payable TYPE NAME").Value
                           allPropertyID = Replace(Replace(Trim(szPropertyID(i)), "'", ""), "'", "")
                        End If
                        .Fields.Item("PaySagePrefix").Value = oRstXls.Fields.Item("PREFIX").Value
            
                        If UCase(oRstXls.Fields.Item("PAYMENT DATES").Value) = "DEFAULT" Then
                           .Fields.Item("PaymentDates").Value = 0
                        Else
                           adoRST_.Open "SELECT DateSetID FROM PaymentDates WHERE NameOfSet = '" & oRstXls.Fields.Item("PAYMENT DATES").Value & "';", oConnAcc, adOpenStatic, adLockReadOnly
                           If adoRST_.EOF Then
                              .Fields.Item("PaymentDates").Value = 0
                           Else
                              .Fields.Item("PaymentDates").Value = adoRST_.Fields.Item("DateSetID").Value
                           End If
                           adoRST_.Close
                        End If
                        .Fields.Item("PropertyID").Value = Trim(Replace(szPropertyID(i), "'", ""))  'adoRST.Fields.Item("PropertyID").Value
                        .Fields.Item("ClientID").Value = FindClient(.Fields.Item("PropertyID").Value, oConnAcc)
                                   If .Fields.Item("ClientID").Value = "" Then
                                        Print #1, "Client Id was not found for the property 'ChargeTypes' excel sheet  : " & .Fields.Item("PropertyID").Value & Chr$(13) + Chr$(10)
                                        MsgBox "Client Id was not found for the property 'ChargeTypes' excel sheet  : " & .Fields.Item("PropertyID").Value
                                        End
                                   End If
                        .Fields.Item("PrintTemplate").Value = "ClientStatement.rpt"
                        .Fields.Item("EmailTemplate").Value = "ClientStatement.rpt"
                                   
                        .Update
                
                        Print #1, "Data - P: " & Replace(szPropertyID(i), "'", "") & " U: " & oRstXls.Fields.Item(1).Value & "" & Chr$(13) + Chr$(10)
                        .Close
                         End If
                    
                     oRstXls.MoveNext
                     If oRstXls.EOF Then Exit Do
                     If oRstXls.Fields.Item(0).Value = "END" Then Exit Do
             Loop
             oRstXls.Close
          Next
        End If 'end if for sztemp length
   
   End With
   imp_PayableTypes = True
   Print #1, "Data writing FINISH: ********** Demand Types **********" & Chr$(13) + Chr$(10)
End Function
Private Function FindClient(szPropertyID, oConnAcc As ADODB.Connection)
    Dim rsClientID As New ADODB.Recordset
     rsClientID.Open "SELECT * FROM Property WHERE PropertyID = '" & szPropertyID & "'", _
                    oConnAcc, adOpenDynamic, adLockOptimistic
                    FindClient = ""
                 If Not rsClientID.EOF Then
                    FindClient = rsClientID("ClientID").Value
                 End If
                If FindClient = "" Then
                   Debug.Print szPropertyID
                End If
                 rsClientID.Close
                 
                 
End Function
Private Function imp_ChargeTypes(oConnEx As ADODB.Connection, oConnAcc As ADODB.Connection) As Boolean
'First we have to import the demand types which property ID's are not Default
 'Secondly we have to import the default
    Print #1, "" & Chr$(13) + Chr$(10)
    Print #1, "" & Chr$(13) + Chr$(10)
    Print #1, "Data writing START: ********** Charge Types **********" & Chr$(13) + Chr$(10)
    Dim oNCsXls As New ADODB.Recordset
    Dim adoRstMdb As New ADODB.Recordset
    Dim oRstXls As New ADODB.Recordset
    Dim adoRST_ As New ADODB.Recordset
    Dim adoRst As New ADODB.Recordset
    Dim iLoop            As Integer
    
    Dim szTemp2           As String
    Dim iKount           As Integer
    Dim szClients        As String
    Dim bFreq            As Boolean
    Dim dtStDt           As Date
    Dim dtEndDt          As Date
   
    Dim szTemp As String
    Dim allPropertyID As String
    Dim szPropertyID() As String
    Dim i As Integer
    lblPlsWait(1).Caption = "While data is uploading: Charge Types"
'  Get the list of the Property in demand type worksheet
   oNCsXls.Open "SELECT [Property] FROM [ChargeTypes$] where [Property] <> 'DEFAULT' AND LEFT([Property],7) <> 'Default' GROUP BY [Property]", oConnEx, adOpenStatic, adLockOptimistic
   szTemp = SQL2String(oNCsXls, 0)
   oNCsXls.Close
   If Len(szTemp) > 0 Then
        szPropertyID = Split(szTemp, ",")
        '  check for the property Column has the valid column values added by anol 20161124
        '   if you use property Name instead of property ID in the [ChargeTypes$] then check the property sheet and return validation result.
            For i = 0 To UBound(szPropertyID)
                oNCsXls.Open "SELECT [PROPERTY ID] FROM [Properties$] WHERE [PROPERTY ID] =" & szPropertyID(i) & ";", oConnEx, adOpenStatic, adLockOptimistic
                If oNCsXls.EOF Then
                        oNCsXls.Close
                        Print #1, "PROPERTY ID is not valid in the ChargeTypes Sheet PROPERTY column.  " & szPropertyID(i) & Chr$(13) + Chr$(10)
                        MsgBox "PROPERTY ID is not valid in the ChargeTypes Sheet PROPERTY column. : " & szPropertyID(i) & Chr$(13) + Chr$(10)
                        Print #1, 0 / 0
                        Exit Function
                Else
                    oNCsXls.Close
                End If
            Next i
      
   DoEvents
'first import charge types for the property that is  specified and one next section import default property
   With adoRstMdb
       For i = 0 To UBound(szPropertyID) 'Loops for the new propertyID that is going to be inserted for new demand type
              .Open "SELECT * FROM ChargeTypes WHERE PropertyID = " & Trim(szPropertyID(i)) & ";", _
                    oConnAcc, adOpenDynamic, adLockOptimistic
              oRstXls.Open "SELECT * FROM [ChargeTypes$] WHERE [Property] = " & Trim(szPropertyID(i)) & ";", oConnEx, adOpenStatic, adLockOptimistic
              Do While Not oRstXls.EOF
                        If Not IsEmpty(oRstXls.Fields.Item("CHARGE TYPE NAME").Value) And Not IsNull(oRstXls.Fields.Item("CHARGE TYPE NAME").Value) Then
                                   .Find "Type = '" & oRstXls.Fields.Item("CHARGE TYPE NAME").Value & "' ", , , 1
                                   If .EOF Then
                                        adoRST_.Open "SELECT MAX(ID) AS A FROM ChargeTypes", oConnAcc, adOpenStatic, adLockReadOnly
                                        iLoop = IIf(adoRST_.EOF, 1, IIf(IsNull(adoRST_.Fields.Item("A").Value), 0, adoRST_.Fields.Item("A").Value) + 1)
                                        adoRST_.Close
                                        .AddNew
                                        .Fields.Item("ID").Value = iLoop
                                        .Fields.Item("FEEType").Value = oRstXls.Fields.Item("CHARGE TYPE NAME").Value
                                   End If
                    
                                   .Fields.Item("Prefix").Value = oRstXls.Fields.Item("PREFIX").Value
                    
                                   If UCase(oRstXls.Fields.Item("PAYMENT DATES").Value) = "DEFAULT" Then
                                      .Fields.Item("PaymentDates").Value = 0
                                   Else
                                      adoRST_.Open "SELECT DateSetID FROM PaymentDates WHERE NameOfSet = '" & oRstXls.Fields.Item("PAYMENT DATES").Value & "';", oConnAcc, adOpenStatic, adLockReadOnly
                                      If adoRST_.EOF Then
                                         .Fields.Item("PaymentDates").Value = 0
                                      Else
                                         .Fields.Item("PaymentDates").Value = adoRST_.Fields.Item("DateSetID").Value
                                      End If
                                      adoRST_.Close
                                   End If
                                   .Fields.Item("PropertyID").Value = Trim(Replace(szPropertyID(i), "'", ""))
                                   .Fields.Item("ClientID").Value = FindClient(.Fields.Item("PropertyID").Value, oConnAcc)
                                   If .Fields.Item("ClientID").Value = "" Then
                                        Print #1, "Client Id was not found for the property 'ChargeTypes' excel sheet  : " & .Fields.Item("PropertyID").Value & Chr$(13) + Chr$(10)
                                        MsgBox "Client Id was not found for the property 'ChargeTypes' excel sheet  : " & .Fields.Item("PropertyID").Value
                                        End
                                   End If
                    
                                   .Update
                    
                                   Print #1, "Data - P: " & Trim(Replace(szPropertyID(i), "'", "")) & " DT: " & oRstXls.Fields.Item(1).Value & "" & Chr$(13) + Chr$(10)
                                
                    
                                    oRstXls.MoveNext
                                    If oRstXls.EOF Then Exit Do
                                    If oRstXls.Fields.Item(0).Value = "END" Then Exit Do
                        End If
            Loop ' Lopp ends for excel sheet CHARGE TYPE NAME loop
'         adoRST.MoveNext
         .Close
         oRstXls.Close
        Next 'End of propertyID loop

'    End If
        End With
    ' 2nd part start
    End If
'  Get the recordset
   ' 2nd part start.................this section import default property
        With adoRstMdb
' oNCsXls.Close
        oNCsXls.Open "SELECT [Property] FROM [ChargeTypes$] where [Property] = 'DEFAULT' OR LEFT([Property],7) = 'Default' GROUP BY [Property]", oConnEx, adOpenStatic, adLockOptimistic
        szTemp2 = SQL2String(oNCsXls, 0)
        If Len(szTemp2) > 0 Then
      'problem was in below line  fixed by anol 25 Sep 2015
        If Len(szTemp) > 0 Then
            adoRst.Open "SELECT [PROPERTY ID] as PropertyID FROM [Properties$] where [PROPERTY ID] NOT IN (" & szTemp & ");", oConnEx, adOpenStatic, adLockReadOnly
        Else
            adoRst.Open "SELECT [PROPERTY ID] as PropertyID FROM [Properties$] ;", oConnEx, adOpenStatic, adLockReadOnly
        End If
        szTemp = SQL2String(adoRst, 0)
        szPropertyID = Split(szTemp, ",")

       For i = 0 To UBound(szPropertyID)
        
        oRstXls.Open "SELECT * FROM [ChargeTypes$] WHERE [Property] = 'DEFAULT' OR LEFT([Property],7) = 'Default' ;", oConnEx, adOpenStatic, adLockOptimistic
         Do While Not oRstXls.EOF
            If Not IsEmpty(oRstXls.Fields.Item("CHARGE TYPE NAME").Value) And Not IsNull(oRstXls.Fields.Item("CHARGE TYPE NAME").Value) Then
                     
                    .Open "SELECT * FROM ChargeTypes WHERE PropertyID = " & Trim(szPropertyID(i)) & ";", _
                          oConnAcc, adOpenDynamic, adLockOptimistic
                           .Find "FEEType = '" & oRstXls.Fields.Item("CHARGE TYPE NAME").Value & "' ", , , 1
        
                        If .EOF Then 'if CHARGE TYPE NAME that exists in excel not found in the database then append it to the database
                           adoRST_.Open "SELECT MAX(ID) AS A FROM ChargeTypes", oConnAcc, adOpenStatic, adLockReadOnly
                           iLoop = IIf(adoRST_.EOF, 1, IIf(IsNull(adoRST_.Fields.Item("A").Value), 0, adoRST_.Fields.Item("A").Value) + 1)
                           adoRST_.Close
                           .AddNew
                           .Fields.Item("ID").Value = iLoop
                           .Fields.Item("FEEType").Value = oRstXls.Fields.Item("CHARGE TYPE NAME").Value
                           .Fields.Item("FeeIC").Value = "X"
                           allPropertyID = Replace(Replace(Trim(szPropertyID(i)), "'", ""), "'", "")
                        End If
                        .Fields.Item("FeeSagePrefix").Value = oRstXls.Fields.Item("PREFIX").Value
            
                        If UCase(oRstXls.Fields.Item("PAYMENT DATES").Value) = "DEFAULT" Then
                           .Fields.Item("PaymentDates").Value = 0
                        Else
                           adoRST_.Open "SELECT DateSetID FROM PaymentDates WHERE NameOfSet = '" & oRstXls.Fields.Item("PAYMENT DATES").Value & "';", oConnAcc, adOpenStatic, adLockReadOnly
                           If adoRST_.EOF Then
                              .Fields.Item("PaymentDates").Value = 0
                           Else
                              .Fields.Item("PaymentDates").Value = adoRST_.Fields.Item("DateSetID").Value
                           End If
                           adoRST_.Close
                        End If
                        .Fields.Item("PropertyID").Value = Trim(Replace(szPropertyID(i), "'", ""))  'adoRST.Fields.Item("PropertyID").Value
                        .Fields.Item("ClientID").Value = FindClient(.Fields.Item("PropertyID").Value, oConnAcc)
                                   If .Fields.Item("ClientID").Value = "" Then
                                        Print #1, "Client Id was not found for the property 'ChargeTypes' excel sheet  : " & .Fields.Item("PropertyID").Value & Chr$(13) + Chr$(10)
                                        MsgBox "Client Id was not found for the property 'ChargeTypes' excel sheet  : " & .Fields.Item("PropertyID").Value
                                        End
                                   End If
                                   
                        .Update
            
                        Print #1, "Data - P: " & Replace(szPropertyID(i), "'", "") & " U: " & oRstXls.Fields.Item(1).Value & "" & Chr$(13) + Chr$(10)
                        .Close
                         End If
                    
                     oRstXls.MoveNext
                     If oRstXls.EOF Then Exit Do
                     If oRstXls.Fields.Item(0).Value = "END" Then Exit Do
             Loop
             oRstXls.Close
          Next
        End If 'end if for sztemp length
   
   End With
   imp_ChargeTypes = True
   Print #1, "Data writing FINISH: ********** Demand Types **********" & Chr$(13) + Chr$(10)
End Function
Private Function imp_demandTypes(oConnEx As ADODB.Connection, oConnAcc As ADODB.Connection) As Boolean
'First we have to import the demand types which property ID's are not Default
 'Secondly we have to import the default
    Print #1, "" & Chr$(13) + Chr$(10)
    Print #1, "" & Chr$(13) + Chr$(10)
    Print #1, "Data writing START: ********** Demand Types **********" & Chr$(13) + Chr$(10)
    Dim oNCsXls As New ADODB.Recordset
    Dim adoRstMdb As New ADODB.Recordset
    Dim oRstXls As New ADODB.Recordset
    Dim adoRST_ As New ADODB.Recordset
    Dim adoRst As New ADODB.Recordset
    Dim iLoop            As Integer
    
    Dim szTemp2           As String
    Dim iKount           As Integer
    Dim szClients        As String
    Dim bFreq            As Boolean
    Dim dtStDt           As Date
    Dim dtEndDt          As Date
   
    Dim szTemp As String
    Dim allPropertyID As String
    Dim szPropertyID() As String
    Dim i As Integer
    lblPlsWait(1).Caption = "While data is uploading: Demand Types"
'  Get the list of the Property in demand type worksheet
   oNCsXls.Open "SELECT [Property] FROM [DemandTypes$] where [Property] <> 'DEFAULT' AND LEFT([Property],7) <> 'Default' GROUP BY [Property]", oConnEx, adOpenStatic, adLockOptimistic
   szTemp = SQL2String(oNCsXls, 0)
   oNCsXls.Close
   If Len(szTemp) > 0 Then
        szPropertyID = Split(szTemp, ",")
'           oNCsXls.Close
        '  check for the property Column has the valid column values added by anol 20161124
        '   if you use property Name instead of property ID in the [DemandTypes$] then check the property sheet and return validation result.
'         i = 0
            For i = 0 To UBound(szPropertyID)
                oNCsXls.Open "SELECT [PROPERTY ID] FROM [Properties$] WHERE [PROPERTY ID] =" & szPropertyID(i) & ";", oConnEx, adOpenStatic, adLockOptimistic
                If oNCsXls.EOF Then
                        oNCsXls.Close
                        Print #1, "PROPERTY ID is not valid in the DemandTypes Sheet PROPERTY column. " & szPropertyID(i) & Chr$(13) + Chr$(10)
                        'added by anol 20161124
                        MsgBox "PROPERTY ID is not valid in the DemandTypes Sheet PROPERTY column. : " & szPropertyID(i) & Chr$(13) + Chr$(10)
                        Print #1, 0 / 0
                        Exit Function
                Else
                    oNCsXls.Close
                End If
            Next i
            
        '  Get the list of the Property which will be set for default demand types
        'Debug.Print "SELECT PropertyID FROM Property WHERE PropertyName NOT IN (" & szTemp & ");"
       
'  Get the recordset
' or [Property] = 'Default' added by anol 24 Sep 2015
   
   ' for testing
'   oRstXls.Open "SELECT * FROM [DemandTypes$] "'
'   oRstXls.MoveNext'
'    oRstXls.Close
  
   'WaitFewSec
   DoEvents
'first import demand types for the property that is  specified and one next section import default property
   With adoRstMdb
       For i = 0 To UBound(szPropertyID) 'Loops for the new propertyID that is going to be inserted for new demand type
       
       If Trim(szPropertyID(i)) = "61AR" Then
            Debug.Print ""
        End If
              .Open "SELECT * FROM DemandTypes WHERE PropertyID = " & Trim(szPropertyID(i)) & ";", _
                    oConnAcc, adOpenDynamic, adLockOptimistic
        'added by anol 26 Jan 2015
              oRstXls.Open "SELECT * FROM [DemandTypes$] WHERE [Property] = " & Trim(szPropertyID(i)) & ";", oConnEx, adOpenStatic, adLockOptimistic
              Do While Not oRstXls.EOF
                        If Not IsEmpty(oRstXls.Fields.Item("DEMAND TYPE NAME").Value) And Not IsNull(oRstXls.Fields.Item("DEMAND TYPE NAME").Value) Then
                                   .Find "Type = '" & oRstXls.Fields.Item("DEMAND TYPE NAME").Value & "' ", , , 1
                                   If .EOF Then
                                        adoRST_.Open "SELECT MAX(ID) AS A FROM DemandTypes", oConnAcc, adOpenStatic, adLockReadOnly
                                        iLoop = IIf(adoRST_.EOF, 1, IIf(IsNull(adoRST_.Fields.Item("A").Value), 0, adoRST_.Fields.Item("A").Value) + 1)
                                        adoRST_.Close
                                        .AddNew
                                        .Fields.Item("ID").Value = iLoop
                                        .Fields.Item("Type").Value = oRstXls.Fields.Item("DEMAND TYPE NAME").Value
                                        .Fields.Item("spare1").Value = GetBankID(Trim(Replace(szPropertyID(i), "'", "")), oConnAcc)      'BANK ID
                                        If .Fields.Item("spare1").Value = "" Or IsNull(.Fields.Item("spare1").Value) Then
                                            'Debug.Print "Error"
                                            Print #1, "Failed to get the BankID for the demandtype : " & oRstXls.Fields.Item("DEMAND TYPE NAME").Value & Chr$(13) + Chr$(10)
                                            MsgBox "Failed to get the BankID for the demandtype : " & oRstXls.Fields.Item("DEMAND TYPE NAME").Value
                                            End
                                        End If
                                   End If
                    
                                   .Fields.Item("Prefix").Value = oRstXls.Fields.Item("PREFIX").Value
                                   .Fields.Item("NominalCodeforAmount").Value = oRstXls.Fields.Item("NOMINAL CODE").Value
                                   If Left(oRstXls.Fields.Item("DEMAND CATEGORY").Value, 1) = "R" Then _
                                      .Fields.Item("CategoryCode").Value = 1
                                   If Left(oRstXls.Fields.Item("DEMAND CATEGORY").Value, 1) = "S" Then _
                                      .Fields.Item("CategoryCode").Value = 2
                                   If Left(oRstXls.Fields.Item("DEMAND CATEGORY").Value, 1) = "I" Then _
                                      .Fields.Item("CategoryCode").Value = 3
                                   If Left(oRstXls.Fields.Item("DEMAND CATEGORY").Value, 1) = "O" Then _
                                      .Fields.Item("CategoryCode").Value = 4
                    
                                   If UCase(oRstXls.Fields.Item("PAYMENT DATES").Value) = "DEFAULT" Then
                                      .Fields.Item("PaymentDates").Value = 0
                                   Else
                                      adoRST_.Open "SELECT DateSetID FROM PaymentDates WHERE NameOfSet = '" & oRstXls.Fields.Item("PAYMENT DATES").Value & "';", oConnAcc, adOpenStatic, adLockReadOnly
                                      If adoRST_.EOF Then
                                         .Fields.Item("PaymentDates").Value = 0
                                      Else
                                         .Fields.Item("PaymentDates").Value = adoRST_.Fields.Item("DateSetID").Value
                                      End If
                                      adoRST_.Close
                                   End If
                                   .Fields.Item("PropertyID").Value = Trim(Replace(szPropertyID(i), "'", ""))
                                   .Fields.Item("DemandReportName").Value = "InvDemandSngMPage.rpt"
                                   .Fields.Item("EmailInvoiceTemplate").Value = "InvDemandSngMPage.rpt"
                    
                                   .Update
                    
                                   Print #1, "Data - P: " & Trim(Replace(szPropertyID(i), "'", "")) & " DT: " & oRstXls.Fields.Item(1).Value & "" & Chr$(13) + Chr$(10)
                                
                    
                                    oRstXls.MoveNext
                                    If oRstXls.EOF Then Exit Do
                                    If oRstXls.Fields.Item(0).Value = "END" Then Exit Do
                        End If
            Loop ' Lopp ends for excel sheet demand type name loop
'         adoRST.MoveNext
         .Close
         oRstXls.Close
        Next 'End of propertyID loop
'        adoRST.Close
'        oRstXls.Close
'    End If
        End With
    ' 2nd part start
    End If
'  Get the recordset
   ' 2nd part start.................this section import default property
        With adoRstMdb
' oNCsXls.Close
        oNCsXls.Open "SELECT [Property] FROM [DemandTypes$] where [Property] = 'DEFAULT' OR LEFT([Property],7) = 'Default' GROUP BY [Property]", oConnEx, adOpenStatic, adLockOptimistic
        szTemp2 = SQL2String(oNCsXls, 0)
        If Len(szTemp2) > 0 Then
      'problem was in below line  fixed by anol 25 Sep 2015
     ' adoRST.Open "SELECT PropertyID, PropertyName FROM Property WHERE PropertyName IN (" & szTemp & ");", oConnAcc, adOpenStatic, adLockReadOnly
        If Len(szTemp) > 0 Then
            adoRst.Open "SELECT [PROPERTY ID] as PropertyID FROM [Properties$] where [PROPERTY ID] NOT IN (" & szTemp & ");", oConnEx, adOpenStatic, adLockReadOnly
        Else
            adoRst.Open "SELECT [PROPERTY ID] as PropertyID FROM [Properties$] ;", oConnEx, adOpenStatic, adLockReadOnly
        End If
'        adoRST.Open "SELECT DISTINCT[PROPERTY ID] as PropertyID FROM [Properties$] AS P LEFT JOIN [DemandTypes$] AS D " & _
'                               "ON P.[PROPERTY ID]<> D.[PROPERTY] where P.[PROPERTY ID]<> D.[PROPERTY];", oConnEx, adOpenStatic, adLockReadOnly
        szTemp = SQL2String(adoRst, 0)
        szPropertyID = Split(szTemp, ",")
'       i = 0
'adoRST.Close
'        For i = 0 To UBound(szPropertyID)
'                 If InStr(1, szPropertyID(i), "'") > 0 Then
'                        allPropertyID = allPropertyID & szPropertyID(i)
'                  End If
'        Next i
'        If Len(allPropertyID) > 0 Then
'            MsgBox "PropertyID in Properties$ excle sheet contains apostrophe :" & allPropertyID, vbInformation, "Warning"
'            Exit Function
'        End If
       For i = 0 To UBound(szPropertyID)
        
        oRstXls.Open "SELECT * FROM [DemandTypes$] WHERE [Property] = 'DEFAULT' OR LEFT([Property],7) = 'Default' ;", oConnEx, adOpenStatic, adLockOptimistic
         Do While Not oRstXls.EOF
            If Not IsEmpty(oRstXls.Fields.Item("DEMAND TYPE NAME").Value) And Not IsNull(oRstXls.Fields.Item("DEMAND TYPE NAME").Value) Then
                    'added by anol 25 Jan 2015
        '            If Not adoRST.EOF Then
        '               adoRST.MoveNext
        '             'fixed by anol 25 Jan 2015
        '           ' adoRST.Find "PropertyName = '" & oRstXls.Fields.Item("Property").Value & "' ", , , 1
        '           adoRST.Find "PropertyID = '" & oRstXls.Fields.Item("Property").Value & "' ", , , 1
        '            If Not adoRST.EOF Then
                     
                    .Open "SELECT * FROM DemandTypes WHERE PropertyID = " & Trim(szPropertyID(i)) & ";", _
                          oConnAcc, adOpenDynamic, adLockOptimistic
                           .Find "Type = '" & oRstXls.Fields.Item("DEMAND TYPE NAME").Value & "' ", , , 1
        '                   End If
        '            End If
        '           If .State = 1 Then
        
                        If .EOF Then 'if demand type name that exists in excel not found in the database then append it to the database
                           adoRST_.Open "SELECT MAX(ID) AS A FROM DemandTypes", oConnAcc, adOpenStatic, adLockReadOnly
                           iLoop = IIf(adoRST_.EOF, 1, IIf(IsNull(adoRST_.Fields.Item("A").Value), 0, adoRST_.Fields.Item("A").Value) + 1)
                           adoRST_.Close
                           .AddNew
                           .Fields.Item("ID").Value = iLoop
                           .Fields.Item("Type").Value = oRstXls.Fields.Item("DEMAND TYPE NAME").Value
'                           szTemp = Replace(szPropertyID(i), "'", "")
'                           .Fields.Item("spare1").Value = GetBankID(Replace(Replace(Trim(szPropertyID(i)), "'", ""), "'", ""), oConnAcc)   'BANK ID
'                           If .Fields.Item("spare1").Value = "" Or IsNull(.Fields.Item("spare1").Value) Then
'                                            'Debug.Print "Error"
'                                             Print #1, "Property: " & allPropertyID & " does not have a bank ID in the client Bank table" & Chr$(13) + Chr$(10)
'                         End If
                           allPropertyID = Replace(Replace(Trim(szPropertyID(i)), "'", ""), "'", "")
                        End If
                       .Fields.Item("spare1").Value = GetBankID(Replace(Replace(Trim(szPropertyID(i)), "'", ""), "'", ""), oConnAcc)   'BANK ID
                           If .Fields.Item("spare1").Value = "" Or IsNull(.Fields.Item("spare1").Value) Then
                                            'Debug.Print "Error"
                                             Print #1, "Property: " & allPropertyID & " does not have a bank ID in the client Bank table" & Chr$(13) + Chr$(10)
                         End If
                        .Fields.Item("Prefix").Value = oRstXls.Fields.Item("PREFIX").Value
                        .Fields.Item("NominalCodeforAmount").Value = oRstXls.Fields.Item("NOMINAL CODE").Value
                        If Left(oRstXls.Fields.Item("DEMAND CATEGORY").Value, 1) = "R" Then _
                           .Fields.Item("CategoryCode").Value = 1
                        If Left(oRstXls.Fields.Item("DEMAND CATEGORY").Value, 1) = "S" Then _
                           .Fields.Item("CategoryCode").Value = 2
                        If Left(oRstXls.Fields.Item("DEMAND CATEGORY").Value, 1) = "I" Then _
                           .Fields.Item("CategoryCode").Value = 3
                        If Left(oRstXls.Fields.Item("DEMAND CATEGORY").Value, 1) = "O" Then _
                           .Fields.Item("CategoryCode").Value = 4
            
                        If UCase(oRstXls.Fields.Item("PAYMENT DATES").Value) = "DEFAULT" Then
                           .Fields.Item("PaymentDates").Value = 0
                        Else
                           adoRST_.Open "SELECT DateSetID FROM PaymentDates WHERE NameOfSet = '" & oRstXls.Fields.Item("PAYMENT DATES").Value & "';", oConnAcc, adOpenStatic, adLockReadOnly
                           If adoRST_.EOF Then
                              .Fields.Item("PaymentDates").Value = 0
                           Else
                              .Fields.Item("PaymentDates").Value = adoRST_.Fields.Item("DateSetID").Value
                           End If
                           adoRST_.Close
                        End If
                        .Fields.Item("PropertyID").Value = Trim(Replace(szPropertyID(i), "'", ""))  'adoRST.Fields.Item("PropertyID").Value
                        .Fields.Item("DemandReportName").Value = "InvDemandSngMPage.rpt"
                        .Fields.Item("EmailInvoiceTemplate").Value = "InvDemandSngMPage.rpt"
                         If .Fields.Item("spare1").Value = "" Or IsNull(.Fields.Item("spare1").Value) Then
                                            'Debug.Print "Error"
                                             Print #1, "Property: " & allPropertyID & " does not have a bank ID in the client Bank table" & Chr$(13) + Chr$(10)
                         End If
                        .Update
            
                        Print #1, "Data - P: " & Replace(szPropertyID(i), "'", "") & " U: " & oRstXls.Fields.Item(1).Value & "" & Chr$(13) + Chr$(10)
                        .Close
                         End If
        '             End If
            
                     oRstXls.MoveNext
                     If oRstXls.EOF Then Exit Do
                     If oRstXls.Fields.Item(0).Value = "END" Then Exit Do
                    
             Loop
             oRstXls.Close
          Next
'          adoRST.Close
        End If 'end if for sztemp length
   
   End With
   imp_demandTypes = True
   Print #1, "Data writing FINISH: ********** Demand Types **********" & Chr$(13) + Chr$(10)
End Function
Private Function import_nominalCategory(oConnEx As ADODB.Connection, oConnAcc As ADODB.Connection) As Boolean
    Dim oNCsXls As New ADODB.Recordset
    Dim adoClient As New ADODB.Recordset
    Dim adoRstMdb As New ADODB.Recordset
    Dim szTemp As String
                   Print #1, "" & Chr$(13) + Chr$(10)
                   Print #1, "" & Chr$(13) + Chr$(10)
                   Print #1, "Data writing START: ********** REPORT CATEGORY **********" & Chr$(13) + Chr$(10)
                   lblPlsWait(1).Caption = "While data is uploading: Report Category"
                   DoEvents
                
                   oNCsXls.Open "SELECT [CLIENT ID] FROM [ReportCategories$] GROUP BY [CLIENT ID];", oConnEx, adOpenDynamic, adLockReadOnly
                   szTemp = SQL2String(oNCsXls, 0)
                   oNCsXls.Close
                
                '  get the access database table
                   adoClient.Open "SELECT * FROM Client WHERE ClientID NOT IN (" & szTemp & ");", oConnAcc, adOpenDynamic, adLockOptimistic
                
                   If Not adoClient.EOF Then
                '  get the recordset in the excel file
                      oNCsXls.Open "SELECT * FROM [ReportCategories$] WHERE [CLIENT ID] = 'DEFAULT';", oConnEx, adOpenDynamic, adLockReadOnly
                
                      With adoRstMdb
                         While Not adoClient.EOF
                            .Open "SELECT * FROM ReportCategory", oConnAcc, adOpenDynamic, adLockOptimistic
                            .Find "ClientID = '" & adoClient.Fields.Item("ClientID").Value & "' ", , , 1
                            If .EOF Then                  'The client does not have category code
                               Do While Not oNCsXls.EOF
                                  .AddNew
                                  .Fields.Item("RecordID").Value = UniqueID()
                                  .Fields.Item("ClientID").Value = adoClient.Fields.Item("ClientID").Value
                                  .Fields.Item("CategoryCode").Value = oNCsXls.Fields.Item("REPORT CATEGORY CODE").Value
                                  .Fields.Item("CategoryName").Value = oNCsXls.Fields.Item("REPORT CATEGORY NAME").Value
                                  .Fields.Item("CatDesc").Value = oNCsXls.Fields.Item("REPORT CATEGORY NAME").Value
                                  .Update
                                  oNCsXls.MoveNext
                                  If oNCsXls.EOF Then Exit Do
                                  If oNCsXls.Fields.Item(0).Value = "END" Then Exit Do
                               Loop
                               oNCsXls.MoveFirst
                               .Close
                            Else
                               If .State = 1 Then .Close
                               .Open "SELECT * FROM ReportCategory WHERE ClientID = '" & adoClient.Fields.Item("ClientID").Value & "' ", oConnAcc, adOpenDynamic, adLockOptimistic
                               Do While Not oNCsXls.EOF
                                  .Find "CategoryCode = '" & oNCsXls.Fields.Item("REPORT CATEGORY CODE").Value & "' ", , , 1
                                  If .EOF Then
                                     .AddNew
                                     .Fields.Item("RecordID").Value = UniqueID()
                                     .Fields.Item("ClientID").Value = adoClient.Fields.Item("ClientID").Value
                                  End If
                                  .Fields.Item("CategoryCode").Value = oNCsXls.Fields.Item("REPORT CATEGORY CODE").Value
                                  .Fields.Item("CategoryName").Value = oNCsXls.Fields.Item("REPORT CATEGORY NAME").Value
                                  .Fields.Item("CatDesc").Value = oNCsXls.Fields.Item("REPORT CATEGORY NAME").Value
                                  .Update
                                  oNCsXls.MoveNext
                                  If oNCsXls.EOF Then Exit Do
                                  If oNCsXls.Fields.Item(0).Value = "END" Then Exit Do
                               Loop
                               oNCsXls.MoveFirst
                            .Close
                            End If
                            adoClient.MoveNext
                         Wend
                      End With
                      oNCsXls.Close
                   End If
                   adoClient.Close
                
                '  IN THE EXCEL IMPORT FILE: IF THERE ANY NOMINALCATEGORY CODES FOUND ALLOCATED TO CLIENTS
                   adoClient.Open "SELECT * FROM Client WHERE ClientID IN (" & szTemp & ");", oConnAcc, adOpenDynamic, adLockOptimistic
                   adoRstMdb.Open "SELECT * FROM ReportCategory", oConnAcc, adOpenDynamic, adLockOptimistic
                   
                   While Not adoClient.EOF
                      oConnAcc.Execute "DELETE * FROM ReportCategory WHERE ClientID = '" & adoClient.Fields.Item("ClientID").Value & "';"
                      
                      oNCsXls.Open "SELECT * FROM [ReportCategories$] WHERE [CLIENT ID] = '" & adoClient.Fields.Item("ClientID").Value & "';", oConnEx, adOpenDynamic, adLockReadOnly
                      Do While oNCsXls.EOF
                         With adoRstMdb
                            .AddNew
                            .Fields.Item("RecordID").Value = UniqueID()
                            .Fields.Item("ClientID").Value = adoClient.Fields.Item("ClientID").Value
                            .Fields.Item("CategoryCode").Value = oNCsXls.Fields.Item("REPORT CATEGORY CODE").Value
                            .Fields.Item("CategoryName").Value = oNCsXls.Fields.Item("REPORT CATEGORY NAME").Value
                            .Fields.Item("CatDesc").Value = oNCsXls.Fields.Item("REPORT CATEGORY NAME").Value
                            .Update
                         End With
                         oNCsXls.MoveNext
                         If oNCsXls.EOF Then Exit Do
                         If oNCsXls.Fields.Item(0).Value = "END" Then Exit Do
                      Loop
                      adoClient.MoveNext
                      oNCsXls.Close
                   Wend
'                   adoClient.Close
'                   adoRstMdb.Close
                   Print #1, "Data writing FINISH: ********** Report Category **********" & Chr$(13) + Chr$(10)
                   import_nominalCategory = True
End Function

