VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Waxy Management Leasing Program - Kingsgate"
   ClientHeight    =   7200
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   11205
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7200
   ScaleWidth      =   11205
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancelExport 
      Caption         =   "&Cancel Export"
      Height          =   1095
      Left            =   6000
      Picture         =   "Main.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1320
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.CommandButton cmdExportData 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Export The Data into &Sage"
      Height          =   1095
      Left            =   3600
      Picture         =   "Main.frx":0974
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1320
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.PictureBox CR1 
      Height          =   480
      Left            =   5520
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   12
      Top             =   5280
      Width           =   1200
   End
   Begin VB.CommandButton cmdSage 
      Caption         =   "Export to &Sage"
      Height          =   1095
      Left            =   8400
      Picture         =   "Main.frx":10A1
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   0
      Width           =   1335
   End
   Begin VB.CommandButton cmdExcel 
      Caption         =   "Export to &Excel"
      Height          =   1095
      Left            =   9720
      Picture         =   "Main.frx":17CE
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   0
      Width           =   1335
   End
   Begin VB.CommandButton cmdDemands 
      Caption         =   "&Demands"
      Height          =   1095
      Left            =   1200
      Picture         =   "Main.frx":1C84
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   0
      Width           =   1215
   End
   Begin VB.CommandButton cmdGlobal 
      Caption         =   "&Global Data"
      Height          =   1095
      Left            =   7200
      Picture         =   "Main.frx":20C6
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   0
      Width           =   1215
   End
   Begin VB.CommandButton cmdShopCentre 
      Caption         =   "Shopping &Centre"
      Height          =   1095
      Left            =   6000
      Picture         =   "Main.frx":2508
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   0
      Width           =   1215
   End
   Begin VB.CommandButton cmdtenants 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Tenants"
      Height          =   1095
      Left            =   2400
      Picture         =   "Main.frx":294A
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   0
      Width           =   1215
   End
   Begin VB.CommandButton cmdUnits 
      Caption         =   "&Units"
      Height          =   1095
      Left            =   3600
      Picture         =   "Main.frx":2D8C
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   0
      Width           =   1215
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   " E&xit Program"
      Height          =   1095
      Left            =   0
      Picture         =   "Main.frx":31CE
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   0
      Width           =   1215
   End
   Begin VB.CommandButton cmdLease 
      Caption         =   "&Lease"
      Height          =   1095
      Left            =   4800
      Picture         =   "Main.frx":3610
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   0
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "PCM Consulting Ltd"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   1455
      Left            =   4320
      TabIndex        =   11
      Top             =   3360
      Width           =   5895
   End
   Begin VB.Image Image1 
      Height          =   1440
      Left            =   2760
      Picture         =   "Main.frx":3A52
      Stretch         =   -1  'True
      Top             =   3120
      Width           =   1440
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuChangePassword 
         Caption         =   "Change Password"
      End
      Begin VB.Menu mnuEditUserNames 
         Caption         =   "Edit User Names"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuLogout 
         Caption         =   "Log out"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuShopCen 
      Caption         =   "Shopping Centre"
      Begin VB.Menu mnuShoppingCentreDetails 
         Caption         =   "View Details"
      End
      Begin VB.Menu mnuSep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuChangeShopCen 
         Caption         =   "Change Shopping Centre"
      End
      Begin VB.Menu mnuDelShopCentre 
         Caption         =   "Delete Shopping Centre"
      End
      Begin VB.Menu mnuSep99 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCompact 
         Caption         =   "Compact Database"
      End
   End
   Begin VB.Menu mnuRecords 
      Caption         =   "Records"
      Begin VB.Menu mnuUnits 
         Caption         =   "Units"
      End
      Begin VB.Menu mnuTenants 
         Caption         =   "Tenants"
      End
      Begin VB.Menu mnuLease 
         Caption         =   "Lease"
      End
      Begin VB.Menu mnuGlobal 
         Caption         =   "Global Data"
      End
      Begin VB.Menu mnuDemands 
         Caption         =   "Demands"
      End
   End
   Begin VB.Menu mnuReports 
      Caption         =   "Reports"
      Begin VB.Menu mnuBreakClause 
         Caption         =   "Break Clause"
      End
      Begin VB.Menu mnuLeaseEnd 
         Caption         =   "Lease End"
      End
      Begin VB.Menu mnuDateFlag 
         Caption         =   "Date Flag"
      End
      Begin VB.Menu mnuRentRenewal 
         Caption         =   "Rent Review"
      End
      Begin VB.Menu mnuRentIncrease 
         Caption         =   "Rent Increase"
      End
      Begin VB.Menu mnuSepp 
         Caption         =   "-"
      End
      Begin VB.Menu mnuStatus 
         Caption         =   "Demand Status"
      End
      Begin VB.Menu mnuSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuComExp 
         Caption         =   "Committed Expenses"
      End
   End
   Begin VB.Menu mnuExport 
      Caption         =   "Export"
      Begin VB.Menu mnuSage 
         Caption         =   "Sage"
         Begin VB.Menu mnuPreUpdate 
            Caption         =   "Pre-update"
         End
         Begin VB.Menu mnuSageExport 
            Caption         =   "Export"
         End
      End
      Begin VB.Menu mnuExcel 
         Caption         =   "Excel"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=========================================================================================
'  Code Page Name
'  Description and Useage
'=========================================================================================
'  Created By: Tania Sánchez
'  Published Date: 31/03/2005
'  WebSite: www.pcmconsulting.com
'  Legal Copyright: Tania Sánchez © 31/03/2005
'=========================================================================================

Option Explicit
Dim Conn1 As New RDO.rdoConnection
Dim Env1 As rdoEnvironment
Dim Envs1 As rdoEnvironments
Dim Rst1 As rdoResultset
Dim Conn2 As New RDO.rdoConnection
Dim Env2 As rdoEnvironment
Dim Envs2 As rdoEnvironments
Dim Rst3 As rdoResultset
Dim SQLStr1 As String
Dim rstDemands As rdoResultset
Dim strDemands As String

'1 = Access
'2 = Sage
Public SageComp As Integer
'=========================================================================================
Private Sub cmdCancelExport_Click()

    cmdCancelExport.Visible = False
    cmdExportData.Visible = False
    Call EnableMenus

End Sub 'cmdCancelExport_Click()
'=========================================================================================
Private Sub cmdDemands_Click()

    Unload Me
    Load frmDemands
    frmDemands.Show

End Sub 'cmdDemands_Click()
'=========================================================================================
Private Sub cmdExcel_Click()

    Call ExcelReport

End Sub 'cmdExcel_Click()
'=========================================================================================
Private Sub cmdExit_Click()

    Call ExitProgram

End Sub 'cmdExit_Click()
'=========================================================================================
Private Sub cmdExportData_Click()

    'Temporary update to prevent export from the test environment
    'not in yet

    Dim rstCheckAccounts1 As rdoResultset
    Dim NumberOfRecordsToExport As Integer
    Dim Count As Integer

    Dim concheck1 As rdoConnection
    Dim ConnCheck1 As rdoConnection
    Dim rstMark As rdoResultset
    
    Set ConnCheck1 = New rdoConnection
    
    Count = 0
    MousePointer = vbHourglass

    ' Dimension Sage Objects
    Dim oApplication As New SageAccountinglib.Application
    Dim oCustomer As SageAccountinglib.Customer
    Dim oCustomers As SageAccountinglib.Customers
    Dim oSalesInvoiceInstrument As SageAccountinglib.SalesInvoiceInstrument

    Conn1.Connect = "DSN=" & Adsn & ";UID=;PWD="
    Conn1.CursorDriver = rdUseIfNeeded
    Conn1.EstablishConnection rdDriverNoPrompt

    SQLStr1 = "SELECT Code FROM ShoppingCentre"
    Set Rst1 = Conn1.OpenResultset(SQLStr1, rdOpenStatic, rdConcurReadOnly)

    If Rst1.EOF = False Then SageComp = Rst1!Code

    Rst1.Close
    
    ' Establish connection and set active company
    oApplication.Connect "PCM", "PCM", "SageLine132"
    oApplication.ActiveCompany = oApplication.Companies(SageComp)

'Check Account and nominal code

    'From Demands
    SQLStr1 = "SELECT DISTINCT SageAccountNumber FROM DemandRecords WHERE ExportedToSage = 'P' ORDER BY SageAccountNumber"
    Set rstCheckAccounts1 = Conn1.OpenResultset(SQLStr1, rdOpenStatic, rdConcurReadOnly)

    If rstCheckAccounts1.EOF = False Then
        While rstCheckAccounts1.EOF = False
            Set oCustomers = oApplication.SalesModule.Customers
            On Error Resume Next
            Set oCustomer = oCustomers(UCase(rstCheckAccounts1!SageAccountNumber))
            If oCustomer Is Nothing Then
                'account not in sage
                Count = Count + 1
                'mark demands not to be exported.
                SQLStr1 = "SELECT ExportedToSage FROM DemandRecords WHERE SageAccountNumber = '" & rstCheckAccounts1!SageAccountNumber & "'"
                Set rstMark = Conn1.OpenResultset(SQLStr1, rdOpenDynamic, rdConcurRowVer)
                While rstMark.EOF = False
                    rstMark.Edit
                    'A means account not in Sage
                    rstMark!ExportedToSage = "A"
                    rstMark.Update
                    rstMark.MoveNext
                Wend
                rstMark.Close
            End If
            rstCheckAccounts1.MoveNext
        Wend
    End If
    rstCheckAccounts1.Close
    
    SQLStr1 = "SELECT DISTINCT NominalCodeforAmount FROM DemandRecords WHERE ExportedToSage = 'P'"
    Set rstCheckAccounts1 = Conn1.OpenResultset(SQLStr1, rdOpenStatic, rdConcurReadOnly)
    
    If rstCheckAccounts1.EOF = False Then
        While rstCheckAccounts1.EOF = False
            ' Extract first NominalCode from collection
            On Error Resume Next
            Dim oNominalCodes As SageAccountinglib.NominalCodes
            Set oNominalCodes = oApplication.NominalModule.NominalCodes
            oNominalCodes.Filter = "ACCOUNT-NUMBER='" & rstCheckAccounts1!NominalCodeforAmount & "'"
            oNominalCodes.Requery
            If oNominalCodes.Count = 0 Then
                'Nominal code not in sage
                Count = Count + 1
                'mark demands not to be exported.
                SQLStr1 = "SELECT ExportedToSage FROM DemandRecords WHERE NominalCodeforAmount = '" & rstCheckAccounts1!NominalCodeforAmount & "'"
                Set rstMark = Conn1.OpenResultset(SQLStr1, rdOpenDynamic, rdConcurRowVer)
                While rstMark.EOF = False
                    rstMark.Edit
                    'A means Nominal code not in Sage
                    rstMark!ExportedToSage = "C"
                    rstMark.Update
                    rstMark.MoveNext
                Wend
                rstMark.Close
            End If
            rstCheckAccounts1.MoveNext
        Wend
    End If
    rstCheckAccounts1.Close
    
    'Select all records to be updated
    strDemands = "SELECT * FROM DemandRecords WHERE ExportedToSage = 'P'"
    Set rstDemands = Conn1.OpenResultset(strDemands, rdOpenDynamic, rdConcurRowVer)

    If rstDemands.EOF = True And rstDemands.BOF = True Then
        MsgBox "There are no demands to be exported into Sage", vbOKOnly + vbInformation, "No Demands"
'        rstDemands.Close
'        Conn1.Close
'        MousePointer = vbDefault
'        Exit Sub
    Else
'
        rstDemands.MoveFirst
    
        While rstDemands.EOF = False
    
            'new sales transaction record
            ' Trap exceptions thrown by Sage Objects
            On Error GoTo ErrH
    
            ' Extract first Customer from collection
            Set oCustomer = oApplication.SalesModule.Customers(UCase(rstDemands!SageAccountNumber))
    
            With oCustomer
    
                ' Check transaction count
                Set oSalesInvoiceInstrument = oApplication.SalesModule.CreateInstrument(CreateSalesInvoice)
    
                If Not oSalesInvoiceInstrument Is Nothing Then
                    With oSalesInvoiceInstrument
                        .Customer = oCustomer
                        .InstrumentNo = rstDemands!Reference
                        .SecondReferenceNo = rstDemands!Reference
                        .EntryType = rstDemands!TransactionType
                        .InstrumentDate = rstDemands!IssueDate 'Mid(rstDemands!IssueDate, 4, 3) & Left(rstDemands!IssueDate, 2) & Right(rstDemands!IssueDate, 3)
                        .DueDate = rstDemands!IssueDate 'Mid(rstDemands!IssueDate, 4, 3) & Left(rstDemands!IssueDate, 2) & Right(rstDemands!IssueDate, 3)
                        .LedgerPeriod = oApplication.SalesModule.Ledger.TradingAccountPeriods(CInt(rstDemands!VATMonth))
                        .NetValue = CDbl(rstDemands!Amount)
                        .AllocatedValue = 0
                        .TaxValue = CDbl(rstDemands!VATAmount)
                        ' Rst3("SALES_CONTROL_VALUE") = CDbl(rstDemands!TotalAmount)
                        ' Rst3("ORIGINAL_EX_RATE") = 1
                        ' Rst3("VAT_MONTH") = CInt(rstDemands!VATMonth)
                        ' Rst3("USER_NUMBER") = 2
                        ' Rst3("UNIQUE_REFERENCE_NO") = NextURN
                        '.Parent = rstDemands!Source
                        .PostedDate = CDate(TodaysDate) 'Mid(TodaysDate, 4, 3) & Left(TodaysDate, 2) & Right(TodaysDate, 3)
                         ' Nominal Analysis Item 1
                        With .NominalAnalysisItems(0)
                            .Amount = CDbl(rstDemands!Amount)
                            .Narrative = CStr(rstDemands!Text)
                            With .NominalSpecification
                                .Reference = CStr(rstDemands!NominalCodeforAmount)
                            End With
                        End With
                        ' Update SalesInvoiceInstrument
                        .Update
                    End With
                End If
    
            End With
    
            Set oSalesInvoiceInstrument = Nothing
    
            rstDemands.Edit
            rstDemands!ExportedToSage = "Y"
            rstDemands.Update
            rstDemands.MoveNext
            'NextUrn = NextUrn + 1
            NumberOfRecordsToExport = NumberOfRecordsToExport + 1
        Wend

        'Rst2.Close
        'Rst3.Close
        'Rst4.Close
        'Rst4a.Close
        'Rst5.Close
        'Rst6.Close
        'Rst7.Close
        'Rst8.Close
    '
    '    'DisconnectSage
    '    Conn2.Close
    '    Env2.Close
    
        MousePointer = vbDefault
    
        MsgBox NumberOfRecordsToExport & " Demands have been exported into Sage.", vbOKOnly + vbInformation, "Export"
    End If
    
    MousePointer = vbDefault
    'Close Resultsets
    rstDemands.Close
    
    If Count > 0 Then
        CR1.ReportFileName = App.Path & "\Missing" & SCID & ".rpt"
        CR1.PrintReport
        
        SQLStr1 = "SELECT ExportedToSage FROM DemandRecords WHERE ExportedToSage <> 'N' AND ExportedToSage <> 'Y'"
        Set Rst1 = Conn1.OpenResultset(SQLStr1, rdOpenDynamic, rdConcurRowVer)
        If Rst1.EOF = False Then
            While Rst1.EOF = False
                Rst1.Edit
                Rst1!ExportedToSage = "N"
                Rst1.Update
                Rst1.MoveNext
            Wend
        End If
        Rst1.Close
    End If
    Conn1.Close

    cmdCancelExport.Visible = False
    cmdExportData.Visible = False
    Call EnableMenus

    Exit Sub

ErrH:
    If Err.Number <> 0 Then
        MsgBox Err.Number & " - " & Err.Description
        Resume Next
    End If

End Sub 'cmdExportData_Click()
'=========================================================================================
Private Sub cmdGlobal_Click()

    Unload Me
    Load frmGlobal
    frmGlobal.Show

End Sub 'cmdGlobal_Click()
'=========================================================================================
Private Sub cmdLease_Click()

    Unload Me
    Load frmLease
    frmLease.Show

End Sub 'cmdLease_Click()
'=========================================================================================
Private Sub cmdSage_Click()

    Call SagePreUpdate

End Sub 'cmdSage_Click()
'=========================================================================================
Private Sub cmdShopCentre_Click()

    Unload Me
    Load frmShoppingCentre
    frmShoppingCentre.Show

End Sub 'cmdShopCentre_Click()
'=========================================================================================
Private Sub cmdtenants_Click()

    Unload Me
    Load frmTenant
    frmTenant.Show

End Sub 'cmdtenants_Click()
'=========================================================================================
Private Sub cmdUnits_Click()

    Unload Me
    Load frmUnit
    frmUnit.Show

End Sub 'cmdUnits_Click()
'=========================================================================================
Private Sub Form_Load()

    Me.Move (Screen.Width - Width) / 2, 0
    Me.Caption = "Waxy Management - " & gCurrentShopCentreName

    Conn1.Connect = "DSN=" & Adsn & ";UID=;PWD="
    Conn1.CursorDriver = rdUseIfNeeded
    Conn1.EstablishConnection rdDriverNoPrompt

    Set Rst1 = Conn1.OpenResultset("SELECT SageDSN FROM ShoppingCentre", rdOpenStatic, rdConcurReadOnly)
    If Rst1.EOF = True And Rst1.BOF = True Then
        MsgBox "You need to Select a Sage Datasource Name", vbInformation + vbOKOnly, "Sage DSN"
    Else
        Rst1.MoveFirst
        Sdsn = Rst1!SageDSN
    End If
    Rst1.Close
    Conn1.Close

    If LCase(User) = "manager" Then mnuEditUserNames.Enabled = True
    If LCase(User) <> "manager" Then mnuEditUserNames.Enabled = False

    Call GetGlobalData

    'MsgBox SCID

End Sub 'Form_Load()
'=========================================================================================
Private Sub mnuBreakClause_Click()

    Dim a As String
    Dim b As String
    Dim c As Integer
    c = 0

    MousePointer = vbHourglass

    CR1.ReportFileName = App.Path & "\Break" & SCID & ".rpt"
    CR1.ReportTitle = "Six Month Break Clause Notification"

    Conn1.Connect = "DSN=" & Adsn & ";UID=;PWD="
    Conn1.CursorDriver = rdUseIfNeeded
    Conn1.EstablishConnection rdDriverNoPrompt

    SQLStr1 = "SELECT SageAccountNumber, BreakDate FROM LeaseDetails WHERE BreakClause = 'Y'"
    Set Rst1 = Conn1.OpenResultset(SQLStr1, rdOpenStatic, rdConcurReadOnly)

    b = DateAdd("m", 6, Date)
    If Rst1.EOF = False Then
        While Rst1.EOF = False
            If Rst1!BreakDate <> "" Then
                '23:09 20/11/2002 Here is where the date selction is made, will extend 3 months back (change 6 to 18 gives years previous as well next six months). Original line below
                'If DateDiff("m", Rst1!BreakDate, b) <= 6 And DateDiff("m", Rst1!BreakDate, b) >= 0 Then
                If DateDiff("m", Rst1!BreakDate, b) <= 18 And DateDiff("m", Rst1!BreakDate, b) >= 0 Then
                    a = a & "{LeaseDetails.SageAccountNumber} = '" & Rst1!SageAccountNumber & "'" & " or" & Chr(13)
                    c = c + 1
                End If
            End If
            Rst1.MoveNext
        Wend
    End If
    Rst1.Close
    Conn1.Close

    If c = 0 Then
        MsgBox "There are no Break Clauses set within the next six months.", vbOKOnly + vbInformation, "Break Clause Report"
        MousePointer = vbDefault
        Exit Sub
    End If

    CR1.SelectionFormula = Left(a, Len(a) - 3)
    CR1.PrintReport

    MousePointer = vbDefault

End Sub 'mnuBreakClause_Click()
'=========================================================================================
Private Sub mnuChangePassword_Click()

    Load frmChangePassword
    Unload Me
    frmChangePassword.Show

End Sub 'mnuChangePassword_Click()
'=========================================================================================
Private Sub mnuChangeShopCen_Click()

    Load frmlogin
    Unload Me
    frmlogin.Show

End Sub 'mnuChangeShopCen_Click()
'=========================================================================================
Private Sub mnuComExp_Click()

    MousePointer = vbHourglass

    CR1.ReportFileName = App.Path & "\ComExp" & SCID & ".rpt"
    CR1.Connect = "UID=PCM;PWD="
    CR1.PrintReport

    MousePointer = vbDefault

End Sub 'mnuComExp_Click()
'=========================================================================================
Private Sub mnuCompact_Click()

    Call dbcompact

End Sub 'mnuCompact_Click()
'=========================================================================================
Private Sub mnuDateFlag_Click()

    Dim a As String
    Dim b As String
    Dim c As Integer
    c = 0
    MousePointer = vbHourglass

    CR1.ReportFileName = App.Path & "\DateFlag" & SCID & ".rpt"
    CR1.ReportTitle = "Six Month Date Flag Notification"

    Conn1.Connect = "DSN=" & Adsn & ";UID=;PWD="
    Conn1.CursorDriver = rdUseIfNeeded
    Conn1.EstablishConnection rdDriverNoPrompt

    SQLStr1 = "SELECT SageAccountNumber, DateFlagDate FROM LeaseDetails WHERE DateFlagDate <> ''"
    Set Rst1 = Conn1.OpenResultset(SQLStr1, rdOpenStatic, rdConcurReadOnly)

    b = DateAdd("m", 6, Date)
    If Rst1.EOF = False Then
        While Rst1.EOF = False
            If DateDiff("m", Rst1!DateFlagDate, b) <= 6 And DateDiff("m", Rst1!DateFlagDate, b) >= 0 Then
                a = a & "{LeaseDetails.SageAccountNumber} = '" & Rst1!SageAccountNumber & "' or" & Chr(13)
                c = c + 1
            End If
            Rst1.MoveNext
        Wend
    End If
    Rst1.Close
    Conn1.Close

    If c = 0 Then
        MsgBox "There are no Date Flags within the next six months.", vbOKOnly + vbInformation, "Date Flag Reprot"
        MousePointer = vbDefault
        Exit Sub
    End If

    CR1.SelectionFormula = Left(a, Len(a) - 3)
    CR1.PrintReport

    MousePointer = vbDefault

End Sub 'mnuDateFlag_Click()
'=========================================================================================
Private Sub mnuDelShopCentre_Click()

    If MsgBox("Do you wish to Delete Shopping Centre - " & gCurrentShopCentreName & "?", vbYesNo, "Delete Shopping Centre") = vbNo Then Exit Sub
    If MsgBox("Are you really sure you wish to delete Shopping Centre " & gCurrentShopCentreName & "?", vbYesNo, "Delete Shopping Centre") = vbNo Then Exit Sub

    'Delete Shopping Centre from WDcontrol
    Conn1.Connect = "DSN=WDControl;UID=;PWD="
    Conn1.CursorDriver = rdUseIfNeeded
    Conn1.EstablishConnection rdDriverNoPrompt

    SQLStr1 = "SELECT * FROM Databases WHERE ID = " & gCurrentShopCentreCode
    Set Rst1 = Conn1.OpenResultset(SQLStr1, rdOpenDynamic, rdConcurRowVer)

    Rst1.Delete
    'Rst1.Update
    Rst1.Close
    Conn1.Close

End Sub 'mnuDelShopCentre_Click()
'=========================================================================================
Private Sub mnuDemands_Click()

    Unload Me
    Load frmDemands
    frmDemands.Show

End Sub 'mnuDemands_Click()
'=========================================================================================
Private Sub mnuEditUserNames_Click()

    Load frmUsers
    Unload Me
    frmUsers.Show

End Sub 'mnuEditUserNames_Click()
'=========================================================================================
Private Sub mnuExcel_Click()

    Call ExcelReport

End Sub 'mnuExcel_Click()
'=========================================================================================
Private Sub mnuExit_Click()

    Call ExitProgram

End Sub 'mnuExit_Click()
'=========================================================================================
Private Sub mnuGlobal_Click()

    Unload Me
    Load frmGlobal
    frmGlobal.Show

End Sub 'mnuGlobal_Click()
'=========================================================================================
Private Sub mnuLease_Click()

    Unload Me
    Load frmLease
    frmLease.Show

    '=========================================================================================
End Sub
Private Sub mnuLeaseEnd_Click() 'mnuLease_Click()
    '=========================================================================================
    Dim b As String
    Dim a As String

    Dim c As Integer
    c = 0

    Dim TempString1 As String
    Dim TempString2 As String

    MousePointer = vbHourglass

    CR1.ReportFileName = App.Path & "\LeaseEnd" & SCID & ".rpt"
    CR1.ReportTitle = "Six Month Lease End Notification"

    Conn1.Connect = "DSN=" & Adsn & ";UID=;PWD="
    Conn1.CursorDriver = rdUseIfNeeded
    Conn1.EstablishConnection rdDriverNoPrompt

    'Inserted CompanyName in to SQL String to allow bring back od incorrect date error message
    SQLStr1 = "SELECT SageAccountNumber, CompanyName, EndDate FROM LeaseDetails"
    Set Rst1 = Conn1.OpenResultset(SQLStr1, rdOpenStatic, rdConcurReadOnly)

    b = DateAdd("m", 6, Date)
    If Rst1.EOF = False Then
        While Rst1.EOF = False
            'Smae princple as for lease break report. Let 6 = 18
            'If DateDiff("m", Rst1!enddate, b) <= 6 And DateDiff("m", Rst1!enddate, b) >= 0 Then

            'verify is date
            If IsDate(Rst1!enddate) Then

                If DateDiff("m", Rst1!enddate, b) <= 18 And DateDiff("m", Rst1!enddate, b) >= 0 Then
                    a = a & "{LeaseDetails.SageAccountNumber} = '" & Rst1!SageAccountNumber & "' or" & Chr(13)
                    c = c + 1
                End If
                Rst1.MoveNext

            Else 'if is not date
                TempString1 = Rst1!SageAccountNumber
                TempString2 = Rst1!CompanyName
                MsgBox ("The Lease End date for " & TempString1 & " - " & TempString2 & " is not a valid date, please correct before running the Lease End Report.")
                Rst1.Close
                Conn1.Close
                MousePointer = vbDefault
                Exit Sub

            End If

        Wend
    End If
    Rst1.Close
    Conn1.Close

    If c = 0 Then
        MsgBox "There are no Lease End dates set within the next six months.", vbOKOnly + vbInformation, "Lease End Report"
        MousePointer = vbDefault
        Exit Sub
    End If

    CR1.SelectionFormula = Left(a, Len(a) - 3)
    CR1.PrintReport

    MousePointer = vbDefault

End Sub 'mnuLeaseEnd_Click() 'mnuLease_Click()
'=========================================================================================
Private Sub mnuLogout_Click()

    If MsgBox("Are you sure you wish to log out?", vbYesNo + vbQuestion, "Log Out") = vbNo Then
        Exit Sub
    Else
        Load frmlogin
        Unload Me
        frmlogin.Show
    End If

End Sub 'mnuLogout_Click()
'=========================================================================================
Private Sub mnuPreUpdate_Click()

    Call SagePreUpdate

End Sub 'mnuPreUpdate_Click()
'=========================================================================================
Private Sub mnuRentIncrease_Click()

    Dim a As String
    Dim b As String
    Dim c As Integer
    c = 0

    MousePointer = vbHourglass

    CR1.ReportFileName = App.Path & "\RentIncrease" & SCID & ".rpt"
    CR1.ReportTitle = "Six Month Rent Increase Notification"

    Conn1.Connect = "DSN=" & Adsn & ";UID=;PWD="
    Conn1.CursorDriver = rdUseIfNeeded
    Conn1.EstablishConnection rdDriverNoPrompt

    SQLStr1 = "SELECT SageAccountNumber, RentIncreaseDate FROM LeaseDetails WHERE RentIncreaseDate <> ''"
    Set Rst1 = Conn1.OpenResultset(SQLStr1, rdOpenStatic, rdConcurReadOnly)

    b = DateAdd("m", 6, Date)
    If Rst1.EOF = False Then
        While Rst1.EOF = False
            If DateDiff("m", Rst1!RentIncreaseDate, b) <= 6 And DateDiff("m", Rst1!RentIncreaseDate, b) >= 0 Then
                a = a & "{LeaseDetails.SageAccountNumber} = '" & Rst1!SageAccountNumber & "' or" & Chr(13)
                c = c + 1
            End If
            Rst1.MoveNext
        Wend
    End If
    Rst1.Close
    Conn1.Close

    If c = 0 Then
        MsgBox "There are no Rent Increase dates set within the next six months.", vbOKOnly + vbInformation, "Rent Increase Report"
        MousePointer = vbDefault
        Exit Sub
    End If

    CR1.SelectionFormula = Left(a, Len(a) - 3)
    CR1.PrintReport

    MousePointer = vbDefault

End Sub 'mnuRentIncrease_Click()
'=========================================================================================
Private Sub mnuRentRenewal_Click()
    Dim a As String
    Dim b As String
    Dim c As Integer
    c = 0

    MousePointer = vbHourglass

    CR1.ReportFileName = App.Path & "\RentReview" & SCID & ".rpt"
    CR1.ReportTitle = "Six Month Rent Review Notification"

    Conn1.Connect = "DSN=" & Adsn & ";UID=;PWD="
    Conn1.CursorDriver = rdUseIfNeeded
    Conn1.EstablishConnection rdDriverNoPrompt

    SQLStr1 = "SELECT SageAccountNumber, RentReviewDate FROM LeaseDetails WHERE RentReviewDate <> ''"
    Set Rst1 = Conn1.OpenResultset(SQLStr1, rdOpenStatic, rdConcurReadOnly)

    b = DateAdd("m", 6, Date)
    If Rst1.EOF = False Then
        While Rst1.EOF = False
            If DateDiff("m", Rst1!RentReviewDate, b) <= 6 And DateDiff("m", Rst1!RentReviewDate, b) >= 0 Then
                a = a & "{LeaseDetails.SageAccountNumber} = '" & Rst1!SageAccountNumber & "' or" & Chr(13)
                c = c + 1
            End If
            Rst1.MoveNext
        Wend
    End If
    Rst1.Close
    Conn1.Close

    If c = 0 Then
        MsgBox "There are no Rent Review dates set within the next six months.", vbOKOnly + vbInformation, "Rent Review Report"
        MousePointer = vbDefault
        Exit Sub
    End If

    CR1.SelectionFormula = Left(a, Len(a) - 3)
    CR1.PrintReport

    MousePointer = vbDefault

End Sub 'mnuRentRenewal_Click()

'=========================================================================================
Private Sub mnuShoppingCentreDetails_Click()

    Unload Me
    Load frmShoppingCentre
    frmShoppingCentre.Show

End Sub 'mnuShoppingCentreDetails_Click()
'=========================================================================================
Private Sub mnuStatus_Click()

    Dim cnnDB As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim SQLStr As String
    Set cnnDB = New ADODB.Connection
    Set rs = New ADODB.Recordset

    cnnDB.Open "DSN=" & Adsn & ";UID=;PWD="
    SQLStr = "SELECT ExportedToSage FROM DemandRecords WHERE IsPrinted = 'Y' AND ExportedToSage = 'N'OR ExportedToSage = 'P'"
    rs.Open SQLStr, cnnDB, adOpenDynamic, adLockPessimistic

    If rs.EOF = False Then
        While rs.EOF = False
            rs!ExportedToSage = "P"
            rs.Update
            rs.MoveNext
        Wend
    End If

    rs.Close
    cnnDB.Close

    Set rs = Nothing
    Set cnnDB = Nothing

    CR1.ReportTitle = "Demand Status Report"
    CR1.ReportFileName = App.Path & "\PreUpdate" & SCID & ".rpt"
    CR1.PrintReport

    Set cnnDB = New ADODB.Connection
    Set rs = New ADODB.Recordset

    cnnDB.Open "DSN=" & Adsn & ";UID=;PWD="
    SQLStr = "SELECT ExportedToSage FROM DemandRecords WHERE ExportedToSage = 'P'"
    rs.Open SQLStr, cnnDB, adOpenDynamic, adLockPessimistic

    If rs.EOF = False Then
        While rs.EOF = False
            rs!ExportedToSage = "N"
            rs.Update
            rs.MoveNext
        Wend
    End If

    rs.Close
    cnnDB.Close

    Set rs = Nothing
    Set cnnDB = Nothing

End Sub 'mnuStatus_Click()
'=========================================================================================
Private Sub mnuTenants_Click()

    Unload Me
    Load frmTenant
    frmTenant.Show

End Sub 'mnuTenants_Click()
'=========================================================================================
Private Sub mnuUnits_Click()

    Unload Me
    Load frmUnit
    frmUnit.Show

End Sub 'mnuUnits_Click()
'=========================================================================================
Public Sub DisableMenus()

    mnuChangePassword.Enabled = False
    mnuEditUserNames.Enabled = False
    mnuExit.Enabled = False
    mnuExcel.Enabled = False
    mnuShoppingCentreDetails.Enabled = False
    mnuChangeShopCen.Enabled = False
    mnuDelShopCentre.Enabled = False
    mnuUnits.Enabled = False
    mnuTenants.Enabled = False
    mnuLease.Enabled = False
    mnuGlobal.Enabled = False
    mnuDemands.Enabled = False
    mnuBreakClause.Enabled = False
    mnuDateFlag.Enabled = False
    mnuLeaseEnd.Enabled = False
    mnuRentRenewal.Enabled = False
    mnuRentIncrease.Enabled = False
    mnuPreUpdate.Enabled = False
    mnuComExp.Enabled = False
    cmdDemands.Enabled = False
    cmdExcel.Enabled = False
    cmdtenants.Enabled = False
    cmdUnits.Enabled = False
    cmdGlobal.Enabled = False
    cmdShopCentre.Enabled = False
    cmdSage.Enabled = False
    cmdLease.Enabled = False

End Sub 'DisableMenus()
'=========================================================================================
Public Sub EnableMenus()

    mnuChangePassword.Enabled = True
    mnuEditUserNames.Enabled = True
    mnuExcel.Enabled = True
    mnuExit.Enabled = True
    mnuShoppingCentreDetails.Enabled = True
    mnuChangeShopCen.Enabled = True
    mnuDelShopCentre.Enabled = True
    mnuUnits.Enabled = True
    mnuTenants.Enabled = True
    mnuLease.Enabled = True
    mnuGlobal.Enabled = True
    mnuDemands.Enabled = True
    mnuBreakClause.Enabled = True
    mnuDateFlag.Enabled = True
    mnuLeaseEnd.Enabled = True
    mnuRentRenewal.Enabled = True
    mnuPreUpdate.Enabled = True
    mnuRentIncrease.Enabled = True
    mnuComExp.Enabled = True
    cmdDemands.Enabled = True
    cmdExcel.Enabled = True
    cmdtenants.Enabled = True
    cmdUnits.Enabled = True
    cmdGlobal.Enabled = True
    cmdShopCentre.Enabled = True
    cmdSage.Enabled = True
    cmdLease.Enabled = True

End Sub 'EnableMenus()
'=========================================================================================
Public Sub SagePreUpdate()

    Dim cnnDB As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim SQLStr As String

    Call DisableMenus

    If MsgBox("Run Pre-Update into Sage Line 100 Sales Ledger Report?", vbOKCancel + vbQuestion, "Pre-Update Report") = vbCancel Then
        Call EnableMenus
        Exit Sub
    End If

    Set cnnDB = New ADODB.Connection
    Set rs = New ADODB.Recordset

    cnnDB.Open "DSN=" & Adsn & ";UID=;PWD="
    SQLStr = "SELECT ExportedToSage FROM DemandRecords WHERE IsPrinted = 'Y' AND ExportedToSage = 'N'OR ExportedToSage = 'P'"
    rs.Open SQLStr, cnnDB, adOpenDynamic, adLockPessimistic

    If rs.EOF = True And rs.BOF = True Then
        MsgBox "There are no demands to be exported into Sage.", vbOKOnly, "No Demands"
        rs.Close
        cnnDB.Close
        Call EnableMenus
        Exit Sub
    End If

    While rs.EOF = False
        rs!ExportedToSage = "P"
        rs.Update
        rs.MoveNext
    Wend

    rs.Close
    cnnDB.Close

    Set rs = Nothing
    Set cnnDB = Nothing

    CR1.ReportTitle = "Pre-Update into Sage Report"
    CR1.ReportFileName = App.Path & "\PreUpdate" & SCID & ".rpt"
    CR1.PrintReport

    cmdCancelExport.Visible = True
    cmdExportData.Visible = True

End Sub 'SagePreUpdate()
'=========================================================================================
Public Sub ExcelReport()

    MousePointer = vbHourglass
    Dim cnnDB As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim SQLStr As String

    Set cnnDB = New ADODB.Connection
    Set rs = New ADODB.Recordset

    cnnDB.Open "DSN=" & Adsn & ";UID=;PWD="
    SQLStr = "SELECT ExportedToExcel FROM DemandRecords WHERE ExportedToSage = 'Y' AND ExportedToExcel = 'N'"
    rs.Open SQLStr, cnnDB, adOpenDynamic, adLockPessimistic

    If rs.EOF = True And rs.BOF = True Then
        MsgBox "There are no demands to be exported into Excel.", vbOKOnly, "No Demands"
        rs.Close
        cnnDB.Close
        Exit Sub
    End If

    While rs.EOF = False
        rs!ExportedToExcel = "C"
        rs.Update
        rs.MoveNext
    Wend

    rs.Close

    CR1.ReportFileName = App.Path & "\ExcelDemands" & SCID & ".rpt"
    CR1.PrintReport

    SQLStr = "SELECT ExportedToExcel FROM DemandRecords WHERE ExportedToExcel = 'C'"
    rs.Open SQLStr, cnnDB, adOpenDynamic, adLockPessimistic

    If rs.EOF = True And rs.BOF = True Then
        rs.Close
        cnnDB.Close
        Exit Sub
    End If

    While rs.EOF = False
        rs!ExportedToExcel = "Y"
        rs.Update
        rs.MoveNext
    Wend

    rs.Close
    cnnDB.Close

    Set rs = Nothing
    Set cnnDB = Nothing

    MousePointer = vbDefault

End Sub 'ExcelReport()
'=========================================================================================
Private Sub Picture7_Click()

End Sub 'Picture7_Click()
'=========================================================================================
Private Sub Picture6_Click()

End Sub 'Picture6_Click()
'=========================================================================================
