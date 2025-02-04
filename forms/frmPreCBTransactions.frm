VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmPreCBTransactions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cashbook Transactions"
   ClientHeight    =   6045
   ClientLeft      =   1125
   ClientTop       =   1935
   ClientWidth     =   12225
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Myriad Web"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPreCBTransactions.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4172.367
   ScaleMode       =   0  'User
   ScaleWidth      =   11479.91
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picClient 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3825
      Left            =   8145
      ScaleHeight     =   3795
      ScaleWidth      =   5850
      TabIndex        =   29
      Top             =   90
      Visible         =   0   'False
      Width           =   5880
      Begin VB.CommandButton cmdPicCLose 
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
         Left            =   5550
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   0
         Width           =   255
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxClient 
         Height          =   3075
         Left            =   45
         TabIndex        =   10
         Top             =   675
         Width           =   5760
         _ExtentX        =   10160
         _ExtentY        =   5424
         _Version        =   393216
         FixedRows       =   0
         FixedCols       =   0
         BackColorFixed  =   13553358
         ForeColorFixed  =   -2147483634
         BackColorSel    =   12648447
         ForeColorSel    =   -2147483630
         BackColorBkg    =   16777215
         GridColor       =   14737632
         WordWrap        =   -1  'True
         GridLinesFixed  =   1
         ScrollBars      =   2
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
      End
      Begin VB.Label lblFlxPayee 
         Caption         =   "EMPTY"
         Height          =   255
         Index           =   4
         Left            =   2115
         TabIndex        =   34
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label lblPayeeFlxConfigured 
         Caption         =   "NOT"
         Height          =   495
         Index           =   4
         Left            =   1515
         TabIndex        =   33
         Top             =   1800
         Width           =   1095
      End
      Begin MSForms.Label lblClientID 
         Height          =   195
         Left            =   120
         TabIndex        =   32
         Top             =   120
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
         Left            =   1620
         TabIndex        =   31
         Top             =   135
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
         TabIndex        =   8
         Top             =   375
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
         TabIndex        =   9
         Top             =   375
         Width           =   4140
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         Size            =   "7302;450"
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
         Left            =   45
         Top             =   75
         Width           =   5490
      End
   End
   Begin VB.CommandButton cmdBC 
      Caption         =   ".."
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   4950
      TabIndex        =   1
      Top             =   585
      Width           =   300
   End
   Begin VB.CommandButton cmdClientList 
      Caption         =   ".."
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   4950
      TabIndex        =   0
      Top             =   90
      Width           =   300
   End
   Begin VB.Frame Frame3 
      Caption         =   "Date options"
      Height          =   465
      Left            =   135
      TabIndex        =   26
      Top             =   1530
      Width           =   5190
      Begin VB.OptionButton Option2 
         Caption         =   "By Financial Year"
         Height          =   195
         Left            =   1530
         TabIndex        =   3
         Top             =   180
         Value           =   -1  'True
         Width           =   1545
      End
      Begin VB.OptionButton Option1 
         Caption         =   "By Date"
         Height          =   195
         Left            =   3195
         TabIndex        =   4
         Top             =   180
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Select the Financial Period"
      Height          =   1410
      Left            =   4770
      TabIndex        =   22
      Top             =   2115
      Width           =   5175
      Begin MSForms.ComboBox cmbPeriodFrom 
         Height          =   285
         Left            =   1800
         TabIndex        =   6
         Top             =   690
         Width           =   1920
         VariousPropertyBits=   1753237531
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "3387;503"
         TextColumn      =   2
         ColumnCount     =   4
         ListRows        =   20
         cColumnInfo     =   4
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         Object.Width           =   "0;1940;0;0"
      End
      Begin MSForms.ComboBox cmbPeriodTo 
         Height          =   285
         Left            =   1800
         TabIndex        =   7
         Top             =   1050
         Width           =   1920
         VariousPropertyBits=   1753237531
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "3387;503"
         TextColumn      =   2
         ColumnCount     =   4
         ListRows        =   20
         cColumnInfo     =   4
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         Object.Width           =   "0;1940;0;0"
      End
      Begin MSForms.ComboBox cmbFinancialYear 
         Height          =   285
         Left            =   1800
         TabIndex        =   5
         Top             =   330
         Width           =   2760
         VariousPropertyBits=   1753237531
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "4868;503"
         TextColumn      =   2
         ColumnCount     =   5
         ListRows        =   20
         cColumnInfo     =   5
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         Object.Width           =   "0;1940;0;0;0"
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000016&
         BackStyle       =   0  'Transparent
         Caption         =   "Financial Year:"
         Height          =   255
         Index           =   66
         Left            =   600
         TabIndex        =   25
         Top             =   345
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Period From:"
         Height          =   255
         Index           =   4
         Left            =   600
         TabIndex        =   24
         Top             =   705
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Period To:"
         Height          =   255
         Index           =   3
         Left            =   600
         TabIndex        =   23
         Top             =   1065
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Date Range"
      Height          =   1215
      Left            =   120
      TabIndex        =   12
      Top             =   2145
      Visible         =   0   'False
      Width           =   5175
      Begin VB.TextBox txtSCYRREnDt 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2400
         TabIndex        =   14
         Top             =   720
         Width           =   1815
      End
      Begin VB.TextBox txtSCYRRStDt 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2400
         TabIndex        =   13
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "End Date"
         Height          =   255
         Index           =   1
         Left            =   840
         TabIndex        =   17
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Start Date"
         Height          =   255
         Index           =   0
         Left            =   840
         TabIndex        =   15
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdSCYRRPrint 
      Caption         =   "&Print"
      Height          =   360
      Left            =   135
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   3750
      Width           =   1335
   End
   Begin VB.CommandButton cmdSCYRRClose 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   360
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   3750
      Width           =   1335
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxCashBook 
      Height          =   1635
      Left            =   600
      TabIndex        =   21
      Top             =   5160
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   2884
      _Version        =   393216
      Cols            =   10
      FixedCols       =   0
      BackColorFixed  =   12632256
      BackColorSel    =   15329508
      ForeColorSel    =   -2147483640
      BackColorBkg    =   16777215
      GridColor       =   -2147483638
      GridColorFixed  =   8421504
      WordWrap        =   -1  'True
      Enabled         =   0   'False
      GridLinesFixed  =   1
      SelectionMode   =   1
      Appearance      =   0
      BandDisplay     =   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   10
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin MSForms.TextBox txtBC 
      Height          =   285
      Left            =   1440
      TabIndex        =   28
      Top             =   585
      Width           =   3510
      VariousPropertyBits=   679495711
      BorderStyle     =   1
      Size            =   "6191;503"
      SpecialEffect   =   0
      FontName        =   "Myriad Web"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox txtClientList 
      Height          =   285
      Left            =   1440
      TabIndex        =   27
      Top             =   90
      Width           =   3510
      VariousPropertyBits=   679495711
      BorderStyle     =   1
      Size            =   "6191;503"
      SpecialEffect   =   0
      FontName        =   "Myriad Web"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label84 
      BackStyle       =   0  'Transparent
      Caption         =   "Transaction Type:"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   20
      Top             =   1080
      Width           =   1275
   End
   Begin MSForms.ComboBox cboTransType 
      Height          =   315
      Left            =   1440
      TabIndex        =   2
      Top             =   1080
      Width           =   3810
      VariousPropertyBits=   1820346395
      DisplayStyle    =   3
      Size            =   "6720;556"
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
      Object.Width           =   "0"
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Client:"
      ForeColor       =   &H80000007&
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   19
      Top             =   120
      Width           =   555
   End
   Begin VB.Label Label84 
      BackStyle       =   0  'Transparent
      Caption         =   "Bank Account:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   11
      Top             =   600
      Width           =   1035
   End
End
Attribute VB_Name = "frmPreCBTransactions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim SEARCHPropertyMODE_ As Boolean
Dim LOOKUPCommand As String
Dim szaBank() As String
Dim sTextBox As String
'Private Sub cboClientID_Change()
'        Dim adoConn As New ADODB.Connection
'        adoConn.Open getConnectionString
'        PrepareCboBC adoConn
'        LoadCmbFinancialYear adoConn
'        adoConn.Close
'End Sub



Private Sub cboTransType_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 13 Then
        Option2.SetFocus
    End If
    
End Sub

Private Sub cmbFinancialYear_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 13 Then
        cmbPeriodFrom.SetFocus
    End If
End Sub

Private Sub cmbPeriodFrom_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 13 Then
        cmbPeriodTo.SetFocus
    End If
End Sub

Private Sub cmbPeriodTo_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 13 Then
        cmdSCYRRPrint.SetFocus
    End If
End Sub

Private Sub cmdBC_Click()
    picClient.Left = 269.029
    picClient.Top = 355.299
    Dim adoConn As New ADODB.Connection
    adoConn.Open getConnectionString
    sTextBox = "2"
    ConfigureFlxBank
    Dim szAllBankBalance As String
    szAllBankBalance = BankAndBalance(adoConn)
    adoConn.Close
    Set adoConn = Nothing
    picClient.Visible = True
    txtSearchClientID.SetFocus
End Sub
Private Sub ConfigureFlxBank()
   flxClient.RowHeight(0) = 0
   flxClient.Cols = 5
   flxClient.ColWidth(0) = 100
   flxClient.ColWidth(1) = 1500
   flxClient.ColWidth(2) = 4500
   flxClient.ColWidth(3) = 0
   flxClient.ColWidth(4) = 0
   txtSearchClientID.Width = 1530
   txtSearchClientName.Visible = True
   flxClient.Clear
   flxClient.Rows = 2
   flxClient.ColAlignment = vbLeftJustify
   lblClientID.Caption = "Bank Code"
   lblClientName.Caption = "Bank Name"
   lblClientID.Width = 1400
   lblClientID.Left = 50
   lblClientName.Width = 2600
   txtSearchClientName.Left = 1620
   txtSearchClientName.text = ""
   txtSearchClientID.text = ""
   txtSearchClientID.Left = 45
End Sub
Private Sub cmdClientList_Click()
     picClient.Left = 269.029
    picClient.Top = 155.299
    sTextBox = "1"
    LoadflxClient
    
    picClient.Visible = True
    txtSearchClientID.SetFocus
End Sub
Private Sub LoadflxClient()
   Dim rRow As Integer
   Dim szSQL As String

   Dim adoConn As New ADODB.Connection
   Dim rstRec As New ADODB.Recordset

   flxClient.RowHeight(0) = 0
   flxClient.Cols = 3
   flxClient.ColWidth(0) = 100
   flxClient.ColWidth(1) = 1500
   flxClient.ColWidth(2) = 4500
   
   
   txtSearchClientID.Width = 1530
   txtSearchClientName.Visible = True
   'picClient.Width = 5295
   'cmdPicCLose.Left = 5010
   
   flxClient.Clear
   flxClient.Rows = 2
   flxClient.ColAlignment(0) = vbLeftJustify
   flxClient.ColAlignment(1) = vbLeftJustify
   flxClient.ColAlignment(2) = vbLeftJustify

   '~~~ Added by Anol Configuring width and position of labels and search boxes.
   lblClientID.Caption = "Client ID"
   lblClientName.Caption = "Client Name"
   lblClientID.Width = 1400
   lblClientID.Left = 50
   lblClientName.Width = 2600
   'lblClientName.Left = lblClientID.Left + flxClient.ColWidth(0)
   txtSearchClientName.Left = 1620
   txtSearchClientName.text = ""
   txtSearchClientID.text = ""
   'txtSearchClientName.Width = 3240
   txtSearchClientID.Left = 45

   adoConn.Open getConnectionString
   szSQL = "SELECT CLIENTID, CLIENTNAME, CT FROM   CLIENT ORDER BY CLIENTID;"

   rstRec.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
           
                 rRow = 1
                While Not rstRec.EOF
                    flxClient.row = 1
                    flxClient.RowSel = 1
                    flxClient.ColSel = 1
                    flxClient.TextMatrix(rRow, 0) = ""
                    flxClient.TextMatrix(rRow, 1) = rstRec.Fields.Item(0).Value
                    flxClient.TextMatrix(rRow, 2) = rstRec.Fields.Item(1).Value
                    flxClient.RowHeight(rRow) = 280
                    rstRec.MoveNext
                    If Not rstRec.EOF Then flxClient.AddItem ""
                    rRow = rRow + 1
                 Wend
          
      
   rstRec.Close
   adoConn.Close
   Set rstRec = Nothing
   Set adoConn = Nothing

End Sub

Private Sub cmdPicCLose_Click()
     picClient.Visible = False
     cmdClientList.SetFocus
End Sub

Private Sub cmdSCYRRClose_Click()
   Unload Me
End Sub

Private Sub cmdSCYRRPrint_Click()
   'Exit Sub
   If txtBC.text = "" Then
      ShowMsgInTaskBar "Please select the bank.", "Y", "N"
      cmdBC.SetFocus
      Exit Sub
   End If
   If txtSCYRREnDt.text = "" Then
      ShowMsgInTaskBar "Please enter the end date.", "Y", "N"
      txtSCYRREnDt.SetFocus
      Exit Sub
   End If
   If cboTransType.text = "" Then
      ShowMsgInTaskBar "Please select the transaction type.", "Y", "N"
      cboTransType.SetFocus
      Exit Sub
   End If
   If txtSCYRREnDt.text = "" Then
      ShowMsgInTaskBar "Please enter the end date.", "Y", "N"
      txtSCYRREnDt.SetFocus
      Exit Sub
   End If

   Dim adoConn As New ADODB.Connection
   Dim adoRst As New ADODB.Recordset
   Dim szSQL As String
   Dim sessionID As String
   Dim reportingDate As String
   'On Error Resume Next
   adoConn.Open getConnectionString

   adoConn.Execute "UPDATE NominalLedger " & _
                   "SET Debit = 0, Credit = 0 " & _
                   "WHERE Type > 0;"

   LoadFlxCashBook adoConn
   createTable adoConn
   sessionID = GetTimeStamp
   reportingDate = Format(DateValue(Now), "dd mmmm yyyy")
   ' '

    adoConn.Execute _
      "DELETE FROM ReportCashbookTrans WHERE SessionID = '" & sessionID & "';"
     'added by anol 20161025
     adoConn.Execute "DELETE FROM ReportCashbookTrans WHERE ReportingDate < #" & reportingDate & "# ;"
     
            szSQL = GetCashBookTransQuery1(txtClientList.Tag, reportingDate, sessionID, CDate(Format(CStr(cmbPeriodFrom.Column(2)))), CDate(Format(CStr(cmbPeriodTo.Column(2)))), txtBC.Tag, CInt(cboTransType.Value))
            adoConn.Execute _
              "INSERT INTO ReportCashbookTrans " & _
              "(ReportingDate, SessionID ,ClientID , RDate, No, Type,SageAccount,Reference ,Details,Debit,Credit,StatementDate ) " & _
                szSQL
           szSQL = GetCashBookTransQuery2(txtClientList.Tag, reportingDate, sessionID, CDate(Format(CStr(cmbPeriodFrom.Column(2)))), CDate(Format(CStr(cmbPeriodTo.Column(2)))), txtBC.Tag, CInt(cboTransType.Value))
            adoConn.Execute _
              "INSERT INTO ReportCashbookTrans " & _
              "(ReportingDate, SessionID ,ClientID , RDate, No, Type,SageAccount,Reference ,Details,Debit,Credit,StatementDate ) " & _
                szSQL
        
            szSQL = GetCashBookTransQuery3(txtClientList.Tag, reportingDate, sessionID, CDate(Format(CStr(cmbPeriodFrom.Column(2)))), CDate(Format(CStr(cmbPeriodTo.Column(2)))), txtBC.Tag, CInt(cboTransType.Value))
            adoConn.Execute _
              "INSERT INTO ReportCashbookTrans " & _
              "(ReportingDate, SessionID ,ClientID , RDate, No, Type,SageAccount,Reference ,Details,Debit,Credit,StatementDate ) " & _
                szSQL
   
      
'MsgBox "DONE"
'       szSQL = "SELECT '" & reportingDate & "' AS ReportingDate, '" & sessionID & "' AS SessionID,  " & _
'    "RDate, Mid({tlbTransactionTypes.CONSTANT}, 4, Len({tlbTransactionTypes.CONSTANT})-3) & CStr({tlbReceipt.SlNumber}, 0)+ tlbReceipt.SlNumber AS No,Type, SageAccountNumber, " & _
'    "Type, Details, iif( type =23, -1*amount,amount)as amt , '', " & _
'    "left({tlbReceipt.ReconNow}, 10),  " & _
'    "From tlbReceipt where Rdate>=#" & fromDate & "# And Rdate<=#" & toDate & "#  AND Type in (3,4,23)"
        
        
   adoConn.Close
   Set adoConn = Nothing

   Dim reportApp As New CRAXDRT.Application
   Dim Report As CRAXDRT.Report

'  All option selected
   Set Report = reportApp.OpenReport(App.Path & szReportPath & "\CashBookTransactions.rpt")
   Report.Database.Tables(1).ConnectionProperties.Item("Database Password") = accessDBPws

   Report.EnableParameterPrompting = False
   Report.DiscardSavedData

   Report.ParameterFields(1).AddCurrentValue txtBC.Tag
    If Option2.Value = True Then
        Report.ParameterFields(2).AddCurrentValue CDate(CStr(cmbPeriodFrom.Column(2)))
        Report.ParameterFields(3).AddCurrentValue CDate(cmbPeriodTo.Column(3))
        Report.ParameterFields(4).AddCurrentValue CalDrCrAcBalance(CDate(Format(CStr(cmbPeriodFrom.Column(2)))), CDate(Format(CStr(cmbPeriodFrom.Column(2)))))
    Else
        Report.ParameterFields(2).AddCurrentValue CDate(Format(txtSCYRRStDt.text, "dd mmmm yyyy"))
        Report.ParameterFields(3).AddCurrentValue CDate(Format(txtSCYRREnDt.text, "dd mmmm yyyy"))
        Report.ParameterFields(4).AddCurrentValue CalDrCrAcBalance(CDate(Format(txtSCYRRStDt.text, "dd mmmm yyyy")), CDate(Format(txtSCYRREnDt.text, "dd mmmm yyyy")))
    End If
   Report.ParameterFields(5).AddCurrentValue CInt(cboTransType.Value)
   Report.ParameterFields(6).AddCurrentValue cboTransType.Column(1)
   If cboTransType.Column(0) = 0 Then
        Report.ParameterFields(7).AddCurrentValue "Net for all transactions"
   Else
        Report.ParameterFields(7).AddCurrentValue "Net " + cboTransType.Column(1)
   End If
   Report.ParameterFields(8).AddCurrentValue sessionID
   Report.ParameterFields(9).AddCurrentValue txtClientList.Tag
   If Option2.Value = True Then
        Report.ParameterFields(10).AddCurrentValue "Finanacial Year"
   Else
        Report.ParameterFields(10).AddCurrentValue "Date Range"
   End If
   
   Load frmReport
   frmReport.LoadReportViewer Report
End Sub

Private Sub createTable(adoConn As ADODB.Connection)
    Dim adoRst As New ADODB.Recordset
    On Error GoTo CreateCashbookTrans

   adoRst.Open "SELECT * FROM ReportCashbookTrans;", adoConn, adOpenStatic, adLockReadOnly
   adoRst.Close

   GoTo alreadyhvtable

CreateCashbookTrans:
   adoConn.Execute _
      "CREATE TABLE ReportCashbookTrans " & _
         "(" & _
            "ReportingDate DateTime  NOT NULL, " & _
            "SessionID     TEXT(100) NOT NULL, " & _
            "ClientID      TEXT(10), " & _
            "RDate   DateTime, " & _
            "No      TEXT(20), " & _
            "Type      BYTE, " & _
            "SageAccount      TEXT(20), " & _
            "Details         TEXT(200), " & _
            "Reference         TEXT(100), " & _
            "Debit        CURRENCY, " & _
            "Credit        CURRENCY, " & _
            "StatementDate        TEXT(20) " & _
            ");"

alreadyhvtable:


End Sub
Private Function CalDrCrAcBalance(dtStart As Date, dtEnd As Date) As Double
   Dim iRow As Integer, cDr As Currency, cCr As Currency

   On Error Resume Next

   For iRow = 1 To flxCashBook.Rows - 1
      If CDate(flxCashBook.TextMatrix(iRow, 1)) >= dtStart And _
            CDate(flxCashBook.TextMatrix(iRow, 1)) <= dtEnd Then
         cDr = cDr + CCur(flxCashBook.TextMatrix(iRow, 6))
         cCr = cCr + CCur(flxCashBook.TextMatrix(iRow, 7))
      End If
   Next iRow

   CalDrCrAcBalance = cDr - cCr
End Function



Private Sub flxClient_Click()
            Dim adoConn As New ADODB.Connection
            adoConn.Open getConnectionString
            If sTextBox = "1" Then
                    txtClientList.Tag = flxClient.TextMatrix(flxClient.row, 1)
                    txtClientList.text = flxClient.TextMatrix(flxClient.row, 2)
                    txtBC.Tag = ""
                    txtBC.text = ""
                    LoadCmbFinancialYear adoConn
                    cmdBC.SetFocus
            ElseIf sTextBox = "2" Then
                    txtBC.Tag = flxClient.TextMatrix(flxClient.row, 1)
                    txtBC.text = flxClient.TextMatrix(flxClient.row, 2)
                    LoadFlxCashBook adoConn
                    cboTransType.SetFocus
'                    txtSCYRRStDt.SetFocus
            End If
            adoConn.Close
            picClient.Visible = False
End Sub



Private Sub flxClient_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        flxClient_Click
    End If
End Sub

Private Sub Form_Load()
    Me.Height = 4845
    Me.Width = 6305
   frmMMain.Arrange vbCascade
   Me.ZOrder 0
    Frame2.Left = 112
    Me.BackColor = MODULEBACKCOLOR
    Frame1.BackColor = MODULEBACKCOLOR
    Frame2.BackColor = MODULEBACKCOLOR
    Frame3.BackColor = MODULEBACKCOLOR
    Option2.BackColor = MODULEBACKCOLOR
    Option1.BackColor = MODULEBACKCOLOR
    
    txtSCYRRStDt.text = "01/01/2000"
    txtSCYRREnDt.text = Format(Now, "dd/mm/yyyy")
    
    Dim szSQL As String
    Dim adoConn As New ADODB.Connection
    Dim TotalRow As Integer, TotalCol As Integer
    Dim i As Integer, j As Integer

    adoConn.Open getConnectionString
    Dim adoRst As New ADODB.Recordset
    szSQL = "SELECT CLIENTID, CLIENTNAME FROM CLIENT order by CLIENTID"
    adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
    If Not adoRst.EOF Then
        txtClientList.Tag = adoRst.Fields("CLIENTID").Value
        txtClientList.text = adoRst.Fields("CLIENTNAME").Value
    End If
    adoRst.Close
    szSQL = "SELECT CB.NominalCode AS BNC, CB.MY_ID AS ID, " & _
                 "N.Name AS BNN, CB.CurrentBalance AS BAL, CB.CLIENT_ID " & _
             "FROM tlbClientBanks AS CB, NominalLedger AS N " & _
             "WHERE N.ClientID = CB.CLIENT_ID AND CB.NominalCode = N.Code AND " & _
                 "CB.CLIENT_ID = '" & txtClientList.Tag & "' " & _
             "GROUP BY CB.NominalCode, CB.MY_ID, N.Name, CB.CurrentBalance, CB.CLIENT_ID;"
    adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
    If Not adoRst.EOF Then
       txtBC.Tag = adoRst.Fields("BNC").Value
       txtBC.text = adoRst.Fields("BNN").Value
    End If
    adoRst.Close
    LoadCmbFinancialYear adoConn
    LoadFlxCashBook adoConn
    Call BankAndBalance(adoConn)

   
   'Call LoadClients(adoConn)
'   Call BankAndBalance(adoConn)
   Call LoadTransType(adoConn)
'    txtClientList.text = "All Client"
'    txtClientList.Tag = "ALL"
'   cboClientID.ListIndex = 0
   Call WheelHook(Me.hWnd)
NoRes:
   adoConn.Close
   Set adoConn = Nothing
End Sub
Private Sub cmbFinancialYear_Change()
        Dim adoConn    As New ADODB.Connection
   Dim adoRst     As New ADODB.Recordset
   Dim TotalRow   As Integer
   Dim TotalCol   As Integer
   Dim szSQL      As String
   Dim Data()     As String
   Dim i          As Integer
   Dim j          As Integer
   Dim k          As Integer                    'Open flag index

   If Not IsNull(cmbFinancialYear.Value) Then
      adoConn.Open getConnectionString
      
      szSQL = "SELECT PeriodID, Period_Descp, P_StDate, P_EndDate, Status " & _
              "FROM   Periods " & _
              "WHERE  FYrID = '" & cmbFinancialYear.Value & "' " & _
              "ORDER BY P_StDate;"

'      Debug.Print szSQL
      
      adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

      If adoRst.EOF Then GoTo NoRes

      TotalRow = adoRst.RecordCount - 1
      TotalCol = adoRst.Fields.Count - 1
      ReDim Data(TotalCol, TotalRow) As String

      k = -1
      For i = 0 To TotalRow
         For j = 0 To TotalCol
            Data(j, i) = IIf(IsNull(adoRst.Fields(j).Value), "", adoRst.Fields(j).Value)
            If k = -1 And j = 4 Then
               If adoRst.Fields("Status").Value Then
                  k = i
'                  dtEnd = CDate(adoRst.Fields("P_EndDate").Value)
               End If
            End If
         Next j
         adoRst.MoveNext
         If adoRst.EOF Then Exit For
      Next i
      
      cmbPeriodFrom.Column() = Data()
      cmbPeriodTo.Column() = Data()
      
      cmbPeriodFrom.ListIndex = 0
      If (cmbPeriodTo.ListCount > 0) Then
         cmbPeriodTo.ListIndex = cmbPeriodTo.ListCount - 1
      End If

      adoConn.Close
      Set adoConn = Nothing
   End If
   Exit Sub

NoRes:
   ShowMsgInTaskBar "Periods are not found. Please contact with system support", "Y", "N"
   Set adoConn = Nothing
End Sub
Private Sub LoadTransType(adoConn As ADODB.Connection)
   Dim adoRst As New ADODB.Recordset
   Dim szSQL As String

   szSQL = "SELECT TYPE_ID, DESCRIPTION " & _
           "FROM   tlbTransactionTypes " & _
           "WHERE  TYPE_ID IN (3, 4, 8, 9, 11, 12) " & _
           "ORDER BY TYPE_ID;"

   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   Dim TotalRow As Integer, TotalCol As Integer
   Dim i As Integer, j As Integer

   TotalRow = adoRst.RecordCount
   TotalCol = adoRst.Fields.Count

   Dim Data() As String

   ReDim Data(2, 8) As String

   Data(0, 0) = "0"
   Data(1, 0) = "All Transactions"
   
   Data(0, 1) = "3"
   Data(1, 1) = "Sales Receipts"
   
   Data(0, 2) = "4"
   Data(1, 2) = "Sales Receipt on Account"
   
   Data(0, 3) = "8"
   Data(1, 3) = "Purchase Payments"
   
   Data(0, 4) = "9"
   Data(1, 4) = "Purchase Payment on Account"
   
   Data(0, 5) = "11"
   Data(1, 5) = "Bank Payments"
   
   Data(0, 6) = "12"
   Data(1, 6) = "Bank Receipts"
   
   Data(0, 7) = "23"
   Data(1, 7) = "Sales Receipt Refunds"
   
   Data(0, 8) = "24"
   Data(1, 8) = "Purchase Payment Refunds"
   
'   For i = 1 To TotalRow - 1
'       For j = 0 To totalcol - 1
'           Data(j, i) = IIf(IsNull(adoRst.Fields(j).Value), "", adoRst.Fields(j).Value)
'       Next j
'       adoRst.MoveNext
'       If adoRst.EOF Then Exit For
'   Next i
   cboTransType.Column() = Data()
   cboTransType.ListIndex = 0
   adoRst.Close
   Set adoRst = Nothing
End Sub
'Private Sub PrepareCboBC(ByVal adoConn As ADODB.Connection)
'   On Error GoTo Error_Handler
'
'   Dim adoRst As New ADODB.Recordset
'   Dim szSQL As String, Data() As String, j As Integer
'   Dim i As Integer, iTotalCol As Integer, iTotalRow As Integer
'
'   If IsNull(cboClientID) Then Exit Sub
'   If IsNull(cboClientID.Value) Then Exit Sub
'   szSQL = "SELECT C.NominalCode AS BNC, " & _
'                  "N.Name AS BNN, C.BANK_AC_NUM " & _
'              "FROM tlbClientBanks AS C,NominalLedger as N " & _
'              "WHERE C.NominalCode = N.Code AND C.CLIENT_ID=N.ClientID AND C.CLIENT_ID = '" & txtClientList.Tag & "' order by C.NominalCode;"
'
'   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'
'   If adoRst.EOF Then GoTo NoRes
'
'   iTotalRow = adoRst.RecordCount
'   iTotalCol = adoRst.Fields.count
'   ReDim Data(iTotalCol - 1, iTotalRow - 1) As String
'   ReDim szNominal(iTotalCol - 1, iTotalRow - 1) As String
'   For i = 0 To iTotalRow
'      For j = 0 To iTotalCol - 1
'         Data(j, i) = IIf(IsNull(adoRst.Fields.Item(j).Value), "", adoRst.Fields.Item(j).Value)
'      Next j
'      adoRst.MoveNext
'      If adoRst.EOF Then Exit For
'   Next i
'   cboBC.Column() = Data()
'
'NoRes:
'   adoRst.Close
'   Set adoRst = Nothing
'   Exit Sub
'
'Error_Handler:
'   ShowMsgInTaskBar ERR.description & "::" & ERR.Number, , "N"
'
'   Set adoRst = Nothing
'End Sub
'Private Sub LoadClients(adoConn As ADODB.Connection)
'   Dim adoRst As New ADODB.Recordset
'   Dim szSQL As String
'
'   szSQL = "SELECT CLIENTID, CLIENTNAME " & _
'           "FROM CLIENT " & _
'           "ORDER BY CLIENTNAME;"
'
'   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'
'   Dim TotalRow As Integer, TotalCol As Integer
'   Dim i As Integer, j As Integer
'
'   TotalRow = adoRst.RecordCount
'   TotalCol = adoRst.Fields.count
'
'   Dim Data() As String
'
'   ReDim Data(TotalCol, TotalRow) As String
'
'   Data(0, 0) = "ALL"
'   Data(1, 0) = "All Clients"
'   For i = 0 To TotalRow
'       For j = 0 To TotalCol - 1
'           Data(j, i) = IIf(IsNull(adoRst.Fields(j).Value), "", adoRst.Fields(j).Value)
'       Next j
'       adoRst.MoveNext
'       If adoRst.EOF Then Exit For
'   Next i
'   cboClientID.Column() = Data()
'
'   adoRst.Close
'   Set adoRst = Nothing
'End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call WheelUnHook(Me.hWnd)
End Sub

Private Sub Option1_Click()
    Frame1.Visible = True
    Frame2.Visible = False
End Sub

Private Sub Option2_Click()
    Frame2.Visible = True
    Frame1.Visible = False
    Frame2.Left = 112
End Sub
Private Sub LoadCmbFinancialYear(adoConn As ADODB.Connection)
   Dim szSQL      As String
   Dim TotalRow   As Integer
   Dim TotalCol   As Integer
   Dim Data()     As String
   Dim i          As Integer
   Dim j          As Integer
   Dim k          As Integer              'Open Flag index
   Dim adoRst     As New ADODB.Recordset
If Trim(txtClientList.text) = "" Then Exit Sub
   szSQL = "SELECT FYrID, FinancialYear, ClientID, FY_StDate, Status " & _
           "FROM   FinancialYear " & _
           "WHERE  ClientID = '" & txtClientList.Tag & "' " & _
           "ORDER BY FY_StDate Desc;"

   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If adoRst.EOF Then GoTo NoRes

   TotalRow = adoRst.RecordCount - 1
   TotalCol = adoRst.Fields.Count - 1
   ReDim Data(TotalCol, TotalRow) As String

   k = -1
   For i = 0 To TotalRow
      For j = 0 To TotalCol
         Data(j, i) = IIf(IsNull(adoRst.Fields(j).Value), "", adoRst.Fields(j).Value)
         If k = -1 And j = 4 Then
            If adoRst.Fields("Status").Value Then
               k = i
'               dtStartPnL = CDate(adoRst.Fields("FY_StDate").Value)
'               dtStartBS = CDate("01 January 2000")
            End If
         End If
      Next j
      adoRst.MoveNext
      If adoRst.EOF Then Exit For
   Next i
   cmbFinancialYear.Column() = Data()
   cmbFinancialYear.ListIndex = k

   adoRst.Close
   Set adoRst = Nothing
   Exit Sub

NoRes:
   Set adoRst = Nothing
   ShowMsgInTaskBar "Financial year has not been created for the client", "Y", "N"
   Exit Sub
End Sub

Private Sub Option2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmbFinancialYear.SetFocus
    End If
End Sub

Private Sub txtSCYRREnDt_Change()
    TextBoxChangeDate txtSCYRREnDt
End Sub

Private Sub txtSCYRREnDt_GotFocus()
   If Len(txtSCYRREnDt.text) < 10 Then txtSCYRREnDt.text = Format(Date, "dd/mm/yyyy")
   SelTxtInCtrl txtSCYRREnDt
End Sub

Private Sub txtSCYRREnDt_KeyPress(KeyAscii As Integer)
    TextBoxKeyPrsDate txtSCYRREnDt, KeyAscii
End Sub

Private Sub txtSCYRREnDt_LostFocus()
    If txtSCYRREnDt.text <> "" Then TextBoxFormatDate txtSCYRREnDt
End Sub

Private Sub txtSCYRRStDt_Change()
    TextBoxChangeDate txtSCYRRStDt
End Sub

Private Sub txtSCYRRStDt_GotFocus()
   If Len(txtSCYRRStDt.text) < 10 Then txtSCYRRStDt.text = Format(Date, "dd/mm/yyyy")
   SelTxtInCtrl txtSCYRRStDt
End Sub

Private Sub txtSCYRRStDt_KeyPress(KeyAscii As Integer)
    TextBoxKeyPrsDate txtSCYRRStDt, KeyAscii
End Sub

Private Sub txtSCYRRStDt_LostFocus()
    If txtSCYRRStDt.text <> "" Then TextBoxFormatDate txtSCYRRStDt
End Sub
'
'Private Sub txtSearchProperty_Change()
'   Dim sFilter_ As String
'
'   sFilter_ = "WHERE PropertyID LIKE '" & Trim(txtSearchProperty.text) & "%' " & _
'                 "ORDER BY PropertyID;"
'   PopulatePropertyLookup sFilter_
'End Sub

'Private Sub UpdateBankList()
'   Dim szaData()  As String
'   Dim i          As Integer
'   Dim k          As Integer  'counter
'
'   If txtClientList.Tag = "ALL" Then
'      cboBC.Clear
'      cboBC.Column() = szaBank()
'      Exit Sub
'   End If
'
'   k = 0
'   For i = 0 To UBound(szaBank, 1) - 1
'      If szaBank(3, i) = txtClientList.Tag Then k = k + 1
'   Next i
'
'   ReDim szaData(3, k - 1) As String
'
'   k = 0
'   For i = 0 To UBound(szaBank, 1) - 1
'      If szaBank(3, i) = txtClientList.Tag Then
'         szaData(0, k) = szaBank(0, i)
'         szaData(1, k) = szaBank(1, i)
'         szaData(2, k) = szaBank(2, i)
'         szaData(3, k) = szaBank(3, i)
'         k = k + 1
'      End If
'   Next i
'   cboBC.Clear
'   cboBC.Column() = szaData()
'End Sub

'Private Function BankAndBalance(adoConn As ADODB.Connection) As String
'   On Error GoTo Error_Handler
'
'   Dim iRec As Integer
'   Dim adoRst As New ADODB.Recordset
'   Dim szSQL As String
'
'   szSQL = "SELECT CB.NominalCode AS BNC, CB.MY_ID AS ID, " & _
'               "N.Name AS BNN, CB.CurrentBalance AS BAL, CB.CLIENT_ID " & _
'           "FROM tlbClientBanks AS CB, NominalLedger AS N " & _
'           "WHERE CB.NominalCode = N.Code AND CB.CLIENT_ID <> '' " & _
'           "GROUP BY CB.NominalCode, CB.MY_ID, N.Name, CB.CurrentBalance, CB.CLIENT_ID;"
''Debug.Print szSQL
'   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'
'   If adoRst.EOF Then
'      MsgBox "Please setup bank account for the client."
'   Else
'      ReDim szaBank(3, adoRst.RecordCount - 1) As String
'
'      While Not adoRst.EOF
'         szaBank(0, iRec) = adoRst.Fields.Item("BNC").Value
'         szaBank(1, iRec) = adoRst.Fields.Item("BNN").Value
'         szaBank(2, iRec) = adoRst.Fields.Item("ID").Value
'         szaBank(3, iRec) = adoRst.Fields.Item("CLIENT_ID").Value
'         iRec = iRec + 1
'
'         BankAndBalance = BankAndBalance & CStr(adoRst.Fields.Item("BAL").Value) & " # "
'         adoRst.MoveNext
'      Wend
''      cboBC.Clear
''      cboBC.Column() = szaBank()
'   End If
'
'   ' Destroy Objects
'   Set adoRst = Nothing
'
''   LoadAdoBank
'
'   Exit Function
'
'   ' Error Handling Code
'Error_Handler:
'   ' Destroy Objects
'   Set adoRst = Nothing
'End Function
Private Function BankAndBalance(adoConn As ADODB.Connection) As String
   On Error GoTo Error_Handler

   Dim iRec As Integer
   Dim adoRst As New ADODB.Recordset
   Dim szSQL As String, szaData() As String
   Dim rRow As Integer
   
  
         szSQL = "SELECT CB.NominalCode AS BNC, CB.MY_ID AS ID, " & _
                  "N.Name AS BNN, CB.CurrentBalance AS BAL, CB.CLIENT_ID " & _
              "FROM tlbClientBanks AS CB, NominalLedger AS N " & _
              "WHERE N.ClientID = CB.CLIENT_ID AND CB.NominalCode = N.Code AND " & _
                  "CB.CLIENT_ID = '" & txtClientList.Tag & "' " & _
              "GROUP BY CB.NominalCode, CB.MY_ID, N.Name, CB.CurrentBalance, CB.CLIENT_ID;"
 
'Debug.Print szSQL
   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If adoRst.EOF Then
   'Modified by anol 22 Feb 2015
            If txtClientList.text <> "Consolidated" Then
                    MsgBox "Please setup your Client Bank Accounts." & Chr(13) & _
                           "Please also check the nominal chart of account for the client."
             End If
             
   Else

                rRow = 1
                While Not adoRst.EOF
                    flxClient.row = 1
                    flxClient.TextMatrix(rRow, 0) = ""
                    flxClient.TextMatrix(rRow, 1) = adoRst.Fields.Item("BNC").Value
                    flxClient.TextMatrix(rRow, 2) = adoRst.Fields.Item("BNN").Value
                    flxClient.TextMatrix(rRow, 3) = adoRst.Fields.Item("ID").Value
                    flxClient.TextMatrix(rRow, 4) = adoRst.Fields.Item("CLIENT_ID").Value
                    flxClient.RowHeight(rRow) = 280
                    adoRst.MoveNext
                    If Not adoRst.EOF Then flxClient.AddItem ""
                    rRow = rRow + 1
                 Wend
   End If

   ' Destroy Objects
   Set adoRst = Nothing

   'LoadAdoBank

   Exit Function

   ' Error Handling Code
Error_Handler:
   ' Destroy Objects
   Set adoRst = Nothing
End Function
Private Sub txtSearchClientID_Change()
    'Updated by anol 22 Dec 2015
   Dim i As Integer

   If Len(txtSearchClientID.text) > 0 Then
        txtSearchClientName.text = ""
   End If

   For i = flxClient.Rows - 1 To 1 Step -1
        flxClient.RowHeight(i) = 240
        If InStr(1, UCase(flxClient.TextMatrix(i, 1)), UCase(txtSearchClientID.text), vbTextCompare) = 0 Then
              flxClient.RowHeight(i) = 0
        End If
        If flxClient.RowHeight(i) = 240 Then
              flxClient.row = i
        End If
   Next i
End Sub

Private Sub txtSearchClientID_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = vbKeyDown Then
           flxClient.SetFocus
    End If
    If KeyCode = 13 Then
           txtSearchClientName.SetFocus
    End If
End Sub

Private Sub txtSearchClientID_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 27 Then
            picClient.Visible = False
          
         
          'If sTextBox = "1" Then
           cmdClientList.SetFocus
'           ElseIf sTextBox = "2" Then
'                cmdproperty.SetFocus
'           ElseIf sTextBox = "3" Then
'                cmdFundLookUp.SetFocus
           'End If
    End If
End Sub

Private Sub txtSearchClientName_Change()
   'Updated by anol 10 Dec 2015
   Dim i As Integer

   If Len(txtSearchClientName.text) > 0 Then
        txtSearchClientID.text = ""
   End If

   For i = flxClient.Rows - 1 To 1 Step -1
        flxClient.RowHeight(i) = 240
        If InStr(1, UCase(flxClient.TextMatrix(i, 2)), UCase(txtSearchClientName.text), vbTextCompare) = 0 Then
            flxClient.RowHeight(i) = 0
        End If
        If flxClient.RowHeight(i) = 240 Then
            flxClient.row = i
        End If
   Next i
End Sub

Private Sub txtSearchClientName_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
     If KeyCode = 13 Then
         flxClient.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        If flxClient.Visible Then
            
            flxClient.SetFocus
        End If
    End If
End Sub
Public Sub LoadFlxCashBook(adoConn As ADODB.Connection)
   Dim szSQL As String, i As Integer, r As Integer, szaTemp() As String
   Dim adoRst As New ADODB.Recordset

'  Column Heading: Trans ID, Trans Type, Date, Ref, Details, Debit, Credit, Reconciled, Statement Date
'                    ^           ^         ^           ^       ^      ^          ^           ^
   szSQL = "SELECT SlNumber AS T_ID, MID(TT.CONSTANT, 4, LEN(TT.CONSTANT)-3) AS Type2, RDate, Details, Amount, " & _
                  "Type as Type1, R.ExtRef AS Rfn, R.ReconNow AS SDate, R.Reconciled, R.SageAccountNumber AS ACC " & _
           "FROM tlbReceipt AS R, tlbTransactionTypes AS TT " & _
           "WHERE (Type = 3 OR Type = 4 OR Type = 23) AND TT.TYPE_ID = type AND " & _
                  "R.BankCode = '" & txtBC.Tag & "' " & _
           "UNION " & _
           "SELECT BP.TRAN_ID AS T_ID, MID(TT.CONSTANT, 4, LEN(TT.CONSTANT)-3) AS Type2, BP.TRAN_DATE AS RDate, " & _
                  "BP.DESCRIPTION AS Details, (BP.NET_AMOUNT + BP.VAT) AS Amount, " & _
                  "BP.TransactionType AS Type1, BP.PROJ_REF AS Rfn, BP.ReconNow AS SDate, BP.Reconciled, " & _
                  "BP.BANK_AC AS ACC " & _
           "FROM tlbBankPayment AS BP, tlbTransactionTypes AS TT " & _
           "WHERE (BP.TransactionType = 11 OR BP.TransactionType = 12) AND " & _
                  "BANK_AC = '" & txtBC.Tag & "' AND BP.TransactionType = TT.TYPE_ID " & _
           "UNION " & _
           "SELECT P.SlNumber AS T_ID, MID(TT.CONSTANT, 4, LEN(TT.CONSTANT)-3) AS Type2, P.PDate AS RDate, Details, " & _
                  "P.Amount, P.Type AS Type1, P.ExtRef AS Rfn, P.ReconNow AS SDate, P.Reconciled, P.SageAccountNumber AS ACC " & _
           "FROM tlbPayment AS P, tlbTransactionTypes AS TT " & _
           "WHERE (P.Type = 8 OR P.Type = 9 OR P.Type = 24) AND " & _
                  "BankCode = '" & txtBC.Tag & "' AND P.Type = TT.TYPE_ID " & _
           "ORDER BY RDate ASC, Type2 ASC, T_ID ASC;"

'Debug.Print szSQL
   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   i = 1
   flxCashBook.Clear
   flxCashBook.Rows = 2
   While Not adoRst.EOF
      flxCashBook.TextMatrix(i, 1) = adoRst.Fields.Item("RDate").Value                            'Date

      flxCashBook.TextMatrix(i, 2) = adoRst.Fields.Item("Type2").Value & _
                                     adoRst.Fields.Item("T_ID").Value                             'Type
      flxCashBook.TextMatrix(i, 3) = adoRst.Fields.Item("ACC").Value                              'Account
      flxCashBook.TextMatrix(i, 4) = adoRst.Fields.Item("Rfn").Value                              'Name
      flxCashBook.TextMatrix(i, 5) = adoRst.Fields.Item("Details").Value                          'Details
      If adoRst.Fields.Item("Type1").Value = "3" Or adoRst.Fields.Item("Type1").Value = "4" Or _
         adoRst.Fields.Item("Type1").Value = "12" Or adoRst.Fields.Item("Type1").Value = "24" Then
         flxCashBook.TextMatrix(i, 6) = Format(adoRst.Fields.Item("Amount").Value, "0.00")         'Debit
      Else
         flxCashBook.TextMatrix(i, 7) = Format(adoRst.Fields.Item("Amount").Value, "0.00")         'Credit
      End If
      flxCashBook.TextMatrix(i, 8) = IIf(Val(adoRst.Fields.Item("Amount").Value) - _
                                         IIf(adoRst.Fields.Item("Reconciled").Value < 0, _
                                         adoRst.Fields.Item("Reconciled").Value * (-1), _
                                         adoRst.Fields.Item("Reconciled").Value) = 0, "YES", _
                                         IIf(IsNull(adoRst.Fields.Item("Reconciled").Value), _
                                         "NO", "PART"))                                            'Reconcialid
      If Not IsNull(adoRst.Fields.Item("SDate").Value) Then
         szaTemp() = Split(adoRst.Fields.Item("SDate").Value, "#")
         flxCashBook.TextMatrix(i, 9) = IIf(szaTemp(1) = "Saved", "", szaTemp(0))                     'Statement Date
         flxCashBook.TextMatrix(i, 8) = IIf(szaTemp(1) = "Saved", "NO", flxCashBook.TextMatrix(i, 8))
      End If

      adoRst.MoveNext
      If Not adoRst.EOF Then flxCashBook.AddItem ""
      i = i + 1
   Wend
   adoRst.Close
   Set adoRst = Nothing

'   For i = 1 To flxCashBook.Rows - 1
'      flxCashBook.row = i
'      If flxCashBook.TextMatrix(i, 8) = "YES" Then
'         For r = 1 To flxCashBook.Cols - 1
'            flxCashBook.col = r
'            flxCashBook.CellBackColor = RGB(162, 185, 224)
'         Next r
'      End If
'   Next i
'   flxCashBook.row = 0
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
