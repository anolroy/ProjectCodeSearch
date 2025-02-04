VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmAgedReports 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Aged Debtors Report"
   ClientHeight    =   5805
   ClientLeft      =   8535
   ClientTop       =   4785
   ClientWidth     =   7365
   Icon            =   "frmAgedReports.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5805
   ScaleWidth      =   7365
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picClient 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4770
      Left            =   990
      ScaleHeight     =   4740
      ScaleWidth      =   6255
      TabIndex        =   18
      Top             =   3690
      Visible         =   0   'False
      Width           =   6285
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
         Left            =   5955
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   0
         Width           =   255
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxClient 
         Height          =   4020
         Left            =   45
         TabIndex        =   10
         Top             =   675
         Width           =   6165
         _ExtentX        =   10874
         _ExtentY        =   7091
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
         TabIndex        =   22
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label lblPayeeFlxConfigured 
         Caption         =   "NOT"
         Height          =   495
         Index           =   4
         Left            =   1515
         TabIndex        =   21
         Top             =   1800
         Width           =   1095
      End
      Begin MSForms.Label lblClientID 
         Height          =   195
         Left            =   120
         TabIndex        =   20
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
         TabIndex        =   19
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
         Width           =   4545
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         Size            =   "8017;450"
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
         Width           =   5850
      End
   End
   Begin VB.Frame Frame2 
      Height          =   2175
      Left            =   45
      TabIndex        =   24
      Top             =   1980
      Width           =   7260
      Begin VB.PictureBox fmeLoading 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         ForeColor       =   &H00FF0000&
         Height          =   450
         Left            =   2430
         ScaleHeight     =   450
         ScaleWidth      =   2655
         TabIndex        =   30
         Top             =   360
         Visible         =   0   'False
         Width           =   2655
         Begin VB.Label lblLoading 
            BackStyle       =   0  'Transparent
            Caption         =   "Please wait while loading . ."
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   330
            Left            =   405
            TabIndex        =   31
            Top             =   90
            Width           =   2745
         End
      End
      Begin VB.CheckBox chkUsePostingDate 
         Caption         =   "Use Posting Date"
         Height          =   195
         Left            =   540
         TabIndex        =   4
         Top             =   900
         Width           =   3975
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "OK"
         Height          =   315
         Left            =   3225
         TabIndex        =   6
         Top             =   1350
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "Cancel"
         Height          =   315
         Left            =   4590
         TabIndex        =   7
         Top             =   1350
         Width           =   1095
      End
      Begin VB.TextBox txtDateTo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   1650
         TabIndex        =   5
         Text            =   "dd/mm/yyyy"
         Top             =   1350
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Date To:"
         Height          =   255
         Left            =   540
         TabIndex        =   29
         Top             =   1395
         Width           =   735
      End
      Begin MSForms.OptionButton optSR 
         Height          =   375
         Left            =   525
         TabIndex        =   2
         Top             =   405
         Width           =   1575
         VariousPropertyBits=   746588179
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DisplayStyle    =   5
         Size            =   "2778;661"
         Value           =   "1"
         Caption         =   "Summary Report"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.OptionButton optDR 
         Height          =   375
         Left            =   3000
         TabIndex        =   3
         Top             =   405
         Width           =   1335
         VariousPropertyBits=   746588179
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DisplayStyle    =   5
         Size            =   "2355;661"
         Value           =   "0"
         Caption         =   "Details Report"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1995
      Left            =   45
      TabIndex        =   23
      Top             =   0
      Width           =   7260
      Begin VB.CommandButton cmdProperty 
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
         Left            =   5850
         TabIndex        =   1
         Top             =   990
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
         Left            =   5865
         TabIndex        =   0
         Top             =   585
         Width           =   300
      End
      Begin VB.Label Label74 
         BackStyle       =   0  'Transparent
         Caption         =   "Type"
         Height          =   255
         Index           =   0
         Left            =   630
         TabIndex        =   28
         Top             =   990
         Width           =   1155
      End
      Begin VB.Label Label84 
         BackStyle       =   0  'Transparent
         Caption         =   "Client:"
         Height          =   255
         Left            =   630
         TabIndex        =   27
         Top             =   630
         Width           =   555
      End
      Begin MSForms.TextBox txtClientList 
         Height          =   330
         Left            =   1635
         TabIndex        =   26
         Top             =   585
         Width           =   4545
         VariousPropertyBits=   679495711
         BorderStyle     =   1
         Size            =   "8017;582"
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtPropertyName 
         Height          =   315
         Left            =   1620
         TabIndex        =   25
         Top             =   990
         Width           =   4545
         VariousPropertyBits=   746604571
         Size            =   "8017;556"
         SpecialEffect   =   6
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   3
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample4 
         Caption         =   "Sample 4"
         Height          =   1785
         Left            =   2100
         TabIndex        =   17
         Top             =   840
         Width           =   2055
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   2
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample3 
         Caption         =   "Sample 3"
         Height          =   1785
         Left            =   1545
         TabIndex        =   16
         Top             =   675
         Width           =   2055
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   1
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample2 
         Caption         =   "Sample 2"
         Height          =   1785
         Left            =   645
         TabIndex        =   15
         Top             =   300
         Width           =   2055
      End
   End
End
Attribute VB_Name = "frmAgedReports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim SEARCHPropertyMODE_ As Boolean
Dim LOOKUPCommand As String
Public szWhoIsCalling As String
Dim sTextBox As String
Private Sub cmdApply_Click()
   MsgBox "Place code here to set options w/o closing dialog!"
End Sub



Private Sub chkUsePostingDate_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtDateTo.SetFocus
'        chkUsePostingDate
    End If
End Sub

Private Sub cmdCancel_Click()
   Unload Me
End Sub







Private Sub cmdClientList_Click()
    
    picClient.Left = 100.029
    picClient.Top = 155.299
    sTextBox = "1"
    loadflxclient
    Frame2.Enabled = False
    Frame1.Enabled = False
    picClient.Visible = True
    txtSearchClientID.SetFocus
End Sub
Private Sub loadflxclient()
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
'   picClient.Height = 4095
'   flxClient.Height = 3345
  ' flxClient.Width = 5175
   
   adoConn.Open getConnectionString
   szSQL = "SELECT CLIENTID, CLIENTNAME, CT FROM   CLIENT ORDER BY CLIENTID;"

   rstRec.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
            flxClient.TextMatrix(1, 0) = ""
            flxClient.TextMatrix(1, 1) = "ALL"
            flxClient.TextMatrix(1, 2) = "All Client"
            flxClient.RowHeight(1) = 280
            flxClient.AddItem ""
            rRow = 2
          
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
Private Sub cmdOK_Click()
    Dim reportApp As New CRAXDRT.Application
    Dim Report As CRAXDRT.Report
    Dim sessionID1 As String
    Dim reportingDate As String
    Dim szWhere As String
    Dim szSQL As String
    Dim sztrans As String
    Dim rsReportAgedcreditor As New ADODB.Recordset
    Dim rsrptTransaction As New ADODB.Recordset
    Dim rsPaytrans As New ADODB.Recordset
    cmdOK.Enabled = False
    fmeLoading.Visible = True
    fmeLoading.Refresh
    SaveSetting "PropertyManagement", "ChoosedOption", "agedCreditor", txtClientList.Tag
    Dim adoConn As New ADODB.Connection
    adoConn.Open getConnectionString
    'issue  449 Fixed by anol 20170829
    adoConn.Execute "UPDATE tlbReceipt SET PostingDate = FORMAT(RDate, 'DD MMMM YYYY') where PostingDate is Null"
    'client ID was null previously and was not writing
    adoConn.Execute "Update tlbReceipt,Units,Property SET tlbReceipt.ClientID=Property.ClientID where tlbReceipt.UnitID=Units.UnitNumber AND " & _
    "Units.PropertyID=Property.PropertyID and tlbReceipt.ClientID is Null"
    Call CreateTableAgedCreditor(adoConn)
    Call CreateTableAgedCreditorSum(adoConn)
    sessionID1 = GetTimeStamp
    reportingDate = Format(DateValue(Now), "dd mmmm yyyy")
    adoConn.Execute "DELETE FROM ReportAgedCreditor ;"
    adoConn.Execute "Delete from reportagedcreditorsum"
    If szWhoIsCalling = "Debtors" Then
            'populating Debtors control reports here
            If txtClientList.Tag <> "ALL" Then
                    szWhere = "tlbreceipt.clientID='" & txtClientList.Tag & "' AND "
            End If
            ' I am inserting blank in Detail  I shall use it for putting invoice number and group. in th place of receipt number
             If chkUsePostingDate.Value Then
             
             '**************************************by posting date debtors report****************************************
             
                    szSQL = "Select '" & reportingDate & "' AS ReportingDate, '" & sessionID1 & "' AS SessionID, ClientID,TransactionID,SLNumber,SageAccountNumber,UnitID," & _
                    "FundID,Type,ref,Amount,'1' as Amountsign,RDate,DDate,PostingDate from tlbreceipt where  " & szWhere & " PostingDate<=#" & Format(txtDateTo.text, "dd MMM yyyy") & "#"
                    
                
                    adoConn.Execute "INSERT INTO ReportAgedCreditor " & _
                     "(ReportingDate,SessionID, ClientID,TransactionID,SLNumber,SageAccountNumber,UnitID,FundID,Type,Details,Amount,Amountsign,PDate,DDate,PostingDate) " & _
                       szSQL
                    

                    adoConn.Execute "Update ReportAgedCreditor A set Amount=amount*(-1) where (A.type=2 or A.type=3 or A.type=4) ;"

                            rsrptTransaction.Open "Select A.ReceiptAmount as Amt,A.Totran,A.fromtran from RptTransactions A , ReportAgedCreditor B,ReportAgedCreditor C where A.totran=C.TransactionID AND " & _
                                    "A.Fromtran=B.TransactionID AND A.AllocDate<=#" & Format(txtDateTo.text, "dd MMM yyyy") & "# AND (C.Type=1 or C.Type=23)", adoConn, adOpenDynamic, adLockReadOnly
                            While Not rsrptTransaction.EOF
                                    'updating the receipt amount to a invoice amount
                                    'reducing amount from invoice which has a receipt
                                    adoConn.Execute "Update ReportAgedCreditor set amount=amount-" & rsrptTransaction("amt").Value & " where (type=1 or type=23) AND TransactionID=" & rsrptTransaction("Totran").Value & " "
                                    'reducing amount from receipt
                                    adoConn.Execute "Update ReportAgedCreditor set Amount=Amount+" & rsrptTransaction("amt").Value & " where (type=2 or type=3 or type=4) AND TransactionID=" & rsrptTransaction("fromtran").Value & " "
                                    rsrptTransaction.MoveNext
                            Wend
                            rsrptTransaction.Close
                            Set rsrptTransaction = Nothing

                                
                    adoConn.Execute "Update ReportAgedCreditor set total=amount where SessionID = '" & sessionID1 & "' and PostingDate<=#" & Format(txtDateTo.text, "dd MMM yyyy") & "# "
                    adoConn.Execute "Update ReportAgedCreditor set current=amount where SessionID = '" & sessionID1 & "' and PostingDate<=#" & Format(txtDateTo.text, "dd MMM yyyy") & _
                                    "# AND PostingDate>=#" & Format(DateAdd("d", -30, Format(txtDateTo.text, "dd MMM yyyy")), "dd MMM yyyy") & "# "
                    adoConn.Execute "Update ReportAgedCreditor set sixty=amount where SessionID = '" & sessionID1 & "' and PostingDate<#" & _
                                     Format(DateAdd("d", -30, Format(txtDateTo.text, "dd MMM yyyy")), "dd MMM yyyy") & "# AND PostingDate>=#" & _
                                     Format(DateAdd("d", -60, Format(txtDateTo.text, "dd MMM yyyy")), "dd MMM yyyy") & "# "
                    adoConn.Execute "Update ReportAgedCreditor set ninety=amount where SessionID = '" & sessionID1 & "' and PostingDate<#" & _
                                    Format(DateAdd("d", -60, Format(txtDateTo.text, "dd MMM yyyy")), "dd MMM yyyy") & "# AND PostingDate>=#" & _
                                     Format(DateAdd("d", -90, Format(txtDateTo.text, "dd MMM yyyy")), "dd MMM yyyy") & "# "
                    adoConn.Execute "Update ReportAgedCreditor set ninetyplus=amount where SessionID = '" & sessionID1 & "' and PostingDate<#" & _
                                    Format(DateAdd("d", -90, Format(txtDateTo.text, "dd MMM yyyy")), "dd MMM yyyy") & "# "
                    
             Else
             
             '************************************************************aged debtor  report by transaction date****************************************************
             
                     szSQL = "Select '" & reportingDate & "' AS ReportingDate, '" & sessionID1 & "' AS SessionID, ClientID,TransactionID,SLNumber,SageAccountNumber,UnitID," & _
                    "FundID,Type,ref,Amount,'1' as Amountsign,RDate,DDate,PostingDate from tlbreceipt where " & szWhere & " RDate<=#" & Format(txtDateTo.text, "dd MMM yyyy") & "#"
                                        adoConn.Execute "INSERT INTO ReportAgedCreditor " & _
                     "(ReportingDate,SessionID, ClientID,TransactionID,SLNumber,SageAccountNumber,UnitID,FundID,Type,Details,Amount,Amountsign,PDate,DDate,PostingDate) " & _
                       szSQL
                    
                    adoConn.Execute "Update ReportAgedCreditor A,RptTransactions B set paytranflag=1,A.AllocationDate=B.AllocDate where A.TransactionID=B.Fromtran and  (type=3 or type=4 or type=2) "
                    
                    adoConn.Execute "Update ReportAgedCreditor A set Amount=amount*(-1) where   (A.type=2 or A.type=3 or A.type=4) ;"

                            rsrptTransaction.Open "Select A.ReceiptAmount as Amt,A.Totran,A.fromtran from RptTransactions A , ReportAgedCreditor B,ReportAgedCreditor C where A.totran=C.TransactionID AND " & _
                                    "A.Fromtran=B.TransactionID AND A.AllocDate<=#" & Format(txtDateTo.text, "dd MMM yyyy") & "# AND (C.Type=1 or C.Type=23) ", adoConn, adOpenDynamic, adLockReadOnly
                            While Not rsrptTransaction.EOF
                                    'updating the receipt amount to a invoice amount
                                    'reducing amount from invoice which has a receipt
                                    adoConn.Execute "Update ReportAgedCreditor set amount=amount-" & rsrptTransaction("amt").Value & " where (type=1 or type=23) AND TransactionID=" & rsrptTransaction("Totran").Value & " "
                                    'reducing amount from receipt
                                    adoConn.Execute "Update ReportAgedCreditor set Amount=Amount+" & rsrptTransaction("amt").Value & " where (type=2 or type=3 or type=4) AND TransactionID=" & rsrptTransaction("fromtran").Value & " "
                                    rsrptTransaction.MoveNext
                            Wend
                            rsrptTransaction.Close
                            Set rsrptTransaction = Nothing
                    adoConn.Execute "Update ReportAgedCreditor set total=amount where SessionID = '" & sessionID1 & "' and PDate<=#" & Format(txtDateTo.text, "dd MMM yyyy") & "# "
                    adoConn.Execute "Update ReportAgedCreditor set current=amount where SessionID = '" & sessionID1 & "' and PDate<=#" & Format(txtDateTo.text, "dd MMM yyyy") & "# AND PDate>=#" & Format(DateAdd("d", -30, Format(txtDateTo.text, "dd MMM yyyy")), "dd MMM yyyy") & "# "
                    adoConn.Execute "Update ReportAgedCreditor set sixty=amount where SessionID = '" & sessionID1 & "' and PDate<#" & Format(DateAdd("d", -30, Format(txtDateTo.text, "dd MMM yyyy")), "dd MMM yyyy") & "# AND PDate>=#" & Format(DateAdd("d", -60, Format(txtDateTo.text, "dd MMM yyyy")), "dd MMM yyyy") & "# "
                    adoConn.Execute "Update ReportAgedCreditor set ninety=amount where SessionID = '" & sessionID1 & "' and PDate<#" & Format(DateAdd("d", -60, Format(txtDateTo.text, "dd MMM yyyy")), "dd MMM yyyy") & "# AND PDate>=#" & Format(DateAdd("d", -90, Format(txtDateTo.text, "dd MMM yyyy")), "dd MMM yyyy") & "# "
                    adoConn.Execute "Update ReportAgedCreditor set ninetyplus=amount where SessionID = '" & sessionID1 & "' and PDate<#" & Format(DateAdd("d", -90, Format(txtDateTo.text, "dd MMM yyyy")), "dd MMM yyyy") & "# "

             End If
              
            'for thee aged debtor summery report
            adoConn.Execute "DELETE from ReportAgedCreditor where (current=0 or current is NULL) AND  (Total=0 or Total  is NULL) AND  (Sixty=0 or Sixty  is NULL) AND  (Ninety=0 or Ninety  is NULL) AND  (NinetyPlus=0 or NinetyPlus  is NULL)"
            If optSR.Value Then
                
                adoConn.Execute "INSERT INTO ReportAgedCreditorsum " & _
             "(ReportingDate,SessionID,ClientID,SageAccountNumber,Total,Current,sixty,ninety,ninetyplus) " & _
                    "SELECT  '" & reportingDate & "' AS ReportingDate, '" & sessionID1 & "' AS SessionID,ClientID,SageAccountNumber,sum(total) as A,sum(current) as B,sum(sixty) as C,sum(ninety) as D,sum(ninetyplus) as E from reportagedcreditor  group by SageAccountNumber,clientID "
                    
                adoConn.Execute "DELETE from reportagedcreditorsum where (current=0 or current is NULL) AND  (Total=0 or Total  is NULL) AND  (Sixty=0 or Sixty  is NULL) AND  (Ninety=0 or Ninety  is NULL) AND  (NinetyPlus=0 or NinetyPlus  is NULL)"
            End If
    Else ' aged creditors report
    
    'Here the Code starts for aged creditors report Types are 6,24 # 7,8,9
            'Note : I am inserting ref instead of detail here
            If txtClientList.Tag <> "ALL" Then
                    szWhere = " tlbpayment.clientID='" & txtClientList.Tag & "' AND "
            End If
'            szWhere = szWhere & " AND (tlbpayment.type=6 or tlbpayment.type=24) AND "
             If chkUsePostingDate.Value Then 'creditors report by posting date Types are 6,24 # 7,8,9
                    szSQL = "Select '" & reportingDate & "' AS ReportingDate, '" & sessionID1 & "' AS SessionID, ClientID,TransactionID,SLNumber,SageAccountNumber,UnitID," & _
                    "FundID,Type,ref,Amount,'1' as Amountsign,PDate,DDate,PostingDate from tlbpayment where " & szWhere & " PostingDate<=#" & Format(txtDateTo.text, "dd MMM yyyy") & "#"
                     
        
                    adoConn.Execute "INSERT INTO ReportAgedCreditor " & _
                     "(ReportingDate,SessionID, ClientID,TransactionID,SLNumber,SageAccountNumber,UnitID,FundID,Type,Details,Amount,Amountsign,PDate,DDate,PostingDate) " & _
                       szSQL
                    adoConn.Execute "Update ReportAgedCreditor A set Amount=amount*(-1) where (A.type=7 or A.type=8 or A.type=9) ;"
                    
                            rsrptTransaction.Open "Select A.PaymentAmount as Amt,A.Fromtran,A.ToTran from PayTransactions A , ReportAgedCreditor paymnt,ReportAgedCreditor inv " & _
                            "where inv.TransactionID=A.totran AND A.AllocDate<=#" & Format(txtDateTo.text, "dd MMM yyyy") & "#  AND A.Fromtran=paymnt.TransactionID  AND (inv.type=6 or inv.type=24) ", adoConn, adOpenDynamic, adLockReadOnly

                            While Not rsrptTransaction.EOF
                            'updating the payment invoice amount
                                    'reducing amount from invoice which has been paid
                                    adoConn.Execute "Update ReportAgedCreditor set Amount=Amount-" & rsrptTransaction("amt").Value & " where (type=6 or type=24) AND  TransactionID=" & rsrptTransaction("totran").Value & " "
                                    'reducing amount from payment
                                    adoConn.Execute "Update ReportAgedCreditor set Amount=Amount+" & rsrptTransaction("amt").Value & " where (type=7 or type=8 or type=9) AND TransactionID=" & rsrptTransaction("fromtran").Value & " "
                                    rsrptTransaction.MoveNext
                            Wend
                            rsrptTransaction.Close
                            Set rsrptTransaction = Nothing

             
             Else 'creditors report with transaction date Types are 6,24 # 7,8,9
                     
                    szSQL = "Select '" & reportingDate & "' AS ReportingDate, '" & sessionID1 & "' AS SessionID, ClientID,TransactionID,SLNumber,SageAccountNumber,UnitID," & _
                    "FundID,Type,ref,Amount,'1' as Amountsign,PDate,DDate,PostingDate from tlbpayment where " & szWhere & " PDate<=#" & Format(txtDateTo.text, "dd MMM yyyy") & "#"
                    adoConn.Execute "INSERT INTO ReportAgedCreditor " & _
                     "(ReportingDate,SessionID, ClientID,TransactionID,SLNumber,SageAccountNumber,UnitID,FundID,Type,Details,Amount,Amountsign,PDate,DDate,PostingDate) " & _
                       szSQL
                    adoConn.Execute "Update ReportAgedCreditor A set Amount=amount*(-1) where (A.type=7 or A.type=8 or A.type=9) ;"
                            
                            rsrptTransaction.Open "Select A.PaymentAmount as Amt,A.Fromtran,A.ToTran from PayTransactions A , ReportAgedCreditor paymnt,ReportAgedCreditor inv " & _
                            "where inv.TransactionID=A.totran AND A.AllocDate<=#" & Format(txtDateTo.text, "dd MMM yyyy") & "#  AND A.Fromtran=paymnt.TransactionID  AND (inv.type=6 or inv.type=24) ", adoConn, adOpenDynamic, adLockReadOnly
                                    
                            While Not rsrptTransaction.EOF
                            'updating the payment invoice amount
                                    Debug.Print rsrptTransaction("fromtran").Value
                                    If rsrptTransaction("fromtran").Value = "5769" Then
                                            Debug.Print rsrptTransaction("fromtran").Value
                                    
                                    End If
                                    'reducing amount from invoice
                                    adoConn.Execute "Update ReportAgedCreditor set amount=amount-" & rsrptTransaction("amt").Value & " where (type=6 or type=24) AND  TransactionID=" & rsrptTransaction("totran").Value & " "
                                    'reducing amount from payment
                                    adoConn.Execute "Update ReportAgedCreditor set amount=amount+" & rsrptTransaction("amt").Value & " where (type=7 or type=8 or type=9) AND TransactionID=" & rsrptTransaction("fromtran").Value & " "
                                    rsrptTransaction.MoveNext
                            Wend
                            rsrptTransaction.Close
                            Set rsrptTransaction = Nothing
             End If
 
            If chkUsePostingDate.Value Then
                    adoConn.Execute "Update ReportAgedCreditor set total=amount where SessionID = '" & sessionID1 & "' and PostingDate<=#" & Format(txtDateTo.text, "dd MMM yyyy") & "# "
                    adoConn.Execute "Update ReportAgedCreditor set current=amount where SessionID = '" & sessionID1 & "' and PostingDate<=#" & Format(txtDateTo.text, "dd MMM yyyy") & "# AND PostingDate>=#" & Format(DateAdd("d", -30, Format(txtDateTo.text, "dd MMM yyyy")), "dd MMM yyyy") & "# "
                    adoConn.Execute "Update ReportAgedCreditor set sixty=amount where SessionID = '" & sessionID1 & "' and PostingDate<#" & Format(DateAdd("d", -30, Format(txtDateTo.text, "dd MMM yyyy")), "dd MMM yyyy") & "# AND PostingDate>=#" & Format(DateAdd("d", -60, Format(txtDateTo.text, "dd MMM yyyy")), "dd MMM yyyy") & "# "
                    adoConn.Execute "Update ReportAgedCreditor set ninety=amount where SessionID = '" & sessionID1 & "' and PostingDate<#" & Format(DateAdd("d", -60, Format(txtDateTo.text, "dd MMM yyyy")), "dd MMM yyyy") & "# AND PostingDate>=#" & Format(DateAdd("d", -90, Format(txtDateTo.text, "dd MMM yyyy")), "dd MMM yyyy") & "# "
                    adoConn.Execute "Update ReportAgedCreditor set ninetyplus=amount where SessionID = '" & sessionID1 & "' and PostingDate<#" & Format(DateAdd("d", -90, Format(txtDateTo.text, "dd MMM yyyy")), "dd MMM yyyy") & "# "
        
            Else
                    adoConn.Execute "Update ReportAgedCreditor set total=amount where SessionID = '" & sessionID1 & "' and PDate<=#" & Format(txtDateTo.text, "dd MMM yyyy") & "# "
                    adoConn.Execute "Update ReportAgedCreditor set current=amount where SessionID = '" & sessionID1 & "' and PDate<=#" & Format(txtDateTo.text, "dd MMM yyyy") & "# AND PDate>#" & Format(DateAdd("d", -30, Format(txtDateTo.text, "dd MMM yyyy")), "dd MMM yyyy") & "# "
                    adoConn.Execute "Update ReportAgedCreditor set sixty=amount where SessionID = '" & sessionID1 & "' and PDate<#" & Format(DateAdd("d", -30, Format(txtDateTo.text, "dd MMM yyyy")), "dd MMM yyyy") & "# AND PDate>=#" & Format(DateAdd("d", -60, Format(txtDateTo.text, "dd MMM yyyy")), "dd MMM yyyy") & "# "
                    adoConn.Execute "Update ReportAgedCreditor set ninety=amount where SessionID = '" & sessionID1 & "' and PDate<#" & Format(DateAdd("d", -60, Format(txtDateTo.text, "dd MMM yyyy")), "dd MMM yyyy") & "# AND PDate>=#" & Format(DateAdd("d", -90, Format(txtDateTo.text, "dd MMM yyyy")), "dd MMM yyyy") & "# "
                    adoConn.Execute "Update ReportAgedCreditor set ninetyplus=amount where SessionID = '" & sessionID1 & "' and PDate<#" & Format(DateAdd("d", -90, Format(txtDateTo.text, "dd MMM yyyy")), "dd MMM yyyy") & "# "
            
            End If
            adoConn.Execute "DELETE from ReportAgedCreditor where (current=0 or current is NULL) AND  (Total=0 or Total  is NULL) AND  (Sixty=0 or Sixty  is NULL) AND  (Ninety=0 or Ninety  is NULL) AND  (NinetyPlus=0 or NinetyPlus  is NULL)"
            If txtPropertyName.Tag = "ALL" Then
            Else
                    adoConn.Execute "DELETE R.* from ReportAgedCreditor R,Supplier S where R.SageAccountNumber=S.SupplierID and S.type<>'" & txtPropertyName.Tag & "'"
            End If
            If optSR.Value Then
                adoConn.Execute "Delete from reportagedcreditorsum"
                adoConn.Execute "INSERT INTO ReportAgedCreditorsum " & _
             "(ReportingDate,SessionID,ClientID,SageAccountNumber,Total,Current,sixty,ninety,ninetyplus) " & _
                    "SELECT  '" & reportingDate & "' AS ReportingDate, '" & sessionID1 & "' AS SessionID,ClientID,SageAccountNumber,sum(total) as A,sum(current) as B,sum(sixty) as C,sum(ninety) as D,sum(ninetyplus) as E from reportagedcreditor  group by SageAccountNumber,ClientID "
                adoConn.Execute "DELETE from reportagedcreditorsum where (current=0 or current is NULL) AND  (Total=0 or Total  is NULL) AND  (Sixty=0 or Sixty  is NULL) AND  (Ninety=0 or Ninety  is NULL) AND  (NinetyPlus=0 or NinetyPlus  is NULL)"
                    
            End If
    End If
    adoConn.Close

   If optSR.Value Then
      If szWhoIsCalling = "Debtors" Then _
         Set Report = reportApp.OpenReport(App.Path & szReportPath & "\ArrearList.rpt")
      If szWhoIsCalling = "Creditors" Then _
         Set Report = reportApp.OpenReport(App.Path & szReportPath & "\CreditList.rpt")
   Else
      If szWhoIsCalling = "Debtors" Then _
         Set Report = reportApp.OpenReport(App.Path & szReportPath & "\ArrearListDetails.rpt")
      If szWhoIsCalling = "Creditors" Then _
         Set Report = reportApp.OpenReport(App.Path & szReportPath & "\CreditListDetails.rpt")
         'CreditListDetails.rpt
   End If
   
   Report.Database.Tables(1).ConnectionProperties.Item("Database Password") = accessDBPws

   Report.EnableParameterPrompting = False
   Report.DiscardSavedData

   Report.ParameterFields(1).AddCurrentValue CDate(txtDateTo.text)
   Report.ParameterFields(2).AddCurrentValue txtClientList.Tag
   Report.ParameterFields(3).AddCurrentValue "ALL" 'txtPropertyName.Tag
   Report.ParameterFields(4).AddCurrentValue User
   Report.ParameterFields(5).AddCurrentValue CBool(chkUsePostingDate.Value)
   Report.ParameterFields(6).AddCurrentValue sessionID1

   
   cmdOK.Enabled = True
   fmeLoading.Visible = False
   fmeLoading.Refresh
   Load frmReport
   frmReport.LoadReportViewer Report

   
End Sub

Private Sub CreateTableAgedCreditor(adoConn As ADODB.Connection)
    
     Dim adoRst As New ADODB.Recordset
     On Error GoTo CreateReportAgedCreditor
       
       adoRst.Open "SELECT * FROM ReportAgedCreditor;", adoConn, adOpenStatic, adLockReadOnly
       adoRst.Close
    
       GoTo alreadycreated
    
CreateReportAgedCreditor:
           adoConn.Execute _
              "CREATE TABLE ReportAgedCreditor ( ReportingDate DateTime  NOT NULL, " & _
                    "SessionID     TEXT(100) NOT NULL, " & _
                    "ClientID      TEXT(10), " & _
                    "TransactionID       LONG, " & _
                    "SLNumber      LONG, " & _
                    "SageAccountNumber TEXT(20) NOT NULL, " & _
                    "UnitID   TEXT(25), " & _
                    "FundID   Number, " & _
                    "Type     Number, " & _
                    "Details      TEXT(250), " & _
                    "Amount       CURRENCY, " & _
                    "Amountsign       number, " & _
                    "PDate         DateTime, " & _
                    "DDate        DateTime, " & _
                    "total        CURRENCY, " & _
                    "current        CURRENCY, " & _
                    "sixty        CURRENCY, " & _
                    "ninety        CURRENCY, " & _
                    "ninetyplus        CURRENCY, " & _
                    "PostingDate        DateTime, " & _
                    "paytranflag  BIT, " & _
                    "AllocationDate      DateTime, " & _
                    "PRIMARY KEY (ReportingDate, SessionID, TransactionID)" & _
                 ");"
        
alreadycreated:
End Sub
Private Sub CreateTableAgedCreditorSum(adoConn As ADODB.Connection)
    
     Dim adoRst As New ADODB.Recordset
     On Error GoTo CreateReportAgedCreditor
       
       adoRst.Open "SELECT * FROM ReportAgedCreditorsum;", adoConn, adOpenStatic, adLockReadOnly
       adoRst.Close
    
       GoTo alreadycreated
    
CreateReportAgedCreditor:
           adoConn.Execute _
              "CREATE TABLE ReportAgedCreditorsum ( ReportingDate DateTime  NOT NULL, " & _
                    "SessionID     TEXT(100) NOT NULL, " & _
                    "ClientID      TEXT(10), " & _
                     "SageAccountNumber TEXT(20) NOT NULL, " & _
                    "total        CURRENCY, " & _
                    "current        CURRENCY, " & _
                    "sixty        CURRENCY, " & _
                    "ninety        CURRENCY, " & _
                    "ninetyplus        CURRENCY " & _
                 ");"
        
alreadycreated:
End Sub

Private Sub cmdPicCLose_Click()
        picClient.Visible = False
    Frame1.Enabled = True
    Frame2.Enabled = True
    cmdClientList.SetFocus
End Sub

Private Sub cmdproperty_Click()
        picClient.Left = 100.029
        picClient.Top = 155.299
        sTextBox = "2"
        If szWhoIsCalling = "Debtors" Then
            LoadPropertyList
        Else
            LoadSupplierType
        End If
        Frame1.Enabled = False
        Frame2.Enabled = False
        picClient.Visible = True
        txtSearchClientID.SetFocus
End Sub

Private Sub flxClient_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Frame1.Enabled = True
        Frame2.Enabled = True
        If sTextBox = "1" Then
                txtClientList.Tag = flxClient.TextMatrix(flxClient.row, 1)
                txtClientList.text = flxClient.TextMatrix(flxClient.row, 2)
                txtPropertyName.text = "ALL Properties"
                txtPropertyName.Tag = "ALL"
               
                FocusControl cmdProperty
                
        End If
        If sTextBox = "2" Then
                txtPropertyName.text = flxClient.TextMatrix(flxClient.row, 2)
                txtPropertyName.Tag = flxClient.TextMatrix(flxClient.row, 1)
                optSR.SetFocus
        End If
        
       
       
        picClient.Visible = False
    End If
End Sub

Private Sub Frame2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.MousePointer = vbArrow
End Sub

Private Sub optDR_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 13 Then
'        txtDateTo.SetFocus
        chkUsePostingDate.SetFocus
    End If
End Sub

Private Sub optSR_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 13 Then
        optDR.SetFocus
    End If
End Sub

Private Sub txtDateTo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdOK.SetFocus
    End If
End Sub

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
Private Sub flxClient_Click()
        Frame1.Enabled = True
        Frame2.Enabled = True
        If sTextBox = "1" Then
                txtClientList.Tag = flxClient.TextMatrix(flxClient.row, 1)
                txtClientList.text = flxClient.TextMatrix(flxClient.row, 2)
                txtPropertyName.text = "ALL Properties"
                txtPropertyName.Tag = "ALL"
               If txtPropertyName.Visible Then
                    cmdProperty.SetFocus
                Else
                     optSR.SetFocus
                End If
                
        End If
        If sTextBox = "2" Then
                txtPropertyName.text = flxClient.TextMatrix(flxClient.row, 2)
                txtPropertyName.Tag = flxClient.TextMatrix(flxClient.row, 1)
                optSR.SetFocus
        End If
        
       
       
        picClient.Visible = False
        
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
          Frame1.Enabled = True
          Frame2.Enabled = True
          If sTextBox = "1" Then
                 cmdClientList.SetFocus
           ElseIf sTextBox = "2" Then
                cmdProperty.SetFocus
          
           End If
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
Private Sub LoadSupplierType()
   Dim rRow As Integer
   Dim szSQL As String

   Dim adoConn As New ADODB.Connection
   Dim rstRec As New ADODB.Recordset
   txtSearchClientID.text = ""
   txtSearchClientName.text = ""
   flxClient.RowHeight(0) = 0
   flxClient.Cols = 3
   flxClient.ColWidth(0) = 0
   flxClient.ColWidth(1) = 1500
   flxClient.ColWidth(2) = 4500
   flxClient.Clear
   flxClient.Rows = 2
   flxClient.ColAlignment(0) = vbLeftJustify
   flxClient.ColAlignment(1) = vbLeftJustify
   flxClient.ColAlignment(2) = vbLeftJustify
   
   txtSearchClientID.Width = 1530
   txtSearchClientName.Visible = True
   'picClient.Width = 5295
   'cmdPicCLose.Left = 5010
   txtSearchClientID.Left = 45
   '~~~ Added by Anol Configuring width and position of labels and search boxes.
   lblClientID.Caption = "Supplier Type"
   lblClientName.Caption = "Supplier Type"
'   lblClientID.Width = 1400
'   lblClientID.Left = 50
'   lblClientName.Width = 2600
'   lblClientName.Left = lblClientID.Left + flxClient.ColWidth(0)
   
   txtSearchClientName.Left = 1620
   txtSearchClientName.text = ""
   txtSearchClientID.text = ""
   'txtSearchClientName.Width = 3240
   txtSearchClientID.Left = 45
'   picClient.Height = 4095
'   flxClient.Height = 3345
'   flxClient.Width = 5175
   
   
           rRow = 1
           flxClient.TextMatrix(rRow, 0) = ""
           flxClient.TextMatrix(rRow, 1) = "ALL"
           flxClient.TextMatrix(rRow, 2) = "ALL Types"
           flxClient.RowHeight(rRow) = 280
           flxClient.AddItem ""
           rRow = 2
           flxClient.TextMatrix(rRow, 1) = "Supplier"
           flxClient.TextMatrix(rRow, 2) = "Supplier"
           flxClient.RowHeight(rRow) = 280
           flxClient.AddItem ""
           rRow = 3
           flxClient.TextMatrix(rRow, 1) = "Client"
           flxClient.TextMatrix(rRow, 2) = "Client"
           flxClient.RowHeight(rRow) = 280
           flxClient.AddItem ""
           rRow = 4
           flxClient.TextMatrix(rRow, 1) = "LLORD"
           flxClient.TextMatrix(rRow, 2) = "LandLord"
           flxClient.RowHeight(rRow) = 280
           flxClient.AddItem ""
           rRow = 5
           flxClient.TextMatrix(rRow, 1) = "Agent"
           flxClient.TextMatrix(rRow, 2) = "Managing Agent"
           flxClient.RowHeight(rRow) = 280
           flxClient.AddItem ""
           
   
   
End Sub
Private Sub LoadPropertyList()
   Dim rRow As Integer
   Dim szSQL As String

   Dim adoConn As New ADODB.Connection
   Dim rstRec As New ADODB.Recordset
   txtSearchClientID.text = ""
   txtSearchClientName.text = ""
   flxClient.RowHeight(0) = 0
   flxClient.Cols = 3
   flxClient.ColWidth(0) = 0
   flxClient.ColWidth(1) = 1500
   flxClient.ColWidth(2) = 4500
   flxClient.Clear
   flxClient.Rows = 2
   flxClient.ColAlignment(0) = vbLeftJustify
   flxClient.ColAlignment(1) = vbLeftJustify
   flxClient.ColAlignment(2) = vbLeftJustify
   
   txtSearchClientID.Width = 1530
   txtSearchClientName.Visible = True
   'picClient.Width = 5295
   'cmdPicCLose.Left = 5010
   txtSearchClientID.Left = 45
   '~~~ Added by Anol Configuring width and position of labels and search boxes.
   lblClientID.Caption = "Property ID"
   lblClientName.Caption = "Property Name"
'   lblClientID.Width = 1400
'   lblClientID.Left = 50
'   lblClientName.Width = 2600
'   lblClientName.Left = lblClientID.Left + flxClient.ColWidth(0)
   
   txtSearchClientName.Left = 1620
   txtSearchClientName.text = ""
   txtSearchClientID.text = ""
   'txtSearchClientName.Width = 3240
   txtSearchClientID.Left = 45
'   picClient.Height = 4095
'   flxClient.Height = 3345
'   flxClient.Width = 5175
   
   
   adoConn.Open getConnectionString
           
        szSQL = "SELECT PropertyID, PropertyName " & _
                    "FROM Property " & _
                    "WHERE ClientID = '" & txtClientList.Tag & "' " & _
                    "ORDER BY PropertyID;"
          
'Debug.Print szSQL
   rstRec.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
            rRow = 1
            flxClient.TextMatrix(rRow, 0) = ""
           flxClient.TextMatrix(rRow, 1) = "ALL"
           flxClient.TextMatrix(rRow, 2) = "ALL Properties"
           flxClient.RowHeight(rRow) = 280
           flxClient.AddItem ""
           rRow = 2
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
Private Sub txtClientList_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 13 Then
        cmdClientList.SetFocus
    End If
End Sub



Private Sub txtPropertyName_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 13 Then
        cmdProperty.SetFocus
    End If
End Sub
Private Sub Form_Activate()
'   frmMMain.Arrange vbCascade
   Me.Caption = "Aged " & szWhoIsCalling & " Report"
'   If szWhoIsCalling = "Debtors" Then
'      Label74(0).Visible = True
'      txtPropertyName.Visible = True
'      cmdProperty.Visible = True
'   Else
'      Label74(0).Visible = False
'      txtPropertyName.Visible = False
'      cmdProperty.Visible = False
'   End If
If szWhoIsCalling = "Debtors" Then
        txtPropertyName.text = "ALL Properties"
        txtPropertyName.Tag = "ALL"
Else
        txtPropertyName.text = "ALL Types"
        txtPropertyName.Tag = "ALL"
End If
End Sub

Private Sub Form_Load()
    txtDateTo.text = Format(Now, "dd/mm/yyyy")
'    Me.Top = 0 '(frmMMain.Height / 2) - (Me.Height / 2)
'    Me.Left = 0 '(frmMMain.Width / 2) - (Me.Width / 2)
'    frmMMain.Arrange vbCascade
'    Me.ZOrder 0
    Me.Height = 5580
    Me.Width = 7485
    Me.BackColor = MODULEBACKCOLOR
    Frame1.BackColor = MODULEBACKCOLOR
    Frame2.BackColor = MODULEBACKCOLOR
    chkUsePostingDate.BackColor = MODULEBACKCOLOR
    Dim adoConn As New ADODB.Connection
    Dim adoRst As New ADODB.Recordset
    Dim szSQL As String, i As Integer, j As Integer
    Dim TotalRow As Integer, TotalCol As Integer
    
    adoConn.Open getConnectionString


    Dim strClientID As String
    strClientID = GetSetting("PropertyManagement", "ChoosedOption", "agedCreditor")
    If Trim(strClientID) <> "" Then
        szSQL = "SELECT CLIENTID, CLIENTNAME FROM CLIENT where CLIENTID='" & Trim(strClientID) & "'"
    Else
        szSQL = "SELECT CLIENTID, CLIENTNAME FROM CLIENT order by CLIENTID"
    End If
    
    adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

    If Not adoRst.EOF Then
        txtClientList.Tag = adoRst.Fields("CLIENTID").Value
        txtClientList.text = adoRst.Fields("CLIENTNAME").Value
        txtPropertyName.text = "ALL Properties"
        txtPropertyName.Tag = "ALL"
    Else
        adoRst.Close
        adoRst.Open "SELECT CLIENTID, CLIENTNAME FROM CLIENT order by CLIENTID", adoConn, adOpenStatic, adLockReadOnly
        txtClientList.Tag = adoRst.Fields("CLIENTID").Value
        txtClientList.text = adoRst.Fields("CLIENTNAME").Value
        txtPropertyName.text = "ALL Properties"
        txtPropertyName.Tag = "ALL"
    End If

    adoRst.Close
    Set adoRst = Nothing
    
    
    adoConn.Close
    Set adoConn = Nothing
End Sub



Private Sub txtDateTo_Change()
   TextBoxChangeDate txtDateTo
End Sub

Private Sub txtDateTo_GotFocus()
   If txtDateTo.text = "dd/mm/yyyy" Then
      txtDateTo.text = ""
      Exit Sub
   End If
   If Len(txtDateTo.text) < 10 Then txtDateTo.text = Format(Date, "dd/mm/yyyy")
   SelTxtInCtrl txtDateTo
End Sub

Private Sub txtDateTo_LostFocus()
   If txtDateTo.text <> "" Then
      TextBoxFormatDate txtDateTo
   Else
      txtDateTo.text = Format(Now, "dd/mm/yyyy")
   End If
End Sub



