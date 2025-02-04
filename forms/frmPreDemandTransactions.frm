VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmPreDemandTransactions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Demand Transaction Report"
   ClientHeight    =   8085
   ClientLeft      =   1125
   ClientTop       =   1935
   ClientWidth     =   8925
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
   Icon            =   "frmPreDemandTransactions.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5580.411
   ScaleMode       =   0  'User
   ScaleWidth      =   8381.038
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Year"
      Height          =   1215
      Left            =   120
      TabIndex        =   14
      Top             =   1680
      Width           =   5175
      Begin VB.TextBox txtSCYRREnDt 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2400
         TabIndex        =   16
         Top             =   720
         Width           =   1815
      End
      Begin VB.TextBox txtSCYRRStDt 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2400
         TabIndex        =   15
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "End Date"
         Height          =   255
         Index           =   1
         Left            =   840
         TabIndex        =   18
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Start Date"
         Height          =   255
         Index           =   0
         Left            =   840
         TabIndex        =   17
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdRefreshData 
      Caption         =   "&Refresh Data"
      Height          =   360
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   3240
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton cmdSCYRRPrint 
      Caption         =   "&Print"
      Default         =   -1  'True
      Height          =   360
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   3240
      Width           =   1335
   End
   Begin VB.CommandButton cmdSCYRRClose 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   360
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3240
      Width           =   1335
   End
   Begin VB.PictureBox fmePropertyLookup 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2025
      Left            =   1080
      ScaleHeight     =   2025
      ScaleWidth      =   4200
      TabIndex        =   9
      Top             =   5160
      Visible         =   0   'False
      Width           =   4196
      Begin VB.CommandButton cmdGridPropertyLookup 
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
         Left            =   3795
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   0
         Width           =   255
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid gridPropertyLookup 
         Height          =   1605
         Left            =   75
         TabIndex        =   3
         Top             =   330
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   2831
         _Version        =   393216
         BackColor       =   16777215
         ForeColor       =   4210752
         Cols            =   9
         FixedCols       =   0
         BackColorFixed  =   16761024
         ForeColorFixed  =   16777215
         BackColorSel    =   13884353
         ForeColorSel    =   0
         BackColorBkg    =   16777215
         BackColorUnpopulated=   15728607
         GridColor       =   12632256
         GridColorFixed  =   4210752
         WordWrap        =   -1  'True
         HighLight       =   2
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
         _Band(0).Cols   =   9
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin MSForms.TextBox txtSearchProperty 
         Height          =   285
         Left            =   80
         TabIndex        =   2
         Top             =   0
         Width           =   1425
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         Size            =   "2514;503"
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
   Begin MSForms.TextBox txtPropertyName 
      Height          =   315
      Left            =   1320
      TabIndex        =   5
      Top             =   825
      Width           =   3915
      VariousPropertyBits=   746604571
      Size            =   "6906;556"
      Value           =   "All Properties"
      SpecialEffect   =   6
      FontName        =   "Myriad Web"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ComboBox cboClientID 
      Height          =   315
      Left            =   1320
      TabIndex        =   0
      Top             =   120
      Width           =   3930
      VariousPropertyBits=   1820346395
      DisplayStyle    =   3
      Size            =   "6932;556"
      TextColumn      =   2
      ColumnCount     =   2
      ListRows        =   20
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      SpecialEffect   =   6
      FontName        =   "Myriad Web"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.CommandButton cmdPropertyLookup 
      Height          =   255
      Left            =   4965
      TabIndex        =   1
      Top             =   1215
      Width           =   255
      Caption         =   """"
      Size            =   "450;450"
      FontName        =   "Myriad Web"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.TextBox txtPropertyID 
      Height          =   315
      Left            =   1320
      TabIndex        =   4
      Top             =   1185
      Width           =   3930
      VariousPropertyBits=   746604575
      BackColor       =   16777215
      MaxLength       =   4
      Size            =   "6932;556"
      Value           =   "ALL"
      SpecialEffect   =   6
      FontName        =   "Myriad Web"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label74 
      BackStyle       =   0  'Transparent
      Caption         =   "Property Name:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   8
      Top             =   825
      Width           =   1155
   End
   Begin VB.Label Label84 
      BackStyle       =   0  'Transparent
      Caption         =   "Client:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   555
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Property Code:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1185
      Width           =   1095
   End
End
Attribute VB_Name = "frmPreDemandTransactions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim SEARCHPropertyMODE_ As Boolean
Dim LOOKUPCommand As String
Public strStatus As String
Private Sub cmdFundLookUp_Click()
   If cboClientID.text = "" Then
      ShowMsgInTaskBar "Please select a client to continue.", , "N"
      Exit Sub
   End If

   fmePropertyLookup.Top = Frame1.Top + Frame1.Height + 5
   fmePropertyLookup.Left = Frame1.Left - (fmePropertyLookup.Width - Frame1.Width) + 200
   fmePropertyLookup.Visible = True
   fmePropertyLookup.ZOrder 0
   gridPropertyLookup.Visible = True
   txtSearchProperty.text = ""
   txtSearchProperty.Enabled = True
   txtSearchProperty.SetFocus

   LOOKUPCommand = "Fund"

   PopulatePropertyLookup IIf(cboClientID.Value = "ALL", "", " WHERE CLIENTID = '" & cboClientID.Value & "'")
End Sub

Private Sub cmdGridPropertyLookup_Click()
   fmePropertyLookup.Visible = False
End Sub

Private Sub cmdPropertyLookup_Click()
   If cboClientID.text = "" Then
      ShowMsgInTaskBar "Please select a client to continue.", , "N"
      Exit Sub
   End If

   fmePropertyLookup.Top = txtPropertyID.Top + txtPropertyID.Height + 5
   fmePropertyLookup.Left = txtPropertyID.Left - (fmePropertyLookup.Width - txtPropertyID.Width) + 200
   fmePropertyLookup.Visible = True
   fmePropertyLookup.ZOrder 0
   gridPropertyLookup.Visible = True
   txtSearchProperty.text = ""
   txtSearchProperty.Enabled = True
   txtSearchProperty.SetFocus

   LOOKUPCommand = "Property"

   PopulatePropertyLookup IIf(cboClientID.Value = "ALL", "", " WHERE CLIENTID = '" & cboClientID.Value & "'")
End Sub

Private Sub cmdRefreshData_Click()
   Me.MousePointer = vbHourglass
   
   Dim adoConn As New ADODB.Connection

   adoConn.Open getConnectionString

   Call ExportData2NominalLedger(adoConn)

   adoConn.Close
   Set adoConn = Nothing
   Me.MousePointer = vbArrow
End Sub

Private Sub cmdSCYRRClose_Click()
   Unload Me
End Sub

Private Sub cmdSCYRRPrint_Click()
If txtSCYRRStDt.text = "" Then
      txtSCYRRStDt.SetFocus
      Exit Sub
   End If
   If txtSCYRREnDt.text = "" Then
      txtSCYRREnDt.SetFocus
      Exit Sub
   End If

   Dim adoConn As New ADODB.Connection
   Dim adoRst As New ADODB.Recordset
   Dim szSQL As String

   'On Error Resume Next
   adoConn.Open getConnectionString

   adoConn.Execute "UPDATE NominalLedger " & _
                   "SET Debit = 0, Credit = 0 " & _
                   "WHERE Type > 0;"

'   UpdateDrCr4NC adoconn

   

   Dim reportApp As New CRAXDRT.Application
   Dim Report As CRAXDRT.Report
If strStatus = "" Then
         Set Report = reportApp.OpenReport(App.Path & szReportPath & "\DemandTransactions.rpt")
         'Resolved by BOSL
         'issue 483 Note 1a
         'Modified by anol
         Dim rs As New ADODB.Recordset
         Dim strSQL As String
      '   strSQL = " SELECT Property.PropertyID, Property.PropertyName, DemandRecords.TransactionType, DemandRecords.DmdSlNo, " & _
      '   "tlbTransactionTypes.DESCRIPTION, DemandRecords.IssueDate, DemandRecords.TenantCompanyName, DemandSplitRecords.Description, " & _
      '   "DemandSplitRecords.DueDate, DemandSplitRecords.Amount, DemandSplitRecords.VATAmount, DemandSplitRecords.TotalAmount " & _
      '   "FROM   ((((tlbTransactionTypes tlbTransactionTypes INNER JOIN DemandRecords DemandRecords ON " & _
      '   "tlbTransactionTypes.TYPE_ID=DemandRecords.TransactionType) INNER JOIN DemandSplitRecords DemandSplitRecords ON " & _
      '   "DemandRecords.DemandID=DemandSplitRecords.DemandID) INNER JOIN LeaseDetails LeaseDetails ON DemandRecords.SageAccountNumber=LeaseDetails.SageAccountNumber) " & _
      '   "INNER JOIN Units Units ON LeaseDetails.UnitNumber=Units.UnitNumber) INNER JOIN Property Property ON Units.PropertyID=Property.PropertyID " & _
      '   "ORDER BY Property.PropertyID"
      
      '    strSQL = "SELECT * " & _
      '   "FROM   ((((tlbTransactionTypes tlbTransactionTypes INNER JOIN DemandRecords DemandRecords ON " & _
      '   "tlbTransactionTypes.TYPE_ID=DemandRecords.TransactionType) INNER JOIN DemandSplitRecords DemandSplitRecords ON " & _
      '   "DemandRecords.DemandID=DemandSplitRecords.DemandID) INNER JOIN LeaseDetails LeaseDetails ON DemandRecords.SageAccountNumber=LeaseDetails.SageAccountNumber) " & _
      '   "INNER JOIN Units Units ON LeaseDetails.UnitNumber=Units.UnitNumber) INNER JOIN Property Property ON Units.PropertyID=Property.PropertyID " & _
      '   "ORDER BY Property.PropertyID"
      ''Where DemandRecords.issueDate >=" & CDate(txtSCYRRStDt.text) & " and DemandRecords.issueDate <=" & CDate(txtSCYRREnDt.text) & "
      '   rs.Open strSQL, adoConn, adOpenDynamic, adLockReadOnly
      '   Report.Database.Tables(1).ConnectionProperties.Item("Database Password") = accessDBPws
      '   Report.Database.SetDataSource rs
      '   rs.Close
      '   Set rs = Nothing
         adoConn.Close
         Set adoConn = Nothing
         Report.EnableParameterPrompting = False
         Report.DiscardSavedData
      
         Report.ParameterFields(1).AddCurrentValue cboClientID.Value
         Report.ParameterFields(2).AddCurrentValue txtPropertyID.text
      '   Report.ParameterFields(3).AddCurrentValue cboVatCode.Value
      '   Report.ParameterFields(4).AddCurrentValue cboVAT_Trans_Type.text
      '   Report.ParameterFields(5).AddCurrentValue ""
         Report.ParameterFields(6).AddCurrentValue CDate(txtSCYRRStDt.text)
         Report.ParameterFields(7).AddCurrentValue CDate(txtSCYRREnDt.text)
         Report.ParameterFields(8).AddCurrentValue cboClientID.Column(1)
         Report.ParameterFields(9).AddCurrentValue txtPropertyName.text
      '   Report.ParameterFields(10).AddCurrentValue cboVatCode.text
         Report.ParameterFields(11).AddCurrentValue gCurrentShopCentreName
      
         Load frmReport
         frmReport.LoadReportViewer Report
   ElseIf strStatus = "DHistory" Then
      'Dim reportApp As New CRAXDRT.Application
      'Dim Report As CRAXDRT.Report
      Dim rep As frmReport
   
      Set Report = reportApp.OpenReport(App.Path & szReportPath & "\PrintDemandHistory.rpt")
      Report.Database.Tables(1).ConnectionProperties.Item("Database Password") = accessDBPws
   
      Report.EnableParameterPrompting = False
      Report.DiscardSavedData
      Report.ParameterFields(1).AddCurrentValue CDate(txtSCYRRStDt.text)
      Report.ParameterFields(2).AddCurrentValue CDate(txtSCYRREnDt.text)
      Report.ParameterFields(3).AddCurrentValue txtPropertyID.text
      Set rep = New frmReport
      Load rep
      rep.LoadReportViewer Report
 End If
End Sub

Private Sub UpdateDrCr4NC(adoConn As ADODB.Connection)
   Dim szSQL As String, szPropSrc As String
   Dim szFundPA As String, szFundDEPT_ID As String, szFundSageDepartment As String
   Dim szFundSA As String, szFundSR As String, szFundBank As String
   Dim adoRst As New ADODB.Recordset, adoNL As New ADODB.Recordset

   If txtPropertyID.text = "ALL" Then
      szPropSrc = ""
   Else
      szPropSrc = "P.PropertyID = '" & txtPropertyID.text & "' AND "
   End If

'   If txtFundNo.text = "ALL" Then
'      szFundDEPT_ID = ""
'      szFundSageDepartment = ""
'      szFundPA = ""
'      szFundSA = ""
'      szFundSR = ""
'      szFundBank = ""
'   Else
'      szFundDEPT_ID = "S.DEPT_ID = '" & txtFundNo.text & "' AND "
'      szFundPA = "P.FundID = " & txtFundNo.text & " AND "
'      szFundSageDepartment = "S.SageDepartment = " & txtFundNo.text & " AND "
'      szFundSA = "R.FundID = " & txtFundNo.text & " AND "
'      szFundSR = "S.SageDepartment = " & txtFundNo.text & " AND "
'      szFundBank = "B.DEPT_ID = '" & txtFundNo.text & "' AND "
'   End If
'--------------------------------------------------------------------------------------------------
'##########                            Purchase Invoices & Credit - PI, PC
'--------------------------------------------------------------------------------------------------
   adoNL.Open "SELECT * FROM NominalLedger;", adoConn, adOpenDynamic, adLockOptimistic

   szSQL = "SELECT S.NOMINAL_CODE AS NC, SUM(S.TOTAL_AMOUNT) AS T, P.TransactionType AS TT " & _
           "FROM tblPurInvSRec AS S INNER JOIN tblPurInv AS P ON S.ParentID = P.MY_ID " & _
           "WHERE " & szFundDEPT_ID & _
              szPropSrc & _
              "P.TRAN_DATE >= #" & Format(txtSCYRRStDt.text, "DD MMMM YYYY") & "# AND " & _
              "P.TRAN_DATE <= #" & Format(txtSCYRREnDt.text, "DD MMMM YYYY") & "# " & _
           "GROUP BY S.NOMINAL_CODE, P.TransactionType;"

   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   While Not adoRst.EOF
      adoNL.Find "Code = '" & adoRst.Fields.Item(0).Value & "'", , , 1
      If adoNL.Fields.Item("DrCr").Value = "Cr" Then
         If adoRst.Fields.Item(2).Value = 6 Then adoNL.Fields.Item("Credit").Value = Val(adoNL.Fields.Item("Credit").Value) + Val(adoRst.Fields.Item("T").Value)
         If adoRst.Fields.Item(2).Value = 7 Then adoNL.Fields.Item("Credit").Value = Val(adoNL.Fields.Item("Credit").Value) - Val(adoRst.Fields.Item("T").Value)
      Else
         If adoRst.Fields.Item(2).Value = 6 Then adoNL.Fields.Item("Debit").Value = Val(adoNL.Fields.Item("Debit").Value) + Val(adoRst.Fields.Item("T").Value)
         If adoRst.Fields.Item(2).Value = 7 Then adoNL.Fields.Item("Debit").Value = Val(adoNL.Fields.Item("Debit").Value) - Val(adoRst.Fields.Item("T").Value)
      End If
      adoRst.MoveNext
      adoNL.Update
   Wend
   adoNL.Close
   adoRst.Close

   adoNL.Open "SELECT * FROM NominalLedger WHERE Code = '2100';", adoConn, adOpenDynamic, adLockOptimistic

   szSQL = "SELECT SUM(S.TOTAL_AMOUNT) AS T, P.TransactionType AS TT " & _
           "FROM tblPurInvSRec AS S INNER JOIN tblPurInv AS P ON S.ParentID = P.MY_ID " & _
           "WHERE " & szFundDEPT_ID & _
              szPropSrc & _
              "P.TRAN_DATE >= #" & Format(txtSCYRRStDt.text, "DD MMMM YYYY") & "# AND " & _
              "P.TRAN_DATE <= #" & Format(txtSCYRREnDt.text, "DD MMMM YYYY") & "# " & _
           "GROUP BY P.TransactionType;"
'Debug.Print szSQL
   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   While Not adoRst.EOF
      If adoNL.Fields.Item("DrCr").Value = "Cr" Then
         If adoRst.Fields.Item(1).Value = 6 Then adoNL.Fields.Item("Credit").Value = Val(adoNL.Fields.Item("Credit").Value) + Val(adoRst.Fields.Item("T").Value)
         If adoRst.Fields.Item(1).Value = 7 Then adoNL.Fields.Item("Credit").Value = Val(adoNL.Fields.Item("Credit").Value) - Val(adoRst.Fields.Item("T").Value)
      Else
         If adoRst.Fields.Item(1).Value = 6 Then adoNL.Fields.Item("Debit").Value = Val(adoNL.Fields.Item("Debit").Value) + Val(adoRst.Fields.Item("T").Value)
         If adoRst.Fields.Item(1).Value = 7 Then adoNL.Fields.Item("Debit").Value = Val(adoNL.Fields.Item("Debit").Value) - Val(adoRst.Fields.Item("T").Value)
      End If
      adoRst.MoveNext
   Wend

   adoNL.Update

   adoNL.Close
   adoRst.Close

'--------------------------------------------------------------------------------------------------
'##########                             Payment on AC - PA
'--------------------------------------------------------------------------------------------------
   adoNL.Open "SELECT * FROM NominalLedger;", adoConn, adOpenDynamic, adLockOptimistic

   szSQL = "SELECT P.BankCode AS NC, SUM(Amount) AS T " & _
           "FROM   tlbPayment AS P " & _
           "WHERE  P.Type = 9 AND " & _
              szFundPA & _
              "P.PDate >= #" & Format(txtSCYRRStDt.text, "DD MMMM YYYY") & "# AND " & _
              "P.PDate <= #" & Format(txtSCYRREnDt.text, "DD MMMM YYYY") & "# " & _
           "GROUP BY P.BankCode;"
   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   While Not adoRst.EOF
      adoNL.Find "Code = '" & adoRst.Fields.Item("NC").Value & "'", , , 1
      If adoNL.Fields.Item("DrCr").Value = "Cr" Then
         adoNL.Fields.Item("Credit").Value = Val(adoNL.Fields.Item("Credit").Value) + _
               IIf(IsNull(adoRst.Fields.Item("T").Value), 0, (adoRst.Fields.Item("T").Value))
      Else
         adoNL.Fields.Item("Debit").Value = Val(adoNL.Fields.Item("Debit").Value) + _
               IIf(IsNull(adoRst.Fields.Item("T").Value), 0, (adoRst.Fields.Item("T").Value))
      End If

      adoRst.MoveNext
      adoNL.Update
   Wend
   adoNL.Close
   adoRst.Close

   adoNL.Open "SELECT * FROM NominalLedger WHERE Code = '2100';", adoConn, adOpenDynamic, adLockOptimistic

   szSQL = "SELECT SUM(Amount) AS T " & _
           "FROM tlbPayment AS P " & _
           "WHERE  P.Type = 9 AND " & _
              szFundPA & _
              "P.PDate >= #" & Format(txtSCYRRStDt.text, "DD MMMM YYYY") & "# AND " & _
              "P.PDate <= #" & Format(txtSCYRREnDt.text, "DD MMMM YYYY") & "#;"

   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If adoNL.Fields.Item("DrCr").Value = "Cr" Then
      adoNL.Fields.Item("Credit").Value = Val(adoNL.Fields.Item("Credit").Value) - _
            IIf(IsNull(adoRst.Fields.Item("T").Value), 0, adoRst.Fields.Item("T").Value)
   Else
'If IsNull(adoRST.Fields.Item("T").Value) Then
      adoNL.Fields.Item("Debit").Value = Val(adoNL.Fields.Item("Debit").Value) - _
            IIf(IsNull(adoRst.Fields.Item("T").Value), 0, adoRst.Fields.Item("T").Value)
   End If

   adoNL.Update

   adoNL.Close
   adoRst.Close

'--------------------------------------------------------------------------------------------------
'##########                             Purchase Payment - PP
'--------------------------------------------------------------------------------------------------
   adoNL.Open "SELECT * FROM NominalLedger;", adoConn, adOpenDynamic, adLockOptimistic

   szSQL = "SELECT SQ.NC, SUM(SQ.A) AS T " & _
           "FROM (" & _
               "SELECT P1.BankCode AS NC, P1.Amount AS A, P1.TransactionID " & _
               "FROM   tlbPayment AS P1, tlbPayment AS P2, PayTransactions AS P, tblPurInvSRec AS S " & _
               "WHERE  P1.Type = 8 AND P1.TransactionID = P.FromTran AND " & _
                  "P2.TransactionID = P.ToTran AND P2.PI = S.ParentID AND " & _
                  szFundDEPT_ID & _
                  "P1.PDate >= #" & Format(txtSCYRRStDt.text, "DD MMMM YYYY") & "# AND " & _
                  "P1.PDate <= #" & Format(txtSCYRREnDt.text, "DD MMMM YYYY") & "# " & _
               "GROUP BY P1.TransactionID, P1.BankCode, P1.Amount" & _
           ") AS SQ " & _
           "GROUP BY SQ.NC;"

   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   While Not adoRst.EOF
      adoNL.Find "Code = '" & adoRst.Fields.Item("NC").Value & "'", , , 1
      If adoNL.Fields.Item("DrCr").Value = "Cr" Then
         adoNL.Fields.Item("Credit").Value = Val(adoNL.Fields.Item("Credit").Value) - _
               IIf(IsNull(adoRst.Fields.Item("T").Value), 0, (adoRst.Fields.Item("T").Value))
      Else
         adoNL.Fields.Item("Debit").Value = Val(adoNL.Fields.Item("Debit").Value) - _
               IIf(IsNull(adoRst.Fields.Item("T").Value), 0, (adoRst.Fields.Item("T").Value))
      End If

      adoRst.MoveNext
      adoNL.Update
   Wend
   adoNL.Close
   adoRst.Close

   adoNL.Open "SELECT * FROM NominalLedger WHERE Code = '2100';", adoConn, adOpenDynamic, adLockOptimistic

   szSQL = "SELECT SUM(SQ.A) AS T " & _
           "FROM (" & _
               "SELECT P1.Amount AS A, P1.TransactionID " & _
               "FROM   tlbPayment AS P1, tlbPayment AS P2, PayTransactions AS P, tblPurInvSRec AS S " & _
               "WHERE  P1.Type = 8 AND P1.TransactionID = P.FromTran AND " & _
                  "P2.TransactionID = P.ToTran AND P2.PI = S.ParentID AND " & _
                  szFundDEPT_ID & _
                  "P1.PDate >= #" & Format(txtSCYRRStDt.text, "DD MMMM YYYY") & "# AND " & _
                  "P1.PDate <= #" & Format(txtSCYRREnDt.text, "DD MMMM YYYY") & "# " & _
               "GROUP BY P1.TransactionID, P1.Amount" & _
           ") AS SQ;"

   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If adoNL.Fields.Item("DrCr").Value = "Cr" Then
      adoNL.Fields.Item("Credit").Value = Val(adoNL.Fields.Item("Credit").Value) - _
            IIf(IsNull(adoRst.Fields.Item(0).Value), 0, (adoRst.Fields.Item(0).Value))
   Else
      adoNL.Fields.Item("Debit").Value = Val(adoNL.Fields.Item("Debit").Value) - _
            IIf(IsNull(adoRst.Fields.Item(0).Value), 0, (adoRst.Fields.Item(0).Value))
   End If

   adoNL.Update

   adoNL.Close
   adoRst.Close

'--------------------------------------------------------------------------------------------------
'##########                             Sales Invoice and Credit - SI & SC
'--------------------------------------------------------------------------------------------------
   adoNL.Open "SELECT * FROM NominalLedger;", adoConn, adOpenDynamic, adLockOptimistic

   szSQL = "SELECT S.NominalCodeforAmount AS NC, SUM(S.TotalAmount) AS T, D.TransactionType AS TT " & _
           "FROM DemandSplitRecords AS S, DemandRecords AS D, Property AS P, Units AS U " & _
           "WHERE " & szFundSageDepartment & " S.DemandID = D.DemandID AND " & _
              szPropSrc & " P.PropertyID = U.PropertyID AND " & _
              "D.IssueDate >= #" & Format(txtSCYRRStDt.text, "DD MMMM YYYY") & "# AND " & _
              "D.IssueDate <= #" & Format(txtSCYRREnDt.text, "DD MMMM YYYY") & "# AND " & _
              "U.UnitNumber = D.UnitNumber " & _
           "GROUP BY S.NominalCodeforAmount, D.TransactionType;"
   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   While Not adoRst.EOF
      adoNL.Find "Code = '" & adoRst.Fields.Item(0).Value & "'", , , 1
      If adoNL.Fields.Item("DrCr").Value = "Cr" Then
         If adoRst.Fields.Item(2).Value = 1 Then _
            adoNL.Fields.Item("Credit").Value = Val(adoNL.Fields.Item("Credit").Value) + _
                  IIf(IsNull(adoRst.Fields.Item("T").Value), 0, (adoRst.Fields.Item("T").Value))
         If adoRst.Fields.Item(2).Value = 2 Then _
            adoNL.Fields.Item("Credit").Value = Val(adoNL.Fields.Item("Credit").Value) - _
                  IIf(IsNull(adoRst.Fields.Item("T").Value), 0, (adoRst.Fields.Item("T").Value))
      Else
         If adoRst.Fields.Item(2).Value = 1 Then _
            adoNL.Fields.Item("Credit").Value = Val(adoNL.Fields.Item("Debit").Value) + (adoRst.Fields.Item("T").Value)
         If adoRst.Fields.Item(2).Value = 2 Then _
            adoNL.Fields.Item("Credit").Value = Val(adoNL.Fields.Item("Debit").Value) - (adoRst.Fields.Item("T").Value)
      End If

      adoRst.MoveNext
      adoNL.Update
   Wend
   adoNL.Close
   adoRst.Close

   adoNL.Open "SELECT * FROM NominalLedger WHERE Code = '1100';", adoConn, adOpenDynamic, adLockOptimistic

   szSQL = "SELECT SUM(S.TotalAmount) AS T, D.TransactionType AS TT " & _
           "FROM DemandSplitRecords AS S, DemandRecords AS D, Property AS P, Units AS U " & _
           "WHERE " & szFundSageDepartment & " S.DemandID = D.DemandID AND " & _
              szPropSrc & _
              "D.IssueDate >= #" & Format(txtSCYRRStDt.text, "DD MMMM YYYY") & "# AND " & _
              "D.IssueDate <= #" & Format(txtSCYRREnDt.text, "DD MMMM YYYY") & "# AND " & _
              "U.UnitNumber = D.UnitNumber AND P.PropertyID = U.PropertyID " & _
           "GROUP BY D.TransactionType;"
'Debug.Print szSQL
   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If Not adoRst.EOF Then
      If adoNL.Fields.Item("DrCr").Value = "Cr" Then
         If adoRst.Fields.Item("TT").Value = 1 Then _
            adoNL.Fields.Item("Credit").Value = Val(adoNL.Fields.Item("Credit").Value) + _
                                                Val(IIf(IsNull(adoRst.Fields.Item("T").Value), 0, _
                                                adoRst.Fields.Item("T").Value))
         If adoRst.Fields.Item("TT").Value = 2 Then _
            adoNL.Fields.Item("Credit").Value = Val(adoNL.Fields.Item("Credit").Value) - _
                                                Val(IIf(IsNull(adoRst.Fields.Item("T").Value), 0, _
                                                adoRst.Fields.Item("T").Value))
      Else
         If adoRst.Fields.Item("TT").Value = 1 Then _
            adoNL.Fields.Item("Debit").Value = Val(adoNL.Fields.Item("Debit").Value) + _
                                               Val(IIf(IsNull(adoRst.Fields.Item("T").Value), 0, _
                                               adoRst.Fields.Item("T").Value))
         If adoRst.Fields.Item("TT").Value = 2 Then _
            adoNL.Fields.Item("Debit").Value = Val(adoNL.Fields.Item("Debit").Value) - _
                                               Val(IIf(IsNull(adoRst.Fields.Item("T").Value), 0, _
                                               adoRst.Fields.Item("T").Value))
      End If

      adoNL.Update
   End If

   adoNL.Close
   adoRst.Close

'--------------------------------------------------------------------------------------------------
'##########                             Sales Receipt on Account - SA
'--------------------------------------------------------------------------------------------------
   adoNL.Open "SELECT * FROM NominalLedger;", adoConn, adOpenDynamic, adLockOptimistic

   szSQL = "SELECT R.BankCode, SUM(Amount) AS T " & _
           "FROM   tlbReceipt AS R " & _
           "WHERE  R.Type = 4 AND " & _
              szFundSA & _
              "R.RDate >= #" & Format(txtSCYRRStDt.text, "DD MMMM YYYY") & "# AND " & _
              "R.RDate <= #" & Format(txtSCYRREnDt.text, "DD MMMM YYYY") & "# " & _
           "GROUP BY R.BankCode;"
   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   While Not adoRst.EOF
      adoNL.Find "Code = '" & adoRst.Fields.Item("BankCode").Value & "'", , , 1
      If adoNL.Fields.Item("DrCr").Value = "Cr" Then
         adoNL.Fields.Item("Credit").Value = Val(adoNL.Fields.Item("Credit").Value) - _
                                             Val(IIf(IsNull(adoRst.Fields.Item("T").Value), 0, _
                                             adoRst.Fields.Item("T").Value))
      Else
         adoNL.Fields.Item("Debit").Value = Val(adoNL.Fields.Item("Debit").Value) - _
                                            Val(IIf(IsNull(adoRst.Fields.Item("T").Value), 0, _
                                            adoRst.Fields.Item("T").Value))
      End If

      adoRst.MoveNext
      adoNL.Update
   Wend
   adoNL.Close
   adoRst.Close

   adoNL.Open "SELECT * FROM NominalLedger WHERE Code = '1100';", adoConn, adOpenDynamic, adLockOptimistic

   szSQL = "SELECT SUM(Amount) AS T " & _
           "FROM tlbReceipt AS R " & _
           "WHERE R.Type = 4 AND " & _
              szFundSA & _
              "R.RDate >= #" & Format(txtSCYRRStDt.text, "DD MMMM YYYY") & "# AND " & _
              "R.RDate <= #" & Format(txtSCYRREnDt.text, "DD MMMM YYYY") & "#;"
'Debug.Print szSQL
   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If adoNL.Fields.Item("DrCr").Value = "Cr" Then
      adoNL.Fields.Item("Credit").Value = Val(adoNL.Fields.Item("Credit").Value) - _
                                          Val(IIf(IsNull(adoRst.Fields.Item("T").Value), 0, _
                                          adoRst.Fields.Item("T").Value))
   Else
      adoNL.Fields.Item("Debit").Value = Val(adoNL.Fields.Item("Debit").Value) - _
                                         Val(IIf(IsNull(adoRst.Fields.Item("T").Value), 0, _
                                         adoRst.Fields.Item("T").Value))
   End If

   adoNL.Update

   adoNL.Close
   adoRst.Close
'--------------------------------------------------------------------------------------------------
'##########                             Sales Receipt - SR
'--------------------------------------------------------------------------------------------------
   adoNL.Open "SELECT * FROM NominalLedger;", adoConn, adOpenDynamic, adLockOptimistic

   szSQL = "SELECT SQ.NC, SUM(SQ.A) AS T " & _
           "FROM (" & _
               "SELECT R1.BankCode AS NC, R1.Amount AS A, R1.TransactionID " & _
               "FROM   tlbReceipt AS R1, tlbReceipt AS R2, RptTransactions AS R, DemandSplitRecords AS S " & _
               "WHERE  R1.Type = 3 AND R1.TransactionID = R.FromTran AND " & _
                  "R2.TransactionID = R.ToTran AND R2.DemandRef = S.DemandID AND " & _
                  szFundSR & _
                  "R1.RDate >= #" & Format(txtSCYRRStDt.text, "DD MMMM YYYY") & "# AND " & _
                  "R1.RDate <= #" & Format(txtSCYRREnDt.text, "DD MMMM YYYY") & "# " & _
               "GROUP BY R1.TransactionID, R1.BankCode, R1.Amount" & _
           ") AS SQ " & _
           "GROUP BY SQ.NC;"

'Debug.Print szSQL
   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   While Not adoRst.EOF
      adoNL.Find "Code = '" & adoRst.Fields.Item(0).Value & "'", , , 1
      If adoNL.Fields.Item("DrCr").Value = "Cr" Then
         adoNL.Fields.Item("Credit").Value = Val(adoNL.Fields.Item("Credit").Value) - _
                                             Val(IIf(IsNull(adoRst.Fields.Item("T").Value), 0, _
                                             adoRst.Fields.Item("T").Value))
      Else
         adoNL.Fields.Item("Debit").Value = Val(adoNL.Fields.Item("Debit").Value) - _
                                            Val(IIf(IsNull(adoRst.Fields.Item("T").Value), 0, _
                                            adoRst.Fields.Item("T").Value))
      End If

      adoRst.MoveNext
      adoNL.Update
   Wend
   adoNL.Close
   adoRst.Close

   adoNL.Open "SELECT * FROM NominalLedger WHERE Code = '1100';", adoConn, adOpenDynamic, adLockOptimistic

   szSQL = "SELECT SUM(SQ.A) AS T " & _
           "FROM (" & _
               "SELECT R1.Amount AS A, R1.TransactionID " & _
               "FROM   tlbReceipt AS R1, tlbReceipt AS R2, RptTransactions AS R, DemandSplitRecords AS S " & _
               "WHERE  R1.Type = 3 AND R1.TransactionID = R.FromTran AND " & _
                  "R2.TransactionID = R.ToTran AND R2.DemandRef = S.DemandID AND " & _
                  szFundSR & _
                  "R1.RDate >= #" & Format(txtSCYRRStDt.text, "DD MMMM YYYY") & "# AND " & _
                  "R1.RDate <= #" & Format(txtSCYRREnDt.text, "DD MMMM YYYY") & "# " & _
               "GROUP BY R1.TransactionID, R1.Amount" & _
           ") AS SQ"
'Debug.Print szSQL
   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If adoNL.Fields.Item("DrCr").Value = "Cr" Then
      adoNL.Fields.Item("Credit").Value = Val(adoNL.Fields.Item("Credit").Value) - Val(adoRst.Fields.Item(0).Value)
   Else
      adoNL.Fields.Item("Debit").Value = Val(adoNL.Fields.Item("Debit").Value) - Val(adoRst.Fields.Item(0).Value)
   End If

   adoNL.Update

   adoNL.Close
   adoRst.Close

'--------------------------------------------------------------------------------------------------
'##########                             Bank Payment and Receipt - BP & BR
'--------------------------------------------------------------------------------------------------
   adoNL.Open "SELECT * FROM NominalLedger;", adoConn, adOpenDynamic, adLockOptimistic

   szSQL = "SELECT B.BANK_AC, (SUM(NET_AMOUNT) + SUM(VAT)) AS T, B.TRANS " & _
           "FROM   tlbBankPayment AS B " & _
           "WHERE " & _
              szFundBank & _
              "B.TRAN_DATE >= #" & Format(txtSCYRRStDt.text, "DD MMMM YYYY") & "# AND " & _
              "B.TRAN_DATE <= #" & Format(txtSCYRREnDt.text, "DD MMMM YYYY") & "# " & _
           "GROUP BY B.BANK_AC, B.TRANS;"
   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   While Not adoRst.EOF
      adoNL.Find "Code = '" & adoRst.Fields.Item(0).Value & "'", , , 1
      If adoRst.Fields.Item(2).Value = "BR" Then
         If adoNL.Fields.Item("DrCr").Value = "Cr" Then
            adoNL.Fields.Item("Credit").Value = Val(adoNL.Fields.Item("Credit").Value) + Val(adoRst.Fields.Item("T").Value)
         Else
            adoNL.Fields.Item("Debit").Value = Val(adoNL.Fields.Item("Debit").Value) + Val(adoRst.Fields.Item("T").Value)
         End If
      End If
      If adoRst.Fields.Item(2).Value = "BP" Then
         If adoNL.Fields.Item("DrCr").Value = "Cr" Then
            adoNL.Fields.Item("Credit").Value = Val(adoNL.Fields.Item("Credit").Value) - Val(adoRst.Fields.Item("T").Value)
         Else
            adoNL.Fields.Item("Debit").Value = Val(adoNL.Fields.Item("Debit").Value) - Val(adoRst.Fields.Item("T").Value)
         End If
      End If

      adoRst.MoveNext
      adoNL.Update
   Wend
'   adoNL.Close
   adoRst.Close

   szSQL = "SELECT B.NOMINAL_CODE, (SUM(NET_AMOUNT) + SUM(VAT)) AS T, B.TRANS " & _
           "FROM   tlbBankPayment AS B " & _
           "WHERE " & _
              szFundBank & _
              "B.TRAN_DATE >= #" & Format(txtSCYRRStDt.text, "DD MMMM YYYY") & "# AND " & _
              "B.TRAN_DATE <= #" & Format(txtSCYRREnDt.text, "DD MMMM YYYY") & "# " & _
           "GROUP BY B.NOMINAL_CODE, B.TRANS;"
'Debug.Print szSQL
   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   While Not adoRst.EOF
      adoNL.Find "Code = '" & adoRst.Fields.Item(0).Value & "'", , , 1
'      If adoNL.Fields.Item("DrCr").Value = "Cr" Then
'         adoNL.Fields.Item("Credit").Value = Val(adoNL.Fields.Item("Credit").Value) + Val(adoRST.Fields.Item("T").Value)
'      Else
'         adoNL.Fields.Item("Debit").Value = Val(adoNL.Fields.Item("Debit").Value) + Val(adoRST.Fields.Item("T").Value)
'      End If
      If adoRst.Fields.Item(2).Value = "BR" Then
         If adoNL.Fields.Item("DrCr").Value = "Cr" Then
            adoNL.Fields.Item("Credit").Value = Val(adoNL.Fields.Item("Credit").Value) + Val(adoRst.Fields.Item("T").Value)
         Else
            adoNL.Fields.Item("Debit").Value = Val(adoNL.Fields.Item("Debit").Value) + Val(adoRst.Fields.Item("T").Value)
         End If
      End If
      If adoRst.Fields.Item(2).Value = "BP" Then
         If adoNL.Fields.Item("DrCr").Value = "Cr" Then
            adoNL.Fields.Item("Credit").Value = Val(adoNL.Fields.Item("Credit").Value) - Val(adoRst.Fields.Item("T").Value)
         Else
            adoNL.Fields.Item("Debit").Value = Val(adoNL.Fields.Item("Debit").Value) - Val(adoRst.Fields.Item("T").Value)
         End If
      End If

      adoRst.MoveNext
      adoNL.Update
   Wend

   adoNL.Close
   adoRst.Close

   Set adoNL = Nothing
   Set adoRst = Nothing
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_Load()
   Dim Data() As String
   frmMMain.Arrange vbCascade
   Me.ZOrder 0
   Me.Height = 4320
   Me.Width = 5595
   Me.BackColor = MODULEBACKCOLOR
   Frame1.BackColor = Me.BackColor

   txtSCYRRStDt.text = "01/01/2000"
   txtSCYRREnDt.text = Format(Now, "dd/mm/yyyy")

   Dim szSQL As String
   Dim adoConn As New ADODB.Connection
   Dim adoRst As New ADODB.Recordset
   Dim TotalRow As Integer, TotalCol As Integer
   Dim i As Integer, j As Integer

   adoConn.Open getConnectionString

   ' Clients
   szSQL = "SELECT CLIENTID, CLIENTNAME FROM CLIENT "
   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   TotalRow = adoRst.RecordCount
   TotalCol = adoRst.Fields.count

   ReDim Data(TotalCol, TotalRow) As String

   Data(0, 0) = "ALL"
   Data(1, 0) = "All Clients"
   For i = 1 To TotalRow
      For j = 0 To TotalCol - 1
         Data(j, i) = IIf(IsNull(adoRst.Fields(j).Value), "", adoRst.Fields(j).Value)
      Next j
      adoRst.MoveNext
      If adoRst.EOF Then Exit For
   Next i
   cboClientID.Column() = Data()
   cboClientID.ListIndex = 0
   adoRst.Close

   Set adoRst = Nothing

   Call WheelHook(Me.hWnd)

NoRes:
   adoConn.Close
   Set adoConn = Nothing
End Sub

Public Function PopulatePropertyLookup(strFilter_ As String)
   Dim conProperty_ As New ADODB.Connection
   Dim rstProperty_ As New ADODB.Recordset
   Dim szSQL As String
   Dim iRow As Integer

   'On Error Resume Next
   conProperty_.Open getConnectionString

   'CLREATE SQL QUERY ON OPTION BUTTON SELECTION
   If LOOKUPCommand = "Property" Then
      szSQL = "SELECT PropertyID, PropertyNAME " _
            & "From Property " & strFilter_
   End If
   If LOOKUPCommand = "Fund" Then szSQL = "SELECT FundID, FundName FROM FUND;"

   rstProperty_.Open szSQL, conProperty_, adOpenStatic, adLockReadOnly

   gridPropertyLookup.Clear
   gridPropertyLookup.Rows = 2
   gridPropertyLookup.Cols = 2
   ConfigurFlexGrid

   iRow = 1
   On Error Resume Next
   While Not rstProperty_.EOF
      gridPropertyLookup.TextMatrix(iRow, 0) = IIf(rstProperty_.Fields.Item(0) = Null, "", rstProperty_.Fields.Item(0))
      gridPropertyLookup.TextMatrix(iRow, 1) = IIf(rstProperty_.Fields.Item(1) = Null, "", rstProperty_.Fields.Item(1))

      rstProperty_.MoveNext
      If Not rstProperty_.EOF Then gridPropertyLookup.AddItem ""
      iRow = iRow + 1
   Wend

   rstProperty_.Close
   conProperty_.Close
   Set rstProperty_ = Nothing
   Set conProperty_ = Nothing
End Function

Private Sub ConfigurFlexGrid()
   fmePropertyLookup.Visible = True
   gridPropertyLookup.Visible = True

   gridPropertyLookup.RowHeight(0) = 255
   gridPropertyLookup.row = 0
   Dim i As Integer
   For i = 0 To gridPropertyLookup.Cols - 1
        gridPropertyLookup.col = i
        gridPropertyLookup.CellFontBold = True
   Next i

   gridPropertyLookup.ColWidth(0) = 800

   If LOOKUPCommand = "Property" Then _
      gridPropertyLookup.TextMatrix(0, 0) = "ID"

   gridPropertyLookup.ColWidth(1) = 2860
   gridPropertyLookup.TextMatrix(0, 1) = "Name"
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call WheelUnHook(Me.hWnd)
End Sub

Private Sub gridPropertyLookup_Click()
   SEARCHPropertyMODE_ = False

'   crash after this line
   If LOOKUPCommand = "Property" Then
      txtPropertyID.text = gridPropertyLookup.TextMatrix(gridPropertyLookup.row, 0)
      txtPropertyName.text = gridPropertyLookup.TextMatrix(gridPropertyLookup.row, 1)
   End If

'    SET OTHERS
   fmePropertyLookup.Visible = False
   SEARCHPropertyMODE_ = True
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

Private Sub txtSearchProperty_Change()
   Dim sFilter_ As String
   
   sFilter_ = "WHERE PropertyID LIKE '" & Trim(txtSearchProperty.text) & "%' " & _
                 "ORDER BY PropertyID;"
   PopulatePropertyLookup sFilter_
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
