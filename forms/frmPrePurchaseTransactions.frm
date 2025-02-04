VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmPrePurchaseTransactions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Purchase Transaction Report"
   ClientHeight    =   6645
   ClientLeft      =   1125
   ClientTop       =   435
   ClientWidth     =   12210
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
   Icon            =   "frmPrePurchaseTransactions.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4586.498
   ScaleMode       =   0  'User
   ScaleWidth      =   11465.82
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picClient 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4770
      Left            =   7875
      ScaleHeight     =   4740
      ScaleWidth      =   5535
      TabIndex        =   23
      Top             =   1395
      Visible         =   0   'False
      Width           =   5565
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
         Left            =   5190
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   0
         Width           =   255
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxClient 
         Height          =   4020
         Left            =   45
         TabIndex        =   11
         Top             =   675
         Width           =   5400
         _ExtentX        =   9525
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
         TabIndex        =   27
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label lblPayeeFlxConfigured 
         Caption         =   "NOT"
         Height          =   495
         Index           =   4
         Left            =   1515
         TabIndex        =   26
         Top             =   1800
         Width           =   1095
      End
      Begin MSForms.Label lblClientID 
         Height          =   195
         Left            =   120
         TabIndex        =   25
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
         TabIndex        =   24
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
         TabIndex        =   9
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
         TabIndex        =   10
         Top             =   375
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
         Left            =   45
         Top             =   75
         Width           =   5085
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Date range"
      Height          =   1215
      Left            =   180
      TabIndex        =   20
      Top             =   1530
      Width           =   5175
      Begin VB.TextBox txtSCYRRStDt 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2400
         TabIndex        =   5
         Top             =   240
         Width           =   1815
      End
      Begin VB.TextBox txtSCYRREnDt 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2400
         TabIndex        =   6
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Start Date"
         Height          =   255
         Index           =   0
         Left            =   840
         TabIndex        =   22
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "End Date"
         Height          =   255
         Index           =   1
         Left            =   840
         TabIndex        =   21
         Top             =   720
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdRefreshData 
      Caption         =   "&Refresh Data"
      Height          =   360
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   4545
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton cmdSCYRRPrint 
      Caption         =   "&Print"
      Default         =   -1  'True
      Height          =   360
      Left            =   750
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3780
      Width           =   1335
   End
   Begin VB.CommandButton cmdSCYRRClose 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   360
      Left            =   3420
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3780
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
      TabIndex        =   17
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
         TabIndex        =   18
         Top             =   0
         Width           =   255
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid gridPropertyLookup 
         Height          =   1605
         Left            =   75
         TabIndex        =   14
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
         TabIndex        =   13
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
      Left            =   4365
      TabIndex        =   1
      Top             =   135
      Width           =   300
   End
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
      Left            =   4365
      TabIndex        =   4
      Top             =   765
      Width           =   300
   End
   Begin VB.CheckBox chkProperty 
      Caption         =   "Excl."
      Height          =   195
      Left            =   4680
      TabIndex        =   28
      Top             =   810
      Width           =   780
   End
   Begin MSForms.TextBox txtPropertyName 
      Height          =   285
      Left            =   2160
      TabIndex        =   3
      Top             =   765
      Width           =   2250
      VariousPropertyBits=   679495711
      BorderStyle     =   1
      Size            =   "3969;503"
      SpecialEffect   =   0
      FontName        =   "Myriad Web"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox txtPropertyID 
      Height          =   285
      Left            =   1350
      TabIndex        =   2
      Top             =   765
      Width           =   765
      VariousPropertyBits=   679495711
      BorderStyle     =   1
      Size            =   "1349;503"
      SpecialEffect   =   0
      FontName        =   "Myriad Web"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox txtClientList 
      Height          =   285
      Left            =   1350
      TabIndex        =   0
      Top             =   135
      Width           =   3015
      VariousPropertyBits=   679495711
      BorderStyle     =   1
      Size            =   "5318;503"
      SpecialEffect   =   0
      FontName        =   "Myriad Web"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label74 
      BackStyle       =   0  'Transparent
      Caption         =   "Property :"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   16
      Top             =   825
      Width           =   1155
   End
   Begin VB.Label Label84 
      BackStyle       =   0  'Transparent
      Caption         =   "Client:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   15
      Top             =   120
      Width           =   555
   End
End
Attribute VB_Name = "frmPrePurchaseTransactions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim SEARCHPropertyMODE_ As Boolean
Dim LOOKUPCommand As String
Public strStatus As String
Dim sTextBox As String
Private Sub cmdFundLookUp_Click()
   If txtClientList.text = "" Then
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

   PopulatePropertyLookup IIf(txtClientList.text = "ALL", "", " WHERE CLIENTID = '" & txtClientList.text & "'")
End Sub

Private Sub chkProperty_Click()
     If chkProperty.Value = 0 Then
        txtPropertyName.text = "ALL Properties"
        txtPropertyID.text = "ALL"
        cmdProperty.Enabled = True
    Else
        txtPropertyName.text = ""
        txtPropertyName.Tag = ""
        txtPropertyID.text = ""
        cmdProperty.Enabled = False
    End If
End Sub

Private Sub cmdGridPropertyLookup_Click()
   fmePropertyLookup.Visible = False
End Sub

Private Sub cmdPropertyLookup_Click()
    picClient.Left = 42.257
    picClient.Top = 31.06
    sTextBox = "2"
    LoadflxProperty
    
    Frame1.Enabled = False
    picClient.Visible = True
    txtSearchClientID.SetFocus
    
'   If txtClientList.text = "" Then
'      ShowMsgInTaskBar "Please select a client to continue.", , "N"
'      Exit Sub
'   End If
'
'   fmePropertyLookup.Top = txtPropertyID.Top + txtPropertyID.Height + 5
'   fmePropertyLookup.Left = txtPropertyID.Left - (fmePropertyLookup.Width - txtPropertyID.Width) + 200
'   fmePropertyLookup.Visible = True
'   fmePropertyLookup.ZOrder 0
'   gridPropertyLookup.Visible = True
'   txtSearchProperty.text = ""
'   txtSearchProperty.Enabled = True
'   txtSearchProperty.SetFocus
'
'   LOOKUPCommand = "Property"
'
'   PopulatePropertyLookup IIf(txtClientList.text = "ALL", "", " WHERE CLIENTID = '" & txtClientList.text & "'")
End Sub
Private Sub LoadflxProperty()
    flxClient.RowHeight(0) = 0
   flxClient.Cols = 3
   flxClient.ColWidth(0) = 1500
   flxClient.ColWidth(1) = 4275
   flxClient.ColWidth(2) = 0
   flxClient.Clear
   flxClient.Rows = 2
   flxClient.ColAlignment(0) = vbLeftJustify
   flxClient.ColAlignment(1) = vbLeftJustify
   flxClient.ColAlignment(2) = vbLeftJustify

  
   lblClientID.Width = 1400
   lblClientID.Left = 50
   lblClientName.Width = 2600
   lblClientName.Left = lblClientID.Left + flxClient.ColWidth(0)

   lblClientID.Caption = "Property ID"
   lblClientName.Caption = "Property Name"
  
   Dim adoConn As New ADODB.Connection
   adoConn.Open getConnectionString
   Dim rRow As Integer
   Dim szSQL As String
   Dim rstRec As New ADODB.Recordset
    rRow = 1
   
    
    If txtClientList.text <> "ALL" Then
        szSQL = "SELECT PropertyID, PropertyName " & _
            "FROM Property " & _
            "WHERE ClientID = '" & txtClientList.text & "' " & _
            "ORDER BY PropertyID;"
    Else
        szSQL = "SELECT PropertyID, PropertyName " & _
            "FROM Property " & _
            "ORDER BY PropertyID;"
    End If
    
   rstRec.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
   flxClient.TextMatrix(1, 0) = "ALL"
   flxClient.TextMatrix(1, 1) = "All Properties"
   flxClient.RowHeight(1) = 280
   flxClient.AddItem ""
   
   flxClient.TextMatrix(2, 0) = "ZZZZ"
   flxClient.TextMatrix(2, 1) = "Common Properties"
   flxClient.RowHeight(2) = 280
   flxClient.AddItem ""
   rRow = 3
   While Not rstRec.EOF
       flxClient.TextMatrix(rRow, 0) = rstRec.Fields.Item(0).Value
       flxClient.TextMatrix(rRow, 1) = rstRec.Fields.Item(1).Value
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

Private Sub cmdproperty_Click()
    picClient.Left = 42.257
    picClient.Top = 31.06
    sTextBox = "2"
    LoadflxProperty
    
    Frame1.Enabled = False
    picClient.Visible = True
    txtSearchClientID.SetFocus
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

   adoConn.Close
   Set adoConn = Nothing

   Dim reportApp As New CRAXDRT.Application
   Dim Report As CRAXDRT.Report

'  All option selected
   If strStatus = "" Then
         Set Report = reportApp.OpenReport(App.Path & szReportPath & "\PurchaseTransactions.rpt")
      
      '   Report.Database.Tables(1).ConnectionProperties.Item("Database Password") = accessDBPws
      
         Report.EnableParameterPrompting = False
         Report.DiscardSavedData
         'Resolved by BOSL
         'issue 483
         'Modified by Anol 29 Sep 2014
         'Report.RecordSelectionFormula = "{tblpurinv.dueDate} >=#" & CDate(txtSCYRRStDt.text) & "# AND {tblpurinv.dueDate}<=#" & CDate(txtSCYRREnDt.text) & "#"
         'End of modification
         Report.ParameterFields(1).AddCurrentValue txtClientList.text 'client ID'in tag there is client name
         Report.ParameterFields(2).AddCurrentValue txtPropertyID.text  ' Property ID
      '   Report.ParameterFields(3).AddCurrentValue cboVatCode.Value
      '   Report.ParameterFields(4).AddCurrentValue cboVAT_Trans_Type.text
      '   Report.ParameterFields(5).AddCurrentValue ""
         Report.ParameterFields(6).AddCurrentValue CDate(txtSCYRRStDt.text)
         Report.ParameterFields(7).AddCurrentValue CDate(txtSCYRREnDt.text)
         Report.ParameterFields(8).AddCurrentValue txtClientList.Tag ''client Name
         Report.ParameterFields(9).AddCurrentValue IIf(txtPropertyID.text = "", "Common Properties", txtPropertyName.text) 'txtPropertyName.text 'property name
      '   Report.ParameterFields(10).AddCurrentValue cboVatCode.text
         Report.ParameterFields(11).AddCurrentValue gCurrentShopCentreName
      
         Load frmReport
         frmReport.LoadReportViewer Report
   Else
   'issue 483
   'Modified by anol 30 Sep
            Dim i As Integer, szMY_ID As String
            Dim rep As frmReport
            'Dim reportApp As New CRAXDRT.Application
            'Dim Report As CRAXDRT.Report
            
            Set Report = reportApp.OpenReport(App.Path & szReportPath & "\PI_List1.rpt")
            Report.Database.Tables(1).ConnectionProperties.Item("Database Password") = accessDBPws
            
            Report.EnableParameterPrompting = False
            Report.DiscardSavedData
      
            Report.ParameterFields(1).AddCurrentValue "Y"
            Report.ParameterFields(2).AddCurrentValue CDate(txtSCYRRStDt.text)
            Report.ParameterFields(3).AddCurrentValue CDate(txtSCYRREnDt.text)
            Report.ParameterFields(4).AddCurrentValue txtPropertyID.text
            Report.ParameterFields(5).AddCurrentValue txtClientList.text
            Report.ParameterFields(6).AddCurrentValue "Purchase Transactions Report"
      
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

Private Sub Form_Activate()
    FocusControl cmdClientList
End Sub

Private Sub Form_Load()
   frmMMain.Arrange vbCascade
   Me.ZOrder 0
   Me.Height = 5670
   Me.Width = 5865
   Me.BackColor = MODULEBACKCOLOR
   Frame1.BackColor = Me.BackColor
   chkProperty.BackColor = Me.BackColor

   txtSCYRRStDt.text = "01/01/2000"
   txtSCYRREnDt.text = Format(Now, "dd/mm/yyyy")
   
   txtClientList.Tag = "ALL Clients"
   txtClientList.text = "ALL"
   txtPropertyID.text = "ALL"
   txtPropertyName.text = "ALL Properties"

   Call WheelHook(Me.hWnd)
   
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

'Private Sub gridPropertyLookup_Click()
'   SEARCHPropertyMODE_ = False
'
''   crash after this line
'   If LOOKUPCommand = "Property" Then
'      txtPropertyID.text = gridPropertyLookup.TextMatrix(gridPropertyLookup.row, 0)
'      txtPropertyName.text = gridPropertyLookup.TextMatrix(gridPropertyLookup.row, 1)
'   End If
'
''    SET OTHERS
'   fmePropertyLookup.Visible = False
'   SEARCHPropertyMODE_ = True
'End Sub

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
'   picClient.Height = 4095
'   flxClient.Height = 3345
  ' flxClient.Width = 5175
   
   adoConn.Open getConnectionString
   szSQL = "SELECT CLIENTID, CLIENTNAME, CT FROM   CLIENT ORDER BY CLIENTID;"

   rstRec.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
           
           rRow = 1
           flxClient.TextMatrix(rRow, 1) = "ALL"
           flxClient.TextMatrix(rRow, 2) = "ALL Clients"
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
Private Sub cmdClientList_Click()
    picClient.Left = 42.257
    picClient.Top = 31.06
    sTextBox = "1"
    LoadflxClient
    
    Frame1.Enabled = False
    picClient.Visible = True
    txtSearchClientID.SetFocus
End Sub

Private Sub cmdPicCLose_Click()
    Frame1.Enabled = True
    
    picClient.Visible = False
    cmdClientList.SetFocus
End Sub

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

Private Sub flxClient_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyUp And flxClient.row = 1 Then
        txtSearchClientID.SetFocus
     End If
End Sub

Private Sub flxClient_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Frame1.Enabled = True
        
        If sTextBox = "1" Then
                txtClientList.Tag = flxClient.TextMatrix(flxClient.row, 1)
                txtClientList.text = flxClient.TextMatrix(flxClient.row, 2)
        End If
        picClient.Visible = False
    End If
    If KeyAscii = 27 Then
         picClient.Visible = False
          Frame1.Enabled = True
          
          If sTextBox = "1" Then
                 cmdClientList.SetFocus
          
           End If
    End If
End Sub
Private Sub flxClient_Click()
        Frame1.Enabled = True
        
        If sTextBox = "1" Then
                txtClientList.Tag = flxClient.TextMatrix(flxClient.row, 2)
                txtClientList.text = flxClient.TextMatrix(flxClient.row, 1)
                txtPropertyID.text = "ALL"
                txtPropertyName.text = "ALL Properties"
               
        End If
        If sTextBox = "2" Then
                txtPropertyID.text = flxClient.TextMatrix(flxClient.row, 0)
                txtPropertyName.text = flxClient.TextMatrix(flxClient.row, 1)
        End If
        picClient.Visible = False
        
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
