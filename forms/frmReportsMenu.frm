VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmReportsMenu 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Generate Reports"
   ClientHeight    =   10080
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10350
   BeginProperty Font 
      Name            =   "Myriad Web"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmReportsMenu.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   10080
   ScaleWidth      =   10350
   Begin MSComctlLib.TreeView tvwReports 
      Height          =   5415
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   9551
      _Version        =   393217
      Indentation     =   441
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "imgFormIcon"
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
   End
   Begin MSComctlLib.ImageList imgFormIcon 
      Left            =   1200
      Top             =   6480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   14
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReportsMenu.frx":0442
            Key             =   ""
            Object.Tag             =   "Client"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReportsMenu.frx":0D1C
            Key             =   ""
            Object.Tag             =   "Property"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReportsMenu.frx":15F6
            Key             =   ""
            Object.Tag             =   "Unit"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReportsMenu.frx":1ED0
            Key             =   ""
            Object.Tag             =   "Lessee"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReportsMenu.frx":2D22
            Key             =   ""
            Object.Tag             =   "Tenant"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReportsMenu.frx":303C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReportsMenu.frx":3356
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReportsMenu.frx":37A8
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReportsMenu.frx":4082
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReportsMenu.frx":4315
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReportsMenu.frx":48C3
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReportsMenu.frx":4CF3
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReportsMenu.frx":5311
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReportsMenu.frx":5717
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmReportsMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 
Public RootName As String
Private TreeLoaded As Boolean

Private Sub LoadReportsTree()
   Dim conReportsTree As New ADODB.Connection
   Dim adoRst     As New ADODB.Recordset
   Dim szSQL      As String

   conReportsTree.Open getConnectionStringUserAccess

   If RootName <> "" Then
      szSQL = "SELECT * FROM TreeRoot " & _
              "WHERE RootVisible AND TreeRoot.RootKey = '" & RootName & "';"
   Else
      szSQL = "SELECT * FROM TreeRoot WHERE RootVisible;"
   End If

   adoRst.Open szSQL, conReportsTree, adOpenStatic, adLockReadOnly

   While Not adoRst.EOF
      tvwReports.Nodes.Add , , adoRst.Fields.Item(0).Value, adoRst.Fields.Item(1).Value, adoRst.Fields.Item("FormIcon").Value, adoRst.Fields.Item("FormIcon").Value
      Me.Icon = imgFormIcon.ListImages.Item(adoRst.Fields.Item("FormIcon").Value).ExtractIcon
      
      adoRst.MoveNext
   Wend
   adoRst.Close
   If RootName = "LA" Then
        szSQL = "SELECT C.* FROM TreeRoot AS R, TreeChildren AS C " & _
              "WHERE C.ChildVisible AND Childkey='LIS' AND R.RootKey = '" & RootName & "' AND " & _
                    "R.RootKey = C.ParentKey;"
        adoRst.Open szSQL, conReportsTree, adOpenStatic, adLockReadOnly
        If adoRst.EOF Then
            conReportsTree.Execute _
            "Insert Into TreeChildren(ChildKey,ParentKey,ChildName,IconNo) values('LIS','LA','Lease Information Summary',7)"
        End If
        adoRst.Close
   Else
   
   End If
   If RootName <> "" Then
      szSQL = "SELECT C.* FROM TreeRoot AS R, TreeChildren AS C " & _
              "WHERE C.ChildVisible AND R.RootKey = '" & RootName & "' AND " & _
                    "R.RootKey = C.ParentKey;"
   Else
      szSQL = "SELECT * FROM TreeChildren WHERE ChildVisible;"
   End If
'Debug.Print szSQL
   adoRst.Open szSQL, conReportsTree, adOpenStatic, adLockReadOnly

   '# Add Cats ---        :     Group
   While Not adoRst.EOF
      tvwReports.Nodes.Add adoRst.Fields.Item(1).Value, tvwChild, adoRst.Fields.Item(0).Value, adoRst.Fields.Item(2).Value, adoRst.Fields.Item("IconNo").Value, adoRst.Fields.Item("IconNo").Value
      
      adoRst.MoveNext
   Wend
   'added by anol 20 Aug 2016
   Dim n As Node
    For Each n In tvwReports.Nodes
        If n.Expanded = False Then n.Expanded = True
    Next
   adoRst.Close
   Set adoRst = Nothing

   conReportsTree.Close
   Set conReportsTree = Nothing
   TreeLoaded = True
End Sub

Private Sub Form_Activate()
   If Not TreeLoaded Then LoadReportsTree
End Sub

Private Sub Form_Load()
   Me.Height = 5910
   Me.Width = 4365
'   frmMMain.Arrange vbCascade
'   Me.ZOrder 0
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   TreeLoaded = False
   UnLoadForm Me
End Sub

Private Sub tvwReports_DblClick()
 On Error Resume Next
   
   Select Case tvwReports.SelectedItem.key
      Case "TR"
         Load frmPreDemandTransactions
         frmPreDemandTransactions.strStatus = ""
         frmPreDemandTransactions.Caption = "Demand Transactions Report"
         frmPreDemandTransactions.Show
         'frmMMain.mnuDemandTransactionReport_Click
      Case "HR"
         'issue 483
         'Modified by anol 30 Sep 2014
         Load frmPreDemandTransactions
         frmPreDemandTransactions.strStatus = "DHistory"
         frmPreDemandTransactions.Caption = "Demand History Report"
         frmPreDemandTransactions.Show
         'frmMMain.mnuDemandHistoryReport_Click
      Case "DA"
         frmMMain.mnuDemandAnalysisReport_Click
      Case "RCC"
         frmMMain.mnuRCC_Click
      Case "RRR"
         frmMMain.mnuRentReceivedReport_Click
      Case "CSS"                                      'Client Summary Statement
         frmMMain.mnuLandlordSummaryStatement_Click
      Case "FS"
         frmMMain.mnuFundSummary_Click
      Case "CT"
         frmMMain.mnuCashbookTransactionReport_Click
      Case "CH"
         frmMMain.mnuCashbookHistoryReport_Click
      Case "TRR"
         frmMMain.mnuRP_Click
'      Case "BRR"
      
      Case "RA"
         frmMMain.mnuReceiptAnalysis_Click
'      Case "SP"
      
'      Case "BP"

      Case "LI"
         frmMMain.mnuLesseeInfo_Click
      Case "LSDTL"
        'Modified by anol 24 Mar 2015
         'frmMMain.mnuLesseeDetails_Click
         frmLDPre.strFrom = "TenantDetailsReport"
         frmLDPre.Show
       Case "LIS"
        'Modified by anol 24 Mar 2015
         'frmMMain.mnuLesseeDetails_Click
         frmLDPre.strFrom = "LeaseInformationReport"
         frmLDPre.Show
      Case "LD"
         frmMMain.mnuLeaseDetails_Click
      Case "LS"
         frmMMain.mnuLesseeStatement_Click
      Case "SASS"
         frmMMain.mnuRAS_Click
      Case "LSL"
         frmMMain.mnuLofSL_Click
      Case "RR"
         frmMMain.mnuRentReviews_Click
      Case "REL"
         frmMMain.mnuExpLease_Click
      Case "SCBR"
         frmMMain.mnuEBR_Click
      Case "BR"
         frmMMain.mnuBR_Click
      Case "SES"
         frmMMain.mnuES_Click
      Case "SBS"
         frmMMain.mnuEBSR_Click
      Case "BVA"
         frmMMain.mnuBAE_Click
      Case "PT"
         Load frmPrePurchaseTransactions
         frmPrePurchaseTransactions.strStatus = ""
         frmPrePurchaseTransactions.Show
      Case "PH"
         'issue 483
         'Modified by anol 30 Sep 2014
         Load frmPrePurchaseTransactions
         frmPrePurchaseTransactions.strStatus = "pHistory"
         frmPrePurchaseTransactions.Caption = "Purchase History Report"
         frmPrePurchaseTransactions.Show
         'frmMMain.mnuPurchaseHistoryReport_Click
      Case "UD"
         frmMMain.mnuUnitDetails_Click
      Case "VU"
         frmMMain.mmuVacUnits_Click
      Case "PL"
         frmMMain.mmuPropertyList_Click
      Case "L_ID"
         frmMMain.mnuChangeLesseeID_Click
      Case "S_ID"
         frmMMain.mnuChangeSupplierID_Click
      Case "C_ID"
         frmMMain.mnuChangeClientID_Click
      Case "P_ID"
         frmMMain.mnuChangePropertyID_Click
      Case "U_ID"
         frmMMain.mnuChangeUnitID_Click
      Case "CBH" 'added by anol 11 Aug 2016
          frmMMain.mnuCashbookHistoryReport_Click
      Case "MGF" 'added by anol 11 Aug 2016
          LoadForm frmPreManagingAgentFee
      Case "LPL" 'added by anol 11 Aug 2016
          LoadForm frmPreLesseeEvery
   End Select
End Sub


