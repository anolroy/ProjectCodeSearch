VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmServiceChargeDetails 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Budget Analysis"
   ClientHeight    =   11280
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6990
   Icon            =   "frmServiceChargeDetails.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   11280
   ScaleWidth      =   6990
   Begin VB.Frame Frame5 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Height          =   5055
      Left            =   135
      TabIndex        =   21
      Top             =   8145
      Visible         =   0   'False
      Width           =   6630
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
         Left            =   6285
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   45
         Width           =   255
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxClientList 
         Height          =   4335
         Left            =   90
         TabIndex        =   13
         Top             =   675
         Width           =   6435
         _ExtentX        =   11351
         _ExtentY        =   7646
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
      Begin MSForms.TextBox TextBox1 
         Height          =   255
         Left            =   5265
         TabIndex        =   11
         Top             =   390
         Width           =   1170
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         Size            =   "2064;450"
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label lblClientID 
         Height          =   195
         Index           =   2
         Left            =   5175
         TabIndex        =   24
         Top             =   180
         Width           =   1185
         VariousPropertyBits=   8388627
         Caption         =   "Balance"
         Size            =   "2090;344"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label lblClientID 
         Height          =   195
         Index           =   1
         Left            =   1635
         TabIndex        =   23
         Top             =   195
         Width           =   1185
         VariousPropertyBits=   8388627
         Caption         =   "Client Name"
         Size            =   "2090;344"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtSearchClientName 
         Height          =   255
         Left            =   1665
         TabIndex        =   10
         Top             =   390
         Width           =   3555
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         Size            =   "6271;450"
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtSearchClientID 
         Height          =   255
         Left            =   315
         TabIndex        =   9
         Top             =   390
         Width           =   1305
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         Size            =   "2302;450"
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label lblClientID 
         Height          =   195
         Index           =   0
         Left            =   360
         TabIndex        =   22
         Top             =   180
         Width           =   1230
         VariousPropertyBits=   8388627
         Caption         =   "Client ID"
         Size            =   "2170;344"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Shape Shape2 
         Height          =   5010
         Left            =   0
         Top             =   0
         Width           =   6585
      End
      Begin VB.Shape Shape4 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C0C0C0&
         BorderStyle     =   6  'Inside Solid
         FillColor       =   &H00E0FFFF&
         FillStyle       =   0  'Solid
         Height          =   240
         Index           =   15
         Left            =   90
         Top             =   135
         Width           =   6345
      End
   End
   Begin VB.Frame Frame1 
      Height          =   7785
      Index           =   1
      Left            =   30
      TabIndex        =   14
      Top             =   -105
      Width           =   6915
      Begin VB.CommandButton cmdSCDBdClose 
         Caption         =   "Save"
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   4005
         TabIndex        =   8
         Top             =   7335
         Width           =   1215
      End
      Begin VB.CommandButton cmdSCDBdCancel 
         Caption         =   "&Cancel"
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   5445
         TabIndex        =   7
         Top             =   7335
         Width           =   1215
      End
      Begin VB.TextBox txtSCFundCode 
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   1
         EndProperty
         Height          =   315
         Left            =   1095
         TabIndex        =   26
         Top             =   225
         Width           =   1305
      End
      Begin VB.CommandButton cmdSCFund 
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
         Left            =   6495
         TabIndex        =   0
         Top             =   225
         Width           =   300
      End
      Begin VB.TextBox txtSCFund 
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   1
         EndProperty
         Height          =   315
         Left            =   2430
         TabIndex        =   25
         Top             =   225
         Width           =   4005
      End
      Begin VB.TextBox txtTotalArea 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   1
         EndProperty
         Height          =   315
         Left            =   1080
         TabIndex        =   1
         Top             =   615
         Width           =   1305
      End
      Begin VB.TextBox txtPpsf 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   1
         EndProperty
         Height          =   315
         Left            =   4845
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   615
         Width           =   1575
      End
      Begin VB.CommandButton cmdNew 
         Caption         =   "&New"
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   120
         TabIndex        =   4
         Top             =   6750
         Width           =   1095
      End
      Begin VB.CommandButton cmdSCDBdDelete 
         Caption         =   "&Delete"
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2290
         TabIndex        =   6
         Top             =   6750
         Width           =   1095
      End
      Begin VB.CommandButton cmdSCDBdEdit 
         Caption         =   "&Edit"
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1200
         TabIndex        =   5
         Top             =   6750
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox txtSCDBudgetTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   1
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
         Height          =   285
         Left            =   5355
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   6735
         Width           =   1200
      End
      Begin VB.TextBox txtRentChargesIDEdit 
         Height          =   285
         Left            =   12720
         TabIndex        =   15
         Top             =   3720
         Visible         =   0   'False
         Width           =   855
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxSCBudgetDetailsAnalysis 
         Height          =   5055
         Left            =   120
         TabIndex        =   3
         Top             =   1440
         Width           =   6675
         _ExtentX        =   11774
         _ExtentY        =   8916
         _Version        =   393216
         ForeColor       =   0
         Cols            =   6
         FixedCols       =   0
         BackColorFixed  =   12632256
         BackColorSel    =   15329508
         ForeColorSel    =   0
         BackColorBkg    =   16777215
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
         _Band(0).Cols   =   6
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin VB.Label lblRentCharges 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fund"
         Height          =   195
         Index           =   3
         Left            =   180
         TabIndex        =   29
         Top             =   255
         Width           =   360
      End
      Begin VB.Label lblRentCharges 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Price/SqFoot"
         Height          =   195
         Index           =   5
         Left            =   3735
         TabIndex        =   28
         Top             =   660
         Width           =   945
      End
      Begin VB.Label lblRentCharges 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Area"
         Height          =   195
         Index           =   6
         Left            =   180
         TabIndex        =   27
         Top             =   615
         Width           =   735
      End
      Begin VB.Label lblRentCharges 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Budget"
         Height          =   195
         Index           =   4
         Left            =   4290
         TabIndex        =   20
         Top             =   6765
         Width           =   915
      End
      Begin VB.Label lblRentCharges 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nominal Name"
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
         Left            =   1440
         TabIndex        =   19
         Top             =   1200
         Width           =   1020
      End
      Begin VB.Label lblRentCharges 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nominal Code"
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
         Left            =   135
         TabIndex        =   18
         Top             =   1200
         Width           =   990
      End
      Begin VB.Label lblRentCharges 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Budget Amount"
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
         Left            =   4680
         TabIndex        =   17
         Top             =   1200
         Width           =   1095
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   120
      Top             =   7560
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmServiceChargeDetails.frx":08CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmServiceChargeDetails.frx":0E64
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmServiceChargeDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Private flgChange As Integer   '0 -> New, 1 -> Edit
Private bDiscard  As Boolean
'Public bNewBudget As Boolean
Public bNewLine   As Boolean
Public bAmended   As Boolean
Dim strCommandSource As String
Public szPropertySelection1 As String
Public szclientID As String

Private Sub cmdGridUnitLookup_Click()
'     tabMain.Enabled = True
'    picMain.Enabled = True
    Frame1(1).Enabled = True
    Frame5.Visible = False
    FocusControl cmdSCFund
End Sub

Private Sub cmdNew_Click()
'   If flxSCBudgetDetailsAnalysis.TextMatrix(1, 0) = "" Then Exit Sub

   On Error Resume Next
    'Resolved by BOSL
    '0000471:  Service Charge Budget
    'Modified by Anol 15 Sep 2014
   Dim iRow As Integer
   If frmServiceCharge.bEDIT = False Then
            For iRow = 1 To frmServiceCharge.flxSCBudgetDetails.Rows - 1
              If txtSCFund.Tag = frmServiceCharge.flxSCBudgetDetails.TextMatrix(iRow, 2) And frmServiceCharge.flxSCBudgetDetails.RowHeight(iRow) = 240 Then
                 ShowMsgInTaskBar "The property already has a budget for this fund.", "Y", "N"
                 'cboSCFund.SetFocus
                 Exit Sub
              End If
            Next iRow
    End If
   bNewLine = True
   Load frmSCDetailsSplit
   frmSCDetailsSplit.txtNCode.text = ""
   frmSCDetailsSplit.txtBudget.text = ""
   frmSCDetailsSplit.Show
   Me.Enabled = False

'   ControlsModeRentBudgetDetails EditMode
'   cboNCode.SetFocus
'   flgChange = 1
End Sub

Private Sub cmdRunFlag_Click()

End Sub

Private Sub cmdSCDBdCancel_Click()
'   If MsgBox("Do you wish to cancel the amendment of the budget analysis?", vbQuestion + vbYesNo, "Cancel") = vbYes Then
'      txtSCDBudgetTotal.text = ""
      Unload Me
'   End If
End Sub

Private Sub cmdSCDBdDelete_Click()
   If MsgBox("Do you wish to delete?", vbQuestion + vbYesNo, "delete") = vbNo Then Exit Sub

   With flxSCBudgetDetailsAnalysis
      .TextMatrix(.row, 5) = "D"
      .RowHeight(.row) = 0
      txtSCDBudgetTotal.text = Format(Val(txtSCDBudgetTotal.text) - Val(.TextMatrix(.row, 4)), "0.00")
   End With
   
    'Issue 471 note 740
    'Modified by anol 05 Nov 2014
    Dim adoConn1 As New ADODB.Connection
    Dim szSQL As String
    Dim adoRst As New ADODB.Recordset
    
    
    
     
   adoConn1.Open getConnectionString
   adoConn1.Execute "Delete FROM GlobalSCDtls WHERE BudgetID = '" & frmServiceCharge.txtBudgetId.text & "' and NC='" & flxSCBudgetDetailsAnalysis.TextMatrix(flxSCBudgetDetailsAnalysis.row, 2) & "';"
   SCSumTotal
   
   szSQL = "SELECT * FROM GlobalSC WHERE BudgetID = '" & frmServiceCharge.txtBudgetId.text & "';"
   adoRst.Open szSQL, adoConn1, adOpenDynamic, adLockOptimistic
   If adoRst.EOF = False Then
      adoRst.Fields.Item("Fund").Value = txtSCFund.Tag
      adoRst.Fields.Item("TotalBudget").Value = txtSCDBudgetTotal.text
      adoRst.Fields.Item("SCArea").Value = txtTotalArea.text
      adoRst.Fields.Item("PPSF").Value = txtPpsf.text
      adoRst.Update
   End If
   frmServiceCharge.flxSCBudgetDetails.TextMatrix(frmServiceCharge.flxSCBudgetDetails.row, 4) = txtSCDBudgetTotal.text
   frmServiceCharge.flxSCBudgetDetails.TextMatrix(frmServiceCharge.flxSCBudgetDetails.row, 6) = txtPpsf.text
   adoConn1.Close
   
   
End Sub

Private Sub cmdSCDBDEdit_Click()
   If flxSCBudgetDetailsAnalysis.TextMatrix(1, 0) = "" Then Exit Sub

   On Error Resume Next

   Load frmSCDetailsSplit
   frmSCDetailsSplit.txtNCode.text = flxSCBudgetDetailsAnalysis.TextMatrix(flxSCBudgetDetailsAnalysis.row, 2)
   frmSCDetailsSplit.txtBudget.text = flxSCBudgetDetailsAnalysis.TextMatrix(flxSCBudgetDetailsAnalysis.row, 4)
   frmSCDetailsSplit.Show
   Me.Enabled = False

'   ControlsModeRentBudgetDetails EditMode
'   cboNCode.SetFocus
'   flgChange = 1
   
   
'   If cmdSCDBdEdit.Caption <> "&Edit" Then
'
'      updateGrid
'      flgChange = 1
'      SCSumTotal
'      ControlsModeRentBudgetDetails DefaultMode
'   End If
End Sub

Private Sub cmdSCDBdClose_Click()
   bDiscard = False

   If txtSCFund.text = "" Then
      ShowMsgInTaskBar "Please select a fund.", "Y", "N"
      cmdSCFund.SetFocus
      Exit Sub
   End If
   If txtTotalArea.text = "" Then
      ShowMsgInTaskBar "Please enter total area.", "Y", "N"
      txtTotalArea.SetFocus
      Exit Sub
   End If
   'Resolved by BOSL
    '0000471:  Service Charge Budget
    'Modified by Anol 15 Sep 2014
    'The condition should be like that when in edit mode it will exclude the selected fund in the whick was double clicked
   Dim iRow As Integer
    If frmServiceCharge.bEDIT = False Then '(When new fund entry mood)
        For iRow = 1 To frmServiceCharge.flxSCBudgetDetails.Rows - 1
            If txtSCFund.Tag = frmServiceCharge.flxSCBudgetDetails.TextMatrix(iRow, 2) And frmServiceCharge.flxSCBudgetDetails.RowHeight(iRow) = 240 Then
               ShowMsgInTaskBar "The property already has a budget for this fund.", "Y", "N"
               'cboSCFund.SetFocus
               Exit Sub
            End If
        Next iRow
    End If
   'Resolved by BOSL
    '0000450:  Service Charge Budget
    'When adding a new Service Charge Budget it is possible to enter information into the Total Budget
    'text box, it should not be. When you fill out all the information (ie. Fund, Total Area, Total Budget)
    'apart from Price/SqFoot and attempt to save the Budget you receive this message
    'Modified by Anol 10 Aug 2014
   If txtPpsf.text = "" Then
      ShowMsgInTaskBar "Press new button and add new service charge details", "Y", "N"
      cmdNew.SetFocus
      Exit Sub
   End If
   
   If txtSCDBudgetTotal.text = "" Then
      ShowMsgInTaskBar "Please enter some budget details.", "Y", "N"
      cmdSCFund.SetFocus
      Exit Sub
   End If

   Dim adoconn As New ADODB.Connection
   Dim adoRst  As New ADODB.Recordset
   Dim szSQL   As String
   'Dim iRow    As Integer

   If frmServiceCharge.bEDIT = False Then
              adoconn.Open getConnectionString
        
              szSQL = "SELECT * FROM GlobalSC;"
              adoRst.Open szSQL, adoconn, adOpenDynamic, adLockOptimistic
        
              adoRst.AddNew
              adoRst.Fields.Item("BudgetID").Value = frmServiceCharge.txtBudgetId.text
              adoRst.Fields.Item("PropertyID").Value = frmServiceCharge.txtPropertyName.Tag
              adoRst.Fields.Item("FinancialYear").Value = frmServiceCharge.txtBudgetYears.Tag
              adoRst.Fields.Item("Fund").Value = txtSCFund.Tag
              adoRst.Fields.Item("TotalBudget").Value = txtSCDBudgetTotal.text
              adoRst.Fields.Item("SCArea").Value = txtTotalArea.text
              adoRst.Fields.Item("PPSF").Value = txtPpsf.text
              adoRst.Update
        
              adoRst.Close
        
              szSQL = "SELECT * FROM GlobalSCDtls;"
              adoRst.Open szSQL, adoconn, adOpenDynamic, adLockOptimistic
        
              For iRow = 1 To flxSCBudgetDetailsAnalysis.Rows - 1
                 adoRst.AddNew
                 adoRst.Fields.Item("BudgetDtlID").Value = UniqueID
                 adoRst.Fields.Item("BudgetID").Value = frmServiceCharge.txtBudgetId.text
        
                 adoRst.Fields.Item("NC").Value = flxSCBudgetDetailsAnalysis.TextMatrix(iRow, 2)
                 adoRst.Fields.Item("NN").Value = flxSCBudgetDetailsAnalysis.TextMatrix(iRow, 3)
                 adoRst.Fields.Item("BudgetAmt").Value = flxSCBudgetDetailsAnalysis.TextMatrix(iRow, 4)
                 adoRst.Update
              Next iRow
        
              adoRst.Close
        
              Set adoRst = Nothing
              adoconn.Close
        
              With frmServiceCharge.flxSCBudgetDetails
                 .AddItem ""
'                  .TextMatrix(.Rows - 1, 0) = frmServiceCharge.txtBudgetId.text
'                 .TextMatrix(.Rows - 1, 1) = frmServiceCharge.txtPropertyName.Tag
'                 .TextMatrix(.Rows - 1, 2) = txtSCFund.Tag
'                 .TextMatrix(.Rows - 1, 3) = txtSCFund.text
'                 .TextMatrix(.Rows - 1, 4) = txtSCDBudgetTotal.text
'                 .TextMatrix(.Rows - 1, 5) = txtTotalArea.text
'                 .TextMatrix(.Rows - 1, 6) = txtPpsf.text
                 
                 .TextMatrix(.Rows - 1, 0) = frmServiceCharge.txtBudgetId.text
                 .TextMatrix(.Rows - 1, 1) = frmServiceCharge.txtPropertyName.Tag
                 .TextMatrix(.Rows - 1, 2) = txtSCFund.Tag
                 .TextMatrix(.Rows - 1, 3) = txtSCFundCode.text
                 .TextMatrix(.Rows - 1, 4) = txtSCFund.text
                 .TextMatrix(.Rows - 1, 5) = txtSCDBudgetTotal.text
                 .TextMatrix(.Rows - 1, 6) = txtTotalArea.text
                 .TextMatrix(.Rows - 1, 7) = txtPpsf.text
                 'Modifed by anol 17 Sep 2014
'                  .flxSCBudgetDetails.TextMatrix(.flxSCBudgetDetails.row, 2) = txtSCFund.Tag
'                 .flxSCBudgetDetails.TextMatrix(.flxSCBudgetDetails.row, 3) = txtSCFundCode.text
'                 .flxSCBudgetDetails.TextMatrix(.flxSCBudgetDetails.row, 4) = txtSCFund.text
'                 .flxSCBudgetDetails.TextMatrix(.flxSCBudgetDetails.row, 5) = txtSCDBudgetTotal.text
'                 .flxSCBudgetDetails.TextMatrix(.flxSCBudgetDetails.row, 6) = txtTotalArea.text
'                 .flxSCBudgetDetails.TextMatrix(.flxSCBudgetDetails.row, 7) = txtPpsf.text
                 
                 'issue 473
                 .TextMatrix(.Rows - 1, 9) = frmServiceCharge.txtBudgetYears.Tag
                 'End of modification
              End With
              frmServiceCharge.SCSumTotal
     Else
   
              adoconn.Open getConnectionString
        
              szSQL = "SELECT * FROM GlobalSCDtls WHERE BudgetID = '" & frmServiceCharge.txtBudgetId.text & "';"
              adoRst.Open szSQL, adoconn, adOpenDynamic, adLockOptimistic
        
        '  szFlxHeader$ = "BudgetDtlID|BudgetID|<NC|<NN|>TotalBudget|Flag"
              For iRow = 1 To flxSCBudgetDetailsAnalysis.Rows - 1
                 If Left(flxSCBudgetDetailsAnalysis.TextMatrix(iRow, 5), 1) = "N" Then
                    adoRst.AddNew
                    adoRst.Fields.Item("BudgetDtlID").Value = UniqueID
                    adoRst.Fields.Item("BudgetID").Value = frmServiceCharge.txtBudgetId.text
                    adoRst.Fields.Item("NC").Value = flxSCBudgetDetailsAnalysis.TextMatrix(iRow, 2)
                    adoRst.Fields.Item("NN").Value = flxSCBudgetDetailsAnalysis.TextMatrix(iRow, 3)
                    adoRst.Fields.Item("BudgetAmt").Value = flxSCBudgetDetailsAnalysis.TextMatrix(iRow, 4)
                    adoRst.Update
                 End If
                 If flxSCBudgetDetailsAnalysis.TextMatrix(iRow, 5) = "A" Then
                    adoRst.Find ("BudgetDtlID = '" & flxSCBudgetDetailsAnalysis.TextMatrix(iRow, 0) & "'"), , , 1
        
                    adoRst.Fields.Item("NC").Value = flxSCBudgetDetailsAnalysis.TextMatrix(iRow, 2)
                    adoRst.Fields.Item("NN").Value = flxSCBudgetDetailsAnalysis.TextMatrix(iRow, 3)
                    adoRst.Fields.Item("BudgetAmt").Value = flxSCBudgetDetailsAnalysis.TextMatrix(iRow, 4)
                    adoRst.Update
                 End If
                 'issue 471 Note 740
'                 If flxSCBudgetDetailsAnalysis.TextMatrix(iRow, 5) = "D" Then
'                    adoRst.Find ("BudgetDtlID = '" & flxSCBudgetDetailsAnalysis.TextMatrix(iRow, 0) & "'"), , , 1
'                    adoRst.Delete
'                    adoRst.Update
'                 End If
              Next iRow
        
              adoRst.Close
        
              szSQL = "SELECT * FROM GlobalSC WHERE BudgetID = '" & frmServiceCharge.txtBudgetId.text & "';"
              adoRst.Open szSQL, adoconn, adOpenDynamic, adLockOptimistic
              'issue 471 Note 740
              If adoRst.EOF = False Then
                  adoRst.Fields.Item("Fund").Value = txtSCFund.Tag
                  adoRst.Fields.Item("TotalBudget").Value = txtSCDBudgetTotal.text
                  adoRst.Fields.Item("SCArea").Value = txtTotalArea.text
                  adoRst.Fields.Item("PPSF").Value = txtPpsf.text
                  adoRst.Update
              End If
              If adoRst.State = 1 Then
                  adoRst.Close
                  Set adoRst = Nothing
              End If
        
              With frmServiceCharge
                'Modified by anol 2021-06-28
                 .flxSCBudgetDetails.TextMatrix(.flxSCBudgetDetails.row, 2) = txtSCFund.Tag
                 .flxSCBudgetDetails.TextMatrix(.flxSCBudgetDetails.row, 3) = txtSCFundCode.text
                 .flxSCBudgetDetails.TextMatrix(.flxSCBudgetDetails.row, 4) = txtSCFund.text
'                 .flxSCBudgetDetails.TextMatrix(.flxSCBudgetDetails.row, 5) = txtSCDBudgetTotal.text
'                 .flxSCBudgetDetails.TextMatrix(.flxSCBudgetDetails.row, 6) = txtTotalArea.text
                 .flxSCBudgetDetails.TextMatrix(.flxSCBudgetDetails.row, 5) = txtTotalArea.text
                 .flxSCBudgetDetails.TextMatrix(.flxSCBudgetDetails.row, 6) = txtSCDBudgetTotal.text
                 .flxSCBudgetDetails.TextMatrix(.flxSCBudgetDetails.row, 7) = txtPpsf.text
                 .SCSumTotal
                 .bEDIT = False
              End With
   End If
   'Resolved by BOSL
   'Modified by anol 15 Sep 2014
   'issue 471
   If adoconn.State = 0 Then
        adoconn.Open getConnectionString
   End If
   'End of modification
   Update_SC_Lease adoconn
      
   adoconn.Close
   Set adoconn = Nothing

   Unload Me
End Sub

Private Sub flxClientList_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        flxClientList_Click
    End If
End Sub

Private Sub txtSearchClientID_Change()
       
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
End Sub
Private Sub Update_SC_Lease(adoconn As ADODB.Connection)
   Dim Rst     As New ADODB.Recordset
   Dim adoRst  As New ADODB.Recordset
   Dim szSQL   As String
   
   
        On Error GoTo ErrHandler:
        
        'Resolved By BOSL.
        'Modified By Asif. Issue: 0000519. Date: 04-Jan-2015
        'Updating the service charge budgets through SQL rather than iteration which is time consuming.
        
        ' Charging Method: 2
        szSQL = "UPDATE LServiceCharges " & _
        "SET " & _
        "LServiceCharges.SCTotal = 0, " & _
        "LServiceCharges.SCAmount = 0 " & _
        "Where " & _
        "LServiceCharges.ChargingMethod = 2;"
        
        adoconn.Execute szSQL
        
'        szSQL = "UPDATE LServiceCharges, GlobalSC AS GSC, LeaseDetails AS L, " & _
'        "Units AS U, Frequencies AS F,  Property AS P " & _
'        "SET " & _
'        "LServiceCharges.SCTotal = (GSC.TotalBudget * LServiceCharges.CMFigure)/100, " & _
'        "LServiceCharges.SCAmount = (GSC.TotalBudget * LServiceCharges.CMFigure / 100) / F.PartOfYear " & _
'        "Where " & _
'        "cstr(GSC.Fund) = LServiceCharges.ServiceChargeDept AND L.UnitNumber = U.UnitNumber AND " & _
'        "L.LeaseID = LServiceCharges.LeaseID AND " & _
'        "LServiceCharges.SCFrequency = F.ID AND U.PropertyID = GSC.PropertyID AND " & _
'        "P.CBY = GSC.FinancialYear AND U.PropertyID = P.PropertyID AND " & _
'        "LServiceCharges.ChargingMethod = 2 and GSC.SCYearEndGenerated=false;"
'
'        adoconn.Execute szSQL
'
'        ' Charging Method: 4
'
'        szSQL = "UPDATE LServiceCharges " & _
'        "SET " & _
'        "LServiceCharges.SCTotal = 0, " & _
'        "LServiceCharges.SCAmount = 0, " & _
'        "LServiceCharges.CMFigure = 0 " & _
'        "Where " & _
'        "LServiceCharges.ChargingMethod = 4;"
'
'        adoconn.Execute szSQL
'
'        szSQL = "UPDATE LServiceCharges, GlobalSC AS GSC, LeaseDetails AS L, " & _
'        "Units AS U, Frequencies AS F,  Property AS P " & _
'        "SET " & _
'        "LServiceCharges.SCTotal = (GSC.PPSF * U.TotalArea), " & _
'        "LServiceCharges.SCAmount = (GSC.PPSF * U.TotalArea)/F.PartOfYear, " & _
'        "LServiceCharges.CMFigure = (GSC.PPSF * U.TotalArea) " & _
'        "Where " & _
'        "cstr(GSC.Fund) = LServiceCharges.ServiceChargeDept AND L.UnitNumber = U.UnitNumber AND " & _
'        "L.LeaseID = LServiceCharges.LeaseID AND " & _
'        "LServiceCharges.SCFrequency = F.ID AND U.PropertyID = GSC.PropertyID AND " & _
'        "P.CBY = GSC.FinancialYear AND U.PropertyID = P.PropertyID AND " & _
'        "LServiceCharges.ChargingMethod = 4 and GSC.SCYearEndGenerated=false;"
'
'        adoconn.Execute szSQL
'******  fund  category not 5
    szSQL = "UPDATE LServiceCharges, GlobalSC AS GSC, LeaseDetails AS L, " & _
    "Units AS U, Frequencies AS F,  Property AS P, Fund FN " & _
    "SET " & _
    "LServiceCharges.SCTotal = (GSC.TotalBudget * LServiceCharges.CMFigure)/100, " & _
    "LServiceCharges.SCAmount = (GSC.TotalBudget * LServiceCharges.CMFigure / 100) / F.PartOfYear " & _
    "Where FN.FundID=GSC.Fund AND " & _
    "cstr(GSC.Fund) = LServiceCharges.ServiceChargeDept AND L.UnitNumber = U.UnitNumber AND " & _
    "L.LeaseID = LServiceCharges.LeaseID AND " & _
    "LServiceCharges.SCFrequency = F.ID AND U.PropertyID = GSC.PropertyID AND " & _
    "P.CBY = GSC.FinancialYear AND U.PropertyID = P.PropertyID AND " & _
    "LServiceCharges.ChargingMethod = 2 AND FN.CategoryCode<>5;"
    
    adoconn.Execute szSQL
    
    '******  fund  category = 5
     szSQL = "UPDATE LServiceCharges, GlobalSC AS GSC, LeaseDetails AS L, " & _
    "Units AS U, Frequencies AS F,  Property AS P, Fund FN " & _
    "SET " & _
    "LServiceCharges.SCTotal = (GSC.TotalBudget * LServiceCharges.CMFigure)/100, " & _
    "LServiceCharges.SCAmount = (GSC.TotalBudget * LServiceCharges.CMFigure / 100) / F.PartOfYear " & _
    "Where  FN.FundID=GSC.Fund AND " & _
    "cstr(GSC.Fund) = LServiceCharges.ServiceChargeDept AND L.UnitNumber = U.UnitNumber AND " & _
    "L.LeaseID = LServiceCharges.LeaseID AND " & _
    "LServiceCharges.SCFrequency = F.ID AND U.PropertyID = GSC.PropertyID AND " & _
    "P.CBY = GSC.FinancialYear AND U.PropertyID = P.PropertyID AND " & _
    "LServiceCharges.ChargingMethod = 2 and GSC.SCYearEndGenerated=false AND FN.CategoryCode=5;"
    
    adoconn.Execute szSQL
    
    
    
    ' Charging Method: 4
    
    szSQL = "UPDATE LServiceCharges " & _
    "SET " & _
    "LServiceCharges.SCTotal = 0, " & _
    "LServiceCharges.SCAmount = 0, " & _
    "LServiceCharges.CMFigure = 0 " & _
    "Where " & _
    "LServiceCharges.ChargingMethod = 4;"
    
    adoconn.Execute szSQL
    
    '******  fund  category not 5
    szSQL = "UPDATE LServiceCharges, GlobalSC AS GSC, LeaseDetails AS L, " & _
    "Units AS U, Frequencies AS F,  Property AS P, Fund FN  " & _
    "SET " & _
    "LServiceCharges.SCTotal = (GSC.PPSF * U.TotalArea), " & _
    "LServiceCharges.SCAmount = (GSC.PPSF * U.TotalArea)/F.PartOfYear, " & _
    "LServiceCharges.CMFigure = (GSC.PPSF * U.TotalArea) " & _
    "Where  FN.FundID=GSC.Fund AND " & _
    "cstr(GSC.Fund) = LServiceCharges.ServiceChargeDept AND L.UnitNumber = U.UnitNumber AND " & _
    "L.LeaseID = LServiceCharges.LeaseID AND " & _
    "LServiceCharges.SCFrequency = F.ID AND U.PropertyID = GSC.PropertyID AND " & _
    "P.CBY = GSC.FinancialYear AND U.PropertyID = P.PropertyID AND " & _
    "LServiceCharges.ChargingMethod = 4 AND FN.CategoryCode<>5;"
    
    adoconn.Execute szSQL
      '******  fund  category = 5
    szSQL = "UPDATE LServiceCharges, GlobalSC AS GSC, LeaseDetails AS L, " & _
    "Units AS U, Frequencies AS F,  Property AS P , Fund FN " & _
    "SET " & _
    "LServiceCharges.SCTotal = (GSC.PPSF * U.TotalArea), " & _
    "LServiceCharges.SCAmount = (GSC.PPSF * U.TotalArea)/F.PartOfYear, " & _
    "LServiceCharges.CMFigure = (GSC.PPSF * U.TotalArea) " & _
    "Where  FN.FundID=GSC.Fund AND " & _
    "cstr(GSC.Fund) = LServiceCharges.ServiceChargeDept AND L.UnitNumber = U.UnitNumber AND " & _
    "L.LeaseID = LServiceCharges.LeaseID AND " & _
    "LServiceCharges.SCFrequency = F.ID AND U.PropertyID = GSC.PropertyID AND " & _
    "P.CBY = GSC.FinancialYear AND U.PropertyID = P.PropertyID AND " & _
    "LServiceCharges.ChargingMethod = 4 and GSC.SCYearEndGenerated=false AND FN.CategoryCode=5;"
    
    adoconn.Execute szSQL
    
        
        MsgBox "The lease service charge budgets are updated successfully."

Exit Sub
ErrHandler:
   MsgBox ERR.Number & " " & ERR.description, vbExclamation + vbOKOnly, "Could not update Service Charge Budget"
End Sub

'
'Private Sub cmdSCDClose_Click()
'   initialiseGrid
'   Me.Hide
'End Sub
'
'Private Sub LoadFlxRCMain()
'   Dim rowIndex As Integer
'   Dim col As Integer
'
'   If frmMMain.IsRibbonVersion Then
'      rowIndex = frmServiceCharge.txtMatrixRow.text
'   Else
'      rowIndex = frmServiceCharge.txtMatrixRow.text
'   End If
'
'   For col = 0 To 59
'      If frmMMain.IsRibbonVersion Then
'      If frmServiceCharge.getDetailsFromMatrix(rowIndex, col).getBudgetDetailID = "" Then
'         Exit For
'      Else
'         addLine frmServiceCharge.getDetailsFromMatrix(rowIndex, col)
'         frmServiceCharge.fillBufferMatrix frmServiceCharge.getDetailsFromMatrix(rowIndex, col), col + 1
'      End If
'      Else
'      If frmServiceCharge.getDetailsFromMatrix(rowIndex, col).getBudgetDetailID = "" Then
'         Exit For
'      Else
'         addLine frmServiceCharge.getDetailsFromMatrix(rowIndex, col)
'         frmServiceCharge.fillBufferMatrix frmServiceCharge.getDetailsFromMatrix(rowIndex, col), col + 1
'      End If
'      End If
'   Next col
'End Sub

Public Sub SCSumTotal()
   Dim iRow As Integer

   txtSCDBudgetTotal.text = "0.00"

   For iRow = 1 To flxSCBudgetDetailsAnalysis.Rows - 1
      If flxSCBudgetDetailsAnalysis.TextMatrix(iRow, 5) <> "D" Then _
         txtSCDBudgetTotal.text = Format(Val(txtSCDBudgetTotal.text) + _
                                         Val(flxSCBudgetDetailsAnalysis.TextMatrix(iRow, 4)), "0.00")
   Next iRow

   If Val(txtTotalArea.text) > 0 And Val(txtSCDBudgetTotal.text) > 0 Then
      txtPpsf.text = Format(Val(txtSCDBudgetTotal.text) / Val(txtTotalArea.text), "0.00")
   End If
End Sub

Private Sub cmdSCDBdNew_Click()

'   updateGrid
   SCSumTotal

'   flgChange = 0
   ControlsModeRentBudgetDetails DefaultMode
'   cmdSCDBdNew.Picture = ImageList1.ListImages.Item(1).Picture
End Sub

Private Sub cmdSCFund_Click()
    Frame1(1).Enabled = False
    strCommandSource = "Fund"
    Call LoadFunds
    Frame5.Top = 30
    Frame5.Left = txtSCFundCode.Left - 1100
    Frame5.Visible = True
    'focuscontrol cmdNCode
End Sub

Private Sub flxClientList_Click()
        If strCommandSource = "Fund" Then
                txtSCFund.Tag = flxClientList.TextMatrix(flxClientList.row, 3) '1 ID,2 code,3 fund Name
                txtSCFund.text = flxClientList.TextMatrix(flxClientList.row, 2)
                txtSCFundCode.text = flxClientList.TextMatrix(flxClientList.row, 1)
                txtTotalArea.text = "1"
                FocusControl txtTotalArea
        End If
        Frame1(1).Enabled = True
        Frame5.Visible = False
        FocusControl flxSCBudgetDetailsAnalysis
End Sub

Private Sub flxSCBudgetDetailsAnalysis_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Me.MousePointer = vbArrow
End Sub

Private Sub flxSCBudgetDetailsAnalysis_DblClick()
    'Resolved by BOSL
    '0000450:  Service Charge Budget
'   If flxSCBudgetDetailsAnalysis.TextMatrix(0, 0) = "" Then Exit Sub

'Resolved by BOSL
'Modified by anol 17 Sep 2014
'issue 471
   If flxSCBudgetDetailsAnalysis.TextMatrix(flxSCBudgetDetailsAnalysis.row, 0) = "" Then
        ShowMsgInTaskBar "There is no budget entered for this budget Year.", , "N"
        Exit Sub
   End If
   bNewLine = False
   On Error Resume Next
   LoadForm frmSCDetailsSplit
   
   frmSCDetailsSplit.txtNCode.text = flxSCBudgetDetailsAnalysis.TextMatrix(flxSCBudgetDetailsAnalysis.row, 2)
   frmSCDetailsSplit.txtNName.text = flxSCBudgetDetailsAnalysis.TextMatrix(flxSCBudgetDetailsAnalysis.row, 3)
   frmSCDetailsSplit.txtBudget.text = flxSCBudgetDetailsAnalysis.TextMatrix(flxSCBudgetDetailsAnalysis.row, 4)

'   frmSCDetailsSplit.Show
   Me.Enabled = False
End Sub

Private Sub ControlsModeRentBudgetDetails(ByVal mode As ComponentMode)
   Select Case mode
      Case ComponentMode.DefaultMode
'         cboNCode.text = ""
'         txtNName.text = ""
'         txtBudget.text = ""

         flxSCBudgetDetailsAnalysis.Enabled = True
         flxSCBudgetDetailsAnalysis.row = 0
         flxSCBudgetDetailsAnalysis.col = 0
'         frmServiceChargeDetails.Show
'         cboNCode.SetFocus
         cmdSCDBdDelete.Enabled = True

      Case ComponentMode.EditMode
'         cboNCode.Locked = False
'         txtBudget.Locked = False
'         cmdSCDBdCancel.Enabled = True
         cmdSCDBdDelete.Enabled = False
         flxSCBudgetDetailsAnalysis.Enabled = False
   End Select
End Sub

Private Sub Form_Load()
'   frmMMain.Arrange vbCascade
'   Me.ZOrder 0
   Me.Width = 7080
   Me.Height = 8475
   Me.BackColor = MODULEBACKCOLOR
   Me.Refresh
   Frame1(1).BackColor = MODULEBACKCOLOR
   bDiscard = True
   bNewLine = False
   bAmended = False

'   Dim adoConn As New ADODB.Connection
'
'   adoConn.Open getConnectionString

'   LoadFund adoConn

   ConfigureFlxBRMain

'   adoConn.Close
'   Set adoConn = Nothing
'   If Not bNewBudget Then LoadFlxSCBudgetDetailsAnalysis
'
'   SCSumTotal
'
'   cmdSCDBdNew.Picture = ImageList1.ListImages.Item(1).Picture
   Call WheelHook(Me.hWnd)
End Sub
Private Sub LoadFunds()
  'My Ideal loading flexgrid component by anol 2020-12-17
  'Learning: inside a picturebox you cannot resize a Textbox, I am I am adding frame and shape to replace this picturebox
   Dim rRow As Integer
   Dim szSQL As String
   Dim iSel As Integer
   Dim adoconn As New ADODB.Connection
   Dim rstRec As New ADODB.Recordset
   Dim rsFundMatrix As New ADODB.Recordset
   'you just change label position then searchbox and grid coulumn will try to fit accordingly
   lblClientID(0).Left = 250
   lblClientID(1).Left = 1365
   lblClientID(2).Left = 3510

   flxClientList.RowHeight(0) = 0
   flxClientList.Cols = 3
   flxClientList.ColWidth(0) = 200
   flxClientList.ColWidth(1) = lblClientID(1).Left - lblClientID(0).Left
   
   
   txtSearchClientID.Width = lblClientID(1).Left - lblClientID(0).Left - 20
   txtSearchClientName.Width = lblClientID(2).Left - lblClientID(1).Left - 20
   TextBox1.Width = cmdGridUnitLookup.Left + cmdGridUnitLookup.Width - lblClientID(2).Left - 20
   
     
   If flxClientList.Cols > 3 Then
        flxClientList.ColWidth(2) = lblClientID(2).Left - lblClientID(1).Left
        txtSearchClientName.Width = lblClientID(2).Left - lblClientID(1).Left - 20
   ElseIf flxClientList.Cols = 3 Then
        flxClientList.ColWidth(2) = cmdGridUnitLookup.Left - lblClientID(1).Left - 200
        txtSearchClientName.Width = cmdGridUnitLookup.Left - lblClientID(1).Left - 200
   End If
   If flxClientList.Cols = 4 Then
        flxClientList.ColWidth(3) = cmdGridUnitLookup.Left + cmdGridUnitLookup.Width - lblClientID(2).Left
        TextBox1.Visible = True
   ElseIf flxClientList.Cols = 3 Then
        flxClientList.ColWidth(3) = 0
        TextBox1.Visible = False
   End If
   
   
   txtSearchClientName.Visible = True

   
   flxClientList.Clear
   flxClientList.Rows = 2
   flxClientList.ColAlignment(0) = vbLeftJustify
   flxClientList.ColAlignment(1) = vbLeftJustify
   flxClientList.ColAlignment(2) = vbLeftJustify
   flxClientList.ColAlignment(3) = vbLeftJustify
   
   lblClientID(0).Caption = "Fund Code"
   lblClientID(1).Caption = "Fund Name"
   lblClientID(2).Caption = ""
   
   txtSearchClientID.Left = lblClientID(0).Left
   txtSearchClientName.Left = lblClientID(1).Left
   
   
   TextBox1.Left = lblClientID(2).Left
   TextBox1.Width = cmdGridUnitLookup.Left - lblClientID(2).Left + 40
   
   txtSearchClientName.text = ""
   txtSearchClientID.text = ""
   TextBox1.text = ""
    
   adoconn.Open getConnectionString
   'szSQL = "SELECT ID, TYPE FROM DemandTypes where PropertyID='" & szPropertySelection1 & "';"
   
   rsFundMatrix.Open "Select isfundAssign from shoppingcentre", adoconn, adOpenStatic, adLockReadOnly
   If rsFundMatrix("isfundAssign").Value = False Then
        iSel = 0
        szSQL = "SELECT FundID, FundName, FundCode,CategoryCode FROM Fund;"
   Else
        iSel = 1
        szSQL = "Select F.* from Fund F,fundMatrix M where F.FundID=M.FundID AND PropertyID='" & _
                szPropertySelection1 & "' and ClientID='" & szclientID & "' and isDeleted=false"
   End If
   rsFundMatrix.Close
   rstRec.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
   If rstRec.EOF Then
        If iSel = 0 Then
            ShowMsgInTaskBar "Fund has not been setup for this company.", , "N"
         Else
            ShowMsgInTaskBar "There are no funds assigned for this property. Please assign a fund.", , "N"
         End If
      flxClientList.Clear
      flxClientList.Rows = 2
   Else
                rRow = 1
                While Not rstRec.EOF
                    flxClientList.row = 1
                    flxClientList.RowSel = 1
                    flxClientList.ColSel = 1
                    flxClientList.TextMatrix(rRow, 0) = ""
                    flxClientList.TextMatrix(rRow, 1) = rstRec.Fields.Item("FundCode").Value
                    flxClientList.TextMatrix(rRow, 2) = rstRec.Fields.Item("FundName").Value
                    flxClientList.TextMatrix(rRow, 3) = rstRec.Fields.Item("FundID").Value
                    flxClientList.RowHeight(rRow) = 280
                    rstRec.MoveNext
                    If Not rstRec.EOF Then flxClientList.AddItem ""
                    rRow = rRow + 1
                 Wend
         
   End If
   rstRec.Close
   adoconn.Close
   Set rstRec = Nothing
   Set adoconn = Nothing
End Sub
Public Sub LoadFlxSCBudgetDetailsAnalysis()
   Dim adoconn As New ADODB.Connection
   Dim adoRst  As New ADODB.Recordset
   Dim szSQL   As String
   Dim iRow    As Integer
''Order by NC' was added by anol on 22 Apr 2015
   szSQL = "SELECT * FROM GlobalSCDtls " & _
           "WHERE  BudgetID = '" & frmServiceCharge.txtBudgetId.text & "' Order by NC;"

   adoconn.Open getConnectionString

   adoRst.Open szSQL, adoconn, adOpenStatic, adLockReadOnly

'   szFlxHeader$ = "BudgetDtlID|BudgetID|<Fund|>FundName|>TotalBudget|>SCArea|>PPSF"
   iRow = 1
   While Not adoRst.EOF
      If Not adoRst.EOF Then flxSCBudgetDetailsAnalysis.AddItem ""
      flxSCBudgetDetailsAnalysis.TextMatrix(iRow, 0) = adoRst.Fields.Item("BudgetDtlID").Value
      flxSCBudgetDetailsAnalysis.TextMatrix(iRow, 1) = adoRst.Fields.Item("BudgetID").Value
      flxSCBudgetDetailsAnalysis.TextMatrix(iRow, 2) = adoRst.Fields.Item("NC").Value
      flxSCBudgetDetailsAnalysis.TextMatrix(iRow, 3) = adoRst.Fields.Item("NN").Value
      flxSCBudgetDetailsAnalysis.TextMatrix(iRow, 4) = Format(adoRst.Fields.Item("BudgetAmt").Value, "0.00")
      iRow = iRow + 1
      adoRst.MoveNext
   Wend

   adoRst.Close
   Set adoRst = Nothing
   adoconn.Close
   Set adoconn = Nothing
End Sub

Private Sub ConfigureFlxBRMain()
   Dim szFlxHeader As String

   flxSCBudgetDetailsAnalysis.Rows = 1
   flxSCBudgetDetailsAnalysis.RowHeight(0) = 0
   flxSCBudgetDetailsAnalysis.Clear
   flxSCBudgetDetailsAnalysis.Cols = 6
   szFlxHeader$ = "BudgetDtlID|BudgetID|<NC|<NN|>TotalBudget|Flag"
   flxSCBudgetDetailsAnalysis.FormatString = szFlxHeader$

   flxSCBudgetDetailsAnalysis.ColWidth(0) = 0
   flxSCBudgetDetailsAnalysis.ColWidth(1) = 0
   flxSCBudgetDetailsAnalysis.ColWidth(2) = lblRentCharges(1).Left - lblRentCharges(0).Left
   flxSCBudgetDetailsAnalysis.ColWidth(3) = lblRentCharges(2).Left - lblRentCharges(1).Left
   flxSCBudgetDetailsAnalysis.ColWidth(4) = flxSCBudgetDetailsAnalysis.Width - lblRentCharges(2).Left - 300
   flxSCBudgetDetailsAnalysis.ColWidth(5) = 0
'   txtSCDBudgetTotal.Width = flxSCBudgetDetailsAnalysis.ColWidth(4)
End Sub

''Private Sub LoadFund(adoConn As ADODB.Connection)
''   ' Error Handler
''   On Error GoTo Error_Handler
''
''   Dim rRow As Integer, iRec As Integer, Data() As String
''   Dim adoRst As New ADODB.Recordset
''   Dim szSQL As String
''
''   szSQL = "SELECT FundID, FundName " & _
''           "FROM Fund " & _
''           "WHERE CategoryCode = 2;"
''
''   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
''
''   If adoRst.EOF Then
''      ShowMsgInTaskBar "Fund has not been setup for this company.", , "N"
''   Else
''      ReDim Data(2, adoRst.RecordCount) As String
''
''      rRow = 0
''      While Not adoRst.EOF
''         Data(0, rRow) = Trim(adoRst.Fields.Item("FundID").Value)
''         Data(1, rRow) = Trim(adoRst.Fields.Item("FundName").Value)
''         rRow = rRow + 1
''         adoRst.MoveNext
''      Wend
''
''      cboSCFund.Clear
''      cboSCFund.Column() = Data()
''   End If
''
''   ' Destroy Objects
''   Set adoRst = Nothing
''
''   Exit Sub
''
''   ' Error Handling Code
''Error_Handler:
''
''   ShowMsgInTaskBar "Error in Loading fund.", , "N"
''   ' Destroy Objects
''   Set adoRst = Nothing
''End Sub
'
'Private Sub initialiseGrid()
'   Call ConfigureFlxBRMain
''   flgChange = 0
'
'   LoadFlxRCMain
'
''   Call LoadNC
'   ControlsModeRentBudgetDetails DefaultMode
'End Sub
'
'Private Sub updateGrid()
'   With flxSCBudgetDetailsAnalysis
'      If flgChange = 0 Then
'         .AddItem ""
'         If .TextMatrix(.Rows - 1, 0) <> "" Then .AddItem ""
'         .TextMatrix(.Rows - 1, 0) = UniqueID()
'         If frmMMain.IsRibbonVersion Then
'         .TextMatrix(.Rows - 1, 1) = frmServiceCharge.txtBudgetId.text
'         Else
'         .TextMatrix(.Rows - 1, 1) = frmServiceCharge.txtBudgetId.text
'         End If
'         .TextMatrix(.Rows - 1, 2) = cboNCode.text
'         .TextMatrix(.Rows - 1, 3) = txtNName.text
'         .TextMatrix(.Rows - 1, 4) = FormatNumber(CDbl(Trim(txtBudget.text)), 2, , , vbDefault)
'         .TextMatrix(.Rows - 1, 5) = "N"        'New
'      End If
'      If flgChange = 1 Then
'         .TextMatrix(.row, 2) = cboNCode.text
'         .TextMatrix(.row, 3) = txtNName.text
'         .TextMatrix(.row, 4) = FormatNumber(CDbl(Trim(txtBudget.text)), 2, , , vbDefault)
'         .TextMatrix(.row, 5) = "M"             'Modified
'      End If
'   End With
'End Sub

Private Sub addLine(ByVal bDetail As clsSCDtl)
   flxSCBudgetDetailsAnalysis.AddItem ""
   flxSCBudgetDetailsAnalysis.TextMatrix(flxSCBudgetDetailsAnalysis.Rows - 1, 0) = bDetail.getBudgetDetailID
   flxSCBudgetDetailsAnalysis.TextMatrix(flxSCBudgetDetailsAnalysis.Rows - 1, 1) = bDetail.getBudgetId
   flxSCBudgetDetailsAnalysis.TextMatrix(flxSCBudgetDetailsAnalysis.Rows - 1, 2) = bDetail.getNCode
   flxSCBudgetDetailsAnalysis.TextMatrix(flxSCBudgetDetailsAnalysis.Rows - 1, 3) = bDetail.getNName
   flxSCBudgetDetailsAnalysis.TextMatrix(flxSCBudgetDetailsAnalysis.Rows - 1, 4) = FormatNumber(bDetail.getBudgetAmount, 2, , , vbDefault)
   If (bDetail.getFlgDel = "D") Then
      flxSCBudgetDetailsAnalysis.TextMatrix(flxSCBudgetDetailsAnalysis.Rows - 1, 5) = bDetail.getFlgDel
      flxSCBudgetDetailsAnalysis.RowHeight(flxSCBudgetDetailsAnalysis.Rows - 1) = 0
   End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Me.MousePointer = vbArrow
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   If frmMMain.IsRibbonVersion Then
   frmServiceCharge.Enabled = True
   Else
   frmServiceCharge.Enabled = True
   End If
'
'   If Not bDiscard Then
'      SaveBudgetAnalysis
'      If frmMMain.IsRibbonVersion Then
''      frmServiceCharge.cmdDetails.Enabled = True
''      frmServiceCharge.txtBudget.Locked = True
'      Else
''      frmServiceCharge.cmdDetails.Enabled = True
''      frmServiceCharge.txtBudget.Locked = True
'      End If
'   Else
'      If frmMMain.IsRibbonVersion Then
''      frmServiceCharge.cmdDetails.Enabled = False
''      frmServiceCharge.txtBudget.Locked = False
'      Else
''      frmServiceCharge.cmdDetails.Enabled = False
''      frmServiceCharge.txtBudget.Locked = False
'      End If
'   End If
End Sub

Public Function initialiseNew()
   cmdSCDBdNew_Click
End Function

Private Sub SaveBudgetAnalysis()
   Dim iRow As Integer
  
   For iRow = 1 To flxSCBudgetDetailsAnalysis.Rows - 1
      If frmMMain.IsRibbonVersion Then
      arraySave iRow
      Else
      arraySave1 iRow
      End If
   Next iRow

   If frmMMain.IsRibbonVersion Then
'      frmServiceCharge.txtBudget.text = txtSCDBudgetTotal.text
   Else
'      frmServiceCharge.txtBudget.text = txtSCDBudgetTotal.text
   End If

'   flgChange = 0
End Sub

Private Sub arraySave(ByVal Index As Integer)
   Dim rowIndex As Integer
   Dim col As Integer
   Dim found As Integer

   With flxSCBudgetDetailsAnalysis
      found = 0
      rowIndex = frmServiceCharge.txtMatrixRow.text
      For col = 0 To 59
         If frmServiceCharge.getDetailsFromMatrix(rowIndex, col).getBudgetDetailID = .TextMatrix(Index, 0) Then
            found = 1
            Exit For
         Else
            If frmServiceCharge.getDetailsFromMatrix(rowIndex, col).getBudgetDetailID = "" Then
               Exit For
            End If
         End If
      Next col

      If (found = 1) Then
         frmServiceCharge.getDetailsFromMatrix(rowIndex, col).setNCode .TextMatrix(Index, 2)
         frmServiceCharge.getDetailsFromMatrix(rowIndex, col).setNName .TextMatrix(Index, 3)
         frmServiceCharge.getDetailsFromMatrix(rowIndex, col).setBudgetAmount .TextMatrix(Index, 4)
         frmServiceCharge.getDetailsFromMatrix(rowIndex, col).setFlgDel .TextMatrix(Index, 5)
      Else
         If col <> 59 And .RowHeight(Index) > 0 Then
            frmServiceCharge.getDetailsFromMatrix(rowIndex, col).setBudgetId frmServiceCharge.txtBudgetId.text
            frmServiceCharge.getDetailsFromMatrix(rowIndex, col).setBudgetDetailId .TextMatrix(Index, 0)
            frmServiceCharge.getDetailsFromMatrix(rowIndex, col).setNCode .TextMatrix(Index, 2)
            frmServiceCharge.getDetailsFromMatrix(rowIndex, col).setNName .TextMatrix(Index, 3)
            frmServiceCharge.getDetailsFromMatrix(rowIndex, col).setBudgetAmount .TextMatrix(Index, 4)
         Else
             ShowMsgInTaskBar "Please save service charge budget before proceeding", , "N"
         End If
      End If
   End With
End Sub

Private Sub arraySave1(ByVal Index As Integer)
   Dim rowIndex As Integer
   Dim col As Integer
   Dim found As Integer

   With flxSCBudgetDetailsAnalysis
      found = 0
      rowIndex = frmServiceCharge.txtMatrixRow.text
      For col = 0 To 59
         If frmServiceCharge.getDetailsFromMatrix(rowIndex, col).getBudgetDetailID = .TextMatrix(Index, 0) Then
            found = 1
            Exit For
         Else
            If frmServiceCharge.getDetailsFromMatrix(rowIndex, col).getBudgetDetailID = "" Then
               Exit For
            End If
         End If
      Next col

      If (found = 1) Then
         frmServiceCharge.getDetailsFromMatrix(rowIndex, col).setNCode .TextMatrix(Index, 2)
         frmServiceCharge.getDetailsFromMatrix(rowIndex, col).setNName .TextMatrix(Index, 3)
         frmServiceCharge.getDetailsFromMatrix(rowIndex, col).setBudgetAmount .TextMatrix(Index, 4)
         frmServiceCharge.getDetailsFromMatrix(rowIndex, col).setFlgDel .TextMatrix(Index, 5)
      Else
         If col <> 59 And .RowHeight(Index) > 0 Then
            frmServiceCharge.getDetailsFromMatrix(rowIndex, col).setBudgetId frmServiceCharge.txtBudgetId.text
            frmServiceCharge.getDetailsFromMatrix(rowIndex, col).setBudgetDetailId .TextMatrix(Index, 0)
            frmServiceCharge.getDetailsFromMatrix(rowIndex, col).setNCode .TextMatrix(Index, 2)
            frmServiceCharge.getDetailsFromMatrix(rowIndex, col).setNName .TextMatrix(Index, 3)
            frmServiceCharge.getDetailsFromMatrix(rowIndex, col).setBudgetAmount .TextMatrix(Index, 4)
         Else
             ShowMsgInTaskBar "Please save service charge budget before proceeding", , "N"
         End If
      End If
   End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
'   Call WheelUnHook(Me.hwnd)
    UnLoadForm Me
End Sub

Private Sub Frame1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Me.MousePointer = vbArrow
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

Private Sub txtSearchClientName_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
        If KeyCode = 13 Then
            FocusControl flxClientList
        End If
End Sub

Private Sub txtSearchClientName_KeyPress(KeyAscii As MSForms.ReturnInteger)
        If KeyAscii = 13 Then
           ' FocusControl flxClientList
        End If
End Sub

Private Sub txtTotalArea_LostFocus()
   computePpsf
End Sub

Private Sub computePpsf()
   On Error GoTo ErrHandler
   
   If Trim(txtSCDBudgetTotal.text) <> "" And Trim(txtTotalArea.text) <> "" Then
      Dim ppsf As Double

      ppsf = CDbl(txtSCDBudgetTotal.text) / CDbl(Trim(txtTotalArea.text))
      txtPpsf.text = FormatNumber(CDbl(ppsf), 2, , , vbDefault)
   End If
   Exit Sub

ErrHandler:
   ShowMsgInTaskBar "Please ensure that the values for Total Budget and Total Area are valid before continuing", , "N"
End Sub




